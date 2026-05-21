"""
core/invoice_parser.py
──────────────────────
Parses supplier invoices (PDF or raster images) and returns a structured list
of line items plus the detected invoice date.

Extraction pipeline
───────────────────
1. PDF  → pdfplumber table extraction (works on most digital/typed invoices)
2. PDF  → PyMuPDF (fitz) text extraction  (fallback for non-table PDFs)
3. Image → pytesseract OCR → text extraction
4. Column-header detection via fuzzy matching → maps raw columns to our schema
5. Post-processing & cleaning of individual cells
"""

from __future__ import annotations

import re
import io
import logging
from pathlib import Path
from typing import Any, Optional

import pdfplumber
try:
    import fitz                      # PyMuPDF (optional fallback)
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False
from PIL import Image
import pytesseract

from core.fuzzy_matcher import FuzzyMatcher

log = logging.getLogger(__name__)

# ─── Column-header synonyms ───────────────────────────────────────────────────
# Maps our canonical field names to the many ways suppliers label them.
HEADER_SYNONYMS: dict[str, list[str]] = {
    "item_name":   ["description", "item", "product", "particulars",
                    "goods", "name", "item name", "product name", "details"],
    "sku_code":    ["sku", "code", "item code", "product code", "article",
                    "part no", "part number", "material code", "model"],
    "hsn_sac":     ["hsn", "sac", "hsn/sac", "hsn code", "sac code",
                    "tariff", "harmonised"],
    "gst_rate":    ["gst", "tax rate", "igst", "cgst", "sgst", "gst%",
                    "gst rate", "tax %", "vat"],
    "rate":        ["rate", "price", "unit price", "mrp", "cost", "unit cost",
                    "basic price", "unit rate"],
    "measurement": ["unit", "uom", "measure", "uom", "pack", "packing",
                    "unit of measure", "nos", "qty unit"],
    "quantity":    ["qty", "quantity", "nos", "no.", "pcs", "pieces",
                    "units", "count", "amount ordered", "ordered qty"],
}

# Regex patterns for dates embedded in invoice text
DATE_PATTERNS = [
    r"\b(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{4})\b",   # DD/MM/YYYY or DD-MM-YYYY
    r"\b(\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})\b",    # YYYY-MM-DD
    r"\b(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*"
    r"\s+\d{4})\b",                                 # 12 January 2024
]

DATE_LABEL_PATTERN = re.compile(
    r"(?:invoice\s*date|date|dated|dt\.?)\s*[:\-]?\s*"
    r"(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{4}|\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})",
    re.IGNORECASE,
)

# Rows that should be skipped (sub-totals, page breaks, etc.)
SKIP_ROW_PATTERNS = [
    re.compile(p, re.IGNORECASE)
    for p in [
        r"^\s*(sub[\s-]?total|total|grand\s*total|amount|balance)\s*$",
        r"^\s*(page|continued|carry\s*forward)\s*",
        r"^\s*\d+\s*$",     # pure-number rows (page numbers, serial only)
    ]
]

# ─── Helpers ──────────────────────────────────────────────────────────────────

def _clean_cell(value: Any) -> str:
    """Convert a cell value to a clean stripped string."""
    if value is None:
        return ""
    return str(value).strip().replace("\n", " ")


def _is_numeric(text: str) -> bool:
    """True if the string represents a number (possibly with commas)."""
    try:
        float(text.replace(",", "").replace(" ", ""))
        return True
    except ValueError:
        return False


def _parse_number(text: str) -> float:
    """
    Parse a number from a messy cell string.

    Handles: commas, currency symbols (₹ $ € £), percentage signs,
    and cells where OCR/column-bleed merges two numbers ('9 18', '5.00 18%').
    Strategy: strip known non-numeric chars, try direct conversion;
    if that fails, extract all numeric tokens and return the first one.
    """
    if text is None:
        return 0.0
    text = str(text).strip()
    if not text:
        return 0.0

    # Remove currency symbols, commas, percent signs
    cleaned = re.sub(r"[₹$€£,%]", "", text).replace(",", "").strip()

    # Direct conversion (the happy path)
    try:
        return float(cleaned)
    except ValueError:
        pass

    # Multiple tokens separated by spaces — pick the FIRST valid number.
    # For '9 18': returns 9.0  (quantity before GST bleed)
    tokens = cleaned.split()
    for token in tokens:
        try:
            return float(token)
        except ValueError:
            continue

    # Last resort: pull out any digit sequence with regex
    nums = re.findall(r"\d+\.?\d*", cleaned)
    if nums:
        try:
            return float(nums[0])
        except ValueError:
            pass

    return 0.0


def _normalise_date(raw: str) -> str:
    """
    Convert a raw date string to DD/MM/YYYY format.
    Returns the original string if conversion fails.
    """
    import dateutil.parser as dup
    try:
        # Replace dots with slashes for consistency
        raw_clean = raw.replace(".", "/").replace("-", "/")
        # Detect YYYY/MM/DD vs DD/MM/YYYY
        parts = raw_clean.split("/")
        if len(parts) == 3 and len(parts[0]) == 4:
            # YYYY/MM/DD → swap
            raw_clean = f"{parts[2]}/{parts[1]}/{parts[0]}"
        dt = dup.parse(raw_clean, dayfirst=True)
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return raw


def _should_skip_row(cells: list[str]) -> bool:
    """Return True for rows that are clearly not line items."""
    combined = " ".join(cells).strip()
    if not combined:
        return True
    for pat in SKIP_ROW_PATTERNS:
        if pat.match(combined):
            return True
    return False


def _extract_date_from_text(text: str) -> Optional[str]:
    """Search raw text for an invoice date, return DD/MM/YYYY or None."""
    # Prioritise labelled dates ("Invoice Date: 12/01/2024")
    m = DATE_LABEL_PATTERN.search(text)
    if m:
        return _normalise_date(m.group(1))

    # Fall back to any date pattern in the first 600 characters
    header_text = text[:600]
    for pat in DATE_PATTERNS:
        m = re.search(pat, header_text, re.IGNORECASE)
        if m:
            return _normalise_date(m.group(1))
    return None


# ─── Column Mapper ────────────────────────────────────────────────────────────

class ColumnMapper:
    """
    Maps raw table headers (from the invoice) to canonical field names
    using fuzzy matching.
    """

    def __init__(self):
        self._matcher = FuzzyMatcher(threshold=72)
        # Flat lookup: synonym → canonical_field
        self._lookup: dict[str, str] = {}
        for field, synonyms in HEADER_SYNONYMS.items():
            for s in synonyms:
                self._lookup[s.lower()] = field

    def map_headers(self, raw_headers: list[str]) -> dict[int, str]:
        """
        Given raw column headers, return {column_index: canonical_field}.
        Columns that don't map to a known field are excluded.
        """
        mapping: dict[int, str] = {}
        all_synonyms = list(self._lookup.keys())

        for idx, h in enumerate(raw_headers):
            h_clean = h.strip().lower() if h else ""
            if not h_clean:
                continue

            # Direct lookup first
            if h_clean in self._lookup:
                mapping[idx] = self._lookup[h_clean]
                continue

            # Fuzzy lookup against all synonyms
            best_syn, score = self._matcher.best_match(h_clean, all_synonyms)
            if best_syn:
                mapping[idx] = self._lookup[best_syn]

        return mapping


# ─── Core Parser Class ────────────────────────────────────────────────────────

class InvoiceParser:
    """
    Accepts a file path (PDF or image) and returns:
      - items: list[dict]  – each dict has canonical field keys
      - date:  str         – invoice date as DD/MM/YYYY (or empty string)
    """

    def __init__(self):
        self._col_mapper = ColumnMapper()

    # ── Public entry point ─────────────────────────────────────────────────

    def parse(self, file_path: str) -> tuple[list[dict], str]:
        """
        Parse an invoice file.

        Returns
        -------
        items : list[dict]
            Each dict: {item_name, sku_code, hsn_sac, gst_rate, rate,
                        measurement, quantity}
        date  : str
            Invoice date in DD/MM/YYYY format, or "" if undetectable.
        """
        path = Path(file_path)
        suffix = path.suffix.lower()

        if suffix == ".pdf":
            return self._parse_pdf(str(path))
        elif suffix in {".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".webp"}:
            return self._parse_image(str(path))
        else:
            raise ValueError(f"Unsupported file type: {suffix}")

    # ── PDF Parsing ────────────────────────────────────────────────────────

    def _parse_pdf(self, path: str) -> tuple[list[dict], str]:
        items: list[dict] = []
        full_text = ""

        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                full_text += page_text + "\n"

                tables = page.extract_tables()
                for table in tables:
                    extracted = self._process_table(table)
                    items.extend(extracted)

        # If no items found via tables, attempt text-based extraction
        if not items:
            log.info("No table found via pdfplumber; falling back to text extraction.")
            items = self._extract_from_text(full_text)

        # If still empty, try PyMuPDF (handles more complex PDFs)
        if not items and FITZ_AVAILABLE:
            log.info("Trying PyMuPDF extraction.")
            items, fitz_text = self._parse_with_fitz(path)
            full_text = fitz_text + "\n" + full_text

        date = _extract_date_from_text(full_text)
        return self._clean_items(items), date or ""

    def _parse_with_fitz(self, path: str) -> tuple[list[dict], str]:
        doc = fitz.open(path)
        full_text = ""
        items: list[dict] = []

        for page in doc:
            full_text += page.get_text("text") + "\n"

        if full_text.strip():
            items = self._extract_from_text(full_text)

        doc.close()
        return items, full_text

    # ── Image Parsing (OCR) ────────────────────────────────────────────────

    def _parse_image(self, path: str) -> tuple[list[dict], str]:
        img = Image.open(path)
        # Upscale small images for better OCR accuracy
        if min(img.width, img.height) < 1200:
            scale = 2400 / min(img.width, img.height)
            img = img.resize(
                (int(img.width * scale), int(img.height * scale)),
                Image.LANCZOS)

        # Convert to grayscale + slight threshold for clarity
        img = img.convert("L")

        ocr_config = r"--oem 3 --psm 6"
        raw_text = pytesseract.image_to_string(img, config=ocr_config)

        # Also try structured data extraction via pytesseract's TSV output
        tsv_df = pytesseract.image_to_data(img, config=ocr_config,
                                           output_type=pytesseract.Output.DICT)

        items = self._extract_from_text(raw_text)
        date  = _extract_date_from_text(raw_text)
        return self._clean_items(items), date or ""

    # ── Table Processing ───────────────────────────────────────────────────

    def _process_table(self, table: list[list]) -> list[dict]:
        """Convert a pdfplumber table (list-of-lists) into line-item dicts."""
        if not table or len(table) < 2:
            return []

        # Detect header row (first non-empty row)
        header_row_idx = 0
        for i, row in enumerate(table):
            cleaned = [_clean_cell(c) for c in row]
            # A header row has mostly non-numeric, non-empty cells
            non_empty = [c for c in cleaned if c]
            numeric   = [c for c in non_empty if _is_numeric(c)]
            if non_empty and len(numeric) / max(len(non_empty), 1) < 0.5:
                header_row_idx = i
                break

        raw_headers = [_clean_cell(c) for c in table[header_row_idx]]
        col_map     = self._col_mapper.map_headers(raw_headers)

        if not col_map:
            return []   # Table has no recognisable headers

        items: list[dict] = []
        for row in table[header_row_idx + 1:]:
            cells = [_clean_cell(c) for c in row]
            if _should_skip_row(cells):
                continue

            item = self._row_to_item(cells, col_map)
            if item and item.get("item_name"):
                items.append(item)

        return items

    def _row_to_item(self, cells: list[str], col_map: dict[int, str]) -> dict:
        """Map a single table row's cells to a canonical item dict."""
        item: dict[str, Any] = {
            "item_name": "", "sku_code": "", "hsn_sac": "",
            "gst_rate": "", "rate": 0.0, "measurement": "", "quantity": 0.0,
        }
        for col_idx, field in col_map.items():
            if col_idx < len(cells):
                val = cells[col_idx]
                if field in ("rate", "quantity"):
                    item[field] = _parse_number(val) if val else 0.0
                else:
                    item[field] = val
        return item

    # ── Text-Based Extraction ──────────────────────────────────────────────

    def _extract_from_text(self, text: str) -> list[dict]:
        """
        Attempt to extract line items from raw text using heuristics.
        Works best when each line item is on a single line.
        """
        items: list[dict] = []
        lines = text.splitlines()

        # Find the line that contains the header row
        header_line_idx, col_positions = self._find_header_line(lines)
        if header_line_idx is None:
            return []

        for line in lines[header_line_idx + 1:]:
            if not line.strip():
                continue
            if _should_skip_row([line]):
                continue

            # Try to split the line into columns using positional alignment
            item = self._parse_text_line(line, col_positions)
            if item and item.get("item_name"):
                items.append(item)

        return items

    def _find_header_line(
        self, lines: list[str]
    ) -> tuple[Optional[int], dict[str, int]]:
        """
        Scan lines to find the invoice table header.
        Returns (line_index, {field: char_position}).
        """
        all_synonyms: list[str] = []
        for syns in HEADER_SYNONYMS.values():
            all_synonyms.extend(syns)
        matcher = FuzzyMatcher(threshold=68)

        for idx, line in enumerate(lines):
            words = re.findall(r"\S+", line)
            if len(words) < 3:
                continue
            hits = sum(1 for w in words
                       if matcher.best_match(w.lower(), all_synonyms)[1] >= 68)
            if hits >= 3:
                # Build positional map: field → character offset in line
                positions: dict[str, int] = {}
                for field, syns in HEADER_SYNONYMS.items():
                    for syn in syns:
                        m = re.search(re.escape(syn), line, re.IGNORECASE)
                        if m:
                            positions[field] = m.start()
                            break
                if len(positions) >= 3:
                    return idx, positions
        return None, {}

    def _parse_text_line(self, line: str, col_positions: dict[str, int]) -> dict:
        """
        Use character offsets from the header to extract cell values from a line.
        Falls back to whitespace splitting if offsets are unavailable.
        """
        if not col_positions:
            parts = line.split()
            if len(parts) < 2:
                return {}
            return {
                "item_name": parts[0] if len(parts) > 0 else "",
                "sku_code": "",
                "hsn_sac": "",
                "gst_rate": "",
                "rate": _parse_number(parts[-2]) if len(parts) >= 2 else 0.0,
                "measurement": "",
                "quantity": _parse_number(parts[-1]) if len(parts) >= 1 else 0.0,
            }

        # Sort fields by their position in the header line
        sorted_fields = sorted(col_positions.items(), key=lambda x: x[1])
        item: dict[str, Any] = {
            "item_name": "", "sku_code": "", "hsn_sac": "",
            "gst_rate": "", "rate": 0.0, "measurement": "", "quantity": 0.0,
        }

        for i, (field, start_pos) in enumerate(sorted_fields):
            end_pos = (sorted_fields[i + 1][1]
                       if i + 1 < len(sorted_fields) else len(line))
            cell = line[start_pos:end_pos].strip() if start_pos < len(line) else ""
            if field in ("rate", "quantity"):
                item[field] = _parse_number(cell) if cell else 0.0
            else:
                item[field] = cell

        return item

    # ── Post-Processing ────────────────────────────────────────────────────

    def _clean_items(self, items: list[dict]) -> list[dict]:
        """Filter and normalise extracted items."""
        cleaned: list[dict] = []
        for item in items:
            # Skip items with no recognisable name
            if not item.get("item_name") or len(item["item_name"]) < 2:
                continue
            # Skip rows that look like headers sneaking through
            if item["item_name"].lower() in {s for ss in HEADER_SYNONYMS.values()
                                             for s in ss}:
                continue
            # Ensure quantity and rate are clean floats (never raw strings)
            item["quantity"] = max(0.0, _parse_number(str(item.get("quantity") or 0)))
            item["rate"]     = max(0.0, _parse_number(str(item.get("rate")     or 0)))

            # Normalise GST rate to a percentage string  "18%"
            gst_raw = str(item.get("gst_rate", "")).replace("%", "").strip()
            if _is_numeric(gst_raw):
                item["gst_rate"] = f"{float(gst_raw):.0f}%"

            # Strip leading serial numbers from item names  ("1. Sugar" → "Sugar")
            item["item_name"] = re.sub(r"^\d+[\.\)]\s*", "",
                                        item["item_name"]).strip()

            cleaned.append(item)
        return cleaned
