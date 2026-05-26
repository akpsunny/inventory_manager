"""
core/invoice_parser.py  — Word-position based invoice parser
"""
from __future__ import annotations

import re, logging
from collections import defaultdict
from pathlib import Path
from typing import Any, Optional

import pdfplumber
try:
    import fitz; FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

from PIL import Image
import pytesseract
from core.fuzzy_matcher import FuzzyMatcher

log = logging.getLogger(__name__)

# ─── Field synonyms ───────────────────────────────────────────────────────────
HEADER_SYNONYMS: dict[str, list[str]] = {
    "item_name":   ["description", "description of goods", "item", "product",
                    "particulars", "goods", "name", "item name", "product name",
                    "details", "narration"],
    "sku_code":    ["sku", "sku code", "code", "item code", "product code",
                    "article", "part no", "part number", "material code", "item no"],
    "hsn_sac":     ["hsn", "sac", "hsn/sac", "hsn code", "sac code", "hsnsac", "tariff"],
    "gst_rate":    ["gst", "gst rate", "gst rate %", "tax rate", "gst%", "tax %"],
    "rate":        ["rate", "price", "unit price", "mrp", "cost", "unit cost", "unit rate"],
    "measurement": ["unit", "uom", "measure", "per", "pack", "packing"],
    "quantity":    ["qty", "quantity", "nos", "no.", "pcs", "pieces", "units", "count"],
}

# Words to skip when scanning for field keywords
IGNORE_HEADER_WORDS = {
    "s", "no", "sr", "sl", "sno", "serial", "#", "of",
    "discount", "disc", "amount", "total", "value",
    "cgst", "sgst", "igst", "tax", "base", "invoice", "the",
}

# Words that signal the END of the item table (ignore columns)
STOP_HEADER_WORDS = {"discount", "disc", "amount", "total", "value", "tax", "base"}

DATE_LABEL_RE = re.compile(
    r"(?:invoice\s*date|date|dated|dt\.?)\s*[:\-]?\s*"
    r"(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{4}|\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})",
    re.IGNORECASE)
DATE_BARE_RE = re.compile(r"\b(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{4})\b")

# Invoice-number patterns — two-pass so 'Invoice No.' wins over 'Bill No.' when
# both appear (FNP invoices have an empty 'EWay Bill No.' field right next to
# the real 'Invoice No.'). [ \t]* after the separator (not \s*) keeps the
# capture on the same line, so an empty labelled field can't bleed into the
# next line's first word.
_INV_NUM_PRIMARY_RE = re.compile(
    r"\b(?:invoice|inv\.?)\s*(?:no\.?|number|#)\s*\.?\s*[:\-=]?[ \t]*"
    r"([A-Za-z0-9][A-Za-z0-9\-\/_\.]{2,})", re.IGNORECASE)
_INV_NUM_FALLBACK_RE = re.compile(
    r"\bbill\s*(?:no\.?|number|#)\s*\.?\s*[:\-=]?[ \t]*"
    r"([A-Za-z0-9][A-Za-z0-9\-\/_\.]{2,})", re.IGNORECASE)

SKIP_ROW_RE = [re.compile(p, re.IGNORECASE) for p in [
    r"sub[\s-]?total|grand\s*total|round\s*off|amount\s*payable",
    r"output\s*(igst|cgst|sgst)",
    r"amount\s+chargeable",
    r"^\s*\d{6,8}\s+\d",
]]


def _extract_invoice_number(text: str) -> str:
    """
    Find the supplier's invoice number from document text.
    Two-pass: prefer 'Invoice No./Number/#' over 'Bill No./Number/#'.
    Returns the stripped number, or '' if none was found.
    """
    m = _INV_NUM_PRIMARY_RE.search(text)
    if m:
        return m.group(1).strip().rstrip(".,;:")
    m = _INV_NUM_FALLBACK_RE.search(text)
    if m:
        return m.group(1).strip().rstrip(".,;:")
    return ""



def _parse_number(text: Any) -> float:
    if text is None: return 0.0
    s = str(text).strip()
    if not s or s in ("-", "–"): return 0.0
    cleaned = re.sub(r"[₹$€£,%]", " ", s).strip()
    try:
        return float(cleaned.replace(",", "").replace(" ", ""))
    except ValueError:
        pass
    for token in cleaned.split():
        t = token.replace(",", "")
        try:
            v = float(t)
            if len(t.replace(".", "")) >= 8 and "." not in t: continue
            return v
        except ValueError:
            continue
    return 0.0


def _clean_cell(v: Any) -> str:
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def _normalise_date(raw: str) -> str:
    try:
        import dateutil.parser as dup
        c = raw.replace(".", "/").replace("-", "/")
        parts = c.split("/")
        if len(parts) == 3 and len(parts[0]) == 4:
            c = f"{parts[2]}/{parts[1]}/{parts[0]}"
        return dup.parse(c, dayfirst=True).strftime("%d/%m/%Y")
    except Exception:
        return raw


def _extract_date(text: str) -> Optional[str]:
    m = DATE_LABEL_RE.search(text)
    if m: return _normalise_date(m.group(1))
    m = DATE_BARE_RE.search(text[:800])
    if m: return _normalise_date(m.group(1))
    return None


def _should_skip(text: str) -> bool:
    return any(p.search(text) for p in SKIP_ROW_RE)


# ─── Word-Position Table Reconstructor ────────────────────────────────────────

class WordTableExtractor:
    """
    Rebuilds a structured table from PDF word X/Y positions.

    Algorithm
    ─────────
    1.  Snap word Y positions to a 4pt grid → group into rows.
    2.  Find the header row using EXACT synonym matches only (avoids false positives
        from fuzzy-matching company names / address words).
    3.  Build column X-ranges from header word positions:
        •  If the same field keyword appears twice with gap > 50pt, the first is
           a compound-header word (e.g. "Gst *Rate*") — update to the far occurrence.
        •  Track the X of first STOP word (Discount / Amount) to cap the last field.
    4.  For each data row assign words to columns by X midpoint.
    5.  Rows with no numeric value in numeric-field columns are "continuation rows"
        (wrapped SKU suffix, item-name overflow) — merged into the previous item.
    """

    Y_SNAP          = 4    # pt — snap tolerance for Y grouping
    COL_HALF_GAP    = 9    # pt — half the gap subtracted from each col's x_start
    REMAP_FAR_THRESHOLD = 50   # pt — re-map a field keyword if it reappears this far right

    def __init__(self):
        # Flat exact-match lookup: synonym → field
        self._syn_to_field: dict[str, str] = {}
        for field, syns in HEADER_SYNONYMS.items():
            for s in syns:
                self._syn_to_field[s.lower()] = field

    # ── Public ─────────────────────────────────────────────────────────────

    def extract(self, page) -> list[dict]:
        words = page.extract_words(x_tolerance=3, y_tolerance=3)
        if not words:
            return []

        y_rows    = self._group_by_y(words)
        sorted_ys = sorted(y_rows.keys())

        header = self._find_header(y_rows, sorted_ys)
        if header is None:
            return []

        header_y_idx, col_ranges, x_max = header

        # Skip a two-row header continuation (e.g. "No  %  %")
        first_data_idx = header_y_idx + 1
        if first_data_idx < len(sorted_ys):
            nxt = [w["text"].strip() for w in y_rows[sorted_ys[first_data_idx]] if w["text"].strip()]
            if nxt and all(len(t) <= 3 for t in nxt):
                first_data_idx += 1

        return self._parse_items(y_rows, sorted_ys, first_data_idx, col_ranges, x_max)

    # ── Y grouping ─────────────────────────────────────────────────────────

    def _group_by_y(self, words: list[dict]) -> dict[int, list[dict]]:
        rows: dict[int, list[dict]] = defaultdict(list)
        for w in words:
            y_key = round(w["top"] / self.Y_SNAP) * self.Y_SNAP
            rows[y_key].append(w)
        return rows

    # ── Header detection ───────────────────────────────────────────────────

    def _find_header(
        self, y_rows: dict, sorted_ys: list
    ) -> Optional[tuple[int, list[tuple], float]]:
        """
        Scan rows for the item table header.

        Requires EXACT synonym matches (no fuzzy) and at least 4 field hits
        including both an item-name field and a quantity field — this prevents
        false positives from company name / address rows.

        Returns (y_index, col_ranges, x_max_boundary).
        """
        for y_idx, y in enumerate(sorted_ys):
            row_words = sorted(y_rows[y], key=lambda w: w["x0"])
            result    = self._try_header(row_words)
            if result:
                field_positions, stop_x = result
                # Require 4+ fields including name + quantity
                if len(field_positions) < 4:
                    continue
                if not (("item_name" in field_positions or "sku_code" in field_positions)
                        and "quantity" in field_positions):
                    continue
                col_ranges = self._build_ranges(field_positions, stop_x)
                return y_idx, col_ranges, stop_x
        return None

    def _try_header(
        self, row_words: list[dict]
    ) -> Optional[tuple[dict[str, float], float]]:
        """
        Try to build a field→x_start map from a row's words using EXACT matches.
        Returns (field_positions, stop_x) or None.

        stop_x: leftmost X of any STOP word (Discount/Amount) — used to cap the
                last recognised field's range so trailing values don't bleed in.
        """
        field_positions: dict[str, float] = {}
        stop_x: float = float("inf")

        for w in row_words:
            wt = w["text"].strip().lower()
            if not wt:
                continue

            # Track stop-word positions (Discount, Amount, etc.)
            if wt in STOP_HEADER_WORDS:
                stop_x = min(stop_x, w["x0"])
                continue

            if wt in IGNORE_HEADER_WORDS:
                continue

            field = self._syn_to_field.get(wt)
            if field is None:
                continue  # Exact match only for header detection

            if field not in field_positions:
                field_positions[field] = w["x0"]
            else:
                # A keyword can appear twice (e.g. "Gst Rate … Rate").
                # Only remap if ANOTHER field has been mapped between the two
                # occurrences — that proves the second occurrence is a separate column.
                # (e.g. "Quantity" sits between the two "Rate" words → real Rate column)
                # This avoids remapping "Description … Goods" which are the same column.
                first_x  = field_positions[field]
                second_x = w["x0"]
                between  = any(
                    first_x < x < second_x
                    for f, x in field_positions.items()
                    if f != field
                )
                if between:
                    field_positions[field] = second_x

        return (field_positions, stop_x) if field_positions else None

    def _build_ranges(
        self, field_positions: dict[str, float], stop_x: float
    ) -> list[tuple]:
        """Convert {field: x_start} → sorted (x_start, x_end, field) tuples."""
        sorted_fields = sorted(field_positions.items(), key=lambda kv: kv[1])
        ranges = []
        for i, (field, x_start) in enumerate(sorted_fields):
            if i + 1 < len(sorted_fields):
                x_end = sorted_fields[i + 1][1]
            else:
                x_end = stop_x  # cap last field at Discount/Amount column start
            ranges.append((x_start - self.COL_HALF_GAP, x_end, field))
        return ranges

    # ── Item parsing ────────────────────────────────────────────────────────

    NUMERIC_FIELDS = {"gst_rate", "quantity", "rate"}

    def _parse_items(
        self,
        y_rows: dict,
        sorted_ys: list,
        first_idx: int,
        col_ranges: list[tuple],
        x_max: float,
    ) -> list[dict]:
        items: list[dict] = []
        current: dict | None = None

        for y in sorted_ys[first_idx:]:
            row_words = sorted(y_rows[y], key=lambda w: w["x0"])

            # Assign words to columns
            row_data: dict[str, list[str]] = defaultdict(list)
            has_numeric = False

            for w in row_words:
                x_mid = (w["x0"] + w["x1"]) / 2
                # Skip words past the stop boundary
                if x_mid >= x_max:
                    continue
                for x_start, x_end, field in col_ranges:
                    if x_start <= x_mid < x_end:
                        row_data[field].append(w["text"])
                        if field in self.NUMERIC_FIELDS:
                            try:
                                float(w["text"].replace(",", ""))
                                has_numeric = True
                            except ValueError:
                                pass
                        break

            row_texts = {f: " ".join(ts).strip() for f, ts in row_data.items() if ts}

            combined = " ".join(row_texts.values())
            if _should_skip(combined):
                if current:
                    items.append(current)
                    current = None
                break

            if has_numeric:
                if current and (current.get("item_name") or current.get("sku_code")):
                    items.append(current)
                current = self._make_item(row_texts)
            elif row_texts and current:
                self._merge(current, row_texts)
            elif row_texts and not current:
                current = self._make_item(row_texts)

        if current and (current.get("item_name") or current.get("sku_code")):
            items.append(current)
        return items

    @staticmethod
    def _make_item(row_texts: dict[str, str]) -> dict:
        item: dict[str, Any] = {
            "item_name": "", "sku_code": "", "hsn_sac": "",
            "gst_rate": "", "rate": 0.0, "measurement": "", "quantity": 0.0,
        }
        for field, text in row_texts.items():
            if field in ("rate", "quantity"):
                item[field] = _parse_number(text)
            else:
                item[field] = text
        return item

    @staticmethod
    def _merge(item: dict, row_texts: dict[str, str]):
        """Append continuation-row text into the current item."""
        for field, text in row_texts.items():
            if field in ("rate", "quantity", "gst_rate"):
                continue
            existing = item.get(field, "")
            sep = "" if field == "sku_code" else " "
            item[field] = (existing + sep + text).strip() if existing else text


# ─── Main Parser ──────────────────────────────────────────────────────────────

class InvoiceParser:

    def __init__(self):
        self._word_extractor = WordTableExtractor()

    def parse(self, file_path: str) -> tuple[list[dict], str]:
        path   = Path(file_path)
        suffix = path.suffix.lower()
        if suffix == ".pdf":
            return self._parse_pdf(str(path))
        elif suffix in {".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".webp"}:
            return self._parse_image(str(path))
        else:
            raise ValueError(f"Unsupported file type: {suffix}")

    # ── Lightweight metadata peek (for duplicate-invoice detection) ────────

    def peek_metadata(self, file_path: str) -> dict:
        """
        Quickly extract identifying metadata (invoice number + date) without
        running the full word-position line-item extractor. Used to dedupe
        invoices before paying the full parse cost.

        Returns
        -------
        dict with keys:
          invoice_number : str  – e.g. 'PF2511DL-0000876', or '' if not found
          invoice_date   : str  – DD/MM/YYYY, or '' if not detected
        Either field may be empty when the document is unusual; callers
        should treat empty values as 'unknown' and fall back to the file
        name + UI-entered date for matching.
        """
        path   = Path(file_path)
        suffix = path.suffix.lower()
        text   = ""

        if suffix == ".pdf":
            try:
                with pdfplumber.open(str(path)) as pdf:
                    # First 2 pages are enough — labelled fields like
                    # 'Invoice No.' and 'Date' live in the page header.
                    for page in pdf.pages[:2]:
                        text += (page.extract_text() or "") + "\n"
            except Exception as e:
                log.warning("peek_metadata: pdfplumber failed (%s)", e)

        elif suffix in {".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".webp"}:
            try:
                img = Image.open(str(path))
                if min(img.width, img.height) < 1200:
                    scale = 2400 / min(img.width, img.height)
                    img = img.resize(
                        (int(img.width*scale), int(img.height*scale)),
                        Image.LANCZOS)
                img = img.convert("L")
                text = pytesseract.image_to_string(
                    img, config=r"--oem 3 --psm 6")
            except Exception as e:
                log.warning("peek_metadata: OCR failed (%s)", e)

        return {
            "invoice_number": _extract_invoice_number(text),
            "invoice_date":   _extract_date(text) or "",
        }

    # ── PDF ────────────────────────────────────────────────────────────────

    def _parse_pdf(self, path: str) -> tuple[list[dict], str]:
        full_text = ""
        all_items: list[dict] = []

        with pdfplumber.open(path) as pdf:
            for pg_num, page in enumerate(pdf.pages):
                pg_text   = page.extract_text() or ""
                full_text += pg_text + "\n"

                if self._is_summary_page(pg_text):
                    log.debug("Skipping summary page %d", pg_num + 1)
                    continue

                # Primary: word-position extraction
                items = self._word_extractor.extract(page)

                # Fallback: pdfplumber table extraction
                if not items:
                    items = self._table_fallback(page)

                log.debug("Page %d: %d items", pg_num + 1, len(items))
                all_items.extend(items)

        # Text fallback
        if not all_items:
            all_items = self._text_fallback(full_text)

        # PyMuPDF last resort
        if not all_items and FITZ_AVAILABLE:
            doc = fitz.open(path)
            txt = "\n".join(p.get_text("text") for p in doc)
            doc.close()
            all_items  = self._text_fallback(txt)
            full_text += txt

        return self._clean_items(all_items), _extract_date(full_text) or ""

    def _is_summary_page(self, text: str) -> bool:
        tl = text.lower()
        hits = sum(1 for ind in ["tax analysis", "tax base", "total tax amount",
                                  "amount chargeable"] if ind in tl)
        if hits >= 2:
            return not bool(re.search(
                r"\b(piece|meter|kilo|nos|pcs|kg|gm|ltr|ml)\b", text, re.IGNORECASE))
        return False

    # ── pdfplumber table fallback ──────────────────────────────────────────

    def _table_fallback(self, page) -> list[dict]:
        syn_to_field: dict[str, str] = {}
        for field, syns in HEADER_SYNONYMS.items():
            for s in syns:
                syn_to_field[s.lower()] = field
        all_syns = list(syn_to_field.keys())
        matcher  = FuzzyMatcher(threshold=68)

        for strat in [{"vertical_strategy": "lines", "horizontal_strategy": "lines"}, {}]:
            try:
                tables = page.extract_tables(strat) if strat else page.extract_tables()
            except Exception:
                continue
            items = []
            for table in tables:
                if not table or len(table) < 2:
                    continue
                # Find header row
                h_idx = None
                for i, row in enumerate(table[:6]):
                    cells = [_clean_cell(c).lower() for c in row if c]
                    hits  = sum(1 for c in cells
                                if c in syn_to_field
                                or matcher.best_match(c, all_syns)[1] >= 68)
                    if hits >= 3:
                        h_idx = i; break
                if h_idx is None:
                    continue
                raw_h   = [_clean_cell(c) for c in table[h_idx]]
                col_map: dict[int, str] = {}
                used:    set[str]       = set()
                for idx, h in enumerate(raw_h):
                    f = syn_to_field.get(h.lower())
                    if f is None:
                        best, _ = matcher.best_match(h.lower(), all_syns)
                        f = syn_to_field.get(best) if best else None
                    if f and f not in used:
                        col_map[idx] = f; used.add(f)
                if "item_name" not in col_map.values() and "sku_code" not in col_map.values():
                    continue
                for row in table[h_idx + 1:]:
                    cells = [_clean_cell(c) for c in row]
                    if not any(cells): continue
                    if _should_skip(" ".join(cells)): break
                    item: dict[str, Any] = {
                        "item_name":"","sku_code":"","hsn_sac":"",
                        "gst_rate":"","rate":0.0,"measurement":"","quantity":0.0}
                    for ci, f in col_map.items():
                        if ci < len(cells):
                            val = cells[ci]
                            if f in ("rate","quantity"): item[f] = _parse_number(val)
                            else: item[f] = val
                    if item.get("item_name") or item.get("sku_code"):
                        items.append(item)
            if items:
                return items
        return []

    # ── Image (OCR) ────────────────────────────────────────────────────────

    def _parse_image(self, path: str) -> tuple[list[dict], str]:
        img = Image.open(path)
        if min(img.width, img.height) < 1200:
            scale = 2400 / min(img.width, img.height)
            img = img.resize((int(img.width*scale), int(img.height*scale)), Image.LANCZOS)
        img      = img.convert("L")
        raw_text = pytesseract.image_to_string(img, config=r"--oem 3 --psm 6")
        return self._clean_items(self._text_fallback(raw_text)), _extract_date(raw_text) or ""

    # ── Text extraction fallback ───────────────────────────────────────────

    def _text_fallback(self, text: str) -> list[dict]:
        syn_to_field: dict[str, str] = {}
        for field, syns in HEADER_SYNONYMS.items():
            for s in syns:
                syn_to_field[s.lower()] = field
        all_syns = list(syn_to_field.keys())
        matcher  = FuzzyMatcher(threshold=68)
        lines    = text.splitlines()
        h_idx    = None
        col_pos: dict[str, int] = {}
        for i, line in enumerate(lines):
            words = re.findall(r"\S+", line)
            hits  = sum(1 for w in words if matcher.best_match(w.lower(), all_syns)[1] >= 68)
            if hits >= 3:
                for field, syns in HEADER_SYNONYMS.items():
                    for syn in syns:
                        m = re.search(re.escape(syn), line, re.IGNORECASE)
                        if m and field not in col_pos:
                            col_pos[field] = m.start(); break
                if len(col_pos) >= 3:
                    h_idx = i; break
        if h_idx is None:
            return []
        sf = sorted(col_pos.items(), key=lambda x: x[1])
        items = []
        for line in lines[h_idx + 1:]:
            if not line.strip() or _should_skip(line): continue
            item: dict[str, Any] = {
                "item_name":"","sku_code":"","hsn_sac":"",
                "gst_rate":"","rate":0.0,"measurement":"","quantity":0.0}
            for i, (field, start) in enumerate(sf):
                end  = sf[i+1][1] if i+1 < len(sf) else len(line)
                cell = line[start:end].strip() if start < len(line) else ""
                if field in ("rate","quantity"): item[field] = _parse_number(cell)
                else: item[field] = cell
            if item.get("item_name") or item.get("sku_code"):
                items.append(item)
        return items

    # ── Post-processing ────────────────────────────────────────────────────

    def _clean_items(self, items: list[dict]) -> list[dict]:
        """
        Basic cleanup — NO deduplication.
        When the same SKU appears multiple times in one invoice (e.g. multiple
        batches of the same product on different pages), both rows are kept and
        update_with_invoice() will correctly SUM the quantities into one master row.
        """
        cleaned: list[dict] = []
        all_syns = {s for ss in HEADER_SYNONYMS.values() for s in ss}
        for item in items:
            name = str(item.get("item_name","")).strip()
            sku  = str(item.get("sku_code","")).strip()
            if not name and not sku: continue
            if name.lower() in all_syns: continue
            if re.match(r"^\d{6,8}$", name): continue   # pure HSN code row
            name = re.sub(r"^\d+[\.\)\s]\s*","", name).strip()
            if not name and not sku: continue
            item["item_name"] = name
            item["quantity"]  = max(0.0, _parse_number(str(item.get("quantity") or 0)))
            item["rate"]      = max(0.0, _parse_number(str(item.get("rate") or 0)))
            gst = str(item.get("gst_rate","")).replace("%","").strip()
            item["gst_rate"]  = f"{_parse_number(gst):.0f}%" if gst and _parse_number(gst)>0 else ""

            # Column-misalignment guard: when the SKU column is empty in the
            # source PDF, the WordTableExtractor sometimes pulls a numeric value
            # from an adjacent column (usually the GST rate) leftward into the
            # SKU slot. Tell-tale signature: the captured "SKU" is a bare 1-2
            # digit integer. Real supplier SKUs are virtually always
            # alphanumeric (CON078, SKUHDR-0001402, PER001, etc.) or at least
            # 3+ digits; a single number like "18", "5", or "12" is almost
            # always GST or some other numeric leak. Clear it so the downstream
            # matcher uses the item name instead. Without this, all such rows
            # collapse into one row per leaked value (FNP invoice had 7 rows
            # merging into a single 'SKU 18' row, losing 6 distinct products).
            if sku.isdigit() and len(sku) <= 2:
                item["sku_code"] = ""

            meas = str(item.get("measurement","")).strip()
            if re.match(r"^\d", meas): item["measurement"] = ""
            cleaned.append(item)
        return cleaned
