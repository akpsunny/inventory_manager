"""
core/consumption_processor.py
──────────────────────────────
Parses a "Consumption Report" file (PDF, Excel, or CSV) and returns a
normalised list of consumed items.

Supported formats
─────────────────
  • .xlsx / .xls  – pandas read_excel
  • .csv           – pandas read_csv
  • .pdf           – pdfplumber tables → fallback text extraction

Expected columns in the report (flexible header matching):
  Item Name  |  SKU Code  |  Quantity  (minimum required)
  Optional:  Period / Date / Notes
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import pdfplumber

from core.fuzzy_matcher import FuzzyMatcher

log = logging.getLogger(__name__)

# ─── Header synonyms for consumption report columns ───────────────────────────
CONS_HEADER_SYNONYMS: dict[str, list[str]] = {
    "item_name": ["description", "item", "product", "particulars",
                  "item name", "product name", "goods", "name", "material"],
    "sku_code":  ["sku", "code", "item code", "product code",
                  "article", "part no", "part number", "material code"],
    "quantity":  ["qty", "quantity", "consumed", "used", "issued",
                  "consumption", "nos", "units", "count", "amount"],
}


def _parse_number(text: Any) -> float:
    """Convert a cell value to a float, return 0.0 on failure."""
    try:
        return float(str(text).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def _clean_str(value: Any) -> str:
    return str(value).strip() if value is not None else ""


# ─── Header mapper ────────────────────────────────────────────────────────────

class ConsHeaderMapper:
    """Maps raw column headers in the consumption report to canonical names."""

    def __init__(self):
        self._matcher = FuzzyMatcher(threshold=68)
        self._lookup: dict[str, str] = {}
        for field, synonyms in CONS_HEADER_SYNONYMS.items():
            for s in synonyms:
                self._lookup[s.lower()] = field

    def map_columns(self, raw_headers: list[str]) -> dict[str, str]:
        """Return {raw_header: canonical_field} for recognised columns."""
        all_synonyms = list(self._lookup.keys())
        result: dict[str, str] = {}

        for h in raw_headers:
            h_norm = str(h).strip().lower()
            if not h_norm:
                continue
            # Direct
            if h_norm in self._lookup:
                result[h] = self._lookup[h_norm]
                continue
            # Fuzzy
            best, score = self._matcher.best_match(h_norm, all_synonyms)
            if best:
                result[h] = self._lookup[best]

        return result


# ─── ConsumptionProcessor ─────────────────────────────────────────────────────

class ConsumptionProcessor:
    """
    Parses a consumption report file and returns a list of dicts:
      [{"item_name": str, "sku_code": str, "quantity": float}, …]
    """

    def __init__(self):
        self._header_mapper = ConsHeaderMapper()

    # ── Public entry point ─────────────────────────────────────────────────

    def parse(self, file_path: str) -> list[dict]:
        """
        Dispatch to the correct parser based on file extension.

        Returns a list of item dicts, always with keys:
          item_name, sku_code, quantity
        """
        path   = Path(file_path)
        suffix = path.suffix.lower()

        if suffix in {".xlsx", ".xls"}:
            items = self._parse_excel(str(path))
        elif suffix == ".csv":
            items = self._parse_csv(str(path))
        elif suffix == ".pdf":
            items = self._parse_pdf(str(path))
        else:
            raise ValueError(f"Unsupported consumption report format: {suffix}")

        cleaned = self._clean(items)
        log.info("Consumption parse: %d items from %s", len(cleaned), path.name)
        return cleaned

    # ── Excel ──────────────────────────────────────────────────────────────

    def _parse_excel(self, path: str) -> list[dict]:
        """Read the first sheet of an Excel file."""
        try:
            # Try to detect header row (first 5 rows)
            for skip in range(5):
                try:
                    df = pd.read_excel(path, skiprows=skip, dtype=str)
                    items = self._df_to_items(df)
                    if items:
                        return items
                except Exception:
                    continue
        except Exception as e:
            log.error("Excel parse error: %s", e)
        return []

    # ── CSV ────────────────────────────────────────────────────────────────

    def _parse_csv(self, path: str) -> list[dict]:
        """Read a CSV file, trying common encodings."""
        for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
            try:
                df = pd.read_csv(path, dtype=str, encoding=enc)
                items = self._df_to_items(df)
                if items:
                    return items
            except Exception:
                continue
        return []

    # ── PDF ────────────────────────────────────────────────────────────────

    def _parse_pdf(self, path: str) -> list[dict]:
        """Extract consumption data from a PDF file."""
        items: list[dict] = []

        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    extracted = self._table_to_items(table)
                    items.extend(extracted)

        # Fallback: text extraction
        if not items:
            items = self._pdf_text_fallback(path)

        return items

    def _pdf_text_fallback(self, path: str) -> list[dict]:
        """Parse raw text from a PDF when no tables are found."""
        items: list[dict] = []
        with pdfplumber.open(path) as pdf:
            full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)

        lines = full_text.splitlines()
        header_idx, col_map = self._detect_header_line(lines)
        if header_idx is None:
            log.warning("Could not detect header line in PDF consumption report.")
            return []

        for line in lines[header_idx + 1:]:
            line = line.strip()
            if not line:
                continue
            item = self._parse_text_line(line, col_map)
            if item:
                items.append(item)

        return items

    def _detect_header_line(
        self, lines: list[str]
    ) -> tuple[Optional[int], dict[str, int]]:
        """Find the header row in a list of text lines."""
        matcher = FuzzyMatcher(threshold=65)
        all_syns = [s for ss in CONS_HEADER_SYNONYMS.values() for s in ss]

        for idx, line in enumerate(lines):
            words = re.findall(r"\S+", line)
            hits  = sum(1 for w in words
                        if matcher.best_match(w.lower(), all_syns)[1] >= 65)
            if hits >= 2:
                # Build character-position map
                col_map: dict[str, int] = {}
                for field, syns in CONS_HEADER_SYNONYMS.items():
                    for syn in syns:
                        m = re.search(re.escape(syn), line, re.IGNORECASE)
                        if m:
                            col_map[field] = m.start()
                            break
                if col_map:
                    return idx, col_map
        return None, {}

    def _parse_text_line(self, line: str, col_map: dict[str, int]) -> Optional[dict]:
        """Extract an item from a single text line using header positions."""
        sorted_fields = sorted(col_map.items(), key=lambda x: x[1])
        item: dict[str, Any] = {"item_name": "", "sku_code": "", "quantity": 0.0}

        for i, (field, start) in enumerate(sorted_fields):
            end   = sorted_fields[i + 1][1] if i + 1 < len(sorted_fields) else len(line)
            cell  = line[start:end].strip() if start < len(line) else ""
            if field == "quantity":
                item["quantity"] = _parse_number(cell)
            else:
                item[field] = cell

        if not item["item_name"]:
            return None
        return item

    # ── DataFrame → item list ──────────────────────────────────────────────

    def _df_to_items(self, df: pd.DataFrame) -> list[dict]:
        """Convert a pandas DataFrame to a list of item dicts."""
        if df.empty:
            return []

        col_map = self._header_mapper.map_columns(list(df.columns))
        if "quantity" not in col_map.values():
            return []   # Can't identify the quantity column

        items: list[dict] = []
        for _, row in df.iterrows():
            item: dict[str, Any] = {"item_name": "", "sku_code": "", "quantity": 0.0}
            for raw_col, canonical in col_map.items():
                val = row.get(raw_col, "")
                if canonical == "quantity":
                    item["quantity"] = _parse_number(val)
                else:
                    item[canonical] = _clean_str(val)
            if item.get("item_name") or item.get("sku_code"):
                items.append(item)

        return items

    def _table_to_items(self, table: list[list]) -> list[dict]:
        """Convert a pdfplumber table to item dicts."""
        if not table or len(table) < 2:
            return []

        # Find header row (first row with non-numeric, non-empty cells)
        header_row = table[0]
        raw_headers = [_clean_str(c) for c in header_row]
        col_map = self._header_mapper.map_columns(raw_headers)

        if "quantity" not in col_map.values():
            return []

        # Build index: canonical_field → column_index
        field_to_idx: dict[str, int] = {}
        for col_idx, raw_h in enumerate(raw_headers):
            if raw_h in col_map:
                field_to_idx[col_map[raw_h]] = col_idx

        items: list[dict] = []
        for row in table[1:]:
            cells = [_clean_str(c) for c in row]
            if not any(cells):
                continue

            item: dict[str, Any] = {"item_name": "", "sku_code": "", "quantity": 0.0}
            for field, idx in field_to_idx.items():
                if idx < len(cells):
                    val = cells[idx]
                    if field == "quantity":
                        item["quantity"] = _parse_number(val)
                    else:
                        item[field] = val

            if item.get("item_name") or item.get("sku_code"):
                items.append(item)

        return items

    # ── Cleaning ───────────────────────────────────────────────────────────

    def _clean(self, items: list[dict]) -> list[dict]:
        """Remove invalid rows, strip whitespace, ensure positive quantities."""
        cleaned: list[dict] = []
        for item in items:
            name = str(item.get("item_name", "")).strip()
            sku  = str(item.get("sku_code",  "")).strip()
            qty  = float(item.get("quantity",  0) or 0)

            # Need at least a name or SKU
            if not name and not sku:
                continue

            # Skip header-like rows
            if name.lower() in {"item name", "description", "product", "particulars"}:
                continue

            # Strip leading serial numbers from names
            name = re.sub(r"^\d+[\.\)]\s*", "", name).strip()

            if qty < 0:
                qty = 0.0

            cleaned.append({
                "item_name": name,
                "sku_code":  sku,
                "quantity":  qty,
            })

        return cleaned
