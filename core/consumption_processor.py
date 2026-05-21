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

Supported column name variants (flexible header matching):
  Item Name   : COMPONENT_NAME, description, item, product, particulars …
  SKU / ID    : COMPONENTID, component id, sku, code, item code …
  Quantity    : QUANTITYUSED, quantity used, qty used, consumed, issued …
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

# ─── Header synonyms (deliberately broad to cover varied supplier formats) ────
CONS_HEADER_SYNONYMS: dict[str, list[str]] = {
    "item_name": [
        "component_name", "component name", "componentname",
        "description", "description of goods", "item", "product",
        "particulars", "item name", "product name", "goods", "name",
        "material", "narration",
    ],
    "sku_code": [
        "componentid", "component id", "component_id",
        "sku", "sku code", "code", "item code", "product code",
        "article", "part no", "part number", "material code", "item no",
    ],
    "quantity": [
        "quantityused", "quantity used", "qty used", "qty_used",
        "quantity_used", "used_qty", "usedqty",
        "consumed", "consumption", "issued", "dispatched",
        "qty", "quantity", "nos", "units", "count",
        "quantityrequired", "quantity required",   # fallback if no "used" col
    ],
}


def _parse_number(value: Any) -> float:
    try:
        return float(str(value).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def _clean_str(value: Any) -> str:
    return str(value).strip() if value is not None else ""


# ─── Header mapper ────────────────────────────────────────────────────────────

class ConsHeaderMapper:
    """Maps raw column headers in the consumption report to canonical names."""

    def __init__(self):
        self._matcher = FuzzyMatcher(threshold=65)
        self._lookup: dict[str, str] = {}
        for field, synonyms in CONS_HEADER_SYNONYMS.items():
            for s in synonyms:
                self._lookup[s.lower()] = field

    def map_columns(self, raw_headers: list[str]) -> dict[str, str]:
        """
        Return {raw_header: canonical_field}.
        For 'quantity' prefer QUANTITYUSED over QUANTITY REQUIRED if both present.
        """
        all_synonyms = list(self._lookup.keys())
        result:       dict[str, str] = {}
        used_fields:  set[str]       = set()

        # Sort headers so that "used" variants come before "required" variants
        # This ensures QUANTITYUSED wins over QUANTITY REQUIRED
        sorted_headers = sorted(
            raw_headers,
            key=lambda h: (0 if "used" in str(h).lower() else 1))

        for h in sorted_headers:
            h_norm = str(h).strip().lower()
            if not h_norm:
                continue

            # Direct lookup
            if h_norm in self._lookup:
                field = self._lookup[h_norm]
                if field not in used_fields:
                    result[h] = field
                    used_fields.add(field)
                continue

            # Fuzzy lookup
            best, score = self._matcher.best_match(h_norm, all_synonyms)
            if best:
                field = self._lookup[best]
                if field not in used_fields:
                    result[h] = field
                    used_fields.add(field)

        return result


# ─── ConsumptionProcessor ─────────────────────────────────────────────────────

class ConsumptionProcessor:
    """
    Parses a consumption report file and returns:
      [{"item_name": str, "sku_code": str, "quantity": float}, …]
    """

    def __init__(self):
        self._mapper = ConsHeaderMapper()

    # ── Public ─────────────────────────────────────────────────────────────

    def parse(self, file_path: str) -> list[dict]:
        path   = Path(file_path)
        suffix = path.suffix.lower()

        if suffix in {".xlsx", ".xls"}:
            items = self._parse_excel(str(path), suffix)
        elif suffix == ".csv":
            items = self._parse_csv(str(path))
        elif suffix == ".pdf":
            items = self._parse_pdf(str(path))
        else:
            raise ValueError(f"Unsupported consumption report format: {suffix}")

        cleaned = self._clean(items)
        log.info("Consumption parse: %d valid items from %s", len(cleaned), path.name)
        return cleaned

    # ── Excel ──────────────────────────────────────────────────────────────

    def _parse_excel(self, path: str, suffix: str) -> list[dict]:
        engine = "xlrd" if suffix == ".xls" else "openpyxl"

        # Try skipping 0–4 header rows to find the real data
        for skip in range(5):
            try:
                df = pd.read_excel(path, engine=engine, skiprows=skip, dtype=str)
                df.columns = [str(c).strip() for c in df.columns]
                items = self._df_to_items(df)
                if items:
                    log.info("Excel parsed (skip=%d): %d items", skip, len(items))
                    return items
            except Exception as e:
                log.debug("Excel skip=%d failed: %s", skip, e)

        return []

    # ── CSV ────────────────────────────────────────────────────────────────

    def _parse_csv(self, path: str) -> list[dict]:
        for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
            try:
                df = pd.read_csv(path, dtype=str, encoding=enc)
                df.columns = [str(c).strip() for c in df.columns]
                items = self._df_to_items(df)
                if items:
                    return items
            except Exception:
                continue
        return []

    # ── PDF ────────────────────────────────────────────────────────────────

    def _parse_pdf(self, path: str) -> list[dict]:
        items: list[dict] = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables():
                    items.extend(self._table_to_items(table))
        if not items:
            items = self._pdf_text_fallback(path)
        return items

    def _pdf_text_fallback(self, path: str) -> list[dict]:
        with pdfplumber.open(path) as pdf:
            text = "\n".join(p.extract_text() or "" for p in pdf.pages)

        lines = text.splitlines()
        header_idx, col_map = self._detect_text_header(lines)
        if header_idx is None:
            return []

        items: list[dict] = []
        for line in lines[header_idx + 1:]:
            line = line.strip()
            if not line:
                continue
            item = self._parse_text_line(line, col_map)
            if item:
                items.append(item)
        return items

    def _detect_text_header(self, lines):
        matcher  = FuzzyMatcher(threshold=62)
        all_syns = [s for ss in CONS_HEADER_SYNONYMS.values() for s in ss]
        for idx, line in enumerate(lines):
            words = re.findall(r"\S+", line)
            hits  = sum(1 for w in words
                        if matcher.best_match(w.lower(), all_syns)[1] >= 62)
            if hits >= 2:
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
        sorted_fields = sorted(col_map.items(), key=lambda x: x[1])
        item: dict[str, Any] = {"item_name": "", "sku_code": "", "quantity": 0.0}
        for i, (field, start) in enumerate(sorted_fields):
            end  = sorted_fields[i + 1][1] if i + 1 < len(sorted_fields) else len(line)
            cell = line[start:end].strip() if start < len(line) else ""
            if field == "quantity":
                item["quantity"] = _parse_number(cell)
            else:
                item[field] = cell
        return item if item["item_name"] else None

    # ── DataFrame → items ──────────────────────────────────────────────────

    def _df_to_items(self, df: pd.DataFrame) -> list[dict]:
        if df.empty:
            return []

        col_map = self._mapper.map_columns(list(df.columns))
        log.debug("Column mapping: %s", col_map)

        if "quantity" not in col_map.values():
            log.debug("No quantity column found in headers: %s", list(df.columns))
            return []

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
        if not table or len(table) < 2:
            return []

        raw_headers = [_clean_str(c) for c in table[0]]
        col_map     = self._mapper.map_columns(raw_headers)

        if "quantity" not in col_map.values():
            return []

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
                    if field == "quantity":
                        item["quantity"] = _parse_number(cells[idx])
                    else:
                        item[field] = cells[idx]
            if item.get("item_name") or item.get("sku_code"):
                items.append(item)
        return items

    # ── Cleaning ───────────────────────────────────────────────────────────

    def _clean(self, items: list[dict]) -> list[dict]:
        cleaned: list[dict] = []
        for item in items:
            name = str(item.get("item_name", "")).strip()
            sku  = str(item.get("sku_code",  "")).strip()
            qty  = _parse_number(item.get("quantity", 0))

            if not name and not sku:
                continue
            # Skip header-like rows
            if name.lower() in {"component_name", "component name", "item name",
                                 "description", "product", "particulars"}:
                continue
            # Strip serial numbers
            name = re.sub(r"^\d+[\.\)]\s*", "", name).strip()
            qty  = max(0.0, qty)

            cleaned.append({"item_name": name, "sku_code": sku, "quantity": qty})

        return cleaned
