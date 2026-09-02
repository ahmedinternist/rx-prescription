"""Drug database backed by a CSV or Excel (.xlsx) file.

The doctor can import/replace the database from any CSV or Excel source (e.g.
exported from another system, a national formulary). The only REQUIRED column
is the drug's name; everything else is optional and used to enrich the
autocomplete + prescription.

Expected columns (case-insensitive, order-independent):
    generic_name   - scientific / generic name (REQUIRED)
    brand_name     - trade / brand name (optional)
    strength       - e.g. "500 mg" (optional)
    form           - e.g. "tablet", "syrup", "capsule" (optional)
    category       - e.g. "antibiotic" (optional)
    notes          - free text (optional)

The database is kept in memory and supports fast prefix/substring search for
the autocomplete widget.
"""
from __future__ import annotations

import csv
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

REQUIRED_COLUMNS = {"generic_name"}
COLUMN_ALIASES = {
    "generic_name": ["generic_name", "generic", "scientific_name", "inn", "name"],
    "brand_name": ["brand_name", "brand", "trade_name", "trade"],
    "strength": ["strength", "dose", "dosage_strength"],
    "form": ["form", "dosage_form", "drug_form", "route"],
    "category": ["category", "class", "group", "atc"],
    "notes": ["notes", "comment", "remarks"],
}


@dataclass
class Drug:
    generic_name: str
    brand_name: str = ""
    strength: str = ""
    form: str = ""
    category: str = ""
    notes: str = ""

    def to_dict(self) -> Dict[str, str]:
        return {
            "generic_name": self.generic_name,
            "brand_name": self.brand_name,
            "strength": self.strength,
            "form": self.form,
            "category": self.category,
            "notes": self.notes,
        }


def _norm(text: str) -> str:
    """Lower-case + strip accents for tolerant matching."""
    text = text.lower().strip()
    return "".join(
        c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c)
    )


def _resolve_columns(header: List[str]) -> Dict[str, str]:
    """Map our canonical field names to the actual column names."""
    lower = {h.strip().lower(): h.strip() for h in header if h.strip()}
    mapping: Dict[str, str] = {}
    for field_name, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in lower:
                mapping[field_name] = lower[alias]
                break
    return mapping


# ---------------------------------------------------------------------------
# File reading helpers (module level) — support CSV and Excel.
# ---------------------------------------------------------------------------
def _read_rows(path: Path) -> List[Dict[str, str]]:
    """Read any supported file (CSV / XLSX) into canonical-field dicts."""
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return _read_xlsx(path)
    if suffix == ".xls":
        raise ValueError("Legacy .xls is not supported. Save the file as .xlsx or CSV first.")
    return _read_csv(path)


def _read_csv(path: Path) -> List[Dict[str, str]]:
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        rows = list(reader)
    if not rows:
        return []
    header = rows[0]
    return _rows_to_dicts(header, rows[1:])


def _read_xlsx(path: Path) -> List[Dict[str, str]]:
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    grid = [["" if c is None else str(c) for c in row]
            for row in ws.iter_rows(values_only=True)]
    wb.close()
    if not grid:
        return []
    grid = [r for r in grid if any(c.strip() for c in r)]
    if not grid:
        return []
    header = grid[0]
    return _rows_to_dicts(header, grid[1:])


def _rows_to_dicts(header: List[str], rows: List[List[str]]) -> List[Dict[str, str]]:
    mapping = _resolve_columns(header)
    if "generic_name" not in mapping:
        mapping["generic_name"] = header[0]
    out = []
    for row in rows:
        if not row or not any(c.strip() for c in row):
            continue
        rec = {k: (row[header.index(v)].strip() if v in header and header.index(v) < len(row) else "")
               for k, v in mapping.items()}
        if rec.get("generic_name", "").strip():
            out.append(rec)
    return out


class DrugDatabase:
    def __init__(self, path: str) -> None:
        self.path = Path(path)
        self.drugs: List[Drug] = []
        self.columns: List[str] = []
        self.load()

    # -- loading ------------------------------------------------------------
    def load(self) -> int:
        """Load drugs from the configured CSV. Returns number loaded."""
        self.drugs = []
        if not self.path.exists():
            return 0
        rows = _read_rows(self.path)
        if not rows:
            return 0
        # Re-derive header from the first source row is not needed; we already
        # normalised to canonical fields.
        self.columns = list(COLUMN_ALIASES.keys())
        for rec in rows:
            name = rec.get("generic_name", "").strip()
            if not name:
                continue
            self.drugs.append(self._dict_to_drug(rec))
        return len(self.drugs)

    # -- import / replace ---------------------------------------------------
    def import_file(self, source_path: str, replace: bool = True) -> int:
        """Import a CSV or Excel (XLS/XLSX) file into the active database.

        replace=True overwrites the current DB; replace=False merges
        (skipping exact duplicate generic names).
        """
        src = Path(source_path)
        if not src.exists():
            raise FileNotFoundError(f"Source file not found: {source_path}")

        rows = _read_rows(src)  # list of dict[str,str] keyed by canonical field
        if not rows:
            return 0

        if replace:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.drugs = [self._dict_to_drug(r) for r in rows]
            self._write(self.drugs)
            return len(self.drugs)

        # merge mode
        self.load()
        existing_names = {_norm(d.generic_name) for d in self.drugs}
        added = 0
        for r in rows:
            d = self._dict_to_drug(r)
            if _norm(d.generic_name) not in existing_names:
                self.drugs.append(d)
                existing_names.add(_norm(d.generic_name))
                added += 1
        self._write(self.drugs)
        return added

    @staticmethod
    def _dict_to_drug(rec: Dict[str, str]) -> "Drug":
        return Drug(
            generic_name=rec.get("generic_name", "").strip(),
            brand_name=rec.get("brand_name", "").strip(),
            strength=rec.get("strength", "").strip(),
            form=rec.get("form", "").strip(),
            category=rec.get("category", "").strip(),
            notes=rec.get("notes", "").strip(),
        )

    def export_csv(self, dest_path: str) -> None:
        """Export the current database to a CSV file."""
        self._write(self.drugs, Path(dest_path))

    def _write(self, drugs: List[Drug], path: Optional[Path] = None) -> None:
        path = path or self.path
        path.parent.mkdir(parents=True, exist_ok=True)
        fields = ["generic_name", "brand_name", "strength", "form", "category", "notes"]
        with open(path, "w", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=fields)
            writer.writeheader()
            for d in drugs:
                writer.writerow(d.to_dict())

    # -- search / autocomplete ---------------------------------------------
    def search(self, query: str, limit: int = 25) -> List[Drug]:
        """Return drugs matching a query (prefix first, then substring)."""
        q = _norm(query)
        if not q:
            return self.drugs[:limit]
        prefix: List[Drug] = []
        contains: List[Drug] = []
        for d in self.drugs:
            hay = _norm(d.generic_name)
            if d.brand_name:
                hay += " " + _norm(d.brand_name)
            if hay.startswith(q):
                prefix.append(d)
            elif q in hay:
                contains.append(d)
        return (prefix + contains)[:limit]

    def all_names(self) -> List[str]:
        return [d.generic_name for d in self.drugs]
