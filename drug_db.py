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
    category       - existing local category/favorite tag (optional)
    therapeutic_group - major therapeutic group used by the dashboard (optional)
    detailed_class - detailed drug class within that group (optional)
    notes          - free text (optional)

The database is kept in memory and supports fast prefix/substring search for
the autocomplete widget.
"""
from __future__ import annotations

import csv
import unicodedata
import xlrd
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

REQUIRED_COLUMNS = {"generic_name"}
COLUMN_ALIASES = {
    "generic_name": ["generic_name", "generic", "generic name", "scientific_name",
                     "scientific name", "scientific", "inn", "name", "drug_name", "drug name",
                     "المادة", "اسم المادة", "اسم الدواء", "الاسم العلمي"],
    "brand_name": ["brand_name", "brand", "trade_name", "trade name", "trade", "الاسم التجاري"],
    "strength": ["strength", "dose", "dosage_strength"],
    "form": ["form", "dosage_form", "drug_form", "route"],
    "category": ["category", "class", "group", "atc"],
    "therapeutic_group": ["therapeutic_group", "therapeutic group", "major_class", "major class",
                            "major classification"],
    "detailed_class": ["detailed_class", "detailed class", "drug_class", "drug class", "sub_class",
                       "sub classification", "sub-classification"],
    "notes": ["notes", "comment", "remarks"],
}


@dataclass
class Drug:
    generic_name: str
    brand_name: str = ""
    strength: str = ""
    form: str = ""
    category: str = ""
    therapeutic_group: str = ""
    detailed_class: str = ""
    notes: str = ""

    def to_dict(self) -> Dict[str, str]:
        return {
            "generic_name": self.generic_name,
            "brand_name": self.brand_name,
            "strength": self.strength,
            "form": self.form,
            "category": self.category,
            "therapeutic_group": self.therapeutic_group,
            "detailed_class": self.detailed_class,
            "notes": self.notes,
        }


def _norm(text: str) -> str:
    """Lower-case + strip accents for tolerant matching."""
    text = text.lower().strip()
    return "".join(
        c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c)
    )


def _header_key(value: object) -> str:
    """Normalise spreadsheet headers without altering the displayed values."""
    return " ".join(str(value).replace("\ufeff", "").replace("\xa0", " ")
                    .replace("_", " ").casefold().split())


def _resolve_columns(header: List[str]) -> Dict[str, str]:
    """Map our canonical field names to the actual column names."""
    lower = {_header_key(h): h for h in header if str(h).strip()}
    mapping: Dict[str, str] = {}
    for field_name, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if _header_key(alias) in lower:
                mapping[field_name] = lower[_header_key(alias)]
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
        return _read_xls(path)
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
    grids: List[List[List[str]]] = []
    try:
        # Workbooks received from pharmacies often contain a cover sheet or a
        # report title above the column headers.  Use the first worksheet that
        # contains recognisable medicine-column headers, not just wb.active.
        for ws in wb.worksheets:
            grid = [["" if c is None else str(c) for c in row]
                    for row in ws.iter_rows(values_only=True)]
            grids.append(grid)
    finally:
        wb.close()
    return _read_grids(grids)


def _read_xls(path: Path) -> List[Dict[str, str]]:
    """Read legacy Excel workbooks using xlrd (which supports .xls BIFF files)."""
    workbook = xlrd.open_workbook(path)
    grids = []
    for sheet in workbook.sheets():
        grids.append([["" if value is None else str(value) for value in sheet.row_values(index)]
                      for index in range(sheet.nrows)])
    return _read_grids(grids)


def _read_grids(grids: List[List[List[str]]]) -> List[Dict[str, str]]:
    """Find a medicine header row across spreadsheet sheets and title rows."""
    fallback_grid: List[List[str]] = []
    for grid in grids:
        grid = [row for row in grid if any(str(cell).strip() for cell in row)]
        if not grid:
            continue
        if not fallback_grid:
            fallback_grid = grid
        for header_index, header in enumerate(grid):
            if "generic_name" in _resolve_columns(header):
                return _rows_to_dicts(header, grid[header_index + 1:])
    # Preserve support for a simple one-column sheet whose first column is the
    # medicine name, even when it does not use a recognised header label.
    if not fallback_grid:
        return []
    return _rows_to_dicts(fallback_grid[0], fallback_grid[1:])


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
            raise ValueError("No medicine rows were found in the selected file.")

        if replace:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.drugs = [self._dict_to_drug(r) for r in rows]
            self._write(self.drugs)
            # Re-read the active local CSV so the autocomplete and class pages
            # immediately use exactly what was saved.
            return self.load()

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

    def clear(self) -> None:
        """Remove all medicines from the active local database safely."""
        self.drugs = []
        self._write(self.drugs)

    @staticmethod
    def _dict_to_drug(rec: Dict[str, str]) -> "Drug":
        return Drug(
            generic_name=rec.get("generic_name", "").strip(),
            brand_name=rec.get("brand_name", "").strip(),
            strength=rec.get("strength", "").strip(),
            form=rec.get("form", "").strip(),
            category=rec.get("category", "").strip(),
            therapeutic_group=rec.get("therapeutic_group", "").strip(),
            detailed_class=rec.get("detailed_class", "").strip(),
            notes=rec.get("notes", "").strip(),
        )

    def export_csv(self, dest_path: str) -> None:
        """Export the current database to a CSV file."""
        self._write(self.drugs, Path(dest_path))

    def _write(self, drugs: List[Drug], path: Optional[Path] = None) -> None:
        path = path or self.path
        path.parent.mkdir(parents=True, exist_ok=True)
        fields = ["generic_name", "brand_name", "strength", "form", "category",
                  "therapeutic_group", "detailed_class", "notes"]
        temporary = path.with_suffix(path.suffix + ".tmp")
        with open(temporary, "w", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=fields)
            writer.writeheader()
            for d in drugs:
                writer.writerow(d.to_dict())
        temporary.replace(path)

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

    def search_scientific(self, query: str, limit: int = 25) -> List[Drug]:
        """Find medicines by the scientific / INN name only."""
        return self._search_field(query, "generic_name", limit)

    def search_trade(self, query: str, limit: int = 25) -> List[Drug]:
        """Find medicines by the generic / trade product name only."""
        return self._search_field(query, "brand_name", limit)

    def search_prescribable(self, query: str, limit: int = 25) -> List[Drug]:
        """Search the name users actually select in the prescription form.

        Modern two-column databases use ``brand_name`` for that field. Older
        and current one-column imports store the product name in
        ``generic_name`` and leave ``brand_name`` empty, so those rows are
        intentionally searched by their only available name.
        """
        q = _norm(query)
        if not q:
            return []
        prefix: List[Drug] = []
        contains: List[Drug] = []
        for drug in self.drugs:
            value = drug.brand_name.strip() or drug.generic_name.strip()
            normalized = _norm(value)
            if normalized.startswith(q):
                prefix.append(drug)
            elif q in normalized:
                contains.append(drug)
        return (prefix + contains)[:limit]

    def _search_field(self, query: str, field_name: str, limit: int) -> List[Drug]:
        q = _norm(query)
        if not q:
            return []
        prefix: List[Drug] = []
        contains: List[Drug] = []
        for drug in self.drugs:
            value = _norm(getattr(drug, field_name, ""))
            if not value:
                continue
            if value.startswith(q):
                prefix.append(drug)
            elif q in value:
                contains.append(drug)
        return (prefix + contains)[:limit]

    def trade_names_for_scientific(self, scientific_name: str) -> List[str]:
        """Return the locally recorded trade names for one scientific name."""
        target = _norm(scientific_name)
        names = {drug.brand_name.strip() for drug in self.drugs
                 if _norm(drug.generic_name) == target and drug.brand_name.strip()}
        return sorted(names, key=str.casefold)

    def all_names(self) -> List[str]:
        return [d.generic_name for d in self.drugs]

    def find_exact(self, name: str) -> Optional[Drug]:
        """Return a database medicine when its generic name is an exact match."""
        target = _norm(name)
        return next((drug for drug in self.drugs if _norm(drug.generic_name) == target), None)

    def update_classification(self, name: str, therapeutic_group: str,
                              detailed_class: str) -> bool:
        """Persist a clinician-reviewed class mapping for one local medicine."""
        drug = self.find_exact(name)
        if drug is None:
            return False
        drug.therapeutic_group = therapeutic_group.strip()
        drug.detailed_class = detailed_class.strip()
        self._write(self.drugs)
        return True
