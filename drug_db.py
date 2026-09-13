"""Drug database imported from CSV/Excel and queried through SQLite.

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

The source file remains portable and editable.  A disposable, indexed SQLite
cache beside it provides fast runtime searches without retaining every row in
memory.  The cache is rebuilt automatically whenever the source changes.
"""
from __future__ import annotations

import csv
import sqlite3
import unicodedata
import xlrd
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Sequence, Union

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
    "class_mappings": ["class_mappings", "class mappings", "multiple_classes",
                       "multiple classes"],
    "mapping_status": ["mapping_status", "mapping status", "classification_status",
                       "classification status", "confidence"],
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
    class_mappings: str = ""
    mapping_status: str = ""
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
            "class_mappings": self.class_mappings,
            "mapping_status": self.mapping_status,
            "notes": self.notes,
        }


DRUG_FIELDS = (
    "generic_name", "brand_name", "strength", "form", "category",
    "therapeutic_group", "detailed_class", "class_mappings",
    "mapping_status", "notes",
)


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


class DrugCollection(Sequence[Drug]):
    """Lazy, list-compatible view over the SQLite drug table."""

    def __init__(self, database: "DrugDatabase") -> None:
        self.database = database

    def __len__(self) -> int:
        return self.database.count()

    def __iter__(self) -> Iterator[Drug]:
        yield from self.database.iter_drugs()

    def __getitem__(self, index: Union[int, slice]):
        if isinstance(index, slice):
            start, stop, step = index.indices(len(self))
            if step != 1:
                return list(self)[index]
            return self.database.fetch_page(start, max(0, stop - start))
        if index < 0:
            index += len(self)
        rows = self.database.fetch_page(index, 1)
        if not rows:
            raise IndexError(index)
        return rows[0]


class DrugDatabase:
    def __init__(self, path: str) -> None:
        self.path = Path(path)
        self.cache_path = self.path.with_suffix(self.path.suffix + ".sqlite3")
        self.drugs: Sequence[Drug] = DrugCollection(self)
        self.columns: List[str] = []
        self.load()

    @contextmanager
    def _connect(self):
        connection = sqlite3.connect(self.cache_path, timeout=20)
        connection.row_factory = sqlite3.Row
        try:
            yield connection
        finally:
            connection.close()

    def _source_signature(self) -> str:
        if not self.path.exists():
            return "missing"
        stat = self.path.stat()
        return f"{self.path.resolve()}|{stat.st_size}|{stat.st_mtime_ns}"

    def _cache_is_current(self) -> bool:
        if not self.cache_path.exists():
            return False
        try:
            with self._connect() as connection:
                value = connection.execute(
                    "SELECT value FROM meta WHERE key='source_signature'"
                ).fetchone()
                return bool(value and value[0] == self._source_signature())
        except sqlite3.Error:
            return False

    def _rebuild_cache(self, drugs: Sequence[Drug]) -> None:
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.cache_path.with_suffix(self.cache_path.suffix + ".tmp")
        if temporary.exists():
            temporary.unlink()
        connection = sqlite3.connect(temporary)
        try:
            connection.executescript("""
                PRAGMA journal_mode=OFF;
                PRAGMA synchronous=OFF;
                CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
                CREATE TABLE drugs (
                    id INTEGER PRIMARY KEY,
                    generic_name TEXT NOT NULL,
                    brand_name TEXT NOT NULL DEFAULT '',
                    strength TEXT NOT NULL DEFAULT '',
                    form TEXT NOT NULL DEFAULT '',
                    category TEXT NOT NULL DEFAULT '',
                    therapeutic_group TEXT NOT NULL DEFAULT '',
                    detailed_class TEXT NOT NULL DEFAULT '',
                    class_mappings TEXT NOT NULL DEFAULT '',
                    mapping_status TEXT NOT NULL DEFAULT '',
                    notes TEXT NOT NULL DEFAULT '',
                    generic_norm TEXT NOT NULL,
                    brand_norm TEXT NOT NULL,
                    prescribable_norm TEXT NOT NULL,
                    search_norm TEXT NOT NULL
                );
                CREATE INDEX idx_drugs_generic ON drugs(generic_norm);
                CREATE INDEX idx_drugs_brand ON drugs(brand_norm);
                CREATE INDEX idx_drugs_prescribable ON drugs(prescribable_norm);
                CREATE INDEX idx_drugs_group ON drugs(therapeutic_group);
                CREATE INDEX idx_drugs_detail ON drugs(detailed_class);
                CREATE INDEX idx_drugs_mapping_status ON drugs(mapping_status);
            """)
            values = []
            for drug in drugs:
                record = drug.to_dict()
                generic = _norm(drug.generic_name)
                brand = _norm(drug.brand_name)
                prescribable = brand or generic
                searchable = " ".join(filter(None, (generic, brand, _norm(drug.strength),
                                                      _norm(drug.form), _norm(drug.category))))
                values.append(tuple(record[field] for field in DRUG_FIELDS)
                              + (generic, brand, prescribable, searchable))
            connection.executemany(
                "INSERT INTO drugs (" + ",".join(DRUG_FIELDS) +
                ",generic_norm,brand_norm,prescribable_norm,search_norm) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", values)
            connection.execute(
                "INSERT INTO meta(key,value) VALUES('source_signature',?)",
                (self._source_signature(),))
            connection.commit()
        finally:
            connection.close()
        # SQLite's backup API can refresh an active cache safely even when a
        # short-lived reader is finishing on Windows (where replacing an open
        # database file would otherwise raise PermissionError).
        source = sqlite3.connect(temporary)
        destination = sqlite3.connect(self.cache_path, timeout=20)
        try:
            source.backup(destination)
        finally:
            destination.close()
            source.close()
        temporary.unlink(missing_ok=True)

    @staticmethod
    def _row_to_drug(row) -> Drug:
        return Drug(**{field: row[field] for field in DRUG_FIELDS})

    # -- loading ------------------------------------------------------------
    def load(self) -> int:
        """Open or rebuild the indexed runtime cache. Returns number loaded."""
        if not self.path.exists():
            self._rebuild_cache([])
            return 0
        self.columns = list(COLUMN_ALIASES.keys())
        if not self._cache_is_current():
            rows = _read_rows(self.path)
            loaded = [self._dict_to_drug(rec) for rec in rows
                      if rec.get("generic_name", "").strip()]
            self._rebuild_cache(loaded)
        self.drugs = DrugCollection(self)
        return self.count()

    def count(self) -> int:
        if not self.cache_path.exists():
            return 0
        with self._connect() as connection:
            return int(connection.execute("SELECT COUNT(*) FROM drugs").fetchone()[0])

    def iter_drugs(self, batch_size: int = 500) -> Iterator[Drug]:
        if not self.cache_path.exists():
            return
        with self._connect() as connection:
            cursor = connection.execute(
                "SELECT " + ",".join(DRUG_FIELDS) + " FROM drugs ORDER BY id")
            while True:
                rows = cursor.fetchmany(batch_size)
                if not rows:
                    break
                for row in rows:
                    yield self._row_to_drug(row)

    def fetch_page(self, offset: int, limit: int) -> List[Drug]:
        if limit <= 0 or offset < 0 or not self.cache_path.exists():
            return []
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT " + ",".join(DRUG_FIELDS) +
                " FROM drugs ORDER BY id LIMIT ? OFFSET ?", (limit, offset)).fetchall()
        return [self._row_to_drug(row) for row in rows]

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
            imported = [self._dict_to_drug(r) for r in rows]
            self._write(imported)
            # Re-read the active local CSV so the autocomplete and class pages
            # immediately use exactly what was saved.
            return self.load()

        # merge mode
        existing = list(self.drugs)
        existing_names = {_norm(d.generic_name) for d in existing}
        added = 0
        for r in rows:
            d = self._dict_to_drug(r)
            if _norm(d.generic_name) not in existing_names:
                existing.append(d)
                existing_names.add(_norm(d.generic_name))
                added += 1
        self._write(existing)
        self.load()
        return added

    def clear(self) -> None:
        """Remove all medicines from the active local database safely."""
        self._write([])
        self.load()

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
            class_mappings=rec.get("class_mappings", "").strip(),
            mapping_status=rec.get("mapping_status", "").strip(),
            notes=rec.get("notes", "").strip(),
        )

    def export_csv(self, dest_path: str) -> None:
        """Export the current database to a CSV file."""
        self._write(self.drugs, Path(dest_path))

    def export_file(self, dest_path: str) -> None:
        """Export the database as CSV, modern XLSX, or legacy XLS."""
        path = Path(dest_path)
        suffix = path.suffix.casefold()
        fields = list(DRUG_FIELDS)
        path.parent.mkdir(parents=True, exist_ok=True)
        if suffix == ".csv":
            self._write(self.drugs, path)
            return
        if suffix == ".xlsx":
            from openpyxl import Workbook
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Drug Database"
            sheet.append(fields)
            for drug in self.drugs:
                record = drug.to_dict()
                sheet.append([record.get(field, "") for field in fields])
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            workbook.save(path)
            return
        if suffix == ".xls":
            import xlwt
            workbook = xlwt.Workbook(encoding="utf-8")
            sheet = workbook.add_sheet("Drug Database")
            for column, field in enumerate(fields):
                sheet.write(0, column, field)
            for row, drug in enumerate(self.drugs, 1):
                record = drug.to_dict()
                for column, field in enumerate(fields):
                    sheet.write(row, column, record.get(field, ""))
            workbook.save(str(path))
            return
        raise ValueError("Choose CSV (.csv), Excel (.xlsx), or legacy Excel (.xls).")

    def _write(self, drugs: List[Drug], path: Optional[Path] = None) -> None:
        path = path or self.path
        path.parent.mkdir(parents=True, exist_ok=True)
        fields = list(DRUG_FIELDS)
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
            return self.fetch_page(0, limit)
        return self._search_column("search_norm", q, limit)

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
        return self._search_column("prescribable_norm", _norm(query), limit)

    def _search_field(self, query: str, field_name: str, limit: int) -> List[Drug]:
        column = "generic_norm" if field_name == "generic_name" else "brand_norm"
        return self._search_column(column, _norm(query), limit)

    def _search_column(self, column: str, query: str, limit: int) -> List[Drug]:
        if not query or limit <= 0:
            return []
        if column not in {"generic_norm", "brand_norm", "prescribable_norm", "search_norm"}:
            raise ValueError("Unsupported search column")
        upper = query + "\uffff"
        fields = ",".join(DRUG_FIELDS)
        with self._connect() as connection:
            rows = connection.execute(
                f"SELECT {fields} FROM drugs WHERE {column}>=? AND {column}<? "
                f"ORDER BY {column}, id LIMIT ?", (query, upper, limit)).fetchall()
            remaining = limit - len(rows)
            if remaining:
                rows += connection.execute(
                    f"SELECT {fields} FROM drugs WHERE instr({column},?)>0 "
                    f"AND NOT ({column}>=? AND {column}<?) "
                    f"ORDER BY {column}, id LIMIT ?",
                    (query, query, upper, remaining)).fetchall()
        return [self._row_to_drug(row) for row in rows]

    def trade_names_for_scientific(self, scientific_name: str) -> List[str]:
        """Return the locally recorded trade names for one scientific name."""
        target = _norm(scientific_name)
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT DISTINCT brand_name FROM drugs WHERE generic_norm=? AND brand_name<>'' "
                "ORDER BY brand_norm", (target,)).fetchall()
        names = {row[0].strip() for row in rows if row[0].strip()}
        return sorted(names, key=str.casefold)

    def all_names(self) -> List[str]:
        with self._connect() as connection:
            return [row[0] for row in connection.execute(
                "SELECT generic_name FROM drugs ORDER BY id")]

    def find_exact(self, name: str) -> Optional[Drug]:
        """Return a database medicine when its generic name is an exact match."""
        target = _norm(name)
        with self._connect() as connection:
            row = connection.execute(
                "SELECT " + ",".join(DRUG_FIELDS) +
                " FROM drugs WHERE generic_norm=? ORDER BY id LIMIT 1", (target,)).fetchone()
        return self._row_to_drug(row) if row else None

    def contains_name(self, generic_name: str = "", brand_name: str = "") -> bool:
        generic, brand = _norm(generic_name), _norm(brand_name)
        if not generic and not brand:
            return False
        clauses, values = [], []
        if generic:
            clauses.append("generic_norm=?")
            values.append(generic)
        if brand:
            clauses.append("brand_norm=?")
            values.append(brand)
        with self._connect() as connection:
            return connection.execute(
                "SELECT 1 FROM drugs WHERE " + " OR ".join(clauses) + " LIMIT 1",
                values).fetchone() is not None

    def update_classification(self, name: str, therapeutic_group: str,
                              detailed_class: str, append: bool = True) -> bool:
        """Persist a clinician-reviewed class mapping for one local medicine."""
        drugs = list(self.drugs)
        drug = next((item for item in drugs if _norm(item.generic_name) == _norm(name)), None)
        if drug is None:
            return False
        self._set_classification(drug, therapeutic_group, detailed_class, append)
        self._write(drugs)
        self.load()
        return True

    def update_classifications(self, names, therapeutic_group: str,
                               detailed_class: str, append: bool = True) -> int:
        """Persist one reviewed classification for several selected medicines."""
        wanted = {_norm(str(name)) for name in names if str(name).strip()}
        if not wanted:
            return 0
        updated = 0
        drugs = list(self.drugs)
        for drug in drugs:
            if _norm(drug.generic_name) not in wanted:
                continue
            self._set_classification(drug, therapeutic_group, detailed_class, append)
            updated += 1
        if updated:
            self._write(drugs)
            self.load()
        return updated

    @staticmethod
    def _stored_class_pairs(drug: Drug) -> List[tuple[str, str]]:
        pairs: List[tuple[str, str]] = []
        for item in str(drug.class_mappings or "").split("|"):
            group, separator, detail = item.partition("::")
            pair = (group.strip(), detail.strip())
            if separator and all(pair) and pair not in pairs:
                pairs.append(pair)
        legacy = (drug.therapeutic_group.strip(), drug.detailed_class.strip())
        if all(legacy) and legacy not in pairs:
            pairs.insert(0, legacy)
        return pairs

    @classmethod
    def _set_classification(cls, drug: Drug, therapeutic_group: str,
                            detailed_class: str, append: bool) -> None:
        pair = (therapeutic_group.strip(), detailed_class.strip())
        pairs = cls._stored_class_pairs(drug) if append else []
        if all(pair) and pair not in pairs:
            pairs.append(pair)
        if pairs:
            drug.therapeutic_group, drug.detailed_class = pairs[0]
        else:
            drug.therapeutic_group, drug.detailed_class = pair
        drug.class_mappings = "|".join(f"{group}::{detail}" for group, detail in pairs)
        drug.mapping_status = "confirmed"
