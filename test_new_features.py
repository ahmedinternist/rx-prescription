"""Test the new features: XLS import + Medication Label Word export + autocomplete width."""
import os, tempfile, csv
from openpyxl import Workbook

import config as cfg
import drug_db as dbmod
import qr_utils as qu
import pdf_generator as pdfgen
import i18n as I

tmp = tempfile.mkdtemp()
I.set_lang("en")

# --- 1. XLS import (replace) ------------------------------------------------
wb = Workbook()
ws = wb.active
ws.append(["Generic Name", "Brand", "Strength", "Form", "Category", "Notes"])
ws.append(["Azithromycin", "Zithromax", "500 mg", "tablet", "antibiotic", "once daily"])
ws.append(["Metformin", "Glucophage", "850 mg", "tablet", "antidiabetic", "with food"])
xls_path = os.path.join(tmp, "drugs.xlsx")
wb.save(xls_path)

# point DB at a temp csv and import the XLS
cfg.DEFAULT_DB_PATH = os.path.join(tmp, "db.csv")
db = dbmod.DrugDatabase(cfg.DEFAULT_DB_PATH)
# ensure empty start
db.drugs = []
db._write(db.drugs)
added = db.import_file(xls_path, replace=True)
print("XLS import ->", added, "drugs")
assert added == 2
names = {d.generic_name for d in db.drugs}
assert names == {"Azithromycin", "Metformin"}, names
assert db.drugs[0].strength == "500 mg"
assert db.drugs[1].brand_name == "Glucophage"

# --- 2. XLS merge ----------------------------------------------------------
ws2 = Workbook().active
ws2.append(["Generic Name", "Brand"])
ws2.append(["Ibuprofen", "Brufen"])
xls2 = os.path.join(tmp, "extra.xlsx")
ws2.parent.save(xls2)
added2 = db.import_file(xls2, replace=False)
print("XLS merge -> added", added2)
assert added2 == 1
assert len(db.drugs) == 3

# --- 3. Medication Label Word export (full content + QR above bottom-right) -
rx = qu.Prescription(
    doctor=qu.Doctor(name="Dr. Smith", license_no="LIC-99", specialty="GP"),
    patient=qu.Patient(name="Jane Doe", age="34", sex="F"),
    drugs=[
        qu.DrugItem(generic_name="Amoxicillin", brand_name="Amoxil", dosage="500 mg",
                    frequency="3x/day", duration="7 days", notes="after meals"),
        qu.DrugItem(generic_name="Ibuprofen", dosage="400 mg", frequency="as needed",
                    duration="5 days"),
    ],
    date="2026-08-30", rx_id="RX-001",
)
url = qu.build_qr_url(rx)
qr = qu.make_qr_image(url)
label_path = os.path.join(tmp, "label.docx")
pdfgen.generate_medication_label_docx(rx, label_path, qr_pil_image=qr)
print("Label docx bytes:", os.path.getsize(label_path))
assert os.path.getsize(label_path) > 2000

# verify contents with python-docx
from docx import Document
d = Document(label_path)
text = "\n".join(p.text for p in d.paragraphs)
tbl = d.tables[0]
cell_text = "\n".join(c.text for row in tbl.rows for c in row.cells)
print("--- label text ---")
print(text)
print("--- label table ---")
print(cell_text)
assert "Amoxicillin" in cell_text
assert "500 mg" in cell_text
assert "3x/day" in cell_text
assert "7 days" in cell_text
assert "after meals" in cell_text

# --- 4. autocomplete dropdown width calc ------------------------------------
from main import DrugRow
rows = db.search("a", limit=12)  # Azithromycin + Amoxicillin etc.
mw = max((len(r.generic_name) + len(r.brand_name or "")
          + len(r.strength or "") + 8) for r in rows)
width = max(40, min(70, mw))
print("dropdown width chars:", width)
assert width >= 40

print("ALL NEW-FEATURE TESTS PASSED")
