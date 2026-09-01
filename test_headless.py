"""Headless smoke test (post bilingual / compact-label changes)."""
import os
import tempfile

import config as cfg
import drug_db as dbmod
import qr_utils as qu
import pdf_generator as pdfgen
import i18n as I

tmp = tempfile.mkdtemp()

# Drug DB
if cfg.DEFAULT_DB_PATH.exists():
    cfg.DEFAULT_DB_PATH.unlink()
cfg.ensure_seed_db()
db = dbmod.DrugDatabase(cfg.DEFAULT_DB_PATH)
assert len(db.drugs) == 27

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

# Full PDF in both languages + A4/Letter/A5
for lang in ("en", "ar"):
    I.set_lang(lang)
    for paper in ("A4", "Letter", "A5"):
        out = os.path.join(tmp, f"full_{lang}_{paper}.pdf")
        pdfgen.generate_prescription_pdf(rx, out, paper_size=paper, qr_pil_image=qr)
        assert os.path.getsize(out) > 2000

# Compact label (Word) in both langs + A5
for lang in ("en", "ar"):
    I.set_lang(lang)
    for paper in ("A4", "A5"):
        out = os.path.join(tmp, f"label_{lang}_{paper}.docx")
        pdfgen.generate_medication_label_docx(rx, out, qr_pil_image=qr)
        assert os.path.getsize(out) > 1500

# DOCX + decode round-trip
pdfgen.generate_prescription_docx(rx, os.path.join(tmp, "rx.docx"), qr_pil_image=qr)
decoded = qu.decode_from_url(url)
assert decoded["doctor"]["name"] == "Dr. Smith"
assert len(decoded["drugs"]) == 2
assert "clinic" not in decoded
# Label QR must decode to the SAME full payload
decoded2 = qu.decode_from_url(url)
assert decoded2["doctor"]["name"] == "Dr. Smith"

print("ALL TESTS PASSED")
