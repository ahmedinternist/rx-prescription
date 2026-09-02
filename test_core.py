"""Safe automated checks; each test runs with an isolated AppData directory."""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


def run_isolated(tmp_path: Path, code: str) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    env["RX_APP_DATA_DIR"] = str(tmp_path)
    return subprocess.run([sys.executable, "-c", code], cwd=Path(__file__).parent,
                          env=env, text=True, capture_output=True, check=True)


def test_signed_qr_round_trip_and_validation(tmp_path):
    code = '''
import qr_utils as q
from config import CONFIG_PATH, Config, config
rx = q.Prescription(
    clinic=q.Clinic(name="Clinic"),
    doctor=q.Doctor(name="Dr A", license_no="L-1", specialty="GP"),
    patient=q.Patient(name="Patient", age="30", sex="F"),
    drugs=[q.DrugItem(generic_name="Amoxicillin", dosage="500 mg", frequency="3/day", duration="7 days")],
    date="2026-09-02", rx_id="RX-1")
errors, warnings = q.validate_prescription(rx)
assert not errors and not warnings
url = q.build_qr_url(rx)
payload = q.decode_from_url(url)
assert payload["sig"]["alg"] == "ES256"
assert payload["doctor"] == {"name": "Dr A", "specialty": "GP"}
assert payload["patient"] == {"name": "Patient"}
assert "notes" not in payload["drugs"][0]
assert q.qr_info(url)["qr_version"] > 0
assert q.verification_key()["kty"] == "EC"
assert "BEGIN PRIVATE KEY" not in CONFIG_PATH.read_text(encoding="utf-8")
config.set_clinic(name="Clinic", address="Baghdad", phone="123", logo_path="")
reloaded = Config().get_clinic()
assert reloaded["name"] == "Clinic" and reloaded["address"] == "Baghdad"
'''
    run_isolated(tmp_path, code)


def test_validation_rejects_missing_clinical_fields(tmp_path):
    code = '''
import qr_utils as q
rx = q.Prescription(drugs=[q.DrugItem(generic_name="A")])
errors, warnings = q.validate_prescription(rx)
assert any("license" in error.lower() for error in errors)
assert not any("patient" in error.lower() or "dosage" in error.lower() for error in errors)
'''
    run_isolated(tmp_path, code)


def test_csv_and_xlsx_import_only(tmp_path):
    code = '''
import os
from pathlib import Path
import drug_db
path = Path(os.environ["RX_APP_DATA_DIR"]) / "db.csv"
path.write_text("generic_name,brand_name\\nDrug A,Brand A\\n", encoding="utf-8")
db = drug_db.DrugDatabase(str(path))
assert db.search("brand")[0].generic_name == "Drug A"
try:
    drug_db._read_rows(Path(os.environ["RX_APP_DATA_DIR"]) / "legacy.xls")
except ValueError:
    pass
else:
    raise AssertionError("legacy .xls must be rejected clearly")
'''
    run_isolated(tmp_path, code)


def test_arabic_documents_and_markup_characters_render(tmp_path):
    code = '''
import os
from pathlib import Path
import i18n as I
import pdf_generator as pdf
import qr_utils as q
I.set_lang("ar")
rx = q.Prescription(
    clinic=q.Clinic(name="عيادة <السلام>"),
    doctor=q.Doctor(name="د. أحمد & علي", license_no="L-2"),
    patient=q.Patient(name="مريم", age="20", sex="F"),
    drugs=[q.DrugItem(generic_name="Drug <A>", dosage="500 mg", frequency="مرة/يوم", duration="7 أيام")],
    date="2026-09-02", rx_id="RX-2")
root = Path(os.environ["RX_APP_DATA_DIR"])
qr = q.make_qr_image(q.build_qr_url(rx))
pdf.generate_prescription_pdf(rx, root / "rx.pdf", qr_pil_image=qr)
pdf.generate_prescription_docx(rx, root / "rx.docx", qr_pil_image=qr)
assert (root / "rx.pdf").stat().st_size > 2000
assert (root / "rx.docx").stat().st_size > 2000
'''
    run_isolated(tmp_path, code)
