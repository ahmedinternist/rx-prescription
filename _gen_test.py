import qr_utils as qu
import pdf_generator as pdfgen  # ensures qrcode import available
from pathlib import Path

RX_BASE = "https://ahmedinternist.github.io/rx-viewer/"

# Sample prescription
rx = qu.Prescription(
    doctor=qu.Doctor(name="Dr. Ahmed Al-Rashid", license_no="SCH-44521", specialty="General Practice"),
    patient=qu.Patient(name="Jane Doe", age="34", sex="F"),
    drugs=[
        qu.DrugItem(generic_name="Amoxicillin", brand_name="Amoxil", dosage="500 mg",
                    frequency="3 times/day", duration="7 days", notes="after meals"),
        qu.DrugItem(generic_name="Omeprazole", brand_name="Losec", dosage="20 mg",
                    frequency="once daily", duration="14 days", notes="before breakfast"),
        qu.DrugItem(generic_name="Ibuprofen", dosage="400 mg",
                    frequency="as needed", duration="5 days", notes="for pain"),
    ],
    date="2026-08-30", rx_id="RX-2026-001",
)

# Build the real QR URL (GitHub Pages base + compressed fragment)
url = qu.build_qr_url(rx, viewer_base=RX_BASE)
frag = url.split("#", 1)[1]

# 1) Sample QR PNG (works once the site is deployed)
qr = qu.make_qr_image(url)
qr.save("test_qr.png")
print("test_qr.png written ->", Path("test_qr.png").resolve())

# 2) viewer_test.html with the sample fragment embedded (works NOW, no internet)
html = Path("viewer.html").read_text(encoding="utf-8")
injection = (
    'const SAMPLE_FRAG = ' + repr(frag) + ';\n'
    'if (!location.hash || location.hash === "#") { location.hash = SAMPLE_FRAG; }\n'
    'decode();'
)
html2 = html.replace("decode();\nwindow.addEventListener('hashchange', decode);",
                     injection + "\nwindow.addEventListener('hashchange', decode);", 1)
Path("viewer_test.html").write_text(html2, encoding="utf-8")
print("viewer_test.html written ->", Path("viewer_test.html").resolve())
print("fragment length:", len(frag))
