import qr_utils as qu
import json

# Build a real QR URL exactly as the app's Settings would, pointing at a
# GitHub Pages root (no trailing hash, no filename).
base = "https://ahmedinternist.github.io/rx-viewer/"
rx = qu.Prescription(
    doctor=qu.Doctor(name="Dr. Ahmed", license_no="SCH-1", specialty="GP"),
    patient=qu.Patient(name="Jane Doe", age="34", sex="F"),
    drugs=[qu.DrugItem(generic_name="Amoxicillin", brand_name="Amoxil",
                        dosage="500 mg", frequency="3x/day",
                        duration="7 days", notes="after meals")],
    date="2026-08-30", rx_id="RX-001",
)
url = qu.build_qr_url(rx, viewer_base=base)
print("URL:", url[:90], "...")
assert "#" in url and url.count("#") == 1, "exactly one hash"
frag = url.split("#", 1)[1]
with open("_frag.txt", "w") as f:
    f.write(frag)
print("QR URL shape OK (no trailing slash is fine for Pages)")
print("fragment length:", len(frag))
