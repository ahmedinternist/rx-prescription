"""Render prescriptions to PDF (full + compact medication label) and to Word.

Two PDF layouts:
  * generate_prescription_pdf  – the full prescription (doctor, patient, drug
    table, QR footer). Paper: A4 / Letter / A5.
  * generate_medication_label_pdf – a compact label: a 7-line blank banner at
    the top (for a stamp / letterhead), then the MEDICATION NAMES ONLY, with
    the QR code pinned to the right-lower border. The QR still encodes the FULL
    prescription, so when scanned the viewer shows doctor/patient/date/Rx too.

Both support switchable paper size (A5 / A4 / Letter) and bidirectional text
(English + Arabic) via arabic-reshaper + python-bidi and a Tahoma/Arial font.
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Optional
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    KeepInFrame,
)

from config import PAPER_SIZES
from qr_utils import Prescription
import i18n as I

# --- Arabic shaping + font registration ------------------------------------
try:
    import arabic_reshaper
    import bidi.algorithm as bidi_algo
    _HAVE_ARABIC = True
except Exception:
    _HAVE_ARABIC = False

_FONT_PATH = None
_FONT_BOLD_PATH = None
for _cand, _bold in [(r"C:/Windows/Fonts/tahoma.ttf", r"C:/Windows/Fonts/tahomabd.ttf"),
                     (r"C:/Windows/Fonts/arial.ttf", r"C:/Windows/Fonts/arialbd.ttf")]:
    if os.path.exists(_cand):
        _FONT_PATH = _cand
        _FONT_BOLD_PATH = _bold if os.path.exists(_bold) else _cand
        break

_FONT_NAME = "Helvetica"
_FONT_BOLD = "Helvetica-Bold"
if _FONT_PATH:
    try:
        pdfmetrics.registerFont(TTFont("AppFont", _FONT_PATH))
        pdfmetrics.registerFont(TTFont("AppFont-Bold", _FONT_BOLD_PATH))
        pdfmetrics.registerFontFamily("AppFont", normal="AppFont", bold="AppFont-Bold",
                                      italic="AppFont", boldItalic="AppFont-Bold")
        _FONT_NAME = "AppFont"
        _FONT_BOLD = "AppFont-Bold"
    except Exception:
        pass


def ar(text: str) -> str:
    """Reshape + reorder Arabic/RTL text for ReportLab."""
    if not text:
        return text
    if _HAVE_ARABIC and any("\u0600" <= ch <= "\u06ff" for ch in text):
        try:
            return bidi_algo.get_display(arabic_reshaper.reshape(text))
        except Exception:
            return text
    return text


def safe(text: str) -> str:
    """Escape user-entered data before it is interpolated into ReportLab markup."""
    return ar(escape(text or ""))


def medicine_name(drug) -> str:
    """Return a complete medicine identity without placeholders or duplication."""
    scientific = (drug.generic_name or "").strip()
    brand = (drug.brand_name or "").strip()
    if scientific and brand and scientific.casefold() != brand.casefold():
        return f"{scientific} ({brand})"
    return scientific or brand or "—"


def _t(key, **kw):
    return ar(I.t(key, **kw))


# Clinical palette
NAVY = colors.HexColor("#0b3d91")
NAVY_SOFT = colors.HexColor("#1e5bbf")
LINE = colors.HexColor("#c9d3e6")
ROW_ALT = colors.HexColor("#f3f6fc")
MUTED = colors.HexColor("#5b6b85")

PAGE_MAP = {"A4": A4, "Letter": letter, "A5": (420.945, 595.276)}
SCALE = {"A4": 1.0, "Letter": 0.94, "A5": 0.82}


def _styles(scale: float):
    ss = getSampleStyleSheet()
    base = ss["Normal"]
    base.fontName = _FONT_NAME
    base.fontSize = 9 * scale
    base.leading = 12 * scale

    title = ParagraphStyle("TitleX", parent=ss["Title"], fontName=_FONT_BOLD,
                           fontSize=22 * scale, leading=26 * scale, alignment=TA_CENTER,
                           textColor=NAVY, spaceAfter=2 * scale)
    sub = ParagraphStyle("Sub", parent=base, fontSize=9 * scale, alignment=TA_CENTER,
                         textColor=NAVY_SOFT)
    h = ParagraphStyle("H", parent=base, fontName=_FONT_BOLD, fontSize=11 * scale,
                       textColor=NAVY, spaceAfter=3 * scale)
    hsmall = ParagraphStyle("HS", parent=base, fontName=_FONT_BOLD, fontSize=9.5 * scale,
                            textColor=NAVY, spaceAfter=2 * scale)
    cell = ParagraphStyle("Cell", parent=base, fontSize=9 * scale, leading=12 * scale)
    cellb = ParagraphStyle("CellB", parent=cell, fontName=_FONT_BOLD)
    label = ParagraphStyle("Label", parent=base, fontName=_FONT_BOLD,
                           fontSize=7 * scale, leading=9 * scale, textColor=NAVY_SOFT)
    small = ParagraphStyle("Small", parent=base, fontSize=7.5 * scale, leading=10 * scale,
                           textColor=MUTED)
    druglabel = ParagraphStyle("DrugL", parent=base, fontName=_FONT_BOLD,
                               fontSize=11 * scale, leading=15 * scale, textColor=NAVY)
    return locals()


def _generate_modern_prescription_pdf(rx: Prescription, output_path: str, paper_size: str,
                                      qr_pil_image=None) -> str:
    """Clean Arabic-first layout used by Preview and Print."""
    if paper_size not in PAGE_MAP:
        paper_size = "A4"
    scale = SCALE.get(paper_size, 1.0)
    margin = 16 * mm * scale
    doc = BaseDocTemplate(str(output_path), pagesize=PAGE_MAP[paper_size], leftMargin=margin,
                          rightMargin=margin, topMargin=margin, bottomMargin=margin)
    doc.addPageTemplates([PageTemplate(id="main", frames=[Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="main")])])
    teal, ink, muted, soft, line = (colors.HexColor("#007f79"), colors.HexColor("#073b3a"),
                                    colors.HexColor("#4f7f80"), colors.HexColor("#f0fbfa"), colors.HexColor("#bff3ec"))
    align = TA_RIGHT if I.get_lang() == "ar" else TA_LEFT
    base = ParagraphStyle("ModernBase", fontName=_FONT_NAME, fontSize=11 * scale, leading=15 * scale, textColor=ink, alignment=align)
    clinic = ParagraphStyle("ModernClinic", parent=base, fontName=_FONT_BOLD, fontSize=19 * scale, leading=24 * scale)
    doctor = ParagraphStyle("ModernDoctor", parent=base, fontName=_FONT_BOLD, fontSize=15 * scale, leading=20 * scale)
    medicine = ParagraphStyle("ModernMedicine", parent=base, fontName=_FONT_BOLD, fontSize=15 * scale, leading=20 * scale)
    small = ParagraphStyle("ModernSmall", parent=base, fontSize=8.5 * scale, leading=12 * scale, textColor=muted)
    story = []
    mark = Paragraph("<font size=28 color='#007f79'><b>Rx</b></font>", ParagraphStyle("RxMark", parent=base, alignment=TA_LEFT))
    if rx.clinic.logo_path and Path(rx.clinic.logo_path).is_file():
        from reportlab.platypus import Image as RLImage
        mark = RLImage(rx.clinic.logo_path, width=22 * mm * scale, height=15 * mm * scale)
    header_lines = [Paragraph(safe(rx.clinic.name or "Prescription"), clinic)]
    if rx.doctor.specialty:
        header_lines.append(Paragraph(safe(rx.doctor.specialty), ParagraphStyle("Specialty", parent=base, fontSize=12 * scale, textColor=muted)))
    if rx.clinic.address or rx.clinic.phone:
        header_lines.append(Paragraph(safe("  |  ".join(x for x in [rx.clinic.address, rx.clinic.phone] if x)), small))
    header = Table([[mark, header_lines]], colWidths=[doc.width * .23, doc.width * .77])
    header.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP"), ("ALIGN", (1,0), (1,0), "RIGHT"), ("BOTTOMPADDING", (0,0), (-1,-1), 7 * scale)]))
    story += [header, Table([[""]], colWidths=[doc.width], style=TableStyle([("LINEABOVE", (0,0), (-1,-1), 1.1, line)])), Spacer(1, 8 * mm * scale)]
    date = Paragraph(f"{_t('pdf_date')} {safe(rx.date)}", small)
    doc_lines = [Paragraph(safe(rx.doctor.name), doctor)]
    if rx.doctor.license_no:
        doc_lines.append(Paragraph(f"{_t('license')}: {safe(rx.doctor.license_no)}", small))
    info = Table([[date, doc_lines]], colWidths=[doc.width*.38, doc.width*.62])
    info.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP"), ("ALIGN", (1,0), (1,0), "RIGHT"), ("BOTTOMPADDING", (0,0), (-1,-1), 8 * scale)]))
    story.append(info)
    # The modern preview already groups patient details in a dedicated panel.
    # Show the name only so it is not repeated as "Patient: Patient Name".
    patient_lines = [f"<b>{safe(rx.patient.name or '—')}</b>"]
    if rx.patient.age: patient_lines.append(f"{_t('age')}: {safe(rx.patient.age)}")
    if rx.diagnosis: patient_lines.append(f"<b>{safe(I.t('diagnosis'))}:</b> {safe(rx.diagnosis)}")
    if rx.patient.allergies: patient_lines.append(f"<b>{safe(I.t('allergies'))}:</b> {safe(rx.patient.allergies)}")
    patient = Table([[Paragraph("<br/>".join(patient_lines), base)]], colWidths=[doc.width])
    patient.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,-1), soft), ("ROUNDEDCORNERS", [15,15,15,15]), ("TOPPADDING", (0,0), (-1,-1), 11*scale), ("BOTTOMPADDING", (0,0), (-1,-1), 11*scale), ("LEFTPADDING", (0,0), (-1,-1), 13*scale), ("RIGHTPADDING", (0,0), (-1,-1), 13*scale)]))
    story += [patient, Spacer(1, 8 * mm * scale)]
    for number, drug_item in enumerate(rx.drugs, 1):
        title = safe(medicine_name(drug_item))
        # Display only values actually chosen for this medication.
        details = " - ".join(
            safe(x)
            for x in [drug_item.dosage, drug_item.frequency, drug_item.duration,
                      drug_item.notes]
            if x
        )
        body = [Paragraph(title, medicine)] + ([Paragraph(details, base)] if details else [])
        pill = Paragraph(f"<b>{number}</b>", ParagraphStyle("Pill", parent=medicine, alignment=TA_CENTER, textColor=teal, fontSize=13*scale))
        row = Table([[pill, body]], colWidths=[12*mm*scale, doc.width-12*mm*scale])
        row.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP"), ("BACKGROUND", (0,0), (0,0), colors.HexColor("#c9f8f1")), ("TOPPADDING", (0,0), (-1,-1), 5*scale), ("BOTTOMPADDING", (0,0), (-1,-1), 8*scale)]))
        story.append(row)
    story += [Spacer(1, 10*mm*scale), Table([[""]], colWidths=[doc.width], style=TableStyle([("LINEABOVE", (0,0), (-1,-1), 1.1, line)])), Spacer(1, 6*mm*scale)]
    if qr_pil_image:
        from reportlab.platypus import Image as RLImage
        qr = RLImage(_qr_path(qr_pil_image), width=42*mm*scale, height=42*mm*scale)
        caption = Paragraph("Scan", base)
        footer = Table([[caption, qr]], colWidths=[doc.width*.58, doc.width*.42])
        footer.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("ALIGN", (1,0), (1,0), "RIGHT"), ("BOX", (1,0), (1,0), .8, line), ("TOPPADDING", (1,0), (1,0), 5*scale), ("BOTTOMPADDING", (1,0), (1,0), 5*scale)]))
        story.append(footer)
    doc.build(story)
    return str(output_path)


def _qr_path(qr_pil_image) -> Optional[str]:
    if qr_pil_image is None:
        return None
    import tempfile as _tf
    _tmp = _tf.NamedTemporaryFile(suffix=".png", delete=False)
    _tmp.close()
    qr_pil_image.save(_tmp.name)
    return _tmp.name


def _apply_docx_language(doc) -> None:
    """Set Word bidirectional metadata for Arabic text in either UI language."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    def has_arabic(text: str) -> bool:
        return any("\u0600" <= char <= "\u08ff" for char in text)

    def has_latin(text: str) -> bool:
        return any(("A" <= char <= "Z") or ("a" <= char <= "z") for char in text)

    paragraphs = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs.extend(cell.paragraphs)
    for paragraph in paragraphs:
        text = paragraph.text
        # Arabic UI is right-to-left throughout.  Under the English UI, make
        # Arabic-only values (such as a table cell) RTL while preserving the
        # English label/value order in mixed paragraphs.
        if I.get_lang() == "ar" or (has_arabic(text) and not has_latin(text)):
            p_pr = paragraph._p.get_or_add_pPr()
            bidi = OxmlElement("w:bidi")
            bidi.set(qn("w:val"), "1")
            p_pr.append(bidi)
        for run in paragraph.runs:
            if not has_arabic(run.text):
                continue
            r_pr = run._r.get_or_add_rPr()
            rtl = OxmlElement("w:rtl")
            rtl.set(qn("w:val"), "1")
            r_pr.append(rtl)
            language = OxmlElement("w:lang")
            language.set(qn("w:bidi"), "ar-IQ")
            r_pr.append(language)


def _word_medication_columns(drugs):
    """Return only medication detail columns that contain a value."""
    fields = [
        ("dosage", I.t("pdf_col_dosage")),
        ("frequency", I.t("pdf_col_freq")),
        ("duration", I.t("pdf_col_duration")),
        ("notes", I.t("pdf_col_notes")),
    ]
    return [(field, label) for field, label in fields
            if any(getattr(drug, field, "").strip() for drug in drugs)]


def _append_word_medication_lines(doc, drugs) -> None:
    """Append numbered medication lines with clear spacing between fields."""
    from docx.shared import Pt

    for number, drug in enumerate(drugs, 1):
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(1)
        paragraph.paragraph_format.line_spacing = 1
        marker = paragraph.add_run(f"{number}.    ")
        marker.bold = True
        name_run = paragraph.add_run(medicine_name(drug))
        name_run.bold = True
        details = [value for value in (drug.dosage, drug.frequency,
                                       drug.duration, drug.notes) if value]
        if details:
            # Separate the medicine name and each supplied clinical detail so
            # the line is easy to scan without reverting to a table.
            field_gap = "      "
            paragraph.add_run(field_gap + field_gap.join(details))


# ---------------------------------------------------------------------------
# Full prescription PDF
# ---------------------------------------------------------------------------
def generate_prescription_pdf(
    rx: Prescription,
    output_path: str,
    paper_size: str = "A4",
    qr_pil_image=None,
) -> str:
    return _generate_modern_prescription_pdf(rx, output_path, paper_size, qr_pil_image)
    if paper_size not in PAGE_MAP:
        paper_size = "A4"
    page = PAGE_MAP[paper_size]
    scale = SCALE.get(paper_size, 1.0)
    s = _styles(scale)

    margin = 16 * mm * scale
    doc = BaseDocTemplate(str(output_path), pagesize=page, leftMargin=margin,
                          rightMargin=margin, topMargin=margin, bottomMargin=margin,
                          title="Prescription", author=rx.doctor.name or "Doctor")
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="main")
    doc.addPageTemplates([PageTemplate(id="main", frames=[frame])])

    story = []
    if rx.clinic.logo_path and Path(rx.clinic.logo_path).is_file():
        from reportlab.platypus import Image as RLImage
        story.append(RLImage(rx.clinic.logo_path, width=28 * mm * scale, height=18 * mm * scale))
    if rx.clinic.name:
        story.append(Paragraph(f"<b>{safe(rx.clinic.name)}</b>", s["title"]))
        clinic_bits = [safe(rx.clinic.address), safe(rx.clinic.phone)]
        clinic_line = "  |  ".join(bit for bit in clinic_bits if bit)
        if clinic_line:
            story.append(Paragraph(clinic_line, s["small"]))
        story.append(Spacer(1, 3 * mm * scale))
    story.append(Paragraph(ar(I.t("pdf_title")), s["title"]))
    story.append(Paragraph(_t("pdf_subtitle"), s["sub"]))
    story.append(Spacer(1, 4 * mm * scale))

    # doctor / patient
    left = [Paragraph(_t("pdf_prescriber"), s["h"]),
            Paragraph(f"<b>{safe(rx.doctor.name or '—')}</b>", s["cell"])]
    if rx.doctor.specialty:
        left.append(Paragraph(f"{_t('specialty')}: {safe(rx.doctor.specialty)}", s["cell"]))
    if rx.doctor.license_no:
        left.append(Paragraph(f"{_t('license')}: {safe(rx.doctor.license_no)}", s["cell"]))

    right = [Paragraph(_t("pdf_patient"), s["h"]),
             Paragraph(f"<b>{safe(rx.patient.name or '—')}</b>", s["cell"])]
    pat = []
    if rx.patient.sex:
        pat.append(f"{_t('sex')}: {safe(rx.patient.sex)}")
    if rx.patient.age:
        pat.append(f"{_t('age')}: {safe(rx.patient.age)}")
    if rx.patient.id_number:
        pat.append(f"ID: {safe(rx.patient.id_number)}")
    if pat:
        right.append(Paragraph("  |  ".join(pat), s["cell"]))
    right.append(Paragraph(f"{_t('pdf_date')} {rx.date or '—'}", s["cell"]))
    if rx.patient.allergies:
        right.append(Paragraph(f"Allergies: {safe(rx.patient.allergies)}", s["cell"]))
    if rx.diagnosis:
        right.append(Paragraph(f"Diagnosis: {safe(rx.diagnosis)}", s["cell"]))
    if rx.refills:
        right.append(Paragraph(f"Refills: {safe(rx.refills)}", s["cell"]))
    if rx.rx_id:
        right.append(Paragraph(f"{_t('pdf_rx')} {rx.rx_id}", s["cell"]))

    info_tbl = Table([[left, right]], colWidths=[doc.width * 0.5, doc.width * 0.5])
    info_tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                  ("LINEBELOW", (0, 0), (-1, -1), 0.8, NAVY),
                                  ("BOTTOMPADDING", (0, 0), (-1, -1), 5 * scale),
                                  ("TOPPADDING", (0, 0), (-1, -1), 2 * scale)]))
    story.append(info_tbl)
    story.append(Spacer(1, 5 * mm * scale))

    # medications table
    story.append(Paragraph(_t("pdf_medications"), s["h"]))
    story.append(Spacer(1, 3 * mm * scale))
    header = [Paragraph(_t("pdf_col_no"), s["cellb"]),
              Paragraph(_t("pdf_col_drug"), s["cellb"]),
              Paragraph(_t("pdf_col_dosage"), s["cellb"]),
              Paragraph(_t("pdf_col_freq"), s["cellb"]),
              Paragraph(_t("pdf_col_duration"), s["cellb"]),
              Paragraph(_t("pdf_col_notes"), s["cellb"])]
    data = [header]
    for i, d in enumerate(rx.drugs, 1):
        primary_name = d.generic_name or d.brand_name or "—"
        drug_txt = f"<b>{safe(primary_name)}</b>"
        if d.generic_name and d.brand_name and d.generic_name.casefold() != d.brand_name.casefold():
            drug_txt += f"<br/><font size=7 color='#5b6b85'>{safe(d.brand_name)}</font>"
        data.append([Paragraph(str(i), s["cell"]),
                     Paragraph(drug_txt, s["cell"]),
                     Paragraph(safe(d.dosage), s["cell"]),
                     Paragraph(safe(d.frequency), s["cell"]),
                     Paragraph(safe(d.duration), s["cell"]),
                     Paragraph(safe(d.notes), s["cell"])])
    cw = [doc.width * x for x in (0.05, 0.30, 0.15, 0.17, 0.15, 0.18)]
    drug_tbl = Table(data, colWidths=cw, repeatRows=1)
    drug_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ROW_ALT]),
        ("TOPPADDING", (0, 0), (-1, -1), 4 * scale),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4 * scale),
        ("LEFTPADDING", (0, 0), (-1, -1), 4 * scale),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4 * scale)]))
    story.append(drug_tbl)
    story.append(Spacer(1, 6 * mm * scale))

    # footer: signature + QR
    footer_left = [Paragraph(_t("pdf_signature"), s["h"]),
                   Spacer(1, 12 * mm * scale),
                   Paragraph(f"<b>{safe(rx.doctor.name)}</b>", s["cell"])]
    if rx.doctor.license_no:
        footer_left.append(Paragraph(f"{_t('license')}: {safe(rx.doctor.license_no)}", s["cell"]))

    footer_right = []
    qp = _qr_path(qr_pil_image)
    if qp:
        from reportlab.platypus import Image as RLImage
        from PIL import Image as PILImage
        iw, ih = PILImage.open(qp).size
        qr_box = 34 * mm * scale
        ratio = qr_box / iw
        qr_img = RLImage(qp, width=iw * ratio, height=ih * ratio)
        qr_caption = Paragraph(_t("pdf_scan"), s["small"])
        qr_block = Table([[qr_img], [qr_caption]], colWidths=[qr_box])
        qr_block.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"),
                                      ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        footer_right.append(qr_block)

    footer_tbl = Table([[footer_left, footer_right]],
                       colWidths=[doc.width * 0.62, doc.width * 0.38])
    footer_tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
                                    ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                                    ("LINEABOVE", (0, 0), (-1, -1), 0.4, colors.HexColor("#9aa7c0")),
                                    ("TOPPADDING", (0, 0), (-1, -1), 4 * scale)]))
    story.append(footer_tbl)
    doc.build(story)
    return str(output_path)


# ---------------------------------------------------------------------------
# Compact medication label as an EDITABLE Word file
# (drug + dosage/frequency/duration/notes + QR a few lines above the bottom-right)
# ---------------------------------------------------------------------------
def generate_medication_label_docx(rx: Prescription, output_path: str,
                                   qr_pil_image=None) -> str:
    """Write a compact, editable Word file with the medication details + QR.

    The QR is placed a few lines above the bottom-right corner (not flush in the
    corner). Scanning it reveals the FULL prescription, while the document body
    only shows the medication content.
    """
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_SECTION

    navy = RGBColor(0x0b, 0x3d, 0x91)
    muted = RGBColor(0x5b, 0x6b, 0x85)

    doc = Document()

    # narrower label-friendly margins
    for s in doc.sections:
        s.top_margin = Cm(1.2)
        s.bottom_margin = Cm(1.2)
        s.left_margin = Cm(1.2)
        s.right_margin = Cm(1.2)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    # Eight blank lines leave one extra writing line for a stamped/printed
    # header while keeping the no-header layout intentionally unbranded.
    for _ in range(8):
        doc.add_paragraph()

    _append_word_medication_lines(doc, rx.drugs)

    # a few blank lines, then the QR a few lines above the bottom-right corner
    for _ in range(3):
        doc.add_paragraph()

    if qr_pil_image is not None:
        import tempfile as _tf
        _tmp = _tf.NamedTemporaryFile(suffix=".png", delete=False)
        _tmp.close()
        qr_pil_image.save(_tmp.name)
        # right-aligned paragraph; QR sized small for a label
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run().add_picture(_tmp.name, width=Inches(1.4))
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        cr = cap.add_run(I.t("pdf_scan"))
        cr.font.size = Pt(8)
        cr.font.color.rgb = muted

    _apply_docx_language(doc)
    doc.save(str(output_path))
    return str(output_path)


# ---------------------------------------------------------------------------
# Editable Word (.docx) export
# ---------------------------------------------------------------------------
def generate_prescription_docx(rx: Prescription, output_path: str,
                              qr_pil_image=None) -> str:
    """Write a fully editable Word document with the same content + QR image."""
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    navy = RGBColor(0x0b, 0x3d, 0x91)
    muted = RGBColor(0x5b, 0x6b, 0x85)

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    # Clinic identity and title
    if rx.clinic.logo_path and Path(rx.clinic.logo_path).is_file():
        logo = doc.add_paragraph()
        logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        logo.add_run().add_picture(rx.clinic.logo_path, width=Inches(0.9))
    if rx.clinic.name:
        clinic_title = doc.add_paragraph()
        clinic_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        clinic_title.add_run(rx.clinic.name).bold = True
        clinic_bits = [rx.clinic.address, rx.clinic.phone]
        if any(clinic_bits):
            clinic_line = doc.add_paragraph("  |  ".join(x for x in clinic_bits if x))
            clinic_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = t.add_run(I.t("pdf_title"))
    run.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = navy
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run(I.t("pdf_subtitle"))
    r.font.color.rgb = muted
    r.font.size = Pt(11)

    def labelled(label, value):
        p = doc.add_paragraph()
        rl = p.add_run(label + "  ")
        rl.bold = True
        rl.font.color.rgb = navy
        p.add_run(value)
        return p

    if rx.doctor.name:
        labelled(I.t("pdf_prescriber") + ":", rx.doctor.name)
    if rx.doctor.specialty:
        labelled(I.t("specialty") + ":", rx.doctor.specialty)
    if rx.doctor.license_no:
        labelled(I.t("license") + ":", rx.doctor.license_no)
    if rx.patient.name:
        labelled(I.t("pdf_patient") + ":", rx.patient.name)
    pat_bits = []
    if rx.patient.sex:
        pat_bits.append(f"{I.t('sex')}: {rx.patient.sex}")
    if rx.patient.age:
        pat_bits.append(f"{I.t('age')}: {rx.patient.age}")
    if pat_bits:
        labelled("", "   |   ".join(pat_bits))
    if rx.patient.id_number:
        labelled("Patient ID:", rx.patient.id_number)
    if rx.patient.allergies:
        labelled("Allergies:", rx.patient.allergies)
    if rx.diagnosis:
        labelled("Diagnosis:", rx.diagnosis)
    if rx.refills:
        labelled("Refills:", rx.refills)
    labelled(I.t("pdf_date"), rx.date or "—")
    if rx.rx_id:
        labelled(I.t("pdf_rx"), rx.rx_id)

    doc.add_paragraph()

    _append_word_medication_lines(doc, rx.drugs)

    doc.add_paragraph()
    sig = doc.add_paragraph()
    sig.add_run(I.t("pdf_signature") + "\n").bold = True
    if rx.doctor.name:
        sig.add_run(rx.doctor.name + "\n")
    if rx.doctor.license_no:
        sig.add_run(f"{I.t('license')}: {rx.doctor.license_no}")

    if qr_pil_image is not None:
        import tempfile as _tf
        _tmp = _tf.NamedTemporaryFile(suffix=".png", delete=False)
        _tmp.close()
        qr_pil_image.save(_tmp.name)
        doc.add_paragraph()
        doc.add_picture(_tmp.name, width=Inches(1.6))
        cap = doc.add_paragraph()
        cr = cap.add_run(I.t("pdf_scan"))
        cr.font.size = Pt(8)
        cr.font.color.rgb = muted

    _apply_docx_language(doc)
    doc.save(str(output_path))
    return str(output_path)
