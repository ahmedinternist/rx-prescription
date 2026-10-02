"""Render full prescription PDFs and editable Word prescription documents.

The PDF and Word exports support A5 and A4 paper. Word output is available with or
without the clinic header. All layouts support English and Arabic text via
arabic-reshaper, python-bidi, and an available Tahoma or Arial font.
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Optional
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
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


PAGE_MAP = {"A4": A4, "A5": PAPER_SIZES["A5"]}
SCALE = {"A4": 1.0, "A5": 0.82}


def _generate_modern_prescription_pdf(
        rx: Prescription, output_path: str, paper_size: str, qr_pil_image=None,
        show_header: bool = True, logo_size: str = "medium",
        margin_mm_value: int = 16) -> str:
    """Clean Arabic-first layout used by Preview and Print."""
    if paper_size not in PAGE_MAP:
        paper_size = "A4"
    scale = SCALE.get(paper_size, 1.0)
    margin = max(8, min(30, int(margin_mm_value))) * mm * scale
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
    if show_header:
        logo_factor = {"small": .75, "medium": 1.0, "large": 1.3}.get(
            str(logo_size).casefold(), 1.0)
        mark = Paragraph("<font size=28 color='#007f79'><b>Rx</b></font>", ParagraphStyle("RxMark", parent=base, alignment=TA_LEFT))
        if rx.clinic.logo_path and Path(rx.clinic.logo_path).is_file():
            from reportlab.platypus import Image as RLImage
            mark = RLImage(
                rx.clinic.logo_path, width=22 * logo_factor * mm * scale,
                height=15 * logo_factor * mm * scale)
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
    if not show_header and rx.doctor.specialty:
        doc_lines.append(Paragraph(safe(rx.doctor.specialty), small))
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
                      drug_item.notes, getattr(drug_item, "quantity", "")]
            if x
        )
        body = [Paragraph(title, medicine)] + ([Paragraph(details, base)] if details else [])
        pill = Paragraph(f"<b>{number}</b>", ParagraphStyle("Pill", parent=medicine, alignment=TA_CENTER, textColor=teal, fontSize=13*scale))
        row = Table([[pill, body]], colWidths=[12*mm*scale, doc.width-12*mm*scale])
        row.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP"), ("BACKGROUND", (0,0), (0,0), colors.HexColor("#c9f8f1")), ("TOPPADDING", (0,0), (-1,-1), 5*scale), ("BOTTOMPADDING", (0,0), (-1,-1), 8*scale)]))
        story.append(row)
    if qr_pil_image:
        story += [Spacer(1, 10*mm*scale), Table([[""]], colWidths=[doc.width], style=TableStyle([("LINEABOVE", (0,0), (-1,-1), 1.1, line)])), Spacer(1, 6*mm*scale)]
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


def _append_word_medication_lines(
        doc, drugs, font_size: int = 10, *, line_spacing: float = 1.0,
        field_gap_mm: float = 8.0, first_line_top_mm: float | None = None,
        page_top_margin_mm: float = 12.0):
    """Append numbered medication lines with clear spacing between fields."""
    from docx.shared import Mm, Pt

    paragraphs = []
    for number, drug in enumerate(drugs, 1):
        paragraph = doc.add_paragraph()
        paragraphs.append(paragraph)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(1)
        paragraph.paragraph_format.line_spacing = max(0.8, min(2.0, float(line_spacing)))
        if number == 1 and first_line_top_mm is not None:
            paragraph.paragraph_format.space_before = Mm(
                max(0.0, float(first_line_top_mm) - float(page_top_margin_mm)))
        marker = paragraph.add_run(f"{number}.    ")
        marker.bold = True
        name_run = paragraph.add_run(medicine_name(drug))
        name_run.bold = True
        details = [value for value in (drug.dosage, drug.frequency,
                                       drug.duration, drug.notes,
                                       getattr(drug, "quantity", "")) if value]
        if details:
            # Separate the medicine name and each supplied clinical detail so
            # the line is easy to scan without reverting to a table.
            field_gap = " " * max(2, int(round(float(field_gap_mm) * 0.75)))
            paragraph.add_run(field_gap + field_gap.join(details))
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
    return paragraphs


def _add_floating_picture(paragraph, image_path: str, *, width_mm: float,
                          left_mm: float, top_mm: float) -> None:
    """Insert a borderless, movable Word picture at an absolute page position."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Mm

    inline_shape = paragraph.add_run().add_picture(image_path, width=Mm(width_mm))
    anchor = inline_shape._inline
    anchor.tag = qn("wp:anchor")
    for name, value in {
            "distT": "0", "distB": "0", "distL": "0", "distR": "0",
            "simplePos": "0", "relativeHeight": "251658240", "behindDoc": "0",
            "locked": "0", "layoutInCell": "1", "allowOverlap": "1"}.items():
        anchor.set(name, value)

    simple_position = OxmlElement("wp:simplePos")
    simple_position.set("x", "0")
    simple_position.set("y", "0")
    horizontal = OxmlElement("wp:positionH")
    horizontal.set("relativeFrom", "page")
    horizontal_offset = OxmlElement("wp:posOffset")
    horizontal_offset.text = str(int(Mm(left_mm)))
    horizontal.append(horizontal_offset)
    vertical = OxmlElement("wp:positionV")
    vertical.set("relativeFrom", "page")
    vertical_offset = OxmlElement("wp:posOffset")
    vertical_offset.text = str(int(Mm(top_mm)))
    vertical.append(vertical_offset)
    anchor.insert(0, simple_position)
    anchor.insert(1, horizontal)
    anchor.insert(2, vertical)

    wrap_none = OxmlElement("wp:wrapNone")
    document_properties = anchor.find(qn("wp:docPr"))
    anchor.insert(list(anchor).index(document_properties), wrap_none)


# ---------------------------------------------------------------------------
# Full prescription PDF
# ---------------------------------------------------------------------------
def generate_prescription_pdf(
    rx: Prescription,
    output_path: str,
    paper_size: str = "A4",
    qr_pil_image=None,
    show_header: bool = True,
    logo_size: str = "medium",
    margin_mm_value: int = 16,
) -> str:
    return _generate_modern_prescription_pdf(
        rx, output_path, paper_size, qr_pil_image, show_header,
        logo_size, margin_mm_value)


# ---------------------------------------------------------------------------
# Compact medication label as an EDITABLE Word file
# (drug + dosage/frequency/duration/notes + QR a few lines above the bottom-right)
# ---------------------------------------------------------------------------
def _set_word_paper_size(doc, paper_size):
    """Do not inherit python-docx's Letter page size from its default template."""
    from docx.shared import Pt
    from docx.oxml.ns import qn
    width, height = PAPER_SIZES.get(paper_size, PAPER_SIZES["A4"])
    for section in doc.sections:
        section.page_width = Pt(width)
        section.page_height = Pt(height)
        # Explicit ISO dimensions are authoritative; remove any stale paper code.
        section._sectPr.pgSz.attrib.pop(qn("w:code"), None)


def generate_medication_label_docx(rx: Prescription, output_path: str,
                                   qr_pil_image=None, paper_size: str = "A4",
                                   font_size: int = 10,
                                   layout_profile: dict | None = None) -> str:
    """Write a compact, editable Word file with the medication details + QR.

    On A5 paper, the QR is a movable borderless picture positioned at the lower
    left with a 20 mm safety margin above the pre-printed footer. Scanning it
    reveals the full prescription, while the document body only shows the
    medication content.
    """
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_SECTION

    navy = RGBColor(0x0b, 0x3d, 0x91)
    muted = RGBColor(0x5b, 0x6b, 0x85)

    doc = Document()
    layout = {
        "horizontal_offset_mm": 0.0, "vertical_offset_mm": 0.0,
        "medication_top_mm": 82.0, "line_spacing": 1.0,
        "field_gap_mm": 8.0, "qr_position": "Left",
        "qr_size_mm": 34.0, "qr_side_margin_mm": 12.0,
        "qr_bottom_margin_mm": 20.0,
    }
    if isinstance(layout_profile, dict):
        layout.update(layout_profile)
    horizontal_offset = max(-25.0, min(25.0, float(
        layout.get("horizontal_offset_mm", 0.0))))
    vertical_offset = max(-25.0, min(25.0, float(
        layout.get("vertical_offset_mm", 0.0))))

    _set_word_paper_size(doc, paper_size)
    # Preserve the usable width while shifting it horizontally for printer
    # registration differences.
    for s in doc.sections:
        s.top_margin = Cm(1.2)
        s.bottom_margin = Cm(1.2)
        s.left_margin = Cm(max(0.3, min(3.0, 1.2 + horizontal_offset / 10.0)))
        s.right_margin = Cm(max(0.3, min(3.0, 1.2 - horizontal_offset / 10.0)))

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    font_size = max(8, min(24, int(font_size)))
    style.font.size = Pt(font_size)

    medication_top = max(25.0, min(160.0, float(
        layout.get("medication_top_mm", 82.0))) + vertical_offset)
    medication_paragraphs = _append_word_medication_lines(
        doc, rx.drugs, font_size=font_size,
        line_spacing=float(layout.get("line_spacing", 1.0)),
        field_gap_mm=float(layout.get("field_gap_mm", 8.0)),
        first_line_top_mm=medication_top, page_top_margin_mm=12.0)

    # On pre-printed A5 paper, use a floating borderless object. It remains
    # selectable and movable in Word without consuming medication-body space.
    if qr_pil_image is not None:
        import tempfile as _tf
        _tmp = _tf.NamedTemporaryFile(suffix=".png", delete=False)
        _tmp.close()
        qr_pil_image.save(_tmp.name)
        if paper_size == "A5":
            paper_width_mm, paper_height_mm = 148.0, 210.0
            qr_size_mm = max(20.0, min(55.0, float(
                layout.get("qr_size_mm", 34.0))))
            side_margin = max(3.0, min(40.0, float(
                layout.get("qr_side_margin_mm", 12.0))))
            position = str(layout.get("qr_position", "Left"))
            if position == "Center":
                qr_left = (paper_width_mm - qr_size_mm) / 2.0
            elif position == "Right":
                qr_left = paper_width_mm - side_margin - qr_size_mm
            else:
                qr_left = side_margin
            qr_left = max(0.0, min(paper_width_mm - qr_size_mm,
                                   qr_left + horizontal_offset))
            bottom_margin = max(5.0, min(60.0, float(
                layout.get("qr_bottom_margin_mm", 20.0))))
            qr_top = max(0.0, min(
                paper_height_mm - qr_size_mm,
                paper_height_mm - bottom_margin - qr_size_mm + vertical_offset))
            anchor_paragraph = medication_paragraphs[0] if medication_paragraphs else doc.add_paragraph()
            _add_floating_picture(
                anchor_paragraph, _tmp.name, width_mm=qr_size_mm,
                left_mm=qr_left, top_mm=qr_top)
            # No caption on A5: the pre-printed footer already occupies the
            # bottom strip and the QR itself is self-explanatory.
        else:
            for _ in range(3):
                doc.add_paragraph()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.add_run().add_picture(_tmp.name, width=Inches(1.4))
            cap = doc.add_paragraph()
            cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
            cr = cap.add_run(I.t("pdf_scan"))
            cr.font.size = Pt(8)
            cr.font.color.rgb = muted

    _apply_docx_language(doc)
    doc.save(str(output_path))
    return str(output_path)


def generate_calibration_docx(output_path: str, *, paper_size: str = "A5",
                              layout_profile: dict | None = None) -> str:
    """Create a patient-free printer alignment sheet with a 10 mm ruler grid."""
    from PIL import Image, ImageDraw, ImageFont
    from docx import Document
    from docx.shared import Mm
    import tempfile as _tf

    points = PAPER_SIZES.get(paper_size, PAPER_SIZES["A5"])
    width_mm = points[0] * 25.4 / 72.0
    height_mm = points[1] * 25.4 / 72.0
    pixels_per_mm = 6
    width_px, height_px = int(width_mm * pixels_per_mm), int(height_mm * pixels_per_mm)
    canvas = Image.new("RGB", (width_px, height_px), "white")
    draw = ImageDraw.Draw(canvas)
    try:
        font = ImageFont.truetype(r"C:/Windows/Fonts/arial.ttf", 14)
        bold = ImageFont.truetype(r"C:/Windows/Fonts/arialbd.ttf", 18)
    except OSError:
        font = bold = ImageFont.load_default()

    def mm_px(value):
        return int(round(float(value) * pixels_per_mm))

    for distance in range(0, int(width_mm) + 1, 10):
        x = mm_px(distance)
        draw.line((x, 0, x, height_px), fill="#d8dee8", width=1)
        if distance:
            draw.text((x + 2, 3), str(distance), fill="#64748b", font=font)
    for distance in range(0, int(height_mm) + 1, 10):
        y = mm_px(distance)
        draw.line((0, y, width_px, y), fill="#d8dee8", width=1)
        if distance:
            draw.text((3, y + 2), str(distance), fill="#64748b", font=font)

    layout = dict(layout_profile or {})
    x_offset = float(layout.get("horizontal_offset_mm", 0.0))
    y_offset = float(layout.get("vertical_offset_mm", 0.0))
    medication_top = float(layout.get("medication_top_mm", 82.0)) + y_offset
    draw.line((mm_px(8 + x_offset), mm_px(medication_top),
               mm_px(width_mm - 8 + x_offset), mm_px(medication_top)),
              fill="#d64045", width=3)
    draw.text((mm_px(10 + x_offset), mm_px(medication_top - 5)),
              "Medication start", fill="#a61b22", font=bold)

    qr_size = float(layout.get("qr_size_mm", 34.0))
    side = float(layout.get("qr_side_margin_mm", 12.0))
    position = layout.get("qr_position", "Left")
    if position == "Center":
        qr_left = (width_mm - qr_size) / 2
    elif position == "Right":
        qr_left = width_mm - side - qr_size
    else:
        qr_left = side
    qr_left = max(0.0, min(width_mm - qr_size, qr_left + x_offset))
    qr_top = max(0.0, min(height_mm - qr_size,
        height_mm - float(layout.get("qr_bottom_margin_mm", 20.0)) - qr_size + y_offset))
    box = (mm_px(qr_left), mm_px(qr_top),
           mm_px(qr_left + qr_size), mm_px(qr_top + qr_size))
    draw.rectangle(box, outline="#0964dc", width=4)
    draw.line((box[0], box[1], box[2], box[3]), fill="#0964dc", width=2)
    draw.line((box[2], box[1], box[0], box[3]), fill="#0964dc", width=2)
    draw.text((box[0], max(0, box[1] - 20)), "QR position", fill="#0964dc", font=bold)
    draw.rectangle((mm_px(5), mm_px(5), width_px - mm_px(5), height_px - mm_px(5)),
                   outline="#111827", width=2)
    draw.text((mm_px(8), mm_px(12)), "Printer calibration test",
              fill="#111827", font=bold)
    draw.text((mm_px(8), mm_px(20)),
              f"{paper_size} | offsets X {x_offset:g} mm  Y {y_offset:g} mm",
              fill="#334155", font=font)

    temporary = _tf.NamedTemporaryFile(suffix=".png", delete=False)
    temporary.close()
    canvas.save(temporary.name)
    doc = Document()
    _set_word_paper_size(doc, paper_size)
    for section in doc.sections:
        section.top_margin = Mm(0)
        section.bottom_margin = Mm(0)
        section.left_margin = Mm(0)
        section.right_margin = Mm(0)
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = paragraph.paragraph_format.space_after = 0
    _add_floating_picture(
        paragraph, temporary.name, width_mm=width_mm, left_mm=0, top_mm=0)
    doc.save(str(output_path))
    return str(output_path)


# ---------------------------------------------------------------------------
# Editable Word (.docx) export
# ---------------------------------------------------------------------------
def _generate_reference_headed_docx(
        rx: Prescription, output_path: str, qr_pil_image=None,
        logo_size: str = "medium", margin_mm_value: int = 16,
        paper_size: str = "A4", font_size: int = 10) -> str:
    """Create the headed Word layout used by the clinical letterhead export.

    The reference layout keeps the clinic identity centred, places a movable
    borderless logo at the upper-left, uses compact left-aligned prescription
    metadata and anchors the QR at the lower-left.  Floating pictures remain
    independently selectable in Word and do not consume body-flow space.
    """
    from docx import Document
    from docx.shared import Mm, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    _set_word_paper_size(doc, paper_size)
    margin_mm = max(8, min(30, int(margin_mm_value)))
    for section in doc.sections:
        section.top_margin = Mm(margin_mm)
        section.bottom_margin = Mm(margin_mm)
        section.left_margin = Mm(margin_mm)
        section.right_margin = Mm(margin_mm)

    font_size = max(8, min(24, int(font_size)))
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(font_size)

    def compact_paragraph(*, alignment=None, before=0, after=0):
        paragraph = doc.add_paragraph()
        if alignment is not None:
            paragraph.alignment = alignment
        paragraph.paragraph_format.space_before = Pt(before)
        paragraph.paragraph_format.space_after = Pt(after)
        paragraph.paragraph_format.line_spacing = 1.0
        return paragraph

    title_text = (rx.clinic.name or rx.doctor.name or "Prescription").strip()
    title = compact_paragraph(alignment=WD_ALIGN_PARAGRAPH.CENTER, after=2)
    title_run = title.add_run(title_text)
    title_run.bold = True
    title_run.font.size = Pt(font_size + 3)

    if rx.clinic.logo_path and Path(rx.clinic.logo_path).is_file():
        logo_width_mm = {"small": 18.0, "medium": 24.0, "large": 30.0}.get(
            str(logo_size).casefold(), 24.0)
        _add_floating_picture(
            title, rx.clinic.logo_path, width_mm=logo_width_mm,
            left_mm=float(margin_mm), top_mm=float(margin_mm))

    if rx.doctor.name and rx.doctor.name.strip().casefold() != title_text.casefold():
        doctor = compact_paragraph(alignment=WD_ALIGN_PARAGRAPH.CENTER, after=1)
        doctor_run = doctor.add_run(rx.doctor.name)
        doctor_run.bold = True
        doctor_run.font.size = Pt(font_size + 1)
    if rx.doctor.specialty:
        specialty = compact_paragraph(alignment=WD_ALIGN_PARAGRAPH.CENTER, after=3)
        specialty_run = specialty.add_run(rx.doctor.specialty)
        specialty_run.bold = True
        specialty_run.font.size = Pt(font_size + 1)
    contact_bits = [value for value in (rx.clinic.address, rx.clinic.phone) if value]
    if contact_bits:
        contact = compact_paragraph(alignment=WD_ALIGN_PARAGRAPH.CENTER, after=10)
        contact.add_run("  |  ".join(contact_bits))

    def labelled_line(label, value="", *, after=3):
        paragraph = compact_paragraph(after=after)
        label_run = paragraph.add_run(label)
        label_run.bold = True
        paragraph.add_run(value)
        return paragraph

    if rx.doctor.license_no:
        labelled_line(I.t("license") + ":  ", rx.doctor.license_no)
    labelled_line(I.t("pdf_patient") + ":  ", rx.patient.name or "")
    date_value = rx.date or "—"
    if rx.rx_id:
        date_value += "     |     " + I.t("pdf_rx") + " " + rx.rx_id
    labelled_line(I.t("pdf_date") + "  ", date_value)

    demographics = []
    if rx.patient.age:
        demographics.append(f"{I.t('age')}: {rx.patient.age}")
    if rx.patient.sex:
        demographics.append(f"{I.t('sex')}: {rx.patient.sex}")
    if demographics:
        compact_paragraph(after=3).add_run("   |   ".join(demographics))
    if rx.patient.id_number:
        labelled_line("Patient ID:  ", rx.patient.id_number)
    if rx.patient.allergies:
        labelled_line(I.t("allergies") + ":  ", rx.patient.allergies)
    if rx.diagnosis:
        labelled_line(I.t("diagnosis") + ":  ", rx.diagnosis)
    if rx.refills:
        labelled_line(I.t("refills") + ":  ", rx.refills)

    medication_gap = compact_paragraph(after=0)
    medication_gap.paragraph_format.space_after = Mm(12)
    medication_paragraphs = _append_word_medication_lines(
        doc, rx.drugs, font_size=font_size)

    if qr_pil_image is not None:
        import tempfile as _tf
        temporary = _tf.NamedTemporaryFile(suffix=".png", delete=False)
        temporary.close()
        qr_pil_image.save(temporary.name)
        _page_width_pt, page_height_pt = PAPER_SIZES.get(
            paper_size, PAPER_SIZES["A4"])
        page_height_mm = float(page_height_pt) * 25.4 / 72.0
        qr_size_mm = 34.0
        qr_top_mm = max(
            0.0, page_height_mm - 20.0 - qr_size_mm)
        anchor = medication_paragraphs[0] if medication_paragraphs else title
        _add_floating_picture(
            anchor, temporary.name, width_mm=qr_size_mm,
            left_mm=float(margin_mm), top_mm=qr_top_mm)

    _apply_docx_language(doc)
    doc.save(str(output_path))
    return str(output_path)


def generate_prescription_docx(
        rx: Prescription, output_path: str, qr_pil_image=None,
        show_header: bool = True, logo_size: str = "medium",
        margin_mm_value: int = 16, paper_size: str = "A4",
        font_size: int = 10) -> str:
    """Write a fully editable Word document with the same content + QR image."""
    if show_header:
        return _generate_reference_headed_docx(
            rx, output_path, qr_pil_image=qr_pil_image,
            logo_size=logo_size, margin_mm_value=margin_mm_value,
            paper_size=paper_size, font_size=font_size)

    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    navy = RGBColor(0x0b, 0x3d, 0x91)
    muted = RGBColor(0x5b, 0x6b, 0x85)

    doc = Document()
    _set_word_paper_size(doc, paper_size)
    margin_inches = max(8, min(30, int(margin_mm_value))) / 25.4
    for section in doc.sections:
        section.top_margin = Inches(margin_inches)
        section.bottom_margin = Inches(margin_inches)
        section.left_margin = Inches(margin_inches)
        section.right_margin = Inches(margin_inches)
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    font_size = max(8, min(24, int(font_size)))
    style.font.size = Pt(font_size)

    # Clinic letterhead (no prescription title/subtitle in the headed export).
    logo_width = {"small": .7, "medium": .9, "large": 1.15}.get(
        str(logo_size).casefold(), .9)
    if show_header and rx.clinic.logo_path and Path(rx.clinic.logo_path).is_file():
        logo = doc.add_paragraph()
        logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        logo.add_run().add_picture(rx.clinic.logo_path, width=Inches(logo_width))
    if show_header and rx.clinic.name:
        clinic_title = doc.add_paragraph()
        clinic_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        clinic_title.add_run(rx.clinic.name).bold = True
    if show_header:
        clinic_bits = [rx.clinic.address, rx.clinic.phone]
        if any(clinic_bits):
            clinic_line = doc.add_paragraph("  |  ".join(x for x in clinic_bits if x))
            clinic_line.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def labelled(label, value):
        p = doc.add_paragraph()
        rl = p.add_run(label + "  ")
        rl.bold = True
        rl.font.color.rgb = RGBColor(0, 0, 0) if show_header else navy
        p.add_run(value)
        return p

    def inline_fields(fields):
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(4)
        for index, (label, value) in enumerate(fields):
            if index:
                paragraph.add_run("     |     ")
            paragraph.add_run(label + "  ").bold = True
            paragraph.add_run(value)
        return paragraph

    if show_header:
        doctor_fields = [
            (I.t("pdf_prescriber") + ":", rx.doctor.name),
            (I.t("specialty") + ":", rx.doctor.specialty),
            (I.t("license") + ":", rx.doctor.license_no),
        ]
        doctor_fields = [(label, value) for label, value in doctor_fields if value]
        if doctor_fields:
            inline_fields(doctor_fields)
        date_fields = [(I.t("pdf_date"), rx.date or "—")]
        if rx.rx_id:
            date_fields.append((I.t("pdf_rx"), rx.rx_id))
        inline_fields(date_fields)
    else:
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
    if not show_header:
        labelled(I.t("pdf_date"), rx.date or "—")
        if rx.rx_id:
            labelled(I.t("pdf_rx"), rx.rx_id)

    doc.add_paragraph()

    _append_word_medication_lines(doc, rx.drugs, font_size=font_size)

    if not show_header:
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
