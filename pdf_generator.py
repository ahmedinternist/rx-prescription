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


def _qr_path(qr_pil_image) -> Optional[str]:
    if qr_pil_image is None:
        return None
    import tempfile as _tf
    _tmp = _tf.NamedTemporaryFile(suffix=".png", delete=False)
    _tmp.close()
    qr_pil_image.save(_tmp.name)
    return _tmp.name


# ---------------------------------------------------------------------------
# Full prescription PDF
# ---------------------------------------------------------------------------
def generate_prescription_pdf(
    rx: Prescription,
    output_path: str,
    paper_size: str = "A4",
    qr_pil_image=None,
) -> str:
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
    story.append(Paragraph(ar(I.t("pdf_title")), s["title"]))
    story.append(Paragraph(_t("pdf_subtitle"), s["sub"]))
    story.append(Spacer(1, 4 * mm * scale))

    # doctor / patient
    left = [Paragraph(_t("pdf_prescriber"), s["h"]),
            Paragraph(f"<b>{ar(rx.doctor.name or '—')}</b>", s["cell"])]
    if rx.doctor.specialty:
        left.append(Paragraph(f"{_t('specialty')}: {ar(rx.doctor.specialty)}", s["cell"]))
    if rx.doctor.license_no:
        left.append(Paragraph(f"{_t('license')}: {ar(rx.doctor.license_no)}", s["cell"]))

    right = [Paragraph(_t("pdf_patient"), s["h"]),
             Paragraph(f"<b>{ar(rx.patient.name or '—')}</b>", s["cell"])]
    pat = []
    if rx.patient.sex:
        pat.append(f"{_t('sex')}: {ar(rx.patient.sex)}")
    if rx.patient.age:
        pat.append(f"{_t('age')}: {ar(rx.patient.age)}")
    if pat:
        right.append(Paragraph("  |  ".join(pat), s["cell"]))
    right.append(Paragraph(f"{_t('pdf_date')} {rx.date or '—'}", s["cell"]))
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
        drug_txt = f"<b>{ar(d.generic_name or '—')}</b>"
        if d.brand_name:
            drug_txt += f"<br/><font size=7 color='#5b6b85'>{ar(d.brand_name)}</font>"
        data.append([Paragraph(str(i), s["cell"]),
                     Paragraph(drug_txt, s["cell"]),
                     Paragraph(ar(d.dosage or ""), s["cell"]),
                     Paragraph(ar(d.frequency or ""), s["cell"]),
                     Paragraph(ar(d.duration or ""), s["cell"]),
                     Paragraph(ar(d.notes or ""), s["cell"])])
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
                   Paragraph(f"<b>{ar(rx.doctor.name or '')}</b>", s["cell"])]
    if rx.doctor.license_no:
        footer_left.append(Paragraph(f"{_t('license')}: {ar(rx.doctor.license_no)}", s["cell"]))

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

    # 7-line blank banner at the top (above the heading / border)
    for _ in range(7):
        doc.add_paragraph()

    h = doc.add_paragraph()
    rh = h.add_run(I.t("pdf_medications"))
    rh.bold = True
    rh.font.size = Pt(13)
    rh.font.color.rgb = navy

    table = doc.add_table(rows=1, cols=5)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    for i, txt in enumerate([I.t("pdf_col_no"), I.t("pdf_col_drug"),
                             I.t("pdf_col_dosage"), I.t("pdf_col_freq"),
                             I.t("pdf_col_duration")]):
        hdr[i].paragraphs[0].add_run(txt).bold = True
    for i, d in enumerate(rx.drugs, 1):
        cells = table.add_row().cells
        cells[0].text = str(i)
        name = d.generic_name or "—"
        if d.brand_name:
            name += f"\n({d.brand_name})"
        cells[1].text = name
        cells[2].text = d.dosage or ""
        cells[3].text = d.frequency or ""
        cells[4].text = d.duration or ""
        # notes, if present, appended under the dosage cell
        if d.notes:
            cells[2].text += f"\n({I.t('notes')}: {d.notes})"

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
    from docx.enum.table import WD_TABLE_ALIGNMENT

    navy = RGBColor(0x0b, 0x3d, 0x91)
    muted = RGBColor(0x5b, 0x6b, 0x85)

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    # Title
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
    labelled(I.t("pdf_date"), rx.date or "—")
    if rx.rx_id:
        labelled(I.t("pdf_rx"), rx.rx_id)

    doc.add_paragraph()

    h = doc.add_paragraph()
    rh = h.add_run(I.t("pdf_medications"))
    rh.bold = True
    rh.font.color.rgb = navy

    table = doc.add_table(rows=1, cols=6)
    table.style = "Light Grid Accent 1"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, txt in enumerate([I.t("pdf_col_no"), I.t("pdf_col_drug"), I.t("pdf_col_dosage"),
                             I.t("pdf_col_freq"), I.t("pdf_col_duration"), I.t("pdf_col_notes")]):
        hdr[i].paragraphs[0].add_run(txt).bold = True
    for i, d in enumerate(rx.drugs, 1):
        cells = table.add_row().cells
        cells[0].text = str(i)
        name = d.generic_name or "—"
        if d.brand_name:
            name += f"\n({d.brand_name})"
        cells[1].text = name
        cells[2].text = d.dosage or ""
        cells[3].text = d.frequency or ""
        cells[4].text = d.duration or ""
        cells[5].text = d.notes or ""

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

    doc.save(str(output_path))
    return str(output_path)
