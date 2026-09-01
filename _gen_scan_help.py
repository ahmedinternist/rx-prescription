"""Generate a printable one-page 'How to scan' instruction sheet for pharmacists.

Embeds the sample QR (test_qr.png) and explains, in English + Arabic, how to
open the full prescription with any phone camera + internet.
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Image,
                                Table, TableStyle)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# Arabic-capable font (Tahoma on Windows) + RTL reshaping, same as pdf_generator.py
import arabic_reshaper
import bidi.algorithm as bidi_algo

_FONT_PATH = None
for _cand, _bold in [(r"C:/Windows/Fonts/tahoma.ttf", r"C:/Windows/Fonts/tahomabd.ttf"),
                     (r"C:/Windows/Fonts/arial.ttf", r"C:/Windows/Fonts/arialbd.ttf")]:
    if os.path.exists(_cand):
        _FONT_PATH = _cand
        _FONT_BOLD = _bold if os.path.exists(_bold) else _cand
        break

if _FONT_PATH:
    pdfmetrics.registerFont(TTFont("AppFont", _FONT_PATH))
    pdfmetrics.registerFont(TTFont("AppFont-Bold", _FONT_BOLD))
    ARFONT = "AppFont"
else:
    ARFONT = "Helvetica"


def ar(text):
    if not text:
        return text
    if any("\u0600" <= ch <= "\u06ff" for ch in text):
        try:
            return bidi_algo.get_display(arabic_reshaper.reshape(text))
        except Exception:
            return text
    return text


OUT = "scan_help.pdf"
QR = "test_qr.png"

NAVY = colors.HexColor("#0b3d91")
SOFT = colors.HexColor("#5b6b85")
LINE = colors.HexColor("#dce3f2")

styles = getSampleStyleSheet()
title = ParagraphStyle("t", parent=styles["Title"], textColor=NAVY, fontSize=20)
sub = ParagraphStyle("s", parent=styles["Normal"], textColor=SOFT, fontSize=10,
                     alignment=TA_CENTER)
h = ParagraphStyle("h", parent=styles["Heading2"], textColor=NAVY, fontSize=12)
body = ParagraphStyle("b", parent=styles["Normal"], fontSize=11, leading=15)
arbody = ParagraphStyle("ab", parent=styles["Normal"], fontName=ARFONT,
                        fontSize=11, leading=18)
foot = ParagraphStyle("f", parent=styles["Normal"], textColor=SOFT, fontSize=8,
                      alignment=TA_CENTER)
foot_ar = ParagraphStyle("fa", parent=foot, fontName=ARFONT)

doc = SimpleDocTemplate(OUT, pagesize=A4, topMargin=18 * mm, bottomMargin=18 * mm,
                        leftMargin=20 * mm, rightMargin=20 * mm,
                        title="How to scan this prescription")
S = []

S.append(Paragraph("How to read this prescription", title))
S.append(Paragraph("Scan the QR code with any phone camera + internet", sub))
S.append(Spacer(1, 8 * mm))

# QR on the left, steps on the right
qr_img = Image(QR, width=70 * mm, height=70 * mm)
steps_en = (
    "<b>1.</b> Open the Camera app on any smartphone.<br/>"
    "<b>2.</b> Point it at the QR code.<br/>"
    "<b>3.</b> Tap the notification that appears to open the link.<br/>"
    "<b>4.</b> The full prescription opens in the browser — doctor, "
    "patient, drugs, dosage and dates."
)
steps_ar = ar(
    "1. افتح تطبيق الكاميرا في أي هاتف ذكي.\n"
    "2. وجهه نحو رمز الاستجابة السريعة (QR).\n"
    "3. اضغط على الإشعار الذي يظهر لفتح الرابط.\n"
    "4. تفتح الوصفة الكاملة في المتصفح — الطبيب والمريض "
    "والأدوية والجرعات والتواريخ."
)
right = Table([[Paragraph(steps_en, body)], [Spacer(1, 4 * mm)],
               [Paragraph(steps_ar, arbody)]], colWidths=[95 * mm])
right.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                           ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
row = Table([[qr_img, right]], colWidths=[75 * mm, 100 * mm])
row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                         ("LEFTPADDING", (0, 0), (0, 0), 0),
                         ("RIGHTPADDING", (1, 0), (1, 0), 0)]))
S.append(row)
S.append(Spacer(1, 6 * mm))

note = Paragraph(
    "The prescription data is stored entirely in the QR code; it is never sent "
    "to a server. Only the viewer page is loaded from the internet.", foot)
S.append(note)
S.append(Spacer(1, 2 * mm))
note_ar = Paragraph(ar(
    "البيانات مخزنة بالكامل داخل رمز الاستجابة السريعة ولا تُرسل إلى أي خادم. "
    "يُحمّل من الإنترنت صفحة العرض فقط."), foot_ar)
S.append(note_ar)

doc.build(S)
print("wrote", OUT)
