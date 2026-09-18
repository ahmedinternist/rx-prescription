"""Prescription desktop app — main GUI (CustomTkinter).

Features
--------
* Bilingual (English / Arabic) UI + documents.
* Prescriber (doctor) and Patient boxes shown SIDE BY SIDE and kept compact.
* Repeatable drug rows: drug NAME on top, with Dosage / Frequency / Duration /
  Notes boxes directly beneath it (in that order), and autocomplete.
* Importable drug database (replace / merge) + export.
* Paper size: A5 / A4.
* Outputs:
    - Preview/Print : full prescription PDF (doctor, patient, drug table, QR).
    - Medication Label : compact Word doc with the full medication content + QR.
    - Export Word : editable .docx of the full prescription.

Run:  python main.py
"""
from __future__ import annotations

import datetime
import csv
import concurrent.futures
import copy
import queue
import functools
import logging
import math
import os
import tempfile
import threading
import tkinter as tk
import tkinter.font as tkfont
import unicodedata
import uuid
import weakref
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import List

import customtkinter as ctk
from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageFilter

import config as cfg
import drug_db as dbmod
import pdf_generator as pdfgen
import qr_utils as qu
import cloud_rx
import i18n as I
import openfda
import drug_classes as classes
import gemini_drug
from workflow_features import calculate_medicine_quantity, compare_prescriptions
from patient_history import PatientHistory

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

APP_TITLE = "Rx Prescription Printer"

logging.basicConfig(filename=cfg.LOG_PATH, level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s")


def _report_exception(exc_type, exc_value, exc_traceback):
    logging.exception("Unexpected application error", exc_info=(exc_type, exc_value, exc_traceback))
    messagebox.showerror(APP_TITLE, f"An unexpected error occurred. Details were saved to:\n{cfg.LOG_PATH}")


tk.Tk.report_callback_exception = staticmethod(_report_exception)

# --- Portable white/light-gray frosted-glass visual system ------------------
# Simulated glass keeps text opaque and needs no Windows compositor/live blur.
BG = "#f6f7f9"
CARD = "#e9eaec"
SURFACE = "#ffffff"
SIDEBAR_SURFACE = "#eef0f3"
TEXT = "#25272a"
TEXT_SECONDARY = "#454c56"
ACCENT = "#0960c7"      # saturated blue, readable on white/light-gray surfaces
PRIMARY = "#0964dc"     # blue fill with opaque white button text
ACCENT_HOVER = "#084fb0"
ACCENT_SOFT = "#e1ebfb"
ICON_BLUE = ACCENT
NAV_ACTIVE = "#144d99"  # deeper blue distinguishes the selected section
NAV_ACTIVE_SOFT = ACCENT_SOFT
LINE = "#cdd2d9"
MUTED = "#59616c"
DANGER = "#b63545"
DANGER_FILL = "#b63545"
DANGER_SOFT = "#fae9e9"
DANGER_HOVER = "#ad3636"
WARNING = "#8b5a0e"
WARNING_SOFT = "#fff3d6"
GOOD = PRIMARY
GOOD_HOVER = ACCENT_HOVER
SUCCESS = "#176b51"

# Keep editable fields solid even inside the simulated frosted panels.
for _field_theme in ("CTkEntry", "CTkComboBox"):
    ctk.ThemeManager.theme[_field_theme].update(
        fg_color=[SURFACE, SURFACE], border_color=[LINE, LINE],
        text_color=[TEXT, TEXT])

# Explicit colors also cover controls that previously inherited the light theme.
ctk.ThemeManager.theme["CTk"].update(fg_color=[BG, BG])
ctk.ThemeManager.theme["CTkToplevel"].update(fg_color=[BG, BG])
ctk.ThemeManager.theme["CTkFrame"].update(
    fg_color=[CARD, CARD], top_fg_color=[CARD, CARD], border_color=[LINE, LINE])
ctk.ThemeManager.theme["CTkLabel"].update(text_color=[TEXT, TEXT])
ctk.ThemeManager.theme["CTkButton"].update(
    fg_color=[PRIMARY, PRIMARY], hover_color=[ACCENT_HOVER, ACCENT_HOVER],
    text_color=["#ffffff", "#ffffff"], text_color_disabled=[MUTED, MUTED], border_color=[LINE, LINE])
ctk.ThemeManager.theme["DropdownMenu"].update(
    fg_color=[SURFACE, SURFACE], hover_color=[ACCENT_SOFT, ACCENT_SOFT], text_color=[TEXT, TEXT])
ctk.ThemeManager.theme["CTkTextbox"].update(
    fg_color=[SURFACE, SURFACE], text_color=[TEXT, TEXT], border_color=[LINE, LINE])
for _choice_theme in ("CTkOptionMenu", "CTkComboBox"):
    ctk.ThemeManager.theme[_choice_theme].update(
        fg_color=[SURFACE, SURFACE], button_color=[PRIMARY, PRIMARY],
        button_hover_color=[ACCENT_HOVER, ACCENT_HOVER], text_color=[TEXT, TEXT])
for _choice_theme in ("CTkCheckBox", "CTkRadioButton", "CTkSwitch"):
    ctk.ThemeManager.theme[_choice_theme].update(
        fg_color=[PRIMARY, PRIMARY], text_color=[TEXT, TEXT], border_color=[LINE, LINE])

PAD = 10                 # compact, readable card padding
CARD_GAP = 4             # consistent gap without stretching short cards
SCROLLBAR_HIDDEN_PAGES = frozenset({
    "prescriber", "patient", "treatment_templates", "interaction_review", "reference",
})
LABEL_FONT = ("Segoe UI", 12)
BRAND_FONT = ("Segoe UI", 13, "bold")
SCIENTIFIC_FONT = ("Segoe UI", 13)
FIELD_HEIGHT = 36        # consistent text-entry and dropdown height
ACTION_HEIGHT = 34       # compact text-action height
ICON_BUTTON_SIZE = 30
LIST_FONT = ("Segoe UI", 30)
PAGE_TITLE_FONT_SIZE = 35
SELECTED_MEDICINE_FONT_SIZE = 35
DASHBOARD_WIDTH = 200
DASHBOARD_COLLAPSED_WIDTH = 62
DASHBOARD_ICON_SIZE = 24
PRESCRIBER_FIELD_WIDTH = 440
PATIENT_NAME_WIDTH = 410
PATIENT_AGE_WIDTH = 76
PATIENT_SEX_WIDTH = 126
PATIENT_SEARCH_WIDTH = 280
PATIENT_RESULTS_HEIGHT = 190
NAV_ICONS = {
    "prescriber": "✚", "patient": "⚕", "medications": "◒",
    "favorites": "★", "drug_classes": "⌬", "treatment_templates": "▥",
    "interaction_review": "⚯",
    "reference": "▤", "settings": "⚒",
}
REFERENCE_TAGS = {
    "indication": ("#e5f5f1", "#167565"),
    "dose": ("#e1ebfb", "#0960c7"),
    "contraindication": ("#fbe9e8", "#ae3d3d"),
    "pregnancy": ("#fff2d8", "#8b5a0e"),
    "renal": ("#f1ebfb", "#7546a3"),
}

@functools.lru_cache(maxsize=12)
def glass_texture(surface="workspace"):
    """Small, cached static texture; resized only after geometry settles."""
    image = Image.new("RGB", (96, 96))
    pixels = image.load()
    for y in range(96):
        for x in range(96):
            u, v = x / 95, y / 95
            blue = math.exp(-((u - .18) ** 2 + (v - .15) ** 2) / .30)
            teal = math.exp(-((u - .86) ** 2 + (v - .84) ** 2) / .35)
            color = (int(246 - 5 * blue - 4 * teal),
                     int(247 - 4 * blue - 3 * teal),
                     int(249 - 2 * blue - 2 * teal))
            # Broad white haze grows towards the lower-left of the shared glass.
            haze = .015 + .055 * v * (1 - .55 * u)
            color = tuple(round(c * (1 - haze) + 255 * haze) for c in color)
            veil = {"workspace": 0, "sidebar": .74, "panel": .88}.get(surface, .88)
            target = tuple(bytes.fromhex((SIDEBAR_SURFACE if surface == "sidebar" else CARD)[1:]))
            pixels[x, y] = tuple(round(c * (1 - veil) + t * veil) for c, t in zip(color, target))
    return image


@functools.lru_cache(maxsize=3)
def glass_reflection(surface="panel"):
    """Cached white-gradient mask, not live blur or window transparency."""
    mask = Image.new("L", (96, 96))
    pixels = mask.load()
    for y in range(96):
        for x in range(96):
            u, v = x / 95, y / 95
            top_glint = .045 * math.exp(-v * 7) * (1 - .5 * u)
            bottom_haze = .075 * v ** 1.8 * (1 - .35 * u)
            strength = min(.085, .012 + top_glint + bottom_haze)
            pixels[x, y] = round(255 * strength)
    return mask


def glass_panel_image(size, surface, region=(0, 0, 1, 1), tint=None):
    """Composite a panel against its portion of the shared window backdrop."""
    width, height, radius, edge = size
    extent = tuple(value * 96 for value in region)
    image = glass_texture("workspace").transform(
        (width, height), Image.Transform.EXTENT, extent, Image.Resampling.BILINEAR)
    veil = {"workspace": 0, "sidebar": .74, "panel": .88, "sheet": .78}.get(surface, .88)
    target = SURFACE if surface == "sheet" else SIDEBAR_SURFACE if surface == "sidebar" else CARD
    image = Image.blend(image, Image.new("RGB", image.size, target), veil)
    if surface != "workspace":
        reflection = glass_reflection(surface).resize(image.size, Image.Resampling.BILINEAR)
        image = Image.composite(Image.new("RGB", image.size, "white"), image, reflection)
    if tint:
        image = Image.blend(image, Image.new("RGB", image.size, tint), .32)
    image = image.convert("RGBA")
    mask = Image.new("L", image.size)
    if width > 2 * edge and height > 2 * edge:
        bounds = (edge, edge, width - edge - 1, height - edge - 1)
        ImageDraw.Draw(mask).rounded_rectangle(bounds, radius=max(0, radius - edge), fill=255)
        if surface != "workspace" and width > 12 and height > 12:
            # One native outer border; only a quiet top-edge white reflection.
            # No inset outline/shadow, which previously looked double-framed.
            ImageDraw.Draw(image).line(
                (edge + radius, edge + 1, width - edge - radius - 1, edge + 1),
                fill=(255, 255, 255, 110), width=1)
    image.putalpha(mask)
    return image


class GlassFrame(ctk.CTkFrame):
    """Portable frosted surface, position-aware and rendered only when visible."""

    def __init__(self, master, *, glass_surface="panel", **kwargs):
        self._glass_surface = glass_surface
        self._glass_job = None
        self._glass_photo = None
        self._glass_size = None
        self._glass_viewport = None
        super().__init__(master, **kwargs)
        self.bind("<Map>", lambda _event: self._queue_glass(), add="+")
        self.bind("<Configure>", lambda _event: self._queue_glass(), add="+")
        self.bind("<Unmap>", lambda _event: self._release_glass(), add="+")
        window = self.winfo_toplevel()
        if not hasattr(window, "_glass_panels"):
            window._glass_panels = weakref.WeakSet()
            def window_resized(event, owner=window):
                if event.widget is owner:
                    for panel in list(owner._glass_panels):
                        panel._queue_glass()
            window.bind("<Configure>", window_resized, add="+")
        window._glass_panels.add(self)
        ancestor = master
        while ancestor is not None:
            if isinstance(ancestor, ctk.CTkScrollableFrame):
                self._glass_viewport = ancestor._parent_canvas
                if not hasattr(ancestor, "_glass_frames"):
                    ancestor._glass_frames = weakref.WeakSet()
                    def scroll_changed(first, last, owner=ancestor):
                        owner._scrollbar.set(first, last)
                        for panel in list(owner._glass_frames):
                            panel._queue_glass()
                    option = "yscrollcommand" if ancestor._orientation == "vertical" else "xscrollcommand"
                    ancestor._parent_canvas.configure(**{option: scroll_changed})
                ancestor._glass_frames.add(self)
                break
            ancestor = getattr(ancestor, "master", None)

    def _queue_glass(self):
        if not self.winfo_exists():
            return
        if self._glass_job is not None:
            self.after_cancel(self._glass_job)
        self._glass_job = self.after(80, self._paint_glass)

    def _release_glass(self):
        if self._glass_job is not None:
            self.after_cancel(self._glass_job)
            self._glass_job = None
        self._canvas.delete("glass_surface")
        self._glass_photo = None
        self._glass_size = None

    def _glass_visible(self):
        if not self.winfo_viewable():
            return False
        viewport = self._glass_viewport
        if viewport is None:
            return True
        x, y = self.winfo_rootx(), self.winfo_rooty()
        vx, vy = viewport.winfo_rootx(), viewport.winfo_rooty()
        return (x < vx + viewport.winfo_width() and x + self.winfo_width() > vx
                and y < vy + viewport.winfo_height() and y + self.winfo_height() > vy)

    def _glass_region(self, width, height):
        window = self.winfo_toplevel()
        ww, wh = max(1, window.winfo_width()), max(1, window.winfo_height())
        x, y = self.winfo_rootx() - window.winfo_rootx(), self.winfo_rooty() - window.winfo_rooty()
        left, top = max(0, min(.999, x / ww)), max(0, min(.999, y / wh))
        right, bottom = max(left + .001, min(1, (x + width) / ww)), max(top + .001, min(1, (y + height) / wh))
        return tuple(round(value, 4) for value in (left, top, right, bottom))

    def _draw(self, no_color_updates=False):
        super()._draw(no_color_updates)
        self._queue_glass()

    def _paint_glass(self, force=False):
        if self._glass_job is not None:
            self.after_cancel(self._glass_job)
        self._glass_job = None
        if not self.winfo_exists():
            return
        if not force and not self._glass_visible():
            self._release_glass()
            return
        width = max(1, round(self._apply_widget_scaling(self._current_width)))
        height = max(1, round(self._apply_widget_scaling(self._current_height)))
        radius = round(self._apply_widget_scaling(self._corner_radius))
        edge = max(0, round(self._apply_widget_scaling(self._border_width)))
        size = (width, height, radius, edge)
        region = self._glass_region(width, height)
        color = self._apply_appearance_mode(self._fg_color)
        tint = color if color not in {BG, CARD, SURFACE, SIDEBAR_SURFACE, "transparent"} else None
        if tint:
            # Tk accepts named gray levels (e.g. gray86) that Pillow does not.
            tint = tuple(channel // 257 for channel in self.winfo_rgb(tint))
        key = (self._glass_surface, size, region, tint)
        if key != self._glass_size:
            root = self._root()
            if not hasattr(root, "_glass_image_cache"):
                root._glass_image_cache = weakref.WeakValueDictionary()
            photo = root._glass_image_cache.get(key)
            if photo is None:
                image = glass_panel_image(size, self._glass_surface, region, tint)
                photo = ImageTk.PhotoImage(image, master=root)
                if len(root._glass_image_cache) >= 64:
                    root._glass_image_cache.pop(next(iter(root._glass_image_cache)), None)
                root._glass_image_cache[key] = photo
            self._glass_photo = photo
            self._glass_size = key
        self._canvas.delete("glass_surface")
        self._canvas.create_image(0, 0, anchor="nw", image=self._glass_photo,
                                  tags="glass_surface")

    def destroy(self):
        if self._glass_job is not None:
            self.after_cancel(self._glass_job)
            self._glass_job = None
        super().destroy()
        self._glass_photo = None


BIDI_DISPLAY_CONTROLS = {
    ord("\u200e"): None,  # LRM
    ord("\u200f"): None,  # RLM
    ord("\u202a"): None,  # LRE
    ord("\u202b"): None,  # RLE
    ord("\u202c"): None,  # PDF
    ord("\u2066"): None,  # LRI
    ord("\u2067"): None,  # RLI
    ord("\u2068"): None,  # FSI
    ord("\u2069"): None,  # PDI
}


def strip_bidi_display_controls(value):
    """Remove visual-only direction markers before storing application data."""
    return str(value or "").translate(BIDI_DISPLAY_CONTROLS)


def first_strong_direction(value):
    """Return rtl/ltr from the first strong Unicode directional character."""
    for char in strip_bidi_display_controls(value):
        direction = unicodedata.bidirectional(char)
        if direction in {"R", "AL"}:
            return "rtl"
        if direction == "L":
            return "ltr"
    return "ltr"


def directional_display_text(value):
    """Give Arabic-first editable text an RTL base without reversing it."""
    logical = strip_bidi_display_controls(value)
    if logical and first_strong_direction(logical) == "rtl":
        return f"\u202b{logical}\u202c"
    return logical


class DirectionalTextBinding:
    """Keep editable bidi presentation separate from the logical saved value."""

    def __init__(self, owner, model_var):
        self.owner = owner
        self.model_var = model_var
        self.display_var = tk.StringVar(
            master=model_var._root, value=directional_display_text(model_var.get()))
        self.widget = None
        self._syncing = False
        self._model_trace = model_var.trace_add("write", self._from_model)
        self._display_trace = self.display_var.trace_add("write", self._from_display)

    def attach(self, widget):
        self.widget = widget
        self._set_left_justified()
        return widget

    def _entry_widget(self):
        return getattr(self.widget, "_entry", self.widget)

    def _set_left_justified(self):
        entry = self._entry_widget()
        if entry is not None:
            try:
                entry.configure(justify="left")
            except (AttributeError, tk.TclError):
                pass

    def _cursor_logical_index(self, displayed):
        entry = self._entry_widget()
        if entry is None:
            return None
        try:
            cursor = int(entry.index(tk.INSERT))
        except (AttributeError, tk.TclError, ValueError):
            return None
        return len(strip_bidi_display_controls(displayed[:cursor]))

    def _restore_cursor(self, logical_index):
        if logical_index is None:
            return
        entry = self._entry_widget()
        if entry is None:
            return
        displayed = self.display_var.get()
        prefix = 1 if displayed.startswith("\u202b") else 0
        maximum = len(displayed) - (1 if displayed.endswith("\u202c") else 0)
        position = min(prefix + logical_index, maximum)
        try:
            entry.icursor(position)
        except (AttributeError, tk.TclError):
            pass

    def _from_model(self, *_args):
        if self._syncing:
            return
        self._syncing = True
        try:
            logical = strip_bidi_display_controls(self.model_var.get())
            if self.model_var.get() != logical:
                self.model_var.set(logical)
            displayed = directional_display_text(logical)
            if self.display_var.get() != displayed:
                self.display_var.set(displayed)
        finally:
            self._syncing = False

    def _from_display(self, *_args):
        if self._syncing:
            return
        displayed = self.display_var.get()
        cursor = self._cursor_logical_index(displayed)
        logical = strip_bidi_display_controls(displayed)
        normalized = directional_display_text(logical)
        self._syncing = True
        try:
            if self.model_var.get() != logical:
                self.model_var.set(logical)
            if displayed != normalized:
                self.display_var.set(normalized)
        finally:
            self._syncing = False
        if displayed != normalized and self.widget is not None:
            self.owner.after_idle(lambda: self._restore_cursor(cursor))

    def dispose(self):
        """Detach variable traces when a dynamically rebuilt editor is removed."""
        try:
            self.model_var.trace_remove("write", self._model_trace)
        except (tk.TclError, ValueError):
            pass
        try:
            self.display_var.trace_remove("write", self._display_trace)
        except (tk.TclError, ValueError):
            pass
        self.widget = None

# Common prescription frequencies.  The editable combo box keeps the list
# convenient while still allowing a clinician to enter a non-standard schedule.
FREQUENCY_OPTIONS = (
    "1x1 (OD / QD)",
    "1x2 (BID)",
    "1x3 (TID)",
    "1x4 (QID)",
    "كل 4 ساعات (Q4H)",
    "كل 6 ساعات (Q6H)",
    "كل 8 ساعات (Q8H)",
    "كل 12 ساعة (Q12H)",
    "عند الحاجة (PRN)",
    "عند الحاجة كل 4 إلى 6 ساعات (PRN q4-6h)",
    "عند الحاجة كل 8 ساعات (PRN q8h)",
    "فوراً / جرعة واحدة (STAT)",
    "1x1 يوم بعد يوم (QOD)",
    "مرة واحدة أسبوعياً (1x/week)",
    "مرتان أسبوعياً (2x/week)",
)

NOTE_OPTIONS = (
    "صباحاً (QAM)",
    "مساءً (QPM)",
    "عند النوم (QHS)",
    "قبل الطعام (AC)",
    "بعد الطعام (PC)",
    "مع الطعام",
)

# The prescriber specialty picker is intentionally editable: these common
# choices speed up entry, while existing profiles and less-common specialties
# remain valid custom text. The selected display label is saved exactly as
# shown so PDF, Word, and QR output continue to use the clinician's wording.
MEDICAL_SPECIALTIES = (
    ("General Practice", "الطب العام"),
    ("Family Medicine", "طب الأسرة"),
    ("Internal Medicine", "الطب الباطني"),
    ("Cardiology", "طب القلب"),
    ("Endocrinology", "طب الغدد الصماء"),
    ("Gastroenterology", "أمراض الجهاز الهضمي"),
    ("Nephrology", "أمراض الكلى"),
    ("Pulmonology", "الأمراض الصدرية"),
    ("Neurology", "طب الأعصاب"),
    ("Psychiatry", "الطب النفسي"),
    ("Pediatrics", "طب الأطفال"),
    ("Obstetrics and Gynecology", "النسائية والتوليد"),
    ("General Surgery", "الجراحة العامة"),
    ("Orthopedics", "جراحة العظام"),
    ("Rheumatology", "أمراض الروماتيزم"),
    ("Dermatology", "الأمراض الجلدية"),
    ("Ophthalmology", "طب العيون"),
    ("ENT", "الأنف والأذن والحنجرة"),
    ("Urology", "المسالك البولية"),
    ("Anesthesiology", "التخدير"),
    ("Emergency Medicine", "طب الطوارئ"),
    ("Infectious Diseases", "الأمراض المعدية"),
    ("Hematology", "أمراض الدم"),
    ("Oncology", "طب الأورام"),
)


def medical_specialty_options(lang=None):
    """Return the standard specialty labels for the active UI language."""
    language = lang or I.get_lang()
    label_index = 1 if language == "ar" else 0
    return [labels[label_index] for labels in MEDICAL_SPECIALTIES]


@functools.lru_cache(maxsize=64)
def _glyph_icon(symbol, color=ICON_BLUE, display_size=25):
    """Render a crisp monochrome navigation glyph in the shared icon color."""
    canvas_size = 64
    image = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    try:
        font_path = Path(os.environ.get("WINDIR", "C:\\Windows")) / "Fonts" / "seguisym.ttf"
        font = ImageFont.truetype(str(font_path), 43)
    except (OSError, ValueError):
        font = ImageFont.load_default()
    bounds = draw.textbbox((0, 0), symbol, font=font)
    x = (canvas_size - (bounds[2] - bounds[0])) / 2 - bounds[0]
    y = (canvas_size - (bounds[3] - bounds[1])) / 2 - bounds[1]
    draw.text((x, y), symbol, font=font, fill=color)
    return ctk.CTkImage(light_image=image, dark_image=image,
                        size=(display_size, display_size))


@functools.lru_cache(maxsize=64)
def _dashboard_icon_artwork(key, color=ICON_BLUE):
    """Draw a font-independent, consistently weighted line-icon family."""
    scale = 3
    image = Image.new("RGBA", (96 * scale, 96 * scale), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    stroke = 6 * scale

    def line(points):
        scaled = [(round(x * scale), round(y * scale)) for x, y in points]
        draw.line(scaled,
                  fill=color, width=stroke, joint="curve")
        for x, y in (scaled[0], scaled[-1]):
            radius = stroke / 2
            draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=color)

    def ellipse(box):
        draw.ellipse(tuple(round(v * scale) for v in box), outline=color, width=stroke)

    def rectangle(box, radius=7):
        draw.rounded_rectangle(tuple(round(v * scale) for v in box),
                               radius=radius * scale, outline=color, width=stroke)

    if key == "prescriber":
        rectangle((18, 22, 78, 82))
        rectangle((36, 12, 60, 30), 5)
        line([(48, 43), (48, 67)])
        line([(36, 55), (60, 55)])
    elif key == "patient":
        ellipse((34, 13, 62, 41))
        draw.arc((18 * scale, 48 * scale, 78 * scale, 102 * scale),
                 180, 360, fill=color, width=stroke)
        line([(21, 75), (21, 81), (75, 81), (75, 75)])
    elif key == "medications":
        rectangle((10, 32, 86, 64), 16)
        line([(48, 34), (48, 62)])
    elif key == "favorites":
        points = []
        for index in range(10):
            angle = -math.pi / 2 + index * math.pi / 5
            radius = 36 if index % 2 == 0 else 16
            points.append((48 + radius * math.cos(angle), 48 + radius * math.sin(angle)))
        line(points + points[:1])
    elif key == "drug_classes":
        rectangle((35, 12, 61, 34), 4)
        line([(48, 34), (48, 49), (22, 49), (22, 62)])
        line([(48, 49), (74, 49), (74, 62)])
        rectangle((10, 62, 35, 83), 4)
        rectangle((61, 62, 86, 83), 4)
    elif key == "treatment_templates":
        rectangle((20, 20, 76, 84))
        rectangle((35, 12, 61, 30), 5)
        for y in (43, 58, 73):
            line([(29, y - 1), (33, y + 3), (39, y - 4)])
            line([(47, y), (65, y)])
    elif key == "interaction_review":
        line([(48, 12), (79, 24), (77, 53), (69, 67), (48, 84),
              (27, 67), (19, 53), (17, 24), (48, 12)])
        line([(48, 32), (48, 53)])
        ellipse((45, 63, 51, 69))
    elif key == "reference":
        line([(46, 25), (34, 18), (12, 18), (12, 71), (34, 71),
              (46, 78), (46, 25), (58, 18), (80, 18), (80, 43)])
        ellipse((55, 48, 79, 72))
        line([(76, 69), (87, 81)])
    elif key in ("settings", "general"):
        points = []
        for index in range(48):
            angle = index * math.tau / 48
            radius = 36 if index % 6 in (1, 2, 3, 4) else 28
            points.append((48 + radius * math.cos(angle), 48 + radius * math.sin(angle)))
        line(points + points[:1])
        ellipse((35, 35, 61, 61))
    elif key == "clinic":
        rectangle((15, 38, 81, 83), 3)
        line([(12, 38), (22, 26), (74, 26), (84, 38)])
        rectangle((35, 10, 61, 35), 3)
        line([(48, 17), (48, 28)])
        line([(42, 23), (54, 23)])
        rectangle((39, 58, 57, 83), 2)
        rectangle((23, 48, 30, 62), 1)
        rectangle((66, 48, 73, 62), 1)
    elif key == "documents":
        line([(21, 12), (59, 12), (77, 30), (77, 84), (21, 84), (21, 12)])
        line([(59, 12), (59, 30), (77, 30)])
        for y in (45, 59, 73):
            line([(33, y), (65, y)])
    elif key == "qr":
        for x, y in ((12, 12), (59, 12), (12, 59)):
            rectangle((x, y, x + 25, y + 25), 2)
            rectangle((x + 10, y + 10, x + 15, y + 15), 1)
        line([(59, 59), (71, 59), (71, 71), (84, 71), (84, 84)])
        line([(59, 74), (59, 84), (69, 84)])
        line([(84, 51), (84, 59)])
    elif key == "database":
        ellipse((18, 12, 78, 34))
        line([(18, 23), (18, 72)])
        line([(78, 23), (78, 72)])
        for y in (37, 61):
            draw.arc(tuple(round(v * scale) for v in (18, y, 78, y + 22)),
                     0, 180, fill=color, width=stroke)
    elif key == "gemini":
        line([(48, 12), (59, 37), (84, 48), (59, 59), (48, 84),
              (37, 59), (12, 48), (37, 37), (48, 12)])
    elif key == "security":
        line([(48, 12), (79, 24), (77, 53), (69, 67), (48, 84),
              (27, 67), (19, 53), (17, 24), (48, 12)])
        draw.arc((37 * scale, 30 * scale, 59 * scale, 56 * scale),
                 180, 360, fill=color, width=stroke)
        rectangle((33, 44, 63, 66), 4)
        line([(48, 52), (48, 58)])
    elif key == "recovery":
        draw.arc((17 * scale, 17 * scale, 81 * scale, 81 * scale),
                 215, 535, fill=color, width=stroke)
        line([(17, 15), (17, 34), (36, 34)])
        line([(48, 31), (48, 49), (61, 57)])
    elif key == "about":
        ellipse((12, 12, 84, 84))
        ellipse((45, 28, 51, 34))
        line([(48, 46), (48, 67)])
    else:
        raise ValueError(f"Unknown dashboard icon: {key}")
    return image.resize((96, 96), Image.Resampling.LANCZOS)


@functools.lru_cache(maxsize=64)
def _dashboard_icon(key, color=ICON_BLUE):
    artwork = _dashboard_icon_artwork(key, color)
    return ctk.CTkImage(light_image=artwork, dark_image=artwork,
                        size=(DASHBOARD_ICON_SIZE, DASHBOARD_ICON_SIZE))


@functools.lru_cache(maxsize=48)
def _ui_font(size=12, weight="normal", family="Segoe UI"):
    """Reuse immutable UI font objects instead of allocating them per result row."""
    return ctk.CTkFont(family=family, size=size, weight=weight)


def dropdown_font(baseline=None, master=None):
    """Default preserves each control's original type; an override is global."""
    baseline = baseline or _ui_font(13)
    ancestor = master
    while ancestor is not None:
        if getattr(ancestor, "_fixed_dropdown_fonts", False):
            return baseline
        ancestor = getattr(ancestor, "master", None)
    size = cfg.config.ui_font_size("dropdown_font_size")
    if not size:
        return baseline
    if isinstance(baseline, ctk.CTkFont):
        return _ui_font(size, baseline.cget("weight"), baseline.cget("family"))
    return _ui_font(size)


class PopupListbox(tk.Listbox):
    def __init__(self, master, **kwargs):
        self._dropdown_font_baseline = kwargs.get("font", ("Segoe UI", 16))
        kwargs["font"] = dropdown_font(self._dropdown_font_baseline, master)
        super().__init__(master, **kwargs)

    def apply_preferences(self):
        self.configure(font=dropdown_font(self._dropdown_font_baseline, self.master))


class VisualMenu(tk.Menu):
    def __init__(self, master, **kwargs):
        self._dropdown_font_baseline = kwargs.get("font", ("Segoe UI", 16))
        kwargs["font"] = dropdown_font(self._dropdown_font_baseline, master)
        super().__init__(master, **kwargs)

    def apply_preferences(self):
        self.configure(font=dropdown_font(self._dropdown_font_baseline, self.master))


class VisualOptionMenu(ctk.CTkOptionMenu):
    def __init__(self, master, **kwargs):
        self._dropdown_font_baseline = kwargs.get("dropdown_font") or _ui_font(13)
        kwargs["dropdown_font"] = dropdown_font(self._dropdown_font_baseline, master)
        super().__init__(master, **kwargs)

    def apply_preferences(self):
        self.configure(dropdown_font=dropdown_font(self._dropdown_font_baseline, self.master))


def attach_search_hint(entry, variable, text):
    """CTk hides native placeholders with StringVar; keep hints out of values."""
    hint = ctk.CTkLabel(entry, text=text, anchor="w", text_color=MUTED,
                        fg_color="transparent", font=_ui_font(13), height=22)
    entry._search_hint = hint

    def refresh(*_args):
        if not entry.winfo_exists():
            return
        if not variable.get() and entry.focus_get() != entry._entry:
            hint.place(x=10, rely=.5, anchor="w")
        else:
            hint.place_forget()

    def focus(_event):
        hint.place_forget()
        entry.focus_set()

    trace_id = variable.trace_add("write", refresh)
    def cleanup(_event):
        try:
            variable.trace_remove("write", trace_id)
        except tk.TclError:
            pass

    hint.bind("<Button-1>", focus)
    entry.bind("<FocusIn>", refresh, add="+")
    entry.bind("<FocusOut>", refresh, add="+")
    entry.bind("<Destroy>", cleanup, add="+")
    refresh()


@functools.lru_cache(maxsize=8)
def _edit_icon(display_size=18):
    """Use the supplied transparent edit artwork consistently on every page."""
    with Image.open(cfg.PROJECT_DIR / "data" / "edit-icon.png") as source:
        artwork = source.convert("RGBA")
    # Preserve the supplied silhouette/transparency; tint with the shared blue.
    alpha = artwork.getchannel("A")
    artwork = Image.new("RGBA", artwork.size, ACCENT)
    artwork.putalpha(alpha)
    return ctk.CTkImage(light_image=artwork, dark_image=artwork,
                        size=(display_size, display_size))


def autocomplete_layout(anchor_rect, screen_rect, requested_width, row_height,
                        count, max_rows=10, align_anchor=False, extra_height=0,
                        cap_width=True):
    """Fit results above/below their field, reserving no unused result rows."""
    ax, ay, aw, ah = anchor_rect
    sx, sy, sw, sh = screen_rect
    inset, border = 10, 8 + extra_height
    bottom = sy + sh - inset
    below = max(0, bottom - (ay + ah + 2))
    above = max(0, ay - 2 - (sy + inset))
    room = max(below, above)
    rows = max(1, min(count, max_rows, int(max(0, room - border) // row_height)))
    height = min(rows * row_height + border, sh - inset * 2)
    width = max(aw, requested_width)
    x = ax
    if cap_width:
        width = min(width, sw - inset * 2)
        x = max(sx + inset, min(ax, sx + sw - width - inset))
        if align_anchor:
            x = max(sx + inset, min(ax, sx + sw - inset - 1))
            width = min(width, sx + sw - inset - x)
    y = ay + ah + 2 if below >= height else ay - height - 2
    y = max(sy + inset, min(y, bottom - height))
    return int(width), int(height), int(x), int(y), rows


def fit_autocomplete_popup(top, anchor, box, count, max_rows=10, align_anchor=False,
                           measure_content=False, width_multiplier=1, cap_width=True):
    """Screen-aware, bordered native popup shared by medicine/quick searches."""
    top.update_idletasks()
    font = tkfont.Font(root=top, font=box.cget("font"))
    row_height = font.metrics("linespace") + 2
    requested_width = box.winfo_reqwidth()
    content_width = 0
    if measure_content:
        content_width = max((font.measure(item) for item in box.get(0, tk.END)), default=0)
        requested_width = max(requested_width, content_width + 36)
    requested_width = max(anchor.winfo_width(), requested_width) * width_multiplier
    layout = autocomplete_layout(
        (anchor.winfo_rootx(), anchor.winfo_rooty(), anchor.winfo_width(), anchor.winfo_height()),
        (0, 0, top.winfo_screenwidth(), top.winfo_screenheight()),
        requested_width, row_height, count, max_rows, align_anchor,
        cap_width=cap_width)
    width, height, x, y, rows = layout
    if measure_content and content_width + 36 > width:
        horizontal = tk.Scrollbar(top, orient="horizontal", command=box.xview)
        top.update_idletasks()
        layout = autocomplete_layout(
            (anchor.winfo_rootx(), anchor.winfo_rooty(), anchor.winfo_width(), anchor.winfo_height()),
            (0, 0, top.winfo_screenwidth(), top.winfo_screenheight()),
            requested_width, row_height, count, max_rows, align_anchor,
            extra_height=horizontal.winfo_reqheight() + 3, cap_width=cap_width)
        width, height, x, y, rows = layout
        horizontal.pack(side="bottom", fill="x", padx=3, pady=(0, 3))
        box.configure(xscrollcommand=horizontal.set)
    box.configure(height=rows)
    if count > rows:
        scrollbar = tk.Scrollbar(top, command=box.yview)
        scrollbar.pack(side="right", fill="y", padx=(0, 3), pady=3)
        box.configure(yscrollcommand=scrollbar.set)
    box.pack(side="left", fill="both", expand=True, padx=3, pady=3)
    top.configure(background=LINE)
    top.geometry(f"{width}x{height}+{x}+{y}")
    top.lift()


@functools.lru_cache(maxsize=32)
def action_icon(symbol, color=ACCENT):
    """Optically balanced desktop action artwork, independent of glyph fonts."""
    image = Image.new("RGBA", (96, 96))
    draw = ImageDraw.Draw(image)
    def line(points):
        draw.line([(round(x * 4), round(y * 4)) for x, y in points], fill=color, width=8, joint="curve")
    if symbol in {"+", "＋"}:
        line([(12, 4), (12, 20)])
        line([(4, 12), (20, 12)])
    elif symbol in {"↑", "↓"}:
        flip = (lambda y: 24 - y) if symbol == "↓" else (lambda y: y)
        line([(12, flip(20)), (12, flip(4))])
        line([(5, flip(11)), (12, flip(4)), (19, flip(11))])
    elif symbol == "🗑":
        line([(4, 6), (20, 6)])
        line([(9, 6), (9, 3), (15, 3), (15, 6)])
        line([(6, 6), (7, 21), (17, 21), (18, 6)])
        line([(10, 10), (10, 17)])
        line([(14, 10), (14, 17)])
    elif symbol == "search":
        draw.ellipse((12, 12, 66, 66), outline=color, width=8)
        line([(16, 16), (22, 22)])
    else:
        points = []
        for index in range(10):
            angle = -math.pi / 2 + index * math.pi / 5
            radius = 10 if index % 2 == 0 else 4.6
            points.append((48 + round(4 * radius * math.cos(angle)),
                           48 + round(4 * radius * math.sin(angle))))
        if symbol == "★":
            draw.polygon(points, fill=color)
        else:
            draw.line(points + points[:1], fill=color, width=8, joint="curve")
    return ctk.CTkImage(image.resize((48, 48), Image.Resampling.LANCZOS), size=(18, 18))


class VisualButton(ctk.CTkButton):
    """Keep action commands/text semantics while giving glyph actions one grid."""
    SYMBOLS = frozenset({"+", "＋", "↑", "↓", "🗑", "★", "☆"})

    def __init__(self, master, **kwargs):
        self._action_symbol = None
        symbol = kwargs.get("text")
        if symbol in self.SYMBOLS and kwargs.get("image") is None:
            self._action_symbol = symbol
            color = kwargs.get("text_color", ACCENT)
            if isinstance(color, (list, tuple)):
                color = color[0]
            kwargs.update(text="", image=action_icon(symbol, color),
                          width=max(30, kwargs.get("width", 30)),
                          height=max(30, kwargs.get("height", 30)))
        super().__init__(master, **kwargs)

    def configure(self, **kwargs):
        if "text" in kwargs:
            symbol = kwargs["text"]
            if symbol in self.SYMBOLS:
                self._action_symbol = symbol
                color = self._apply_appearance_mode(kwargs.get("text_color", self.cget("text_color")))
                kwargs.update(text="", image=action_icon(symbol, color))
            elif self._action_symbol is not None:
                self._action_symbol = None
                kwargs.setdefault("image", None)
        super().configure(**kwargs)

    def cget(self, attribute_name):
        if attribute_name == "text" and self._action_symbol is not None:
            return self._action_symbol
        return super().cget(attribute_name)


class FieldFocus:
    """Color-only focus feedback; never changes geometry or replaces bindings."""
    def _bind_field_focus(self):
        self._focus_border = None
        self.bind("<FocusIn>", self._field_focused, add="+")
        self.bind("<FocusOut>", self._field_blurred, add="+")

    def _field_focused(self, _event=None):
        color = self.cget("border_color")
        if color == LINE or color == [LINE, LINE]:
            self._focus_border = color
            self.configure(border_color=ACCENT)

    def _field_blurred(self, _event=None):
        if self._focus_border is not None and self.cget("border_color") == ACCENT:
            self.configure(border_color=self._focus_border)
        self._focus_border = None


class VisualEntry(FieldFocus, ctk.CTkEntry):
    def __init__(self, master, **kwargs):
        kwargs.setdefault("border_width", 1)
        super().__init__(master, **kwargs)
        self._bind_field_focus()


class VisualComboBox(FieldFocus, ctk.CTkComboBox):
    def __init__(self, master, **kwargs):
        kwargs.setdefault("border_width", 1)
        self._dropdown_font_baseline = kwargs.get("dropdown_font") or _ui_font(13)
        kwargs["dropdown_font"] = dropdown_font(self._dropdown_font_baseline, master)
        super().__init__(master, **kwargs)
        self._bind_field_focus()

    def apply_preferences(self):
        self.configure(dropdown_font=dropdown_font(self._dropdown_font_baseline, self.master))


def polish_toolbar(frame):
    """One compact rhythm for search and text actions, not icon hit areas."""
    for widget in frame.winfo_children():
        if isinstance(widget, (ctk.CTkEntry, ctk.CTkComboBox, ctk.CTkOptionMenu)):
            widget.configure(height=FIELD_HEIGHT, corner_radius=9)
        elif isinstance(widget, VisualButton) and widget._action_symbol is None and widget.cget("image") is None:
            widget.configure(height=FIELD_HEIGHT, corner_radius=9)


class DrugRow(GlassFrame):
    """One medication row with linked generic/trade and scientific-name inputs."""

    def __init__(self, master, db, on_change, on_remove,
                 on_move_up, on_move_down, on_drag, on_reference, **kwargs):
        super().__init__(master, **kwargs)
        self.db = db
        self.on_change = on_change
        self.on_remove = on_remove
        self.on_move_up = on_move_up
        self.on_move_down = on_move_down
        self.on_drag = on_drag
        self.on_reference = on_reference
        self._ac_top = None
        self._ac_listbox = None
        self._form = ""
        self._brand = ""
        self._category = ""
        self._bidi_bindings = []
        self._drag_start_y = 0
        self._drag_moved = False
        self._move_menu = None
        self._autocomplete_job = None
        self._autocomplete_token = 0

        self.name_var = tk.StringVar()
        self.trade_var = tk.StringVar()
        names = ctk.CTkFrame(self, fg_color="transparent")
        names.pack(fill="x", padx=PAD, pady=(6, CARD_GAP))
        names.grid_columnconfigure((1, 2), weight=1, uniform="medication_names")
        self.number_badge = ctk.CTkLabel(
            names, text="1.", width=28, height=36,
            fg_color="transparent", text_color=MUTED,
            font=ctk.CTkFont(weight="bold", size=12))
        self.number_badge.grid(row=0, column=0, sticky="n", padx=(0, 4), pady=(30, 0))
        header_actions = ctk.CTkFrame(names, fg_color="transparent")
        header_actions.grid(row=0, column=3, sticky="n", padx=(4, 0), pady=(33, 0))
        self.move_up_button = VisualButton(
            header_actions, text="↑", width=ICON_BUTTON_SIZE, height=ICON_BUTTON_SIZE, corner_radius=7,
            fg_color="transparent", text_color=ACCENT, border_width=0,
            hover_color=ACCENT_SOFT, font=ctk.CTkFont(size=18),
            command=lambda: self.on_move_up(self))
        self.move_up_button.pack(side="left", padx=2)
        self.move_down_button = VisualButton(
            header_actions, text="↓", width=ICON_BUTTON_SIZE, height=ICON_BUTTON_SIZE,
            corner_radius=7, fg_color="transparent", text_color=ACCENT,
            hover_color=ACCENT_SOFT, font=ctk.CTkFont(size=18),
            command=lambda: self.on_move_down(self))
        self.move_down_button.pack(side="left", padx=2)
        for button in (self.move_up_button, self.move_down_button):
            button.bind("<Button-3>", self._show_move_menu)
            button.bind("<Shift-F10>", self._show_move_menu)
        # The plain row number retains drag sorting and the keyboard/context menu.
        self.drag_handle = self.number_badge
        self.drag_handle.bind("<ButtonPress-1>", self._drag_start)
        self.drag_handle.bind("<B1-Motion>", self._drag_motion)
        self.drag_handle.bind("<ButtonRelease-1>", self._drag_end)
        self.drag_handle.bind("<Button-3>", self._show_move_menu)
        self.drag_handle.bind("<Shift-F10>", self._show_move_menu)
        VisualButton(
            header_actions, text="🗑", width=ICON_BUTTON_SIZE, height=ICON_BUTTON_SIZE, corner_radius=7,
            fg_color="transparent", text_color=DANGER, border_width=0,
            hover_color=DANGER_SOFT, font=ctk.CTkFont(size=14),
            command=self.on_remove).pack(side="left", padx=(2, 0))
        trade_col = ctk.CTkFrame(names, fg_color="transparent")
        trade_col.grid(row=0, column=1, sticky="ew", padx=(0, 4))
        science_col = ctk.CTkFrame(names, fg_color="transparent")
        science_col.grid(row=0, column=2, sticky="ew", padx=(4, 0))
        ctk.CTkLabel(trade_col, text=I.t("generic_trade_name"), anchor="w", text_color=MUTED,
                     font=LABEL_FONT).pack(fill="x", pady=(0, 2))
        self.trade_entry = VisualEntry(trade_col, textvariable=self.trade_var,
                                        fg_color=SURFACE, text_color=TEXT,
                                        placeholder_text=I.t("generic_trade_name"), height=36,
                                        corner_radius=8, border_color=LINE,
                                        font=BRAND_FONT)
        self.trade_entry.pack(fill="x")
        self.trade_entry.bind("<KeyRelease>", self._on_trade_type)
        science_heading = ctk.CTkFrame(science_col, fg_color="transparent")
        science_heading.pack(fill="x", pady=(0, 2))
        ctk.CTkLabel(science_heading, text=I.t("scientific_name"), anchor="w",
                     text_color=MUTED,
                     font=LABEL_FONT).pack(side="left")
        science_input_row = ctk.CTkFrame(science_col, fg_color="transparent")
        science_input_row.pack(fill="x")
        self.name_entry = VisualEntry(
            science_input_row, textvariable=self.name_var,
            fg_color=SURFACE, text_color=TEXT,
            placeholder_text=I.t("scientific_name"), height=36,
            corner_radius=8, border_color=LINE,
            font=SCIENTIFIC_FONT)
        self.name_entry.pack(side="left", fill="x", expand=True)
        self.reference_button = VisualButton(
            science_input_row, text="!", width=ICON_BUTTON_SIZE, height=ICON_BUTTON_SIZE, corner_radius=15,
            fg_color=WARNING_SOFT, text_color=WARNING, border_width=1,
            border_color=WARNING, hover_color=WARNING_SOFT,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=lambda: self.on_reference(self, False))
        # Each name box searches only its matching imported database column.
        self.name_entry.bind("<KeyRelease>", self._on_scientific_type)
        for entry in (self.trade_entry, self.name_entry):
            entry.bind("<Down>", self._ac_down)
            entry.bind("<Up>", self._ac_up)
            entry.bind("<Return>", self._ac_choose_current)
            entry.bind("<Escape>", lambda _event: self._hide_ac())
        self.trade_picker = VisualOptionMenu(
            trade_col, values=[I.t("choose_linked_trade_name")], height=28,
            fg_color=ACCENT_SOFT, text_color=ACCENT, button_color=PRIMARY,
            button_hover_color=ACCENT_HOVER, command=self._choose_linked_trade)

        self.class_badge = ctk.CTkLabel(
            self, text="", height=22, corner_radius=11, fg_color=ACCENT_SOFT,
            text_color=ACCENT, font=ctk.CTkFont(size=10, weight="bold"), anchor="w")

        details = ctk.CTkFrame(self, fg_color="transparent")
        self.details = details
        details.pack(fill="x", padx=PAD, pady=(0, 6))
        # Balanced 20% / 20% / 15% / 45% proportions for dosage, frequency,
        # duration and notes.
        for column, weight in enumerate((4, 4, 3, 8, 4)):
            details.grid_columnconfigure(column, weight=weight)

        self.dosage_var = tk.StringVar()
        self.freq_var = tk.StringVar()
        self.dur_var = tk.StringVar()
        self.notes_var = tk.StringVar()
        self.quantity_var = tk.StringVar()
        self._quantity_manual = False
        self.dosage_entry = self._box(details, 0, I.t("dosage"),
                                      I.t("ph_dosage"), self.dosage_var)
        self.freq_entry = self._frequency_box(details, 1, I.t("frequency"),
                                              self.freq_var)
        self.dur_entry = self._box(details, 2, I.t("duration"),
                                   I.t("ph_duration"), self.dur_var)
        self.notes_entry = self._notes_box(details, 3, I.t("notes"),
                                           self.notes_var)
        quantity_col = ctk.CTkFrame(details, fg_color="transparent")
        quantity_col.grid(row=0, column=4, sticky="ew", padx=4)
        quantity_head = ctk.CTkFrame(quantity_col, fg_color="transparent")
        quantity_head.pack(fill="x", pady=(0, 2))
        ctk.CTkLabel(quantity_head, text=I.t("quantity"), anchor="w",
                     text_color=MUTED, font=LABEL_FONT).pack(
                         side="left", fill="x", expand=True)
        VisualButton(quantity_head, text="↻", width=25, height=22,
                      fg_color="transparent", text_color=ACCENT,
                      hover_color=ACCENT_SOFT, command=self.recalculate_quantity).pack(side="right")
        self.quantity_entry = VisualEntry(
            quantity_col, textvariable=self.quantity_var, height=FIELD_HEIGHT,
            corner_radius=9, border_color=LINE, placeholder_text=I.t("quantity_auto"),
            font=ctk.CTkFont(size=11), justify="left")
        self.quantity_entry.pack(fill="x")
        self.quantity_entry.bind("<KeyRelease>", self._quantity_edited)
        for variable in (self.dosage_var, self.freq_var, self.dur_var):
            variable.trace_add("write", lambda *_: self._auto_quantity())

        self._reference_drug = ""
        self._reference_display = ""
        self.reference_panel = ctk.CTkFrame(
            self, fg_color=WARNING_SOFT, border_color=LINE,
            border_width=1, corner_radius=10)
        reference_head = ctk.CTkFrame(self.reference_panel, fg_color="transparent")
        reference_head.pack(fill="x", padx=10, pady=(8, 4))
        ctk.CTkLabel(
            reference_head, text=I.t("gemini_reference_title"), text_color=WARNING,
            font=ctk.CTkFont(size=12, weight="bold"), anchor="w").pack(
                side="left", fill="x", expand=True)
        VisualButton(
            reference_head, text=I.t("gemini_refresh"), width=72, height=28,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT,
            command=lambda: self.on_reference(self, True)).pack(side="right", padx=2)
        VisualButton(
            reference_head, text=I.t("gemini_copy"), width=58, height=28,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=self.copy_reference).pack(side="right", padx=2)
        VisualButton(
            reference_head, text="×", width=30, height=28,
            fg_color="transparent", text_color=MUTED, hover_color=ACCENT_SOFT,
            command=self.close_reference).pack(side="right", padx=2)
        self.reference_status = ctk.CTkLabel(
            self.reference_panel, text="", text_color=MUTED,
            font=ctk.CTkFont(size=10), anchor="w")
        self.reference_status.pack(fill="x", padx=10, pady=(0, 4))
        self.reference_text = ctk.CTkTextbox(
            self.reference_panel, height=250, corner_radius=8, wrap="word",
            fg_color=CARD, border_width=1, border_color=LINE,
            font=ctk.CTkFont(size=13))
        self.reference_text.pack(fill="x", padx=10, pady=(0, 10))
        self.reference_text.configure(state="disabled")
        self.name_var.trace_add("write", self._scientific_name_changed)
        self._sync_reference_button()

    def _drag_start(self, event):
        self._drag_start_y = event.y_root
        self._drag_moved = False
        self.on_drag(self, "start", event.y_root)

    def _drag_motion(self, event):
        if abs(event.y_root - self._drag_start_y) >= 4:
            self._drag_moved = True
            self.on_drag(self, "move", event.y_root)

    def _drag_end(self, event):
        self.on_drag(self, "end", event.y_root)
        self.after_idle(lambda: setattr(self, "_drag_moved", False))

    def _show_move_menu(self, event=None):
        """Show reliable, large text-only reorder actions."""
        if self._drag_moved:
            return "break"
        self._close_move_menu()
        menu = VisualMenu(
            self, tearoff=False, font=("Segoe UI", 16, "bold"),
            background=CARD, foreground=ACCENT,
            activebackground=ACCENT_SOFT, activeforeground=ACCENT,
            borderwidth=0, activeborderwidth=0, relief="flat")
        self._move_menu = menu
        menu.add_command(
            label=f"  {I.t('move_up')}  ", command=lambda: self.on_move_up(self))
        menu.add_command(
            label=f"  {I.t('move_down')}  ", command=lambda: self.on_move_down(self))
        x = (event.x_root if event is not None
             else self.drag_handle.winfo_rootx() - 120)
        y = (event.y_root if event is not None
             else self.drag_handle.winfo_rooty() + self.drag_handle.winfo_height() + 2)
        try:
            menu.tk_popup(max(0, x), max(0, y))
        finally:
            menu.grab_release()
        return "break"

    def _close_move_menu(self):
        menu = self._move_menu
        self._move_menu = None
        if menu is not None and menu.winfo_exists():
            menu.destroy()

    def _scientific_name_changed(self, *_args):
        self._form = ""
        current = self.name_var.get().strip()
        self._sync_reference_button()
        if self._reference_drug and current.casefold() != self._reference_drug.casefold():
            self.close_reference()

    def _sync_reference_button(self):
        if self.name_var.get().strip():
            if not self.reference_button.winfo_manager():
                self.reference_button.pack(side="right", padx=(4, 0))
        elif self.reference_button.winfo_manager():
            self.reference_button.pack_forget()

    def _show_reference_panel(self):
        if not self.reference_panel.winfo_manager():
            self.reference_panel.pack(fill="x", padx=PAD, pady=(0, 8), before=self.details)

    def _set_reference_text(self, value):
        self._reference_display = value
        self.reference_text.configure(state="normal")
        self.reference_text.delete("1.0", tk.END)
        self.reference_text.insert("1.0", value)
        self.reference_text.configure(state="disabled")

    def show_reference_loading(self, drug_name):
        self._reference_drug = drug_name
        self._show_reference_panel()
        self.reference_button.configure(state="disabled")
        self.reference_status.configure(text=I.t("gemini_loading"), text_color=WARNING)
        self._set_reference_text("")

    def show_reference_result(self, result):
        if not result.grounded:
            mode = I.t("gemini_ungrounded")
            status_color = WARNING
        else:
            mode = I.t("gemini_cached") if result.cache_hit else I.t("gemini_live")
            status_color = ACCENT
        self.reference_status.configure(
            text=f"{mode} · {I.t('gemini_checked', date=result.checked_at)}",
            text_color=status_color)
        text = result.text
        if not result.grounded:
            text = I.t("gemini_ungrounded_warning") + "\n\n" + text
        if result.sources:
            text += "\n\n" + I.t("gemini_sources")
            for number, (title, url) in enumerate(result.sources, 1):
                text += f"\n{number}. {title}\n   {url}"
        elif result.grounded:
            text += "\n\n" + I.t("gemini_no_sources")
        self._set_reference_text(text)
        self.reference_button.configure(state="normal")

    def show_reference_error(self, detail):
        self._show_reference_panel()
        self.reference_status.configure(text=I.t("gemini_error", detail=detail),
                                        text_color=DANGER)
        self._set_reference_text("")
        self.reference_button.configure(state="normal")

    def copy_reference(self):
        if not self._reference_display:
            return
        self.clipboard_clear()
        self.clipboard_append(self._reference_display)
        messagebox.showinfo(APP_TITLE, I.t("gemini_copied"))

    def close_reference(self):
        self.reference_panel.pack_forget()
        self._reference_drug = ""

    def _box(self, parent, column, label, ph, var):
        col = ctk.CTkFrame(parent, fg_color="transparent")
        col.grid(row=0, column=column, sticky="ew", padx=4)
        ctk.CTkLabel(col, text=label, anchor="w", text_color=MUTED,
                     font=LABEL_FONT).pack(fill="x", pady=(0, 2))
        binding = DirectionalTextBinding(self, var)
        self._bidi_bindings.append(binding)
        e = VisualEntry(col, textvariable=binding.display_var,
                         height=FIELD_HEIGHT, corner_radius=9,
                         border_color=LINE, placeholder_text=ph,
                         font=ctk.CTkFont(size=11), justify="left")
        binding.attach(e)
        e.pack(fill="x")
        e.bind("<KeyRelease>", lambda ev: self.on_change())
        return e

    def _frequency_box(self, parent, column, label, var):
        """An editable clinical-frequency picker, matching the other fields."""
        col = ctk.CTkFrame(parent, fg_color="transparent")
        col.grid(row=0, column=column, sticky="ew", padx=4)
        ctk.CTkLabel(col, text=label, anchor="w", text_color=MUTED,
                     font=LABEL_FONT).pack(
                         fill="x", pady=(0, 2))
        binding = DirectionalTextBinding(self, var)
        self._bidi_bindings.append(binding)
        picker = VisualComboBox(
            col, values=[directional_display_text(value) for value in FREQUENCY_OPTIONS],
            variable=binding.display_var, height=FIELD_HEIGHT,
            corner_radius=9, border_width=1, border_color=LINE, fg_color=SURFACE,
            button_color=ACCENT_SOFT, button_hover_color=LINE,
            dropdown_fg_color=SURFACE, dropdown_hover_color=ACCENT_SOFT,
            font=ctk.CTkFont(size=13), dropdown_font=_ui_font(16),
            justify="left",
            command=lambda _value: self.on_change())
        binding.attach(picker)
        picker.pack(fill="x")
        picker.bind("<KeyRelease>", lambda ev: self.on_change())
        return picker

    def _notes_box(self, parent, column, label, var):
        """An editable notes picker with common administration instructions."""
        col = ctk.CTkFrame(parent, fg_color="transparent")
        col.grid(row=0, column=column, sticky="ew", padx=4)
        ctk.CTkLabel(col, text=label, anchor="w", text_color=MUTED,
                     font=LABEL_FONT).pack(
                         fill="x", pady=(0, 2))
        binding = DirectionalTextBinding(self, var)
        self._bidi_bindings.append(binding)
        picker = VisualComboBox(
            col, values=[directional_display_text(value) for value in NOTE_OPTIONS],
            variable=binding.display_var, height=FIELD_HEIGHT,
            corner_radius=9, border_width=1, border_color=LINE, fg_color=SURFACE,
            button_color=ACCENT_SOFT, button_hover_color=LINE,
            dropdown_fg_color=SURFACE, dropdown_hover_color=ACCENT_SOFT,
            font=ctk.CTkFont(size=13), dropdown_font=_ui_font(16),
            justify="left",
            command=lambda _value: self.on_change())
        binding.attach(picker)
        picker.pack(fill="x")
        picker.bind("<KeyRelease>", lambda ev: self.on_change())
        return picker

    def _quantity_edited(self, _event=None):
        self._quantity_manual = bool(self.quantity_var.get().strip())
        self.on_change()

    def _auto_quantity(self):
        if self._quantity_manual:
            return
        self.quantity_var.set(calculate_medicine_quantity(
            self.dosage_var.get(), self.freq_var.get(), self.dur_var.get(), self._form))

    def recalculate_quantity(self):
        self._quantity_manual = False
        self._auto_quantity()
        self.on_change()

    def set_position(self, number):
        """Keep the plain numbered name row accurate after reordering."""
        self.number_badge.configure(text=f"{number}.")

    # autocomplete (sticky, large)
    def _on_scientific_type(self, event=None):
        if event and event.keysym in {
                "Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R",
                "Control_L", "Control_R", "Alt_L", "Alt_R"}:
            return
        self.on_change()
        self._update_class_badge()
        self._hide_trade_picker()
        q = self.name_var.get().strip()
        if len(q) < 1:
            self._hide_ac()
            return
        self._schedule_autocomplete(q, "scientific")

    def _on_trade_type(self, event=None):
        # The deferred worker uses db.search_prescribable for this brand/trade field.
        if event and event.keysym in {
                "Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R",
                "Control_L", "Control_R", "Alt_L", "Alt_R"}:
            return
        self._form = ""
        self.on_change()
        self._hide_trade_picker()
        q = self.trade_var.get().strip()
        if len(q) < 1:
            self._hide_ac()
            return
        self._schedule_autocomplete(q, "trade")

    def _schedule_autocomplete(self, query, mode):
        """Debounce typing and perform the indexed query away from Tk's UI thread."""
        if self._autocomplete_job is not None:
            try:
                self.after_cancel(self._autocomplete_job)
            except (tk.TclError, ValueError):
                pass
        self._autocomplete_token += 1
        token = self._autocomplete_token
        self._autocomplete_job = self.after(
            180, lambda: self._run_autocomplete(query, mode, token))

    def _run_autocomplete(self, query, mode, token):
        self._autocomplete_job = None
        app = self.winfo_toplevel()
        search = (self.db.search_scientific if mode == "scientific"
                  else self.db.search_prescribable)

        def finished(matches):
            current = (self.name_var.get().strip() if mode == "scientific"
                       else self.trade_var.get().strip())
            if token != self._autocomplete_token or current != query:
                return
            if matches:
                self._show_ac(matches, mode=mode)
            else:
                self._hide_ac()

        if hasattr(app, "submit_background"):
            app.submit_background(lambda: search(query, limit=20), finished,
                                  silent=True)
        else:
            finished(search(query, limit=20))

    def _show_ac(self, matches, mode):
        self._hide_ac()
        top = tk.Toplevel(self)
        top.wm_overrideredirect(True)
        anchor = self.name_entry if mode == "scientific" else self.trade_entry
        x = anchor.winfo_rootx()
        y = anchor.winfo_rooty() + anchor.winfo_height()
        top.geometry(f"+{x}+{y}")
        self._ac_top = top
        self._ac_matches = matches
        self._ac_mode = mode
        self._ac_index = -1
        lb = PopupListbox(top,
                        height=min(10, len(matches)),
                        font=LIST_FONT,
                        bg=CARD, fg=TEXT, relief="flat", borderwidth=0,
                        activestyle="none", exportselection=False,
                        highlightthickness=1, highlightbackground=LINE,
                        selectbackground=PRIMARY, selectforeground="white")
        max_chars = max((len(d.generic_name) + len(d.brand_name or "")
                         + len(d.strength or "") + 8) for d in matches)
        width = max(30, min(62, max_chars + 2))
        lb.configure(width=width)
        for d in matches:
            label = (d.generic_name if mode == "scientific"
                     else (d.brand_name or d.generic_name))
            linked = (d.brand_name if mode == "scientific"
                      else (d.generic_name if d.brand_name else ""))
            if linked:
                label += f"  ({linked})"
            if d.strength:
                label += f"  [{d.strength}]"
            lb.insert(tk.END, label)
        lb.bind("<<ListboxSelect>>", lambda e: self._ac_choose())
        fit_autocomplete_popup(top, anchor, lb, len(matches))
        self._ac_listbox = lb
        if lb.size():
            lb.selection_set(0)
            lb.see(0)
            self._ac_index = 0

    def _hide_ac(self, after=0):
        if after:
            self.after(after, self._real_hide)
        else:
            self._real_hide()

    def _real_hide(self):
        if getattr(self, "_ac_top", None):
            try:
                self._ac_top.destroy()
            except Exception:
                pass
            self._ac_top = None
            self._ac_listbox = None

    def _ac_down(self, event=None):
        if not self._ac_listbox:
            return
        self._ac_index += 1
        if self._ac_index >= self._ac_listbox.size():
            self._ac_index = 0
        self._ac_listbox.selection_clear(0, tk.END)
        self._ac_listbox.selection_set(self._ac_index)
        self._ac_listbox.see(self._ac_index)

    def _ac_up(self, event=None):
        if not self._ac_listbox:
            return
        self._ac_index -= 1
        if self._ac_index < 0:
            self._ac_index = self._ac_listbox.size() - 1
        self._ac_listbox.selection_clear(0, tk.END)
        self._ac_listbox.selection_set(self._ac_index)
        self._ac_listbox.see(self._ac_index)

    def _ac_choose_current(self, event=None):
        if self._ac_listbox and self._ac_index >= 0:
            self._ac_choose()
            return "break"
        return None

    def _ac_choose(self, event=None):
        if not getattr(self, "_ac_listbox", None):
            return
        idx = self._ac_listbox.curselection()
        if not idx:
            return
        d = self._ac_matches[idx[0]]
        mode = self._ac_mode
        if mode == "trade":
            self.trade_var.set(d.brand_name or d.generic_name)
            self.name_var.set(d.generic_name if d.brand_name else "")
        else:
            self.name_var.set(d.generic_name)
            self.trade_var.set(d.brand_name if d.brand_name else "")
        self._brand = self.trade_var.get()
        self._category = d.category
        self._form = d.form
        self._real_hide()
        self._update_class_badge(d)
        if mode == "scientific":
            self._show_trade_picker(d.generic_name, selected=d.brand_name)
        else:
            self._hide_trade_picker()
        self._auto_quantity()
        self.on_change()

    def _show_trade_picker(self, scientific_name, selected=""):
        choices = self.db.trade_names_for_scientific(scientific_name)
        if not choices:
            self._hide_trade_picker()
            return
        self.trade_picker.configure(values=choices)
        self.trade_picker.set(selected if selected in choices else choices[0])
        self.trade_var.set(self.trade_picker.get())
        if not self.trade_picker.winfo_manager():
            self.trade_picker.pack(fill="x", pady=(4, 0))

    def _hide_trade_picker(self):
        if self.trade_picker.winfo_manager():
            self.trade_picker.pack_forget()

    def _choose_linked_trade(self, value):
        self.trade_var.set(value)
        self._brand = value
        self.on_change()

    def _update_class_badge(self, drug=None):
        """Keep therapeutic-class metadata out of the prescription entry form."""
        # Classification is useful for browsing and mapping, but it is not a
        # prescription instruction.  Do not show it beneath the medicine name.
        self.class_badge.pack_forget()

    def get_data(self) -> qu.DrugItem:
        return qu.DrugItem(
            generic_name=self.name_var.get().strip(),
            brand_name=self.trade_var.get().strip(),
            dosage=self.dosage_var.get().strip(),
            frequency=self.freq_var.get().strip(),
            duration=self.dur_var.get().strip(),
            notes=self.notes_var.get().strip(),
            quantity=self.quantity_var.get().strip(),
        )


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1280x840")
        self.minsize(1040, 720)
        self.configure(fg_color=BG)
        I.set_lang(cfg.config.language)
        self.protocol("WM_DELETE_WINDOW", self.confirm_close)
        cfg.ensure_seed_db()
        self.db = dbmod.DrugDatabase(cfg.config.drug_db_path)
        try:
            cfg.config.maybe_create_automatic_backup()
        except Exception:
            logging.exception("Automatic backup could not be created")
        self.patient_history = PatientHistory()
        self.rows: List[DrugRow] = []
        self._medication_patient_id = ""
        self._executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=3, thread_name_prefix="rx-worker")
        self._background_tasks = 0
        self._background_callbacks = queue.SimpleQueue()
        self._closing = False
        self._export_busy = False
        self._document_lock = threading.Lock()
        self._loaded_pages = set()
        self._classification_cache = None
        self._latest_query_tokens = {}
        self._favorite_ac_job = None
        self._favorite_ac_token = 0
        self._favorite_render_limit = 60
        self._favorite_filter_signature = None
        self._word_preview_job = None
        self._comparison_job = None
        self._favorite_refresh_job = None
        self._class_search_job = None
        self.word_preview_visible = False
        self._build_ui()

        self.after(40, self._drain_background_callbacks)

    def destroy(self):
        self._closing = True
        if hasattr(self, "_executor"):
            self._executor.shutdown(wait=False, cancel_futures=True)
        super().destroy()

    def _drain_background_callbacks(self):
        if self._closing:
            return
        for _ in range(64):
            try:
                callback = self._background_callbacks.get_nowait()
            except queue.Empty:
                break
            if self._closing:
                return
            try:
                callback()
            except Exception as exc:
                self.report_callback_exception(type(exc), exc, exc.__traceback__)
        if not self._closing:
            self.after(40, self._drain_background_callbacks)

    def _build_ui(self):
        self.paper_var = tk.StringVar(value=cfg.config.paper_size)
        self.lang_var = tk.StringVar(value=cfg.config.language)
        self.profile_var = tk.StringVar(value=cfg.config.get("active_profile", "Default"))

        # Dashboard: persistent navigation on the left; one data-entry page at a time.
        self.workspace = GlassFrame(self, glass_surface="workspace", fg_color=BG,
                                     corner_radius=14)
        self.workspace.pack(fill="both", expand=True, padx=4, pady=(2, 4))
        self.dashboard = GlassFrame(self.workspace, glass_surface="sidebar",
                                      width=DASHBOARD_WIDTH, fg_color=SIDEBAR_SURFACE,
                                      border_color=LINE, border_width=1, corner_radius=12)
        self.dashboard.pack(side="left", fill="y", padx=(4, 6), pady=4)
        self.dashboard.pack_propagate(False)
        self.dashboard_collapsed = bool(cfg.config.get("dashboard_collapsed", False))
        dashboard_heading = ctk.CTkFrame(self.dashboard, fg_color="transparent")
        dashboard_heading.pack(fill="x", padx=8, pady=(8, 6))
        self.dashboard_title = ctk.CTkLabel(
            dashboard_heading, text=I.t("dashboard"), text_color=TEXT,
            font=ctk.CTkFont(weight="bold", size=16), anchor="w")
        self.dashboard_title.pack(side="left", fill="x", expand=True)
        self.dashboard_toggle = VisualButton(
            dashboard_heading, text="‹", width=ICON_BUTTON_SIZE, height=ICON_BUTTON_SIZE,
            fg_color="transparent", text_color=ACCENT, hover_color=ACCENT_SOFT,
            font=_ui_font(18, "bold"), command=self.toggle_dashboard)
        self.dashboard_toggle.pack(side="right")
        self.page_buttons = {}
        self._dashboard_button_labels = {}
        self._dashboard_action_buttons = {}
        self._page_button_icons = {}
        self._dashboard_action_icons = []
        self._add_page_button("prescriber", I.t("prescriber_details"))
        self._add_page_button("patient", I.t("patient_details"))
        self._add_page_button("medications", I.t("medication_entry"))
        self._add_page_button("favorites", I.t("favorite_drugs"))
        self._add_page_button("drug_classes", I.t("drug_classes"))
        self._add_page_button("treatment_templates", I.t("treatment_templates"))
        self._add_page_button("interaction_review", I.t("interaction_review"))
        self._add_page_button("reference", I.t("online_drug_reference"))
        self._add_dashboard_action("settings", I.t("settings"), self.open_settings)
        self.dashboard_footer = ctk.CTkLabel(self.dashboard,
                     text=I.t("local_encrypted", version=cfg.APP_VERSION),
                     text_color=MUTED, font=ctk.CTkFont(size=10),
                     justify="left", anchor="w", wraplength=155)
        self.dashboard_footer.pack(side="bottom", fill="x", padx=10, pady=(4, 6))
        database_card = ctk.CTkFrame(
            self.dashboard, fg_color=ACCENT_SOFT, border_color=LINE,
            border_width=1, corner_radius=10)
        database_card.pack(side="bottom", fill="x", padx=10, pady=(4, 6))
        self.dashboard_database_card = database_card
        ctk.CTkLabel(
            database_card, text=I.t("settings_database"), text_color=MUTED,
            font=ctk.CTkFont(size=10, weight="bold"), anchor="w").pack(
                fill="x", padx=10, pady=(7, 0))
        self.db_label = ctk.CTkLabel(
            database_card, text=I.t("db_count", n=len(self.db.drugs)),
            text_color=ACCENT, font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w")
        self.db_label.pack(fill="x", padx=10, pady=(0, 7))
        self._apply_dashboard_density()

        self.content = ctk.CTkFrame(self.workspace, fg_color="transparent")
        self.content.pack(side="left", fill="both", expand=True, padx=(0, 4), pady=4)
        self._build_quick_prescribe()
        self.scroll = ctk.CTkScrollableFrame(self.content, fg_color=BG, corner_radius=8)
        self.scroll.pack(fill="both", expand=True)
        # CustomTkinter defaults to 30 px per scroll unit on Windows and then
        # requests many units per wheel event.  A smaller increment keeps fast
        # wheel/touchpad movement smooth when many medication cards are visible.
        self.scroll._parent_canvas.configure(yscrollincrement=12)

        self.build_forms()

        # action bar (fixed footer)
        self.action = GlassFrame(self, height=44, fg_color=SURFACE,
                                   glass_surface="sheet", border_color=LINE, border_width=0,
                                   corner_radius=8)
        self.action.pack(side="bottom", fill="x", padx=4, pady=(0, 4))
        # left cluster
        VisualButton(self.action, text=I.t("preview"),
                      width=_ui_font(12).measure(I.t("preview")) + 22, height=ACTION_HEIGHT,
                      font=_ui_font(12),
                      corner_radius=9, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.preview).pack(side="left", padx=6)
        VisualButton(self.action, text=I.t("clear"),
                      width=_ui_font(12).measure(I.t("clear")) + 22, height=ACTION_HEIGHT,
                      font=_ui_font(12),
                      corner_radius=9, fg_color=CARD, text_color=MUTED,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.clear_all).pack(side="left", padx=6)
        self.busy_label = ctk.CTkLabel(
            self.action, text="", text_color=MUTED,
            font=ctk.CTkFont(size=10), anchor="w")
        # right cluster (Word exports)
        VisualButton(self.action, text=I.t("export_compact"),
                      width=_ui_font(12).measure(I.t("export_compact")) + 22,
                      height=ACTION_HEIGHT, font=_ui_font(12),
                      corner_radius=9, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.export_label).pack(side="right", padx=6)
        VisualButton(self.action, text=I.t("export_word"),
                      width=_ui_font(12).measure(I.t("export_word")) + 22,
                      height=ACTION_HEIGHT, font=_ui_font(12),
                      corner_radius=9, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.export_word).pack(side="right", padx=6)

        self.add_row()
        self.load_profile_into_ui()
        self.show_page("prescriber")
        self.apply_ui_font_preferences()

    def _build_quick_prescribe(self):
        bar = GlassFrame(self.content, fg_color=CARD, border_color=LINE,
                           border_width=1, corner_radius=10)
        bar.pack(fill="x", padx=3, pady=(0, 4))
        self._quick_search_bar = bar
        ctk.CTkLabel(bar, text="", image=action_icon("search"), width=34, text_color=ACCENT,
                     font=ctk.CTkFont(size=18, weight="bold")).pack(side="left", padx=(7, 0))
        self.quick_prescribe_var = tk.StringVar()
        self.quick_prescribe_entry = VisualEntry(
            bar, textvariable=self.quick_prescribe_var,
            placeholder_text=I.t("quick_prescribe_placeholder"), height=FIELD_HEIGHT,
            border_width=0, fg_color="transparent", font=ctk.CTkFont(size=13))
        self.quick_prescribe_entry.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=4)
        self.quick_prescribe_entry.bind(
            "<FocusIn>", lambda _event: bar.configure(border_color=ACCENT))
        self.quick_prescribe_entry.bind(
            "<FocusOut>", lambda _event: bar.configure(border_color=LINE))
        self._quick_results = []
        self._quick_popup = None
        self._quick_job = None
        self._quick_hide_job = None
        self.quick_prescribe_var.trace_add("write", lambda *_: self._schedule_quick_prescribe())
        self.quick_prescribe_entry.bind("<Down>", lambda _event: self._quick_move(1))
        self.quick_prescribe_entry.bind("<Up>", lambda _event: self._quick_move(-1))
        self.quick_prescribe_entry.bind("<Return>", self._quick_choose)
        self.quick_prescribe_entry.bind("<Escape>", lambda _event: self._hide_quick_results())
        attach_search_hint(self.quick_prescribe_entry, self.quick_prescribe_var,
                           I.t("quick_prescribe_placeholder"))
        self.bind_all("<Control-k>", self._focus_quick_prescribe)
        self.bind_all("<Button-1>", self._quick_click_outside, add="+")

    def _focus_quick_prescribe(self, _event=None):
        self.quick_prescribe_entry.focus_set()
        self.quick_prescribe_entry.select_range(0, tk.END)
        return "break"

    def _schedule_quick_prescribe(self):
        if self._quick_job is not None:
            try:
                self.after_cancel(self._quick_job)
            except (tk.TclError, ValueError):
                pass
        query = self.quick_prescribe_var.get().strip()
        if len(query) < 2:
            self._hide_quick_results()
            return
        self._quick_job = self.after(160, lambda: self._run_quick_prescribe(query))

    def _run_quick_prescribe(self, query):
        self._quick_job = None
        token = self._latest_query_tokens.get("quick", 0) + 1
        self._latest_query_tokens["quick"] = token

        def worker():
            needle = query.casefold()
            results = []
            for record in self.patient_history.search(query)[:8]:
                results.append((I.t("quick_patient"), record.get("name", ""), "patient", record))
            for template in cfg.config.treatment_templates():
                if needle in template.get("disease", "").casefold():
                    results.append((I.t("quick_template"), template["disease"], "template", template))
            for favorite in cfg.config.medication_favorites():
                name = favorite.get("brand_name") or favorite.get("generic_name", "")
                if favorite.get("pinned") and needle in name.casefold():
                    results.append((I.t("quick_starred"), name, "favorite", favorite.get("id", "")))
            for code in classes.GROUPS:
                group_name = I.t("class_" + code)
                if needle in group_name.casefold():
                    results.append((I.t("quick_class"), group_name, "major", code))
                for detail in classes.subclasses_for(code):
                    if needle in detail.casefold():
                        results.append((I.t("quick_class"), detail, "detail", (code, detail)))
            return results[:30]

        self.submit_background(
            worker, lambda results: self._render_quick_results(query, token, results), silent=True)

    def _render_quick_results(self, query, token, results):
        if token != self._latest_query_tokens.get("quick") or query != self.quick_prescribe_var.get().strip():
            return
        self._hide_quick_results()
        if not results:
            return
        top = tk.Toplevel(self)
        top.wm_overrideredirect(True)
        top.geometry(f"+{self.quick_prescribe_entry.winfo_rootx()}+{self.quick_prescribe_entry.winfo_rooty() + self.quick_prescribe_entry.winfo_height()}")
        box = PopupListbox(top, height=min(10, len(results)), width=70,
                         font=("Segoe UI", 21), bg=CARD, fg=TEXT,
                         relief="flat", borderwidth=0, selectbackground=PRIMARY,
                         highlightthickness=1, highlightbackground=LINE,
                         selectforeground="white", activestyle="none", exportselection=False)
        for kind, label, _action, _payload in results:
            box.insert(tk.END, f"{kind}  ·  {label}")
        fit_autocomplete_popup(top, self._quick_search_bar, box, len(results), align_anchor=True,
                               measure_content=True, width_multiplier=2, cap_width=False)
        box.selection_set(0)
        box.bind("<Double-Button-1>", self._quick_choose)
        box.bind("<Return>", self._quick_choose)
        self._quick_popup, self._quick_listbox, self._quick_results = top, box, results
        self._quick_hide_job = self.after(6000, self._expire_quick_results)

    def _expire_quick_results(self):
        self._quick_hide_job = None
        self._hide_quick_results()

    @staticmethod
    def _point_inside_widget(widget, x_root, y_root):
        try:
            return (widget.winfo_exists()
                    and widget.winfo_rootx() <= x_root < widget.winfo_rootx() + widget.winfo_width()
                    and widget.winfo_rooty() <= y_root < widget.winfo_rooty() + widget.winfo_height())
        except tk.TclError:
            return False

    def _quick_click_outside(self, event):
        popup = getattr(self, "_quick_popup", None)
        if popup is None:
            return
        if (self._point_inside_widget(popup, event.x_root, event.y_root)
                or self._point_inside_widget(
                    self.quick_prescribe_entry, event.x_root, event.y_root)):
            return
        self._hide_quick_results()

    def _hide_quick_results(self):
        hide_job = getattr(self, "_quick_hide_job", None)
        self._quick_hide_job = None
        if hide_job is not None:
            try:
                self.after_cancel(hide_job)
            except (tk.TclError, ValueError):
                pass
        if getattr(self, "_quick_popup", None):
            try:
                self._quick_popup.destroy()
            except tk.TclError:
                pass
        self._quick_popup = None
        self._quick_listbox = None

    def _quick_move(self, step):
        box = getattr(self, "_quick_listbox", None)
        if not box or not box.size():
            return "break"
        selected = box.curselection()
        index = (selected[0] if selected else 0) + step
        index = max(0, min(box.size() - 1, index))
        box.selection_clear(0, tk.END)
        box.selection_set(index)
        box.see(index)
        return "break"

    def _quick_choose(self, _event=None):
        box = getattr(self, "_quick_listbox", None)
        selected = box.curselection() if box else ()
        if not selected or selected[0] >= len(self._quick_results):
            return "break"
        _kind, _label, action, payload = self._quick_results[selected[0]]
        self._hide_quick_results()
        self.quick_prescribe_var.set("")
        if action == "patient":
            self.show_page("patient")
            self._load_patient_record(payload)
        elif action == "template":
            self.show_page("treatment_templates")
            self._load_treatment_template_record(payload)
        elif action == "favorite":
            self.use_favorite(payload)
        elif action == "major":
            self.show_page("drug_classes")
            self.show_subclass_page(payload)
        elif action == "detail":
            self.show_page("drug_classes")
            self.show_detail_medicines_page(*payload)
        return "break"

    def confirm_close(self):
        """Require an explicit confirmation before closing the desktop app."""
        if messagebox.askyesno(
                I.t("confirm_close_title"), I.t("confirm_close_message"), parent=self):
            if not self._confirm_treatment_leave():
                return
            self._executor.shutdown(wait=False, cancel_futures=True)
            self.destroy()

    def submit_background(self, worker, on_success=None, *, on_error=None,
                          label="Working…", silent=False):
        """Run blocking work without freezing Tk; callbacks always run on Tk."""
        if not silent:
            self._background_tasks += 1
            if hasattr(self, "busy_label"):
                self.busy_label.configure(text=label)
                self.busy_label.pack(side="left", padx=4)
        future = self._executor.submit(worker)

        def completed(done):
            try:
                result = done.result()
                error = None
            except Exception as exc:
                result, error = None, exc

            def deliver():
                if self._closing:
                    return
                if not silent:
                    self._background_tasks = max(0, self._background_tasks - 1)
                    if hasattr(self, "busy_label") and self._background_tasks == 0:
                        self.busy_label.configure(text="")
                        self.busy_label.pack_forget()
                if error is None:
                    if on_success:
                        on_success(result)
                elif on_error:
                    on_error(error)
                elif not silent:
                    logging.error(
                        "Background operation failed",
                        exc_info=(type(error), error, error.__traceback__))
                    messagebox.showerror(APP_TITLE, str(error), parent=self)

            if not self._closing:
                self._background_callbacks.put(deliver)

        future.add_done_callback(completed)
        return future

    # -- forms ---------------------------------------------------------------
    def section(self, parent, title):
        f = GlassFrame(parent, fg_color=CARD, border_color=LINE,
                         border_width=1, corner_radius=12)
        f.pack(fill="x", padx=2, pady=CARD_GAP)
        if title:
            ctk.CTkLabel(f, text=title, font=ctk.CTkFont(weight="bold", size=14),
                         text_color=TEXT, anchor="w").pack(
                             anchor="w", padx=PAD, pady=(6, 3))
        return f

    def page_header(self, parent, title, subtitle):
        """A consistent title block makes every dashboard page easy to scan."""
        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(fill="x", padx=4, pady=(0, 3))
        ctk.CTkLabel(header, text=title, text_color=TEXT,
                     font=ctk.CTkFont(weight="bold", size=PAGE_TITLE_FONT_SIZE),
                     anchor="w").pack(fill="x")
        if subtitle:
            ctk.CTkLabel(header, text=subtitle, text_color=MUTED,
                         font=ctk.CTkFont(size=11), anchor="w", justify="left",
                         wraplength=720).pack(fill="x", pady=(2, 0))
        return header

    def _add_page_button(self, key, text):
        normal_icon = _dashboard_icon(key)
        selected_icon = _dashboard_icon(key, NAV_ACTIVE)
        row = ctk.CTkFrame(self.dashboard, fg_color="transparent", corner_radius=8)
        row.pack(fill="x", padx=6, pady=2)
        indicator = ctk.CTkFrame(row, width=3, height=22, corner_radius=1,
                                 fg_color=NAV_ACTIVE)
        button = VisualButton(
            row, text=text, image=normal_icon, compound="left",
            anchor="w", height=36, corner_radius=8,
            fg_color="transparent", text_color=TEXT_SECONDARY, hover_color=ACCENT_SOFT,
            font=_ui_font(12),
            command=lambda page=key: self.show_page(page))
        button._image_label_spacing = 10
        button.configure(anchor="w")
        button.pack(side="left", fill="x", expand=True, padx=(5, 0))
        button._selection_indicator = indicator
        button._navigation_row = row
        self.page_buttons[key] = button
        self._dashboard_button_labels[key] = text
        self._page_button_icons[key] = (normal_icon, selected_icon)

    def _add_dashboard_action(self, key, text, command):
        icon = _dashboard_icon(key)
        self._dashboard_action_icons.append(icon)
        button = VisualButton(
            self.dashboard, text=text, image=icon, compound="left",
            anchor="w", height=36,
            corner_radius=8, fg_color="transparent", text_color=ACCENT,
            hover_color=ACCENT_SOFT, font=_ui_font(12),
            command=command)
        button._image_label_spacing = 10
        button.configure(anchor="w")
        button.pack(fill="x", padx=(11, 6), pady=2)
        self._dashboard_action_buttons[key] = button
        self._dashboard_button_labels[key] = text

    def toggle_dashboard(self):
        self.dashboard_collapsed = not self.dashboard_collapsed
        self._apply_dashboard_density()
        cfg.config.set("dashboard_collapsed", self.dashboard_collapsed)

    def _apply_dashboard_density(self):
        collapsed = self.dashboard_collapsed
        self.dashboard.configure(width=(DASHBOARD_COLLAPSED_WIDTH if collapsed
                                        else DASHBOARD_WIDTH))
        self.dashboard_toggle.configure(text="›" if collapsed else "‹")
        for key, button in {**self.page_buttons, **self._dashboard_action_buttons}.items():
            button.configure(text="" if collapsed else self._dashboard_button_labels[key],
                             anchor="center" if collapsed else "w", width=36)
        if collapsed:
            self.dashboard_title.pack_forget()
            self.dashboard_footer.pack_forget()
            self.dashboard_database_card.pack_forget()
        else:
            self.dashboard_title.pack(side="left", fill="x", expand=True,
                                      before=self.dashboard_toggle)
            self.dashboard_footer.pack(side="bottom", fill="x", padx=10, pady=(4, 6))
            self.dashboard_database_card.pack(side="bottom", fill="x", padx=10, pady=(4, 6))

    def show_page(self, key):
        if (key != "treatment_templates" and
                getattr(self, "active_page", "") == "treatment_templates" and
                not self._confirm_treatment_leave()):
            return False
        if hasattr(self, "medication_subpage"):
            self.close_medication_subpage()
        for name, page in self.pages.items():
            page.pack_forget()
            normal_icon, selected_icon = self._page_button_icons[name]
            self.page_buttons[name].configure(
                fg_color=NAV_ACTIVE_SOFT if name == key else "transparent",
                text_color=NAV_ACTIVE if name == key else TEXT_SECONDARY,
                image=selected_icon if name == key else normal_icon,
                font=_ui_font(12, "bold" if name == key else "normal"),
            )
            button = self.page_buttons[name]
            button._navigation_row.configure(
                fg_color=NAV_ACTIVE_SOFT if name == key else "transparent")
            if name == key:
                button._selection_indicator.place(x=0, rely=0.5, anchor="w")
            else:
                button._selection_indicator.place_forget()
        self.pages[key].pack(fill="both", expand=True, padx=4, pady=4)
        # Hide only the visual rail; the canvas and wheel scrolling remain active.
        if key in SCROLLBAR_HIDDEN_PAGES:
            self.scroll._scrollbar.grid_remove()
        else:
            self.scroll._scrollbar.grid()
        self.active_page = key
        if key not in self._loaded_pages:
            self._loaded_pages.add(key)
            self.after_idle(lambda page=key: self._load_page_data(page))

    def _load_page_data(self, key):
        """Populate expensive pages only when the clinician first opens them."""
        if key == "patient":
            self.refresh_patient_history()
        elif key == "favorites":
            self.refresh_favorites_page()
        elif key == "drug_classes":
            self.refresh_class_overview()
        elif key == "treatment_templates":
            self.refresh_treatment_template_menu()
            self.refresh_treatment_drug_results()
            self.render_treatment_template_drugs()

    def field(self, parent, label, var, width=260, placeholder=None, expand=True,
              directional=False):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=PAD, pady=4)
        ctk.CTkLabel(row, text=label, width=130, anchor="w",
                     text_color=MUTED, font=ctk.CTkFont(size=11),
                     height=FIELD_HEIGHT).pack(side="left")
        display_var = var
        binding = None
        if directional:
            binding = DirectionalTextBinding(self, var)
            self._bidi_bindings.append(binding)
            display_var = binding.display_var
        entry = VisualEntry(
            row, textvariable=display_var, width=width, height=FIELD_HEIGHT,
            corner_radius=9, border_color=LINE, placeholder_text=placeholder,
            justify="left")
        if binding:
            binding.attach(entry)
        if expand:
            entry.pack(side="left", fill="x", expand=True)
        else:
            entry.pack(side="left")
        return row

    def specialty_field(self, parent, label, var, width=260):
        """Editable and searchable selector for common medical specialties."""
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=PAD, pady=4)
        ctk.CTkLabel(
            row, text=label, width=130, anchor="w", text_color=MUTED,
            font=ctk.CTkFont(size=11), height=FIELD_HEIGHT,
        ).pack(side="left")
        binding = DirectionalTextBinding(self, var)
        self._bidi_bindings.append(binding)
        self.specialty_menu = VisualComboBox(
            row, values=self._specialty_values(var.get()),
            variable=binding.display_var,
            width=width, height=FIELD_HEIGHT, corner_radius=9,
            border_width=2, border_color=ACCENT, fg_color=CARD,
            button_color=PRIMARY, button_hover_color=ACCENT_HOVER,
            dropdown_fg_color=CARD, dropdown_hover_color=ACCENT_SOFT,
            font=ctk.CTkFont(size=13), dropdown_font=ctk.CTkFont(size=13),
            justify="left",
            command=self._on_specialty_selected,
        )
        binding.attach(self.specialty_menu)
        self.specialty_menu.pack(side="left")
        self.specialty_menu.bind("<KeyRelease>", self._filter_specialties)
        self.specialty_menu.bind(
            "<FocusOut>", lambda _event: self._refresh_specialty_values())
        return row

    def _specialty_values(self, current=""):
        choices = medical_specialty_options()
        custom = strip_bidi_display_controls(current).strip()
        if custom and custom not in choices and custom != I.t("other_specialty"):
            choices.insert(0, custom)
        choices.append(I.t("other_specialty"))
        return [directional_display_text(choice) for choice in choices]

    def _refresh_specialty_values(self, current=None):
        if not getattr(self, "specialty_menu", None):
            return
        value = self.doctor_vars["specialty"].get() if current is None else current
        self.specialty_menu.configure(values=self._specialty_values(value))

    def _filter_specialties(self, event=None):
        if event and event.keysym in {
                "Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R",
                "Control_L", "Control_R", "Alt_L", "Alt_R"}:
            return
        query = self.doctor_vars["specialty"].get().strip().casefold()
        choices = medical_specialty_options()
        matches = [choice for choice in choices if query in choice.casefold()]
        matches.append(I.t("other_specialty"))
        self.specialty_menu.configure(
            values=[directional_display_text(choice) for choice in matches])

    def _on_specialty_selected(self, value):
        value = strip_bidi_display_controls(value)
        if value == I.t("other_specialty"):
            self.doctor_vars["specialty"].set("")
            self.after_idle(self.specialty_menu.focus_set)
            return
        self.doctor_vars["specialty"].set(value)

    def sex_field(self, parent, label, var):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=PAD, pady=4)
        ctk.CTkLabel(row, text=label, width=130, anchor="w",
                     text_color=MUTED, font=ctk.CTkFont(size=11),
                     height=FIELD_HEIGHT).pack(side="left")
        opts = [I.t("sex_m"), I.t("sex_f")]
        code_of = {I.t("sex_m"): "M", I.t("sex_f"): "F"}
        cur = var.get()
        sel = next((lab for lab, c in code_of.items() if c == cur), opts[0])
        om = VisualOptionMenu(row, values=opts, width=120, height=FIELD_HEIGHT,
                               corner_radius=9, fg_color=ACCENT_SOFT,
                               text_color=ACCENT, button_color=PRIMARY,
                               button_hover_color=ACCENT_HOVER,
                               dropdown_hover_color=ACCENT_SOFT,
                               command=lambda v, code_of=code_of, var=var: var.set(code_of.get(v, "")))
        om.set(sel)
        om.pack(side="left")
        return row

    def build_forms(self):
        self._bidi_bindings = []
        self.doctor_vars = {k: tk.StringVar() for k in ["name", "license_no", "specialty"]}
        self.patient_vars = {k: tk.StringVar() for k in ["name", "age", "sex"]}
        self._patient_status_suspend = False
        self._patient_saved_snapshot = ("", "", "")
        self.pages = {
            "prescriber": ctk.CTkFrame(self.scroll, fg_color="transparent"),
            "patient": ctk.CTkFrame(self.scroll, fg_color="transparent"),
            "medications": ctk.CTkFrame(self.scroll, fg_color="transparent"),
            "favorites": ctk.CTkFrame(self.scroll, fg_color="transparent"),
            "drug_classes": ctk.CTkFrame(self.scroll, fg_color="transparent"),
            "treatment_templates": ctk.CTkFrame(self.scroll, fg_color="transparent"),
            "interaction_review": ctk.CTkFrame(self.scroll, fg_color="transparent"),
            "reference": ctk.CTkFrame(self.scroll, fg_color="transparent"),
        }

        self.page_header(self.pages["prescriber"], I.t("prescriber_details"), "")
        d = self.section(self.pages["prescriber"], "")
        self.field(
            d, I.t("f_name"), self.doctor_vars["name"],
            PRESCRIBER_FIELD_WIDTH, expand=False, directional=True)
        self.field(
            d, I.t("license"), self.doctor_vars["license_no"],
            PRESCRIBER_FIELD_WIDTH, expand=False)
        self.specialty_field(
            d, I.t("specialty"), self.doctor_vars["specialty"],
            PRESCRIBER_FIELD_WIDTH)
        VisualButton(
            d, text=I.t("save_profile"), width=150, height=ACTION_HEIGHT,
            corner_radius=9, fg_color=GOOD, hover_color=GOOD_HOVER,
            command=self.save_profile).pack(
                anchor="w", padx=(PAD + 130, PAD), pady=(2, PAD))

        self.page_header(self.pages["patient"], I.t("patient_details"), "")
        p = self.section(self.pages["patient"], "")

        patient_fields = ctk.CTkFrame(p, fg_color="transparent")
        patient_fields.pack(fill="x", padx=PAD, pady=(6, CARD_GAP))
        # The three patient controls use predictable compact widths.
        patient_fields.grid_columnconfigure(0, weight=0)
        patient_fields.grid_columnconfigure(1, weight=0)
        patient_fields.grid_columnconfigure(2, weight=0)

        name_col = ctk.CTkFrame(patient_fields, fg_color="transparent")
        name_col.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkLabel(name_col, text=I.t("f_name"), text_color=MUTED,
                     font=ctk.CTkFont(size=11), anchor="w").pack(fill="x", pady=(0, 3))
        patient_name_binding = DirectionalTextBinding(self, self.patient_vars["name"])
        self._bidi_bindings.append(patient_name_binding)
        self.patient_name_entry = VisualEntry(
            name_col, textvariable=patient_name_binding.display_var,
            width=PATIENT_NAME_WIDTH, height=FIELD_HEIGHT,
            corner_radius=9, border_color=LINE,
            font=_ui_font(cfg.config.ui_font_size("patient_name_font_size")), justify="left")
        patient_name_binding.attach(self.patient_name_entry)
        self.patient_name_entry.pack(fill="x")

        age_col = ctk.CTkFrame(patient_fields, fg_color="transparent")
        age_col.grid(row=0, column=1, sticky="ew", padx=6)
        ctk.CTkLabel(age_col, text=I.t("age"), text_color=MUTED,
                     font=ctk.CTkFont(size=11), anchor="w").pack(fill="x", pady=(0, 3))
        self.patient_age_entry = VisualEntry(
            age_col, textvariable=self.patient_vars["age"], width=PATIENT_AGE_WIDTH,
            height=FIELD_HEIGHT,
            corner_radius=9, border_color=LINE, font=ctk.CTkFont(size=14),
            justify="left")
        self.patient_age_entry.pack(anchor="w")

        sex_col = ctk.CTkFrame(patient_fields, fg_color="transparent")
        sex_col.grid(row=0, column=2, sticky="ew", padx=(6, 0))
        ctk.CTkLabel(sex_col, text=I.t("sex"), text_color=MUTED,
                     font=ctk.CTkFont(size=11), anchor="w").pack(fill="x", pady=(0, 3))
        sex_labels = [I.t("sex_m"), I.t("sex_f")]
        self._patient_sex_codes = {I.t("sex_m"): "M", I.t("sex_f"): "F"}
        self.patient_sex_menu = VisualOptionMenu(
            sex_col, values=sex_labels, width=PATIENT_SEX_WIDTH,
            height=FIELD_HEIGHT, corner_radius=9,
            fg_color=ACCENT_SOFT, text_color=ACCENT, button_color=PRIMARY,
            button_hover_color=ACCENT_HOVER, dropdown_hover_color=ACCENT_SOFT,
            command=self._set_patient_sex)
        self.patient_sex_menu.set(I.t("sex_m"))
        self.patient_sex_menu.pack(anchor="w")

        patient_actions = ctk.CTkFrame(p, fg_color="transparent")
        patient_actions.pack(fill="x", padx=PAD, pady=(1, PAD))
        patient_new_button = VisualButton(
            patient_actions, text="＋", width=ICON_BUTTON_SIZE, height=ICON_BUTTON_SIZE,
            fg_color="transparent", text_color=ACCENT, border_width=0,
            hover_color=ACCENT_SOFT, font=ctk.CTkFont(size=20, weight="bold"),
            command=self.new_patient)
        patient_new_button.pack(side="left", padx=(0, 4))
        patient_save_button = VisualButton(
            patient_actions, text="💾", width=ICON_BUTTON_SIZE, height=ICON_BUTTON_SIZE,
            fg_color="transparent", text_color=ACCENT, border_width=0,
            hover_color=ACCENT_SOFT, font=ctk.CTkFont(size=17),
            command=self.save_patient_history)
        patient_save_button.pack(side="left", padx=4)
        self.patient_delete_button = VisualButton(
            patient_actions, text="🗑", width=ICON_BUTTON_SIZE, height=ICON_BUTTON_SIZE,
            fg_color="transparent", text_color=DANGER, border_width=0,
            hover_color=DANGER_SOFT, font=ctk.CTkFont(size=17),
            state="disabled", command=self.delete_selected_patient)
        self.patient_delete_button.pack(side="left", padx=4)
        self.patient_status_label = ctk.CTkLabel(
            patient_actions, text="", width=92, height=30,
            fg_color=ACCENT_SOFT, text_color=ACCENT, corner_radius=15,
            font=ctk.CTkFont(size=11, weight="bold"))
        self.patient_status_label.pack(side="right", padx=(8, 0))
        for variable in self.patient_vars.values():
            variable.trace_add("write", self._on_patient_form_change)
        self._set_patient_status("new")

        history = self.section(self.pages["patient"], I.t("patient_history"))
        patient_history_grid = ctk.CTkFrame(history, fg_color="transparent")
        patient_history_grid.pack(fill="x", padx=PAD, pady=(0, PAD))
        patient_history_grid.grid_columnconfigure(0, weight=0)
        patient_history_grid.grid_columnconfigure(1, weight=1)
        self.patient_search_var = tk.StringVar()
        self._patient_search_job = None
        self.patient_search_var.trace_add(
            "write", lambda *_: self._schedule_patient_history_refresh())
        patient_finder = ctk.CTkFrame(patient_history_grid, fg_color="transparent")
        patient_finder.grid(row=0, column=0, sticky="nw", padx=(0, 10))
        patient_search_binding = DirectionalTextBinding(self, self.patient_search_var)
        self._bidi_bindings.append(patient_search_binding)
        self.patient_search_entry = VisualEntry(
            patient_finder, textvariable=patient_search_binding.display_var,
            width=PATIENT_SEARCH_WIDTH,
            height=FIELD_HEIGHT, placeholder_text=I.t("search_patients"),
            border_color=ACCENT, corner_radius=9, font=ctk.CTkFont(size=16),
            justify="left")
        patient_search_binding.attach(self.patient_search_entry)
        self.patient_search_entry.pack(anchor="w", pady=(0, 6))
        result_shell = ctk.CTkFrame(
            patient_finder, width=PATIENT_SEARCH_WIDTH,
            height=PATIENT_RESULTS_HEIGHT, fg_color=SURFACE,
            border_color=LINE, border_width=1, corner_radius=9)
        result_shell.pack_propagate(False)
        result_shell.pack(anchor="w")
        self.patient_history_list = PopupListbox(result_shell, height=5,
                                               font=LIST_FONT, bg=SURFACE, fg=TEXT,
                                               relief="flat", borderwidth=0, highlightthickness=0,
                                               selectbackground=PRIMARY, selectforeground="white",
                                               activestyle="none")
        self.patient_history_list.pack(fill="both", expand=True, padx=4, pady=4)
        self.patient_history_list.bind("<<ListboxSelect>>", self.select_patient_history)
        self.patient_history_list.bind("<Double-Button-1>", self.load_selected_patient)
        self.patient_search_entry.bind("<Return>", self.load_selected_patient)
        self.patient_search_entry.bind("<Escape>", self.clear_patient_search)
        prescriptions_panel = ctk.CTkFrame(patient_history_grid, fg_color="transparent")
        prescriptions_panel.grid(row=0, column=1, sticky="nsew")
        ctk.CTkLabel(prescriptions_panel, text=I.t("previous_prescriptions"), text_color=ACCENT,
                     font=ctk.CTkFont(size=11, weight="bold"), anchor="w").pack(
                         fill="x", pady=(0, 3))
        self.patient_comparison_body = ctk.CTkFrame(
            prescriptions_panel, fg_color=SURFACE, border_color=LINE,
            border_width=1, corner_radius=9)
        self.patient_comparison_body.pack(fill="x", pady=(0, 5))
        self.patient_prescriptions_body = ctk.CTkFrame(
            prescriptions_panel, fg_color="transparent")
        self.patient_prescriptions_body.pack(fill="x")
        self._loaded_patient_id = ""
        self._expanded_prescription_ids = set()
        self._current_history_record = None
        self.patient_name_entry.bind("<Return>", lambda _event: self.patient_age_entry.focus_set())
        self.patient_age_entry.bind("<Return>", lambda _event: self.patient_sex_menu.focus_set())
        self._show_empty_prescriptions()
        self.refresh_prescription_comparison()

        self.page_header(self.pages["medications"], I.t("medication_entry"), "")
        self.word_preview_visible = False
        self.medication_main = ctk.CTkFrame(self.pages["medications"], fg_color="transparent")
        self.medication_main.pack(fill="both", expand=True)
        self.medication_subpage = ctk.CTkFrame(self.pages["medications"], fg_color="transparent")
        subpage_bar = ctk.CTkFrame(self.medication_subpage, fg_color="transparent")
        subpage_bar.pack(fill="x", pady=(4, 8))
        VisualButton(
            subpage_bar, text=I.t("back"), width=70, height=32,
            fg_color=CARD, text_color=ACCENT, border_color=LINE, border_width=1,
            hover_color=ACCENT_SOFT, command=self.close_medication_subpage).pack(side="left")
        self.medication_subpage_title = ctk.CTkLabel(
            subpage_bar, text="", font=_ui_font(18, "bold"), text_color=TEXT)
        self.medication_subpage_title.pack(side="left", padx=10)
        medication_toolbar = ctk.CTkFrame(
            self.medication_main, fg_color="transparent")
        medication_toolbar.pack(fill="x", padx=2, pady=(0, 4))
        VisualButton(
            medication_toolbar, text=I.t("add_drug"),
            width=_ui_font(12).measure(I.t("add_drug")) + 22,
            height=32, font=_ui_font(12), corner_radius=8, fg_color=SURFACE,
            text_color=ACCENT, border_color=LINE, border_width=1,
            hover_color=ACCENT_SOFT, command=self.add_row).pack(side="left", padx=(0, 5))
        self.medication_favorites_button = VisualButton(
            medication_toolbar, text=I.t("starred_drugs"),
            width=_ui_font(12).measure(I.t("starred_drugs")) + 22,
            height=32, font=_ui_font(12), corner_radius=8, fg_color=SURFACE,
            text_color=ACCENT, border_color=LINE, border_width=1,
            hover_color=ACCENT_SOFT,
            command=self.toggle_medication_favorite_picker)
        self.medication_favorites_button.pack(side="left")
        VisualButton(
            medication_toolbar, text=I.t("save_patient_prescription"),
            width=_ui_font(12).measure(I.t("save_patient_prescription")) + 22,
            height=32, font=_ui_font(12), corner_radius=8,
            fg_color=SURFACE, text_color=ACCENT, border_color=LINE, border_width=1,
            hover_color=ACCENT_SOFT,
            command=self.save_prescription_for_patient).pack(side="left", padx=(5, 0))

        self.medication_favorite_panel = ctk.CTkFrame(
            self.medication_subpage, fg_color=SURFACE, border_color=LINE,
            border_width=1, corner_radius=12)
        favorite_picker_head = ctk.CTkFrame(
            self.medication_favorite_panel, fg_color="transparent")
        favorite_picker_head.pack(fill="x", padx=10, pady=(8, 5))
        self.medication_favorite_search_var = tk.StringVar()
        self.medication_favorite_search_var.trace_add(
            "write", lambda *_: self.refresh_medication_favorite_picker())
        self.medication_favorite_search_entry = VisualEntry(
            favorite_picker_head, textvariable=self.medication_favorite_search_var,
            placeholder_text=I.t("search_starred_drugs"), height=36,
            border_color=LINE, corner_radius=8)
        self.medication_favorite_search_entry.pack(
            side="left", fill="x", expand=True, padx=(0, 6))
        VisualButton(
            favorite_picker_head, text="×", width=34, height=32,
            fg_color="transparent", text_color=MUTED, hover_color=ACCENT_SOFT,
            command=self.close_medication_subpage).pack(side="right")
        # Let the result container follow its cards instead of reserving a
        # fixed blank area when only one or two starred drugs are available.
        self.medication_favorite_results = ctk.CTkFrame(
            self.medication_favorite_panel, fg_color="transparent")
        self.medication_favorite_results.pack(fill="x", padx=8, pady=(0, 8))
        self.medication_favorite_results.grid_columnconfigure(
            (0, 1), weight=1, uniform="starred_medicine_cards")

        dr = self.section(self.medication_main, "")
        self.medications_section = dr
        self.drugs_frame = ctk.CTkFrame(dr, fg_color="transparent")
        self.drugs_frame.pack(fill="x", padx=6, pady=(0, 2))

        self.page_header(self.pages["interaction_review"], I.t("interaction_review"),
                         I.t("page_interaction_help"))
        review = self.section(self.pages["interaction_review"], "")
        review.pack_configure(pady=(12, CARD_GAP))
        ctk.CTkLabel(review, text=I.t("interaction_review_tip"), text_color=MUTED,
                     justify="left", wraplength=720, anchor="w").pack(fill="x", padx=PAD, pady=(12, 6))
        self.medscape_button = VisualButton(
            review, text=I.t("check_medscape"), height=ACTION_HEIGHT, width=250,
            fg_color=CARD, text_color=ACCENT, border_width=1,
            border_color=LINE, hover_color=ACCENT_SOFT,
            command=self.check_interactions_in_medscape)
        self.medscape_button.pack(anchor="w", padx=PAD, pady=(0, 8))

        preview = self.section(self.medication_subpage, "")
        self.word_preview_section = preview
        preview.pack_forget()
        self.word_preview_toggle = VisualButton(
            medication_toolbar, text=I.t("show_word_preview"),
            width=max(_ui_font(12).measure(I.t(key)) for key in (
                "show_word_preview", "hide_word_preview")) + 22,
            height=32, font=_ui_font(12),
            fg_color=SURFACE, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=self.toggle_word_preview)
        self.word_preview_toggle.pack(side="left", padx=(5, 0))
        self.word_preview_body = ctk.CTkFrame(preview, fg_color="transparent")

        self.page_header(self.pages["reference"], I.t("online_drug_reference"),
                         "")
        safety = self.section(self.pages["reference"], "")
        safety.pack_configure(pady=(0, CARD_GAP))
        ctk.CTkLabel(safety, text=I.t("openfda_connected"), text_color=MUTED,
                     justify="left", wraplength=720, anchor="w").pack(fill="x", padx=PAD, pady=(12, 8))
        reference_search = ctk.CTkFrame(safety, fg_color="transparent")
        reference_search.pack(fill="x", padx=PAD, pady=(0, 6))
        self.reference_search_var = tk.StringVar()
        self.reference_search_entry = VisualEntry(
            reference_search, textvariable=self.reference_search_var,
            placeholder_text=I.t("openfda_search_placeholder"), height=FIELD_HEIGHT,
            border_color=LINE, font=_ui_font(14))
        self.reference_search_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.reference_search_entry.bind("<Return>", self.lookup_openfda_search)
        self.reference_search_button = VisualButton(
            reference_search, text=I.t("openfda_search_button"), width=90,
            height=ACTION_HEIGHT, fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            command=self.lookup_openfda_search)
        self.reference_search_button.pack(side="left")
        reference_actions = ctk.CTkFrame(safety, fg_color="transparent")
        reference_actions.pack(fill="x", padx=PAD, pady=(0, 6))
        self.reference_lookup_button = VisualButton(
            reference_actions, text=I.t("lookup_current_medicines"), height=ACTION_HEIGHT, width=180,
            fg_color=CARD, text_color=ACCENT, border_width=1,
            border_color=LINE, hover_color=ACCENT_SOFT,
            command=self.lookup_openfda_labels)
        self.reference_lookup_button.pack(side="left")
        self.reference_refresh_button = VisualButton(
            reference_actions, text=I.t("reference_refresh"), height=ACTION_HEIGHT, width=90,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=self.refresh_openfda_labels)
        self.reference_refresh_button.pack(side="left", padx=5)
        VisualButton(
            reference_actions, text=I.t("reference_clear_cache"), height=ACTION_HEIGHT,
            width=100, fg_color="transparent", text_color=MUTED,
            hover_color=ACCENT_SOFT, command=self.clear_openfda_cache).pack(side="left")
        self.reference_status = ctk.CTkLabel(safety, text=I.t("openfda_ready"),
                                             text_color=MUTED, anchor="w")
        self.reference_status.pack(fill="x", padx=PAD, pady=(0, 6))
        self.reference_cards = ctk.CTkFrame(safety, fg_color="transparent")
        self.reference_cards.pack(fill="x", padx=PAD, pady=(0, PAD))
        self._reference_last_medicines = []
        self._show_reference_placeholder()

        self.page_header(self.pages["favorites"], I.t("favorite_drugs"), "")
        favorites = self.section(self.pages["favorites"], "")
        favorite_toolbar = ctk.CTkFrame(favorites, fg_color="transparent")
        favorite_toolbar.pack(fill="x", padx=PAD, pady=(6, 6))
        self.favorite_search_var = tk.StringVar()
        self.favorite_search_var.trace_add(
            "write", lambda *_: self._schedule_favorites_refresh())
        VisualEntry(
            favorite_toolbar, textvariable=self.favorite_search_var,
            height=FIELD_HEIGHT, placeholder_text=I.t("favorite_search"),
            border_color=LINE, font=ctk.CTkFont(size=15)).pack(
                side="left", fill="x", expand=True, padx=(0, 8))
        self.favorite_sort_var = tk.StringVar(value=I.t("favorite_sort_used"))
        self.favorite_sort_menu = VisualOptionMenu(
            favorite_toolbar,
            values=[I.t("favorite_sort_used"), I.t("favorite_sort_recent"),
                    I.t("favorite_sort_name"), I.t("favorite_sort_name_reverse")],
            variable=self.favorite_sort_var, width=145, height=FIELD_HEIGHT,
            fg_color=ACCENT_SOFT, text_color=ACCENT, button_color=PRIMARY,
            button_hover_color=ACCENT_HOVER, font=ctk.CTkFont(size=13),
            command=lambda _value: self.refresh_favorites_page())
        self.favorite_sort_menu.pack(side="left", padx=(0, 5))
        VisualButton(
            favorite_toolbar, text=I.t("new_favorite"), width=125,
            height=ACTION_HEIGHT, fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=13),
            command=self.new_favorite).pack(side="left")
        polish_toolbar(favorite_toolbar)

        self.favorite_category_var = tk.StringVar(value=I.t("all_categories"))
        self._favorite_selected_ids = set()
        self._favorite_card_widgets = {}
        self._favorite_highlight_id = None
        self.favorite_category_chips = ctk.CTkFrame(favorites, fg_color="transparent")
        self.favorite_category_chips.pack(fill="x", padx=PAD, pady=(0, 8))
        self._favorite_category_width = 0
        self.favorite_category_chips.bind(
            "<Configure>", self._on_favorite_category_resize)

        self.favorite_notice = ctk.CTkFrame(
            favorites, fg_color=WARNING_SOFT, border_width=1,
            border_color=LINE, corner_radius=9)
        self.favorite_notice_label = ctk.CTkLabel(
            self.favorite_notice, text="", text_color=WARNING, anchor="w",
            font=ctk.CTkFont(size=12, weight="bold"))
        self.favorite_notice_label.pack(side="left", fill="x", expand=True, padx=12, pady=8)
        self.favorite_undo_button = VisualButton(
            self.favorite_notice, text=I.t("favorite_undo"), width=80, height=30,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=self.undo_delete_favorite)
        self.favorite_undo_timer = None
        self._deleted_favorite = None

        self.favorite_vars = {key: tk.StringVar() for key in
                              ("generic_name", "brand_name", "category",
                               "dosage", "frequency", "duration", "notes")}
        self.favorite_edit_index = None
        self.favorite_editor = ctk.CTkFrame(
            favorites, fg_color=SURFACE, border_color=LINE,
            border_width=1, corner_radius=12)
        self.favorite_editor.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.favorite_editor_title = ctk.CTkLabel(
            self.favorite_editor, text=I.t("new_favorite"), text_color=ACCENT,
            font=ctk.CTkFont(size=18, weight="bold"), anchor="w")
        self.favorite_editor_title.grid(
            row=0, column=0, columnspan=3, sticky="ew", padx=14, pady=(12, 8))
        VisualButton(
            self.favorite_editor, text="×", width=34, height=30,
            fg_color="transparent", text_color=MUTED, hover_color=ACCENT_SOFT,
            command=self.close_favorite_editor).grid(
                row=0, column=3, sticky="e", padx=12, pady=(10, 6))

        editor_fields = (
            ("brand_name", I.t("brand_name"), None),
            ("generic_name", I.t("scientific_name"), None),
            ("category", I.t("category"), "categories"),
            ("dosage", I.t("dosage"), None),
            ("frequency", I.t("frequency"), FREQUENCY_OPTIONS),
            ("duration", I.t("duration"), None),
            ("notes", I.t("notes"), NOTE_OPTIONS),
        )
        for index, (key, label, values) in enumerate(editor_fields):
            row = 1 + (index // 4) * 2
            column = index % 4
            ctk.CTkLabel(
                self.favorite_editor, text=label, text_color=MUTED, anchor="w",
                font=_ui_font(11)).grid(
                    row=row, column=column, sticky="ew", padx=6, pady=(2, 3))
            if values == "categories":
                widget = VisualComboBox(
                    self.favorite_editor, values=[""],
                    variable=self.favorite_vars[key], height=FIELD_HEIGHT,
                    border_color=LINE, fg_color=CARD,
                    button_color=PRIMARY, button_hover_color=ACCENT_HOVER,
                    dropdown_fg_color=CARD, dropdown_hover_color=ACCENT_SOFT,
                    dropdown_text_color=TEXT, dropdown_font=_ui_font(16),
                    font=ctk.CTkFont(size=12))
                self.favorite_category_combo = widget
            elif values:
                widget = VisualComboBox(
                    self.favorite_editor, values=list(values),
                    variable=self.favorite_vars[key], height=FIELD_HEIGHT,
                    border_color=LINE, fg_color=CARD,
                    button_color=PRIMARY, button_hover_color=ACCENT_HOVER,
                    dropdown_fg_color=CARD, dropdown_hover_color=ACCENT_SOFT,
                    dropdown_text_color=TEXT, dropdown_font=_ui_font(16),
                    font=ctk.CTkFont(size=12))
            else:
                widget = VisualEntry(
                    self.favorite_editor, textvariable=self.favorite_vars[key],
                    height=FIELD_HEIGHT, border_color=LINE,
                    font=ctk.CTkFont(
                        size=12, weight="bold" if key == "brand_name" else "normal"))
                if key == "generic_name":
                    self.favorite_generic_entry = widget
                elif key == "brand_name":
                    self.favorite_brand_entry = widget
            widget.grid(row=row + 1, column=column, sticky="ew", padx=6, pady=(0, 8))

        self._favorite_ac_top = None
        self._favorite_ac_listbox = None
        self._favorite_ac_matches = []
        self._favorite_ac_mode = ""
        self._favorite_ac_index = -1
        self.favorite_brand_entry.bind("<KeyRelease>", self._on_favorite_brand_type)
        self.favorite_brand_entry.bind("<Down>", self._favorite_ac_down)
        self.favorite_brand_entry.bind("<Up>", self._favorite_ac_up)
        self.favorite_brand_entry.bind("<Return>", self._favorite_ac_choose_current)
        self.favorite_brand_entry.bind(
            "<Escape>", lambda _event: self._hide_favorite_autocomplete())
        self.favorite_generic_entry.bind(
            "<FocusIn>", lambda _event: self._hide_favorite_autocomplete())

        preview_row = ctk.CTkFrame(self.favorite_editor, fg_color="transparent")
        preview_row.grid(row=5, column=0, columnspan=4, sticky="ew", padx=6, pady=(2, 12))
        preview_row.grid_columnconfigure(0, weight=1)
        self.favorite_preview_label = ctk.CTkLabel(
            preview_row, text="", height=42, corner_radius=9,
            fg_color=CARD, text_color=TEXT, anchor="w", padx=12,
            font=ctk.CTkFont(size=14, weight="bold"))
        self.favorite_preview_label.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        VisualButton(
            preview_row, text=I.t("save_changes"), width=130, height=42,
            fg_color=GOOD, hover_color=GOOD_HOVER,
            command=self.save_favorite_changes).grid(row=0, column=1)
        for variable in self.favorite_vars.values():
            variable.trace_add("write", lambda *_: self.update_favorite_preview())

        self.favorite_selection_bar = ctk.CTkFrame(
            favorites, fg_color=ACCENT_SOFT, corner_radius=9)
        self.favorite_selection_label = ctk.CTkLabel(
            self.favorite_selection_bar, text="", text_color=ACCENT,
            font=ctk.CTkFont(size=12, weight="bold"))
        self.favorite_selection_label.pack(side="left", padx=10, pady=5)
        VisualButton(
            self.favorite_selection_bar, text=I.t("favorite_clear_selection"),
            width=72, height=28, fg_color="transparent", text_color=ACCENT,
            hover_color=LINE, command=self.clear_favorite_selection).pack(
                side="right", padx=(3, 7), pady=4)
        VisualButton(
            self.favorite_selection_bar, text=I.t("favorite_add_selected"),
            width=130, height=28, fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            command=self.use_selected_favorites).pack(side="right", padx=3, pady=4)
        self.favorite_cards = ctk.CTkFrame(favorites, fg_color="transparent")
        self.favorite_cards.pack(fill="x", expand=False, padx=PAD, pady=(0, PAD))
        self.favorite_cards.grid_columnconfigure((0, 1, 2), weight=1, uniform="favorite")
        self._favorite_card_columns = 3
        self.favorite_cards.bind("<Configure>", self._on_favorite_cards_resize)

        self.class_page_header = self.page_header(
            self.pages["drug_classes"], I.t("drug_classes"), "")
        self.class_overview = ctk.CTkFrame(self.pages["drug_classes"], fg_color="transparent")
        self.class_overview.pack(fill="both", expand=True)
        class_page = self.section(self.class_overview, "")
        top = ctk.CTkFrame(class_page, fg_color="transparent")
        top.pack(fill="x", padx=PAD, pady=(10, 6))
        self.class_breadcrumb = ctk.CTkFrame(top, fg_color="transparent")
        self.class_breadcrumb.pack(side="left", fill="x", expand=True)
        self.class_breadcrumb_group = VisualButton(
            self.class_breadcrumb, text="", height=30, width=40,
            fg_color="transparent", text_color=ACCENT,
            hover_color=ACCENT_SOFT, anchor="w",
            font=ctk.CTkFont(size=13, weight="bold"))
        self.class_breadcrumb_group.pack(side="left")
        self.class_breadcrumb_separator = ctk.CTkLabel(
            self.class_breadcrumb, text="›", width=22, text_color=MUTED,
            font=ctk.CTkFont(size=14, weight="bold"))
        self.class_breadcrumb_separator.pack(side="left")
        self.class_breadcrumb_detail = VisualButton(
            self.class_breadcrumb, text="", height=30, width=40,
            fg_color="transparent", text_color=ACCENT,
            hover_color=ACCENT_SOFT, anchor="w",
            font=ctk.CTkFont(size=13, weight="bold"))
        self.class_breadcrumb_detail.pack(side="left")
        VisualButton(top, text=I.t("manage_class_mappings"), height=ACTION_HEIGHT, width=180,
                      fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
                      hover_color=ACCENT_SOFT, command=self.show_class_mapping_editor).pack(
                          side="right", padx=(10, 0))
        VisualButton(top, text=I.t("show_all_detailed_classes"), height=ACTION_HEIGHT, width=210,
                      fg_color=CARD, text_color=ACCENT, border_width=1,
                      border_color=LINE, hover_color=ACCENT_SOFT,
                      command=self.show_all_detailed_classes).pack(side="right", padx=(10, 0))
        polish_toolbar(top)
        self.class_search_var = tk.StringVar()
        self.class_search_var.trace_add(
            "write", lambda *_: self._schedule_class_browser_refresh())
        VisualEntry(
            class_page, textvariable=self.class_search_var, height=FIELD_HEIGHT,
            placeholder_text=I.t("search_classes_medicines"),
            border_color=LINE, corner_radius=9).pack(
                fill="x", padx=PAD, pady=(0, 4))
        self.class_unclassified_filter_var = tk.BooleanVar(value=False)

        self.class_browser = ctk.CTkFrame(class_page, fg_color="transparent")
        self.class_browser.pack(fill="both", expand=True, padx=PAD, pady=(0, PAD))
        self.class_left_panel = GlassFrame(
            self.class_browser, width=320, fg_color=SURFACE, border_color=LINE,
            border_width=1, corner_radius=12)
        self.class_left_panel.pack(side="left", fill="y", padx=(0, 8))
        self.class_left_panel.pack_propagate(False)
        ctk.CTkLabel(
            self.class_left_panel, text=I.t("major_therapeutic_groups"), text_color=ACCENT,
            font=ctk.CTkFont(size=14, weight="bold"), anchor="w").pack(
                fill="x", padx=10, pady=(10, 5))
        self.class_group_list = ctk.CTkScrollableFrame(
            self.class_left_panel, width=292, height=470, fg_color="transparent")
        self.class_group_list.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        self.class_right_panel = GlassFrame(
            self.class_browser, fg_color=SURFACE, border_color=LINE,
            border_width=1, corner_radius=12)
        self.class_right_panel.pack(side="left", fill="both", expand=True)
        ctk.CTkLabel(
            self.class_right_panel, text=I.t("detailed_drug_classes"), text_color=ACCENT,
            font=ctk.CTkFont(size=14, weight="bold"), anchor="w").pack(
                fill="x", padx=10, pady=(10, 5))
        self.class_detail_list = ctk.CTkScrollableFrame(
            self.class_right_panel, height=470, fg_color="transparent")
        self.class_detail_list.pack(
            fill="both", expand=True, padx=6, pady=(0, 6))

        self.class_unclassified_panel = ctk.CTkFrame(
            self.class_browser, fg_color=SURFACE, border_color=LINE,
            border_width=1, corner_radius=12)
        self.class_unclassified_heading = ctk.CTkLabel(
            self.class_unclassified_panel, text="", text_color=ACCENT,
            font=ctk.CTkFont(size=14, weight="bold"), anchor="w")
        self.class_unclassified_heading.pack(fill="x", padx=12, pady=(10, 5))
        self.class_unclassified_results = ctk.CTkScrollableFrame(
            self.class_unclassified_panel, height=470, fg_color="transparent")
        self.class_unclassified_results.pack(
            fill="both", expand=True, padx=8, pady=(0, 8))
        self.class_unclassified_results.grid_columnconfigure(
            (0, 1), weight=1, uniform="unclassified_drugs")

        self.class_buttons = {}
        self.class_tiles = {}
        self.class_count_badges = {}
        for code in classes.GROUPS:
            tile = ctk.CTkFrame(self.class_group_list, fg_color="transparent")
            button = VisualButton(
                tile, text=I.t("class_" + code), height=ACTION_HEIGHT, corner_radius=9,
                anchor="w", fg_color=SURFACE, text_color=TEXT,
                border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
                command=lambda selected=code: self.schedule_class_group_selection(selected))
            tile.grid_columnconfigure(0, weight=1)
            button.grid(row=0, column=0, sticky="ew")
            button.bind(
                "<Double-Button-1>",
                lambda _event, selected=code: self.open_therapeutic_group(selected))
            count_badge = ctk.CTkLabel(
                tile, text="0", width=28, height=28, corner_radius=14,
                fg_color=ACCENT_SOFT, text_color=ACCENT,
                font=ctk.CTkFont(size=11, weight="bold"))
            count_badge.grid(row=0, column=1, padx=(5, 0))
            self.class_buttons[code] = button
            self.class_tiles[code] = tile
            self.class_count_badges[code] = count_badge
        self.class_group_empty_label = ctk.CTkLabel(
            self.class_group_list, text=I.t("class_search_no_results"),
            text_color=MUTED, anchor="w")
        self._selected_therapeutic_group = classes.GROUPS[0]
        self._selected_detailed_class = (
            classes.subclasses_for(classes.GROUPS[0])[0]
            if classes.subclasses_for(classes.GROUPS[0]) else None)
        self.class_detail_buttons = {}
        self.class_visible_drugs = []
        self._class_visible_drugs = []
        self.class_subpage = ctk.CTkFrame(self.pages["drug_classes"], fg_color="transparent")
        self._active_subclass_group = None
        self._build_treatment_templates_page()

    # -- treatment templates ----------------------------------------------
    def _build_treatment_templates_page(self):
        """Build reusable disease regimens from medicines in the active database."""
        page = self.pages["treatment_templates"]
        self.page_header(page, I.t("treatment_templates"), "")
        navigation = ctk.CTkFrame(page, fg_color="transparent")
        navigation.pack(fill="x", padx=PAD, pady=(0, 6))
        self.treatment_new_view_button = VisualButton(
            navigation, text="＋ " + I.t("treatment_new_view"), width=150,
            height=ACTION_HEIGHT, fg_color=CARD, text_color=ACCENT,
            border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
            command=self.new_treatment_template)
        self.treatment_new_view_button.pack(side="left", padx=(0, 6))
        self.treatment_saved_view_button = VisualButton(
            navigation, text=I.t("treatment_saved_view"),
            image=_dashboard_icon("treatment_templates"), compound="left",
            width=170, height=ACTION_HEIGHT, fg_color=ACCENT_SOFT,
            text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=self.show_treatment_saved_templates)
        self.treatment_saved_view_button.pack(side="left")
        editor = self.section(page, "")
        self.treatment_editor_view = editor

        toolbar = ctk.CTkFrame(editor, fg_color="transparent")
        toolbar.pack(fill="x", padx=PAD, pady=(14, 7))
        VisualButton(toolbar, text=I.t("back"), width=65, height=ACTION_HEIGHT,
                      fg_color=CARD, text_color=ACCENT, border_width=1,
                      border_color=LINE, hover_color=ACCENT_SOFT,
                      command=self.show_treatment_saved_templates).pack(side="left", padx=(0, 7))
        ctk.CTkLabel(toolbar, text=I.t("treatment_disease"), text_color=MUTED,
                     font=_ui_font(12), anchor="w").pack(side="left", padx=(0, 7))
        self.treatment_disease_var = tk.StringVar()
        self.treatment_template_selector_var = tk.StringVar(
            value=I.t("treatment_select_template"))
        self.treatment_template_selector = VisualComboBox(
            toolbar, values=[I.t("treatment_select_template")],
            variable=self.treatment_template_selector_var, width=280,
            height=ACTION_HEIGHT, fg_color=ACCENT_SOFT, text_color=ACCENT,
            border_color=LINE, dropdown_fg_color=CARD,
            dropdown_hover_color=ACCENT_SOFT,
            button_color=PRIMARY, button_hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=12), dropdown_font=ctk.CTkFont(size=12),
            justify="left",
            command=self.load_treatment_template)
        self.treatment_template_selector.pack(side="left", padx=(0, 5))
        self.treatment_template_selector.bind(
            "<KeyRelease>", self._filter_treatment_template_menu)
        self.treatment_template_selector.bind(
            "<Return>", self._load_treatment_template_from_search)
        self.treatment_template_selector.bind(
            "<FocusOut>", self._on_treatment_template_focus_out)
        self.treatment_template_selector.bind(
            "<FocusIn>", self._on_treatment_template_focus_in)
        self._treatment_template_popup = None
        self._treatment_template_suggestion_buttons = []
        save_button = VisualButton(
            toolbar, text="💾", width=ICON_BUTTON_SIZE, height=ICON_BUTTON_SIZE,
            fg_color="transparent", text_color=ACCENT, border_width=0,
            hover_color=ACCENT_SOFT,
            font=ctk.CTkFont(size=17), command=self.save_treatment_template)
        save_button.pack(side="left", padx=3)
        self.treatment_delete_button = VisualButton(
            toolbar, text="🗑", width=ICON_BUTTON_SIZE,
            height=ICON_BUTTON_SIZE, fg_color="transparent", text_color=DANGER,
            border_width=0, hover_color=DANGER_SOFT,
            font=ctk.CTkFont(size=17),
            command=self.delete_treatment_template)
        self.treatment_delete_button.pack(side="left", padx=3)
        VisualButton(
            toolbar, text=I.t("treatment_use_rx"), width=105,
            height=ACTION_HEIGHT, fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            command=self.use_treatment_template).pack(side="right")
        polish_toolbar(toolbar)

        search_row = ctk.CTkFrame(editor, fg_color="transparent")
        search_row.pack(fill="x", padx=PAD, pady=(0, 5))
        ctk.CTkLabel(
            search_row, text=I.t("treatment_add_from_database"), width=125,
            text_color=MUTED, anchor="w",
            font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")
        self.treatment_drug_search_var = tk.StringVar()
        self.treatment_drug_search_var.trace_add(
            "write", lambda *_: self._schedule_treatment_drug_search())
        self.treatment_drug_search_entry = VisualEntry(
            search_row, textvariable=self.treatment_drug_search_var,
            height=FIELD_HEIGHT, border_color=LINE, corner_radius=9,
            placeholder_text=I.t("treatment_search_database"),
            font=ctk.CTkFont(size=14))
        self.treatment_drug_search_entry.pack(side="left", fill="x", expand=True)

        self.treatment_drug_results = ctk.CTkFrame(
            editor, fg_color=SURFACE, border_color=LINE,
            border_width=1, corner_radius=9)

        selected_head = ctk.CTkFrame(editor, fg_color="transparent")
        self.treatment_selected_head = selected_head
        selected_head.pack(fill="x", padx=PAD, pady=(2, 3))
        ctk.CTkLabel(
            selected_head, text=I.t("treatment_selected_medicines"),
            text_color=ACCENT, anchor="w",
            font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        self.treatment_count_label = ctk.CTkLabel(
            selected_head, text="0", width=28, height=24, corner_radius=12,
            fg_color=ACCENT_SOFT, text_color=ACCENT,
            font=ctk.CTkFont(size=11, weight="bold"))
        self.treatment_count_label.pack(side="left", padx=7)

        self.treatment_medications_frame = ctk.CTkFrame(
            editor, fg_color="transparent")
        self.treatment_medications_frame.pack(fill="x", padx=PAD, pady=(0, PAD))
        self.treatment_template_status = ctk.CTkLabel(
            editor, text="", text_color=MUTED, anchor="w",
            font=ctk.CTkFont(size=11))
        self.treatment_template_status.pack(fill="x", padx=PAD, pady=(0, 8))

        self._treatment_template_id = ""
        self._treatment_template_drugs = []
        self._treatment_template_item_vars = []
        self._treatment_template_bindings = []
        self._treatment_search_job = None
        self._treatment_template_lookup = {}
        self._treatment_view = "saved"
        self._capture_treatment_baseline()
        editor.pack_forget()
        self.treatment_saved_view = self.section(page, "")
        self.treatment_saved_search_var = tk.StringVar()
        self.treatment_saved_search_var.trace_add(
            "write", lambda *_: self.refresh_saved_treatment_templates())
        VisualEntry(
            self.treatment_saved_view, textvariable=self.treatment_saved_search_var,
            placeholder_text=I.t("treatment_saved_search"), height=FIELD_HEIGHT,
            border_color=LINE).pack(fill="x", padx=PAD, pady=(10, 6))
        self.treatment_saved_cards = ctk.CTkFrame(self.treatment_saved_view, fg_color="transparent")
        self.treatment_saved_cards.pack(fill="x", padx=PAD, pady=(0, PAD))
        self._treatment_saved_expanded_id = ""
        self._treatment_saved_highlight_id = ""
        self._treatment_saved_limit = 24

    @staticmethod
    def _treatment_draft_signature(disease, medicines):
        return (strip_bidi_display_controls(disease).strip(), tuple(
            tuple(sorted((key, str(value)) for key, value in medicine.items()
                         if not key.startswith("_"))) for medicine in medicines))

    def _capture_treatment_baseline(self):
        self._treatment_baseline_disease = self.treatment_disease_var.get()
        self._treatment_baseline_selector = self.treatment_template_selector_var.get()
        self._treatment_baseline_drugs = [dict(item) for item in self._treatment_template_drugs]
        self._treatment_baseline_signature = self._treatment_draft_signature(
            self._treatment_baseline_disease, self._treatment_baseline_drugs)

    def _confirm_treatment_leave(self):
        if getattr(self, "_treatment_view", "saved") != "editor":
            return True
        current = self._treatment_draft_signature(
            self._current_treatment_disease(), self._treatment_template_drugs)
        if current == self._treatment_baseline_signature:
            return True
        if not messagebox.askyesno(I.t("treatment_unsaved_title"),
                                   I.t("treatment_unsaved_message"), parent=self):
            return False
        self.treatment_disease_var.set(self._treatment_baseline_disease)
        self.treatment_template_selector_var.set(self._treatment_baseline_selector)
        self._treatment_template_drugs = [dict(item) for item in self._treatment_baseline_drugs]
        self.render_treatment_template_drugs()
        return True

    def _switch_treatment_view(self, view):
        self._hide_treatment_template_suggestions()
        self.treatment_editor_view.pack_forget()
        self.treatment_saved_view.pack_forget()
        target = self.treatment_editor_view if view == "editor" else self.treatment_saved_view
        target.pack(fill="x", pady=CARD_GAP)
        self._treatment_view = view
        self.treatment_new_view_button.configure(fg_color=ACCENT_SOFT if view == "editor" else CARD)
        self.treatment_saved_view_button.configure(fg_color=ACCENT_SOFT if view == "saved" else CARD)
        self.scroll._parent_canvas.yview_moveto(0)

    def show_treatment_saved_templates(self, selected_id=""):
        if not self._confirm_treatment_leave():
            return False
        self._switch_treatment_view("saved")
        if selected_id:
            self._treatment_saved_highlight_id = selected_id
            self._treatment_saved_limit = 24
            self.treatment_saved_search_var.set("")
        self.refresh_saved_treatment_templates()
        if selected_id:
            card = getattr(self, "_treatment_highlight_card", None)
            if card is not None:
                self.after_idle(lambda target=card: self._scroll_medication_row_into_view(target))
        return True

    def refresh_saved_treatment_templates(self):
        if not hasattr(self, "treatment_saved_cards"):
            return
        self._treatment_highlight_card = None
        for child in self.treatment_saved_cards.winfo_children():
            child.destroy()
        query = self.treatment_saved_search_var.get().strip().casefold()
        templates = sorted(cfg.config.treatment_templates(), key=lambda item: (
            strip_bidi_display_controls(item.get("disease", "")).casefold(), str(item.get("id", ""))))
        matches = [item for item in templates if query in
                   strip_bidi_display_controls(item.get("disease", "")).casefold()]
        if not matches:
            ctk.CTkLabel(self.treatment_saved_cards, text=I.t("treatment_saved_empty"),
                         text_color=MUTED, anchor="w").pack(fill="x", pady=8)
            return
        limit = max(24, self._treatment_saved_limit)
        highlighted = next((index for index, item in enumerate(matches)
                            if item["id"] == self._treatment_saved_highlight_id), -1)
        if highlighted >= limit:
            limit = highlighted + 1
        for template in matches[:limit]:
            template_id = template["id"]
            selected = template_id == self._treatment_saved_highlight_id
            card = GlassFrame(self.treatment_saved_cards,
                                fg_color=ACCENT_SOFT if selected else CARD,
                                border_color=ACCENT if selected else LINE,
                                border_width=1, corner_radius=10)
            card.pack(fill="x", pady=3)
            if selected:
                self._treatment_highlight_card = card
            header = ctk.CTkFrame(card, fg_color="transparent")
            header.pack(fill="x", padx=8, pady=5)
            actions = ctk.CTkFrame(header, fg_color="transparent")
            actions.pack(side="right")
            for symbol, image, command in (
                    ("+", None,
                     lambda item=template: self.use_saved_treatment_template(item)),
                    ("", _edit_icon(18),
                     lambda item=template: self.edit_saved_treatment_template(item)),
                    ("🗑", None,
                     lambda item=template: self.delete_saved_treatment_template(item))):
                button = VisualButton(actions, text=symbol, image=image, width=30, height=30,
                                       fg_color="transparent", border_width=0,
                                       text_color=DANGER if symbol == "🗑" else ACCENT,
                                       hover_color=ACCENT_SOFT, font=_ui_font(18), command=command)
                button.pack(side="left", padx=2)
            VisualButton(header, text=directional_display_text(I.t(
                "treatment_template_option", disease=template["disease"], n=len(template["medications"]))),
                height=32, anchor="w", fg_color="transparent", text_color=TEXT,
                hover_color=ACCENT_SOFT, font=_ui_font(14, "bold"),
                command=lambda picked=template_id: self.toggle_saved_treatment_template(picked)).pack(
                    side="left", fill="x", expand=True)
            if self._treatment_saved_expanded_id == template_id:
                for index, medicine in enumerate(template["medications"], 1):
                    primary, secondary = self._treatment_medicine_title(medicine)
                    name = primary + (" · " + secondary if secondary else "")
                    regimen = self._treatment_regimen_summary(medicine)
                    text = f"{index}. {name}" + ("   ·   " + regimen if regimen else "")
                    if medicine.get("alternative_to_previous"):
                        text = I.t("treatment_or") + "  " + text
                    label = ctk.CTkLabel(card, text=directional_display_text(text), anchor="w",
                                         justify="left", wraplength=650, text_color=TEXT, font=_ui_font(12))
                    label.pack(fill="x", padx=12, pady=(0, 5))
                    label.bind("<Configure>", lambda event, target=label:
                               target.configure(wraplength=max(100, event.width - 10)))
        if len(matches) > limit:
            VisualButton(self.treatment_saved_cards, text=I.t("load_more"), width=110,
                          height=32, command=self._load_more_saved_treatments).pack(anchor="w", pady=5)

    def _load_more_saved_treatments(self):
        self._treatment_saved_limit += 24
        self.refresh_saved_treatment_templates()

    def toggle_saved_treatment_template(self, template_id):
        self._treatment_saved_expanded_id = "" if self._treatment_saved_expanded_id == template_id else template_id
        self.refresh_saved_treatment_templates()

    def edit_saved_treatment_template(self, template):
        self._load_treatment_template_record(template)

    def use_saved_treatment_template(self, template):
        self._show_treatment_apply_preview(medicines=template["medications"])

    def delete_saved_treatment_template(self, template):
        if not messagebox.askyesno(I.t("treatment_delete"),
                                   I.t("treatment_delete_confirm"), parent=self):
            return
        if cfg.config.remove_treatment_template(template["id"]):
            cfg.config.add_recovery_item("template", template["disease"], template)
        self.refresh_treatment_template_menu()
        self.refresh_saved_treatment_templates()

    def _set_treatment_status(self, text, color=MUTED):
        self.treatment_template_status.configure(text=text, text_color=color)

    def refresh_treatment_template_menu(self, selected_id=""):
        templates = sorted(
            cfg.config.treatment_templates(),
            key=lambda item: (
                strip_bidi_display_controls(item.get("disease", "")).casefold(),
                str(item.get("id", "")).casefold()))
        lookup = {}
        values = [I.t("treatment_select_template")]
        for index, template in enumerate(templates, 1):
            label = I.t(
                "treatment_template_option", disease=template["disease"],
                n=len(template["medications"]))
            if label in lookup:
                label = f"{label} · {index}"
            lookup[label] = template["id"]
            values.append(label)
        self._treatment_template_lookup = lookup
        self._treatment_template_all_values = values
        self.treatment_template_selector.configure(values=values)
        selected_label = next(
            (label for label, template_id in lookup.items()
             if template_id == selected_id), I.t("treatment_select_template"))
        self.treatment_template_selector_var.set(selected_label)
        self.refresh_saved_treatment_templates()

    def _restore_treatment_template_values(self):
        if hasattr(self, "treatment_template_selector"):
            self.treatment_template_selector.configure(
                values=getattr(
                    self, "_treatment_template_all_values",
                    [I.t("treatment_select_template")]))

    def _filter_treatment_template_menu(self, event=None):
        """Filter the editable saved-template dropdown by disease name."""
        if event and event.keysym == "Escape":
            self._hide_treatment_template_suggestions()
            return
        if event and event.keysym in {
                "Up", "Down", "Return", "Tab", "Shift_L", "Shift_R",
                "Control_L", "Control_R", "Alt_L", "Alt_R"}:
            return
        query = strip_bidi_display_controls(
            self.treatment_template_selector_var.get()).strip().casefold()
        typed_value = strip_bidi_display_controls(
            self.treatment_template_selector_var.get()).strip()
        if (typed_value
                and typed_value not in self._treatment_template_lookup
                and typed_value not in {
                    I.t("treatment_select_template"),
                    I.t("treatment_no_template_match")}):
            self.treatment_disease_var.set(typed_value)
        if not query or query == I.t("treatment_select_template").casefold():
            values = getattr(
                self, "_treatment_template_all_values",
                [I.t("treatment_select_template")])
        else:
            values = [label for label in self._treatment_template_lookup
                      if query in strip_bidi_display_controls(label).casefold()]
        self.treatment_template_selector.configure(values=values or [I.t(
            "treatment_no_template_match")])
        matches = [value for value in values if value in self._treatment_template_lookup]
        if query and query != I.t("treatment_select_template").casefold() and matches:
            self._show_treatment_template_suggestions(matches)
        else:
            self._hide_treatment_template_suggestions()

    def _show_treatment_template_suggestions(self, matches):
        """Open saved treatment plans automatically beneath the top selector."""
        self._hide_treatment_template_suggestions()
        matches = sorted(
            matches,
            key=lambda value: strip_bidi_display_controls(value).casefold())
        if not matches:
            return
        self.update_idletasks()
        popup = tk.Toplevel(self)
        popup.withdraw()
        popup.wm_overrideredirect(True)
        try:
            popup.attributes("-topmost", True)
        except tk.TclError:
            pass
        width = max(540, int(self.treatment_template_selector.winfo_width() * 1.5))
        width = min(width, popup.winfo_screenwidth() - 20)
        x = self.treatment_template_selector.winfo_rootx()
        x = max(10, min(x, popup.winfo_screenwidth() - width - 10))
        anchor_top = self.treatment_template_selector.winfo_rooty()
        anchor_bottom = anchor_top + self.treatment_template_selector.winfo_height()
        shell = ctk.CTkFrame(
            popup, fg_color=CARD, border_color=ACCENT,
            border_width=2, corner_radius=8)
        shell.pack(fill="both", expand=True)
        self._treatment_template_popup = popup
        self._treatment_template_suggestion_buttons = []
        visible_rows = min(len(matches), 7)
        container = shell
        is_scrollable = len(matches) > 3
        if is_scrollable:
            container = ctk.CTkScrollableFrame(
                shell, fg_color=CARD, corner_radius=6,
                scrollbar_button_color=PRIMARY,
                scrollbar_button_hover_color=ACCENT_HOVER)
            container.pack(fill="both", expand=True, padx=3, pady=3)
        for index, label in enumerate(matches):
            button = VisualButton(
                container, text=directional_display_text(label), height=46,
                corner_radius=5, fg_color="transparent", text_color=TEXT,
                hover_color=ACCENT_SOFT, anchor="w",
                font=dropdown_font(_ui_font(18)),
                command=lambda value=label: self._choose_treatment_template(value))
            bottom_inset = 4 if index == len(matches) - 1 else 0
            button.pack(fill="x", padx=4, pady=(4, bottom_inset))
            self._treatment_template_suggestion_buttons.append(button)
        popup.update_idletasks()
        screen_height = popup.winfo_screenheight()
        max_height = max(96, screen_height - 24)
        if is_scrollable:
            row_height = max(
                button.winfo_reqheight()
                for button in self._treatment_template_suggestion_buttons) + 8
            height = min(
                visible_rows * row_height + 18,
                max(180, int(screen_height * 0.55)), max_height)
        else:
            height = min(max(48, shell.winfo_reqheight() + 12), max_height)
        below_y = anchor_bottom + 2
        if below_y + height <= screen_height - 12:
            y = below_y
        elif anchor_top - height >= 12:
            y = anchor_top - height - 2
        else:
            y = max(12, min(below_y, screen_height - height - 12))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.deiconify()
        popup.lift()

    def _hide_treatment_template_suggestions(self):
        popup = getattr(self, "_treatment_template_popup", None)
        self._treatment_template_popup = None
        self._treatment_template_suggestion_buttons = []
        if popup is not None:
            try:
                if popup.winfo_exists():
                    popup.destroy()
            except tk.TclError:
                pass

    def _on_treatment_template_focus_out(self, _event=None):
        self._restore_treatment_template_values()
        self.after(160, self._hide_treatment_template_suggestions)

    def _on_treatment_template_focus_in(self, _event=None):
        current = strip_bidi_display_controls(
            self.treatment_template_selector_var.get()).strip()
        if current in {
                I.t("treatment_select_template"),
                I.t("treatment_no_template_match")}:
            self.treatment_template_selector_var.set("")

    def _choose_treatment_template(self, label):
        self._hide_treatment_template_suggestions()
        self.treatment_template_selector_var.set(label)
        self.load_treatment_template(label)

    def _load_treatment_template_from_search(self, _event=None):
        self._hide_treatment_template_suggestions()
        value = strip_bidi_display_controls(
            self.treatment_template_selector_var.get()).strip()
        if value in self._treatment_template_lookup:
            self.load_treatment_template(value)
            return "break"
        matches = [label for label in self._treatment_template_lookup
                   if value.casefold() in strip_bidi_display_controls(label).casefold()]
        if len(matches) == 1:
            self.treatment_template_selector_var.set(matches[0])
            self.load_treatment_template(matches[0])
        return "break"

    def new_treatment_template(self):
        if not self._confirm_treatment_leave():
            return False
        self._hide_treatment_template_suggestions()
        self._treatment_template_id = ""
        self.treatment_disease_var.set("")
        self.treatment_drug_search_var.set("")
        self._treatment_template_drugs = []
        self.refresh_treatment_template_menu()
        self.treatment_template_selector_var.set("")
        self.render_treatment_template_drugs()
        self._capture_treatment_baseline()
        self._switch_treatment_view("editor")
        self._set_treatment_status(I.t("treatment_new_ready"), ACCENT)
        self.treatment_template_selector.focus_set()

    def load_treatment_template(self, label):
        self._hide_treatment_template_suggestions()
        template_id = self._treatment_template_lookup.get(label, "")
        if not template_id:
            return
        template = next(
            (item for item in cfg.config.treatment_templates()
             if item["id"] == template_id), None)
        if not template:
            return
        self._load_treatment_template_record(template)

    def _load_treatment_template_record(self, template):
        """Load a saved template directly, including all of its medicines."""
        if not self._confirm_treatment_leave():
            self.treatment_template_selector_var.set(self.treatment_disease_var.get())
            return False
        template_id = template.get("id", "")
        self._treatment_template_id = template_id
        self.treatment_disease_var.set(template["disease"])
        self.treatment_template_selector_var.set(template["disease"])
        self._treatment_template_drugs = [
            dict(item, _editor_open=False) for item in template["medications"]]
        self.render_treatment_template_drugs()
        self._capture_treatment_baseline()
        self._switch_treatment_view("editor")
        self._set_treatment_status(I.t("treatment_loaded"), ACCENT)

    def _current_treatment_disease(self):
        """Use the top selector as both saved-template search and disease input."""
        entered = strip_bidi_display_controls(
            self.treatment_template_selector_var.get()).strip()
        if (entered and entered not in self._treatment_template_lookup
                and entered not in {
                    I.t("treatment_select_template"),
                    I.t("treatment_no_template_match")}):
            self.treatment_disease_var.set(entered)
        return self.treatment_disease_var.get().strip()

    def _schedule_treatment_drug_search(self):
        if self._treatment_search_job is not None:
            try:
                self.after_cancel(self._treatment_search_job)
            except (tk.TclError, ValueError):
                pass
        self._treatment_search_job = self.after(160, self.refresh_treatment_drug_results)

    def refresh_treatment_drug_results(self):
        self._treatment_search_job = None
        for child in self.treatment_drug_results.winfo_children():
            child.destroy()
        query = self.treatment_drug_search_var.get().strip()
        if not query:
            self.treatment_drug_results.pack_forget()
            return
        if not self.treatment_drug_results.winfo_manager():
            self.treatment_drug_results.pack(
                fill="x", padx=(PAD + 125, PAD), pady=(0, 7),
                before=self.treatment_selected_head)
        token = self._latest_query_tokens.get("treatment_drug", 0) + 1
        self._latest_query_tokens["treatment_drug"] = token
        self.submit_background(
            lambda: self.db.search_prescribable(query, 8),
            lambda matches: self._render_treatment_drug_results(
                query, token, matches), silent=True)

    def _render_treatment_drug_results(self, query, token, matches):
        if (token != self._latest_query_tokens.get("treatment_drug")
                or self.treatment_drug_search_var.get().strip() != query):
            return
        for child in self.treatment_drug_results.winfo_children():
            child.destroy()
        if not matches:
            ctk.CTkLabel(
                self.treatment_drug_results, text=I.t("treatment_no_database_match"),
                text_color=MUTED, anchor="w").pack(fill="x", padx=10, pady=7)
            return
        for drug in matches:
            brand = drug.brand_name.strip() or drug.generic_name.strip()
            scientific = drug.generic_name.strip() if drug.brand_name.strip() else ""
            text = brand + (f"  ·  {scientific}" if scientific else "")
            VisualButton(
                self.treatment_drug_results, text="+  " + text, height=32,
                fg_color="transparent", text_color=TEXT, anchor="w",
                hover_color=ACCENT_SOFT,
                command=lambda item=drug: self.add_treatment_database_drug(item)).pack(
                    fill="x", padx=4, pady=1)

    def add_treatment_database_drug(self, drug):
        medicine = {
            "generic_name": drug.generic_name.strip(),
            "brand_name": drug.brand_name.strip(),
            "dosage": "", "frequency": "", "duration": "", "notes": "",
            "alternative_to_previous": False, "_editor_open": True,
        }
        identity = (medicine["generic_name"].casefold(), medicine["brand_name"].casefold())
        if any((item["generic_name"].casefold(), item["brand_name"].casefold()) == identity
               for item in self._treatment_template_drugs):
            self._set_treatment_status(I.t("treatment_duplicate_drug"), WARNING)
            return
        self._treatment_template_drugs.append(medicine)
        self.treatment_drug_search_var.set("")
        self.render_treatment_template_drugs()
        self._set_treatment_status(I.t("treatment_drug_added"), ACCENT)

    def _update_treatment_drug_field(self, medicine, key, variable):
        medicine[key] = variable.get().strip()

    def _set_treatment_alternative(self, medicine, variable):
        medicine["alternative_to_previous"] = bool(variable.get())
        self.render_treatment_template_drugs()

    def toggle_treatment_drug_editor(self, index):
        if 0 <= index < len(self._treatment_template_drugs):
            medicine = self._treatment_template_drugs[index]
            medicine["_editor_open"] = not bool(medicine.get("_editor_open"))
            self.render_treatment_template_drugs()

    @staticmethod
    def _treatment_choice_groups(medicines):
        """Group sequential `OR` alternatives with the medicine before them."""
        groups = []
        for medicine in medicines:
            if medicine.get("alternative_to_previous") and groups:
                groups[-1].append(medicine)
            else:
                groups.append([medicine])
        return groups

    @staticmethod
    def _treatment_medicine_title(medicine):
        brand = str(medicine.get("brand_name", "")).strip()
        scientific = str(medicine.get("generic_name", "")).strip()
        return brand or scientific, scientific if brand and scientific.casefold() != brand.casefold() else ""

    @staticmethod
    def _treatment_regimen_summary(medicine):
        parts = []
        for key, label in (("dosage", I.t("dosage")),
                           ("frequency", I.t("frequency")),
                           ("duration", I.t("duration")),
                           ("notes", I.t("notes"))):
            value = str(medicine.get(key, "")).strip()
            if value:
                parts.append(f"{label}: {value}")
        return "   ·   ".join(parts)

    def render_treatment_template_drugs(self):
        for binding in self._treatment_template_bindings:
            binding.dispose()
        self._treatment_template_bindings = []
        for child in self.treatment_medications_frame.winfo_children():
            child.destroy()
        self._treatment_template_item_vars = []
        self.treatment_count_label.configure(text=str(len(self._treatment_template_drugs)))
        if not self._treatment_template_drugs:
            ctk.CTkLabel(
                self.treatment_medications_frame, text=I.t("treatment_no_medicines"),
                text_color=MUTED, anchor="w").pack(fill="x", pady=10)
            return
        for index, medicine in enumerate(self._treatment_template_drugs):
            card = GlassFrame(
                self.treatment_medications_frame, fg_color=CARD,
                border_color=LINE, border_width=1, corner_radius=9)
            card.pack(fill="x", pady=3)
            header = ctk.CTkFrame(card, fg_color="transparent")
            header.pack(fill="x", padx=9, pady=(5, 3))
            ctk.CTkLabel(
                header, text=f"{index + 1}.", text_color=TEXT,
                font=ctk.CTkFont(size=13, weight="bold"), anchor="w").pack(side="left")

            actions = ctk.CTkFrame(header, fg_color="transparent")
            actions.pack(side="right")
            VisualButton(
                actions, text=I.t("treatment_remove"), width=68, height=27,
                fg_color="transparent", text_color=DANGER, hover_color=DANGER_SOFT,
                command=lambda position=index: self.remove_treatment_drug(position)).pack(
                    side="right", padx=(4, 0))
            VisualButton(
                actions, text="↓", width=30, height=27, fg_color="transparent",
                text_color=ACCENT, hover_color=ACCENT_SOFT,
                command=lambda position=index: self.move_treatment_drug(position, 1)).pack(
                    side="right", padx=2)
            VisualButton(
                actions, text="↑", width=30, height=27, fg_color="transparent",
                text_color=ACCENT, hover_color=ACCENT_SOFT,
                command=lambda position=index: self.move_treatment_drug(position, -1)).pack(
                    side="right", padx=2)
            treatment_edit_button = VisualButton(
                actions, text="", image=_edit_icon(18), width=30, height=27,
                fg_color="transparent", text_color=ACCENT,
                border_width=0, hover_color=ACCENT_SOFT,
                font=ctk.CTkFont(size=15, weight="bold"),
                command=lambda position=index: self.toggle_treatment_drug_editor(position))
            treatment_edit_button.pack(side="right", padx=2)

            if medicine.get("alternative_to_previous"):
                ctk.CTkLabel(
                    header, text=I.t("treatment_or"), width=32, height=22,
                    corner_radius=11, fg_color=WARNING_SOFT, text_color=WARNING,
                    font=ctk.CTkFont(size=10, weight="bold")).pack(
                        side="left", padx=(7, 3))
            primary, secondary = self._treatment_medicine_title(medicine)
            ctk.CTkLabel(
                header, text=primary or I.t("drug"), text_color=TEXT,
                font=ctk.CTkFont(size=14, weight="bold"), anchor="w").pack(
                    side="left", padx=(7, 0))
            if secondary:
                ctk.CTkLabel(
                    header, text=f"  ·  {secondary}", text_color=MUTED,
                    font=ctk.CTkFont(size=13), anchor="w").pack(side="left")

            summary = self._treatment_regimen_summary(medicine)
            if summary:
                ctk.CTkLabel(
                    header, text=f"   ·   {summary}", text_color=MUTED,
                    anchor="w", justify="left",
                    font=ctk.CTkFont(size=11)).pack(side="left", padx=(4, 0))

            variables = {}
            if not medicine.get("_editor_open"):
                self._treatment_template_item_vars.append(variables)
                continue
            name_fields = ctk.CTkFrame(card, fg_color="transparent")
            name_fields.pack(fill="x", padx=9, pady=(2, 5))
            name_fields.grid_columnconfigure((0, 1), weight=1, uniform="treatment_names")
            for column, (key, label) in enumerate((
                    ("brand_name", I.t("generic_trade_name")),
                    ("generic_name", I.t("scientific_name")))):
                column_frame = ctk.CTkFrame(name_fields, fg_color="transparent")
                column_frame.grid(
                    row=0, column=column, sticky="ew",
                    padx=(0 if column == 0 else 4, 0 if column == 1 else 4))
                ctk.CTkLabel(
                    column_frame, text=label, text_color=MUTED, anchor="w",
                    font=ctk.CTkFont(size=10, weight="bold")).pack(
                        fill="x", pady=(0, 2))
                variable = tk.StringVar(value=medicine[key])
                variable.trace_add(
                    "write", lambda *_args, item=medicine, field=key, var=variable:
                    self._update_treatment_drug_field(item, field, var))
                variables[key] = variable
                VisualEntry(
                    column_frame, textvariable=variable, height=36, corner_radius=8,
                    border_color=LINE,
                    font=ctk.CTkFont(size=13, weight="bold" if key == "brand_name" else "normal")
                ).pack(fill="x")
            fields = ctk.CTkFrame(card, fg_color="transparent")
            fields.pack(fill="x", padx=9, pady=(0, 7))
            fields.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="regimen")
            for column, (key, label, options) in enumerate((
                    ("dosage", I.t("dosage"), None),
                    ("frequency", I.t("frequency"), FREQUENCY_OPTIONS),
                    ("duration", I.t("duration"), None),
                    ("notes", I.t("notes"), NOTE_OPTIONS))):
                column_frame = ctk.CTkFrame(fields, fg_color="transparent")
                column_frame.grid(
                    row=0, column=column, sticky="ew",
                    padx=(0 if column == 0 else 3, 0 if column == 3 else 3))
                ctk.CTkLabel(
                    column_frame, text=label, text_color=MUTED, anchor="w",
                    font=ctk.CTkFont(size=10, weight="bold")).pack(
                        fill="x", pady=(0, 2))
                variable = tk.StringVar(value=medicine[key])
                variable.trace_add(
                    "write", lambda *_args, item=medicine, field=key, var=variable:
                    self._update_treatment_drug_field(item, field, var))
                variables[key] = variable
                binding = DirectionalTextBinding(self, variable)
                self._treatment_template_bindings.append(binding)
                if options:
                    widget = VisualComboBox(
                        column_frame,
                        values=[directional_display_text(value) for value in options],
                        variable=binding.display_var, height=34, corner_radius=8,
                        border_width=1, border_color=ACCENT, fg_color=CARD,
                        button_color=PRIMARY, button_hover_color=ACCENT_HOVER,
                        dropdown_fg_color=CARD, dropdown_hover_color=ACCENT_SOFT,
                        font=ctk.CTkFont(size=12),
                        dropdown_font=ctk.CTkFont(size=12), justify="left")
                else:
                    widget = VisualEntry(
                        column_frame, textvariable=binding.display_var,
                        height=34, corner_radius=8, border_color=LINE,
                        font=ctk.CTkFont(size=12), justify="left")
                binding.attach(widget)
                widget.pack(fill="x")

            relationship = ctk.CTkFrame(card, fg_color="transparent")
            relationship.pack(fill="x", padx=9, pady=(0, 7))
            alternative_var = tk.BooleanVar(
                value=bool(medicine.get("alternative_to_previous")))
            variables["alternative_to_previous"] = alternative_var
            alternative_check = ctk.CTkCheckBox(
                relationship, text=I.t("treatment_alternative_previous"),
                variable=alternative_var, height=26, fg_color=PRIMARY,
                hover_color=ACCENT_HOVER, border_color=LINE,
                font=ctk.CTkFont(size=11),
                command=lambda item=medicine, var=alternative_var:
                    self._set_treatment_alternative(item, var))
            if index == 0:
                alternative_check.configure(state="disabled")
            alternative_check.pack(side="left")
            self._treatment_template_item_vars.append(variables)

    def remove_treatment_drug(self, index):
        if 0 <= index < len(self._treatment_template_drugs):
            self._treatment_template_drugs.pop(index)
            if self._treatment_template_drugs:
                self._treatment_template_drugs[0]["alternative_to_previous"] = False
            self.render_treatment_template_drugs()

    def move_treatment_drug(self, index, delta):
        target = index + delta
        if not (0 <= index < len(self._treatment_template_drugs)
                and 0 <= target < len(self._treatment_template_drugs)):
            return
        self._treatment_template_drugs[index], self._treatment_template_drugs[target] = (
            self._treatment_template_drugs[target], self._treatment_template_drugs[index])
        if self._treatment_template_drugs:
            self._treatment_template_drugs[0]["alternative_to_previous"] = False
        self.render_treatment_template_drugs()

    def save_treatment_template(self):
        disease = self._current_treatment_disease()
        if not disease:
            self._set_treatment_status(I.t("treatment_disease_required"), DANGER)
            self.treatment_template_selector.focus_set()
            return
        if not self._treatment_template_drugs:
            self._set_treatment_status(I.t("treatment_medicine_required"), DANGER)
            self.treatment_drug_search_entry.focus_set()
            return
        template_id = self._treatment_template_id or uuid.uuid4().hex
        saved_id = cfg.config.save_treatment_template({
            "id": template_id, "disease": disease,
            "variant": "",
            "medications": self._treatment_template_drugs,
        })
        if not saved_id:
            self._set_treatment_status(I.t("treatment_save_failed"), DANGER)
            return
        self._treatment_template_id = saved_id
        self.refresh_treatment_template_menu(saved_id)
        self._capture_treatment_baseline()
        self._set_treatment_status(I.t("treatment_saved"), ACCENT)
        self.show_treatment_saved_templates(selected_id=saved_id)

    def delete_treatment_template(self):
        if not self._treatment_template_id:
            self._set_treatment_status(I.t("treatment_select_first"), WARNING)
            return
        if not messagebox.askyesno(
                I.t("treatment_delete"), I.t("treatment_delete_confirm"), parent=self):
            return
        deleted = next((item for item in cfg.config.treatment_templates()
                        if item.get("id") == self._treatment_template_id), None)
        cfg.config.remove_treatment_template(self._treatment_template_id)
        if deleted:
            cfg.config.add_recovery_item(
                "template", deleted.get("disease", I.t("treatment_templates")), deleted)
        self._capture_treatment_baseline()
        self.new_treatment_template()
        self._set_treatment_status(I.t("treatment_deleted"), ACCENT)
        self.show_treatment_saved_templates()

    @staticmethod
    def _treatment_review_rows(templates):
        rows = []
        for template in templates:
            for step, medicine in enumerate(template.get("medications", []), 1):
                rows.append([
                    template.get("disease", ""), step,
                    "OR" if medicine.get("alternative_to_previous") else "Standard",
                    medicine.get("brand_name", ""), medicine.get("generic_name", ""),
                    medicine.get("dosage", ""), medicine.get("frequency", ""),
                    medicine.get("duration", ""), medicine.get("notes", ""),
                ])
        return rows

    @staticmethod
    def _read_treatment_templates_file(path):
        """Read editable treatment templates from CSV, XLSX, or legacy XLS."""
        source = Path(path)
        suffix = source.suffix.casefold()
        if suffix == ".csv":
            with source.open("r", encoding="utf-8-sig", newline="") as handle:
                source_rows = list(csv.reader(handle))
        elif suffix == ".xlsx":
            from openpyxl import load_workbook
            workbook = load_workbook(source, read_only=True, data_only=True)
            try:
                source_rows = list(workbook.active.iter_rows(values_only=True))
            finally:
                workbook.close()
        elif suffix == ".xls":
            import xlrd
            workbook = xlrd.open_workbook(source)
            sheet = workbook.sheet_by_index(0)
            source_rows = [sheet.row_values(index) for index in range(sheet.nrows)]
        else:
            raise ValueError(I.t("treatment_import_format"))

        def normalized(value):
            return "".join(
                character for character in str(value or "").casefold()
                if character.isalnum())

        translated_columns = {
            "treatment_disease": "disease",
            "treatment_relationship": "relationship",
            "generic_trade_name": "brand_name",
            "scientific_name": "generic_name",
            "dosage": "dosage",
            "frequency": "frequency",
            "duration": "duration",
            "notes": "notes",
        }
        aliases = {}
        for language in ("en", "ar"):
            strings = I.STRINGS.get(language, {})
            for translation_key, field_name in translated_columns.items():
                aliases[normalized(strings.get(translation_key, translation_key))] = field_name
        aliases.update({
            "disease": "disease", "indication": "disease",
            "brandname": "brand_name", "tradename": "brand_name",
            "generictradename": "brand_name", "drugname": "brand_name",
            "genericname": "generic_name", "scientificname": "generic_name",
            "activeingredient": "generic_name", "relationship": "relationship",
            "dose": "dosage", "dosage": "dosage", "frequency": "frequency",
            "duration": "duration", "note": "notes", "notes": "notes",
        })
        rows = iter(source_rows)
        header = next(rows, None)
        if not header:
            raise ValueError(I.t("treatment_import_columns"))
        columns = {}
        for index, value in enumerate(header):
            field_name = aliases.get(normalized(value))
            if field_name and field_name not in columns:
                columns[field_name] = index
        if "disease" not in columns or not {
                "brand_name", "generic_name"}.intersection(columns):
            raise ValueError(I.t("treatment_import_columns"))

        def cell(row, field_name):
            index = columns.get(field_name)
            if index is None or index >= len(row) or row[index] is None:
                return ""
            return str(row[index]).strip()

        templates = {}
        disease_order = []
        current_disease = ""
        for row in rows:
            entered_disease = cell(row, "disease")
            if entered_disease:
                current_disease = entered_disease
            if not current_disease:
                continue
            brand_name = cell(row, "brand_name")
            generic_name = cell(row, "generic_name")
            if not brand_name and not generic_name:
                continue
            disease_key = current_disease.casefold()
            if disease_key not in templates:
                templates[disease_key] = {
                    "disease": current_disease, "medications": []}
                disease_order.append(disease_key)
            relationship = normalized(cell(row, "relationship"))
            medicines = templates[disease_key]["medications"]
            medicines.append({
                "brand_name": brand_name,
                "generic_name": generic_name,
                "dosage": cell(row, "dosage"),
                "frequency": cell(row, "frequency"),
                "duration": cell(row, "duration"),
                "notes": cell(row, "notes"),
                "alternative_to_previous": bool(medicines) and relationship in {
                    "or", "alternative", "بديل", "أو"},
            })
        return sorted(
            (templates[key] for key in disease_order),
            key=lambda item: item["disease"].casefold())

    @staticmethod
    def _read_treatment_templates_xlsx(path):
        """Compatibility wrapper retained for older tests and integrations."""
        return App._read_treatment_templates_file(path)

    def import_treatment_templates_xlsx(self, parent=None, on_done=None):
        dialog_parent = parent or self
        path = filedialog.askopenfilename(
            parent=dialog_parent, title=I.t("treatment_import_database"),
            filetypes=[
                ("Treatment database", "*.csv *.xlsx *.xls"),
                ("CSV", "*.csv"), ("Excel workbook", "*.xlsx"),
                ("Legacy Excel workbook", "*.xls")])
        if not path:
            return 0
        def loaded(templates):
            if not templates:
                messagebox.showwarning(
                    I.t("treatment_import_database"), I.t("treatment_import_invalid"),
                    parent=dialog_parent)
                return
            replace = messagebox.askyesnocancel(
                I.t("treatment_import_database"), I.t("treatment_import_choice"),
                parent=dialog_parent)
            if replace is None:
                return

            def saved(count):
                self.refresh_treatment_template_menu()
                self.show_treatment_saved_templates()
                if on_done:
                    on_done()

            self.submit_background(
                lambda: cfg.config.merge_treatment_templates(
                    templates, replace=bool(replace)), saved,
                on_error=lambda exc: messagebox.showerror(
                    I.t("treatment_import_database"), str(exc), parent=dialog_parent),
                label=I.t("treatment_import_database") + "…")

        self.submit_background(
            lambda: self._read_treatment_templates_file(path), loaded,
            on_error=lambda exc: messagebox.showerror(
                I.t("treatment_import_database"), str(exc), parent=dialog_parent),
            label=I.t("treatment_import_database") + "…")
        return 0

    @staticmethod
    def _write_treatment_templates_file(path, headers, rows):
        """Write treatment-plan rows as CSV, XLSX, or legacy XLS."""
        suffix = Path(path).suffix.casefold()
        if suffix == ".csv":
            with open(path, "w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerow(headers)
                writer.writerows(rows)
            return
        if suffix == ".xlsx":
            from openpyxl import Workbook
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Treatment Templates"
            sheet.append(headers)
            for row in rows:
                sheet.append(row)
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for column in sheet.columns:
                width = min(45, max(12, max(len(str(cell.value or ""))
                                            for cell in column) + 2))
                sheet.column_dimensions[column[0].column_letter].width = width
            workbook.save(path)
            return
        if suffix == ".xls":
            import xlwt
            workbook = xlwt.Workbook(encoding="utf-8")
            sheet = workbook.add_sheet("Treatment Templates")
            for column, heading in enumerate(headers):
                sheet.write(0, column, heading)
            for row_index, row in enumerate(rows, 1):
                for column, value in enumerate(row):
                    sheet.write(row_index, column, value)
            workbook.save(path)
            return
        raise ValueError(I.t("treatment_import_format"))

    def export_treatment_templates_review(self, parent=None, on_done=None):
        dialog_parent = parent or self
        templates = cfg.config.treatment_templates()
        if not templates:
            self._set_treatment_status(I.t("treatment_no_saved_templates"), WARNING)
            messagebox.showwarning(
                I.t("treatment_export_database"), I.t("treatment_no_saved_templates"),
                parent=dialog_parent)
            return ""
        path = filedialog.asksaveasfilename(
            parent=dialog_parent, title=I.t("treatment_export_database"),
            defaultextension=".xlsx",
            filetypes=[
                ("Excel workbook", "*.xlsx"), ("CSV", "*.csv"),
                ("Legacy Excel workbook", "*.xls")])
        if not path:
            return ""
        headers = [
            I.t("treatment_disease"), I.t("treatment_step"),
            I.t("treatment_relationship"),
            I.t("generic_trade_name"), I.t("scientific_name"),
            I.t("dosage"), I.t("frequency"), I.t("duration"), I.t("notes"),
        ]
        rows = self._treatment_review_rows(templates)
        def finished(_result):
            self._set_treatment_status(I.t("treatment_export_done", path=path), ACCENT)
            if on_done:
                on_done()

        self.submit_background(
            lambda: self._write_treatment_templates_file(path, headers, rows),
            finished,
            on_error=lambda exc: messagebox.showerror(
                I.t("treatment_export_database"), str(exc), parent=dialog_parent),
            label=I.t("treatment_export_database") + "…")
        return path

    def use_treatment_template(self):
        if not self._treatment_template_drugs:
            self._set_treatment_status(I.t("treatment_medicine_required"), DANGER)
            return
        self._show_treatment_apply_preview()

    def _show_treatment_apply_preview(self, medicines=None):
        """Preview mandatory steps and choose one medicine from every OR group."""
        groups = self._treatment_choice_groups(
            self._treatment_template_drugs if medicines is None else medicines)
        dialog = ctk.CTkToplevel(self)
        dialog.title(I.t("treatment_apply_preview"))
        dialog.geometry("760x590")
        dialog.minsize(620, 460)
        dialog.transient(self)
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        ctk.CTkLabel(
            dialog, text=I.t("treatment_apply_preview"), text_color=ACCENT,
            font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(
                fill="x", padx=18, pady=(16, 2))
        ctk.CTkLabel(
            dialog, text=I.t("treatment_apply_preview_help"), text_color=MUTED,
            font=ctk.CTkFont(size=12), anchor="w", justify="left",
            wraplength=700).pack(fill="x", padx=18, pady=(0, 10))
        body = ctk.CTkScrollableFrame(dialog, fg_color=BG, corner_radius=10)
        body.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        selections = []
        for step, group in enumerate(groups, 1):
            card = GlassFrame(
                body, fg_color=CARD, border_color=LINE,
                border_width=1, corner_radius=9)
            card.pack(fill="x", pady=4)
            heading = I.t("treatment_step_number", n=step)
            if len(group) > 1:
                heading += "  ·  " + I.t("treatment_choose_one")
            ctk.CTkLabel(
                card, text=heading, text_color=ACCENT, anchor="w",
                font=ctk.CTkFont(size=12, weight="bold")).pack(
                    fill="x", padx=10, pady=(7, 3))
            selected = tk.IntVar(value=0)
            selections.append((group, selected))
            for option, medicine in enumerate(group):
                row = ctk.CTkFrame(card, fg_color="transparent")
                row.pack(fill="x", padx=10, pady=(1, 5))
                primary, secondary = self._treatment_medicine_title(medicine)
                title = primary + (f"  ·  {secondary}" if secondary else "")
                summary = self._treatment_regimen_summary(medicine)
                text = title if not summary else f"{title}\n{summary}"
                if len(group) > 1:
                    ctk.CTkRadioButton(
                        row, text=text, variable=selected, value=option,
                        fg_color=PRIMARY, hover_color=ACCENT_HOVER,
                        text_color=TEXT, font=ctk.CTkFont(size=12),
                        command=lambda: None).pack(
                            fill="x", padx=2, pady=2, anchor="w")
                else:
                    ctk.CTkLabel(
                        row, text=text, text_color=TEXT, anchor="w",
                        justify="left", font=ctk.CTkFont(size=12)).pack(
                            fill="x", padx=2, pady=2)
        actions = ctk.CTkFrame(dialog, fg_color="transparent")
        actions.pack(fill="x", padx=18, pady=(0, 14))
        VisualButton(
            actions, text=I.t("settings_cancel"), width=90, height=ACTION_HEIGHT,
            fg_color=CARD, text_color=MUTED, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=dialog.destroy).pack(side="left")
        VisualButton(
            actions, text=I.t("treatment_add_current"), width=150,
            height=ACTION_HEIGHT, fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            command=lambda: self._apply_treatment_selection(
                selections, False, dialog)).pack(side="right", padx=(6, 0))
        VisualButton(
            actions, text=I.t("treatment_replace_current"), width=165,
            height=ACTION_HEIGHT, fg_color=CARD, text_color=ACCENT,
            border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
            command=lambda: self._apply_treatment_selection(
                selections, True, dialog)).pack(side="right")
        dialog.grab_set()
        dialog.focus_set()

    def _apply_treatment_selection(self, selections, replace, dialog):
        if not self._confirm_treatment_leave():
            return
        selected = [group[choice.get()] for group, choice in selections]
        if replace:
            for row in self.rows:
                row.destroy()
            self.rows.clear()
        elif len(self.rows) == 1:
            row = self.rows[0]
            if not any((row.name_var.get().strip(), row.trade_var.get().strip(),
                        row.dosage_var.get().strip(), row.freq_var.get().strip(),
                        row.dur_var.get().strip(), row.notes_var.get().strip())):
                row.destroy()
                self.rows.clear()
        first_added = None
        for medicine in selected:
            payload = {key: str(medicine.get(key, "")) for key in (
                "generic_name", "brand_name", "dosage", "frequency", "duration", "notes")}
            row = self.add_row(data=qu.DrugItem(**payload))
            first_added = first_added or row
        dialog.grab_release()
        dialog.destroy()
        self.show_page("medications")
        self.on_any_change()
        if first_added is not None:
            self.after_idle(lambda row=first_added: self._scroll_medication_row_into_view(row))

    def _scroll_medication_row_into_view(self, row):
        """Move the shared page canvas to the first medicine inserted from a template."""
        canvas = getattr(getattr(self, "scroll", None), "_parent_canvas", None)
        if canvas is None or not row.winfo_exists():
            return
        try:
            self.update_idletasks()
            bounds = canvas.bbox("all")
            if not bounds or bounds[3] <= bounds[1]:
                return
            visible_offset = row.winfo_rooty() - canvas.winfo_rooty()
            absolute_y = canvas.canvasy(0) + visible_offset
            fraction = (absolute_y - bounds[1]) / (bounds[3] - bounds[1])
            canvas.yview_moveto(max(0.0, min(1.0, fraction - 0.04)))
        except tk.TclError:
            pass

    # -- drug rows ----------------------------------------------------------
    def add_drug_and_show(self):
        self.show_page("medications")
        self.add_row()

    def add_row(self, data=None):
        row = DrugRow(self.drugs_frame, self.db, self.on_any_change,
                      lambda: self.remove_row(row),
                      lambda r: self.move_row(r, -1), lambda r: self.move_row(r, 1),
                      self.drag_row,
                      lambda r, force=False: self.query_gemini_drug(r, force),
                      fg_color=CARD, border_color=LINE, border_width=1, corner_radius=12)
        if data:
            row.name_var.set(data.generic_name)
            row.trade_var.set(data.brand_name)
            row._brand = data.brand_name
            matched_drug = self.db.find_exact(data.generic_name)
            row._form = matched_drug.form if matched_drug is not None else ""
            row.dosage_var.set(data.dosage)
            row.freq_var.set(data.frequency)
            row.dur_var.set(data.duration)
            row.notes_var.set(data.notes)
            saved_quantity = getattr(data, "quantity", "")
            row.quantity_var.set(saved_quantity)
            row._quantity_manual = bool(saved_quantity.strip())
            if not saved_quantity.strip():
                row.recalculate_quantity()
        row.pack(fill="x", padx=2, pady=4)
        self.rows.append(row)
        self._number_drug_rows()
        if data:
            row._update_class_badge(matched_drug)
        self.on_any_change()
        return row

    def toggle_medication_favorite_picker(self):
        self.open_medication_subpage("starred")
        self.refresh_medication_favorite_picker()
        self.after_idle(self.medication_favorite_search_entry.focus_set)

    def refresh_medication_favorite_picker(self):
        if not hasattr(self, "medication_favorite_results"):
            return
        for child in self.medication_favorite_results.winfo_children():
            child.destroy()
        query = self.medication_favorite_search_var.get().strip().casefold()
        favorites = cfg.config.medication_favorites()
        indexed = [(index, favorite) for index, favorite in enumerate(favorites)
                   if favorite.get("pinned")]
        indexed.sort(key=lambda pair: (
            -int(pair[1].get("use_count", 0)),
            self._favorite_sort_name(pair[1])))
        shown = 0
        for index, favorite in indexed:
            searchable = " ".join(str(favorite.get(key, "")) for key in (
                "brand_name", "generic_name", "category", "dosage",
                "frequency", "duration", "notes")).casefold()
            if query and query not in searchable:
                continue
            result_row = ctk.CTkFrame(
                self.medication_favorite_results, fg_color=CARD,
                border_color=LINE, border_width=1, corner_radius=8)
            result_row.grid(row=shown // 2, column=shown % 2, sticky="ew",
                            padx=CARD_GAP, pady=CARD_GAP)
            result_row.grid_columnconfigure(0, weight=1)
            brand = favorite.get("brand_name", "").strip()
            scientific = favorite.get("generic_name", "").strip()
            name = brand or scientific
            if brand and scientific and brand.casefold() != scientific.casefold():
                name = f"{brand}  ·  {scientific}"
            regimen = self._favorite_regimen_line(favorite)
            text = name if not regimen else f"{name}   ·   {regimen}"
            line = ctk.CTkLabel(
                result_row, text="", width=1, text_color=TEXT, anchor="w", justify="left",
                font=_ui_font(12), wraplength=0)
            line.grid(row=0, column=0, sticky="ew", padx=(8, 4), pady=4)
            result_row.bind("<Configure>", lambda event, target=line, full=text:
                            self.fit_starred_medicine_line(target, full, event.width - 54))
            add_button = VisualButton(
                result_row, text="+", width=30, height=30,
                fg_color="transparent", text_color=ACCENT, hover_color=ACCENT_SOFT,
                font=_ui_font(22), border_width=0,
                command=lambda item_id=favorite.get("id", index):
                    self.use_favorite_from_medication(item_id))
            add_button.grid(row=0, column=1, padx=(0, 6), pady=4)
            shown += 1
            if shown >= 12:
                break
        if not shown:
            ctk.CTkLabel(
                self.medication_favorite_results,
                text=(I.t("starred_drugs_empty") if not indexed and not query
                      else I.t("favorite_no_matches")),
                text_color=MUTED, anchor="w").grid(
                    row=0, column=0, columnspan=2, sticky="ew", padx=8, pady=8)

    @staticmethod
    def fit_starred_medicine_line(label, text, available_width):
        """Keep two-column cards on one line with an ellipsis when necessary."""
        font = label.cget("font")
        available_width = max(0, available_width)
        if font.measure(text) <= available_width:
            label.configure(text=text)
            return
        low, high = 0, len(text)
        while low < high:
            middle = (low + high + 1) // 2
            if font.measure(text[:middle] + "…") <= available_width:
                low = middle
            else:
                high = middle - 1
        label.configure(text=text[:low].rstrip() + "…")

    def use_favorite_from_medication(self, identifier):
        self.use_favorite(identifier)
        self.close_medication_subpage()
        self.medication_favorite_search_var.set("")

    def query_gemini_drug(self, row, force=False):
        """Retrieve one grounded reference without blocking the Tk event loop."""
        name = row.name_var.get().strip()
        try:
            name = gemini_drug.normalize_drug_name(name)
        except gemini_drug.GeminiDrugError as exc:
            row.show_reference_error(str(exc))
            return
        if not cfg.config.gemini_enabled or not cfg.config.gemini_api_key:
            messagebox.showinfo(I.t("gemini_lookup"), I.t("gemini_not_configured"))
            GeminiSettingsWindow(self)
            return
        row.show_reference_loading(name)

        def worker():
            try:
                result = gemini_drug.GeminiDrugClient(cfg.config.gemini_api_key).fetch(
                    name, force=force)
            except Exception as exc:
                logging.exception("Gemini drug-reference lookup failed")
                self.after(0, lambda: self._finish_gemini_error(row, str(exc)))
                return
            self.after(0, lambda: self._finish_gemini_result(row, name, result))

        threading.Thread(target=worker, daemon=True).start()

    @staticmethod
    def _row_is_available(row):
        try:
            return bool(row.winfo_exists())
        except Exception:
            return False

    def _finish_gemini_result(self, row, requested_name, result):
        if (self._row_is_available(row)
                and row.name_var.get().strip().casefold() == requested_name.casefold()):
            row.show_reference_result(result)

    def _finish_gemini_error(self, row, detail):
        if self._row_is_available(row):
            row.show_reference_error(detail)

    def _number_drug_rows(self):
        for number, item in enumerate(self.rows, 1):
            item.set_position(number)

    def refresh_favorites_page(self):
        if not hasattr(self, "favorite_cards") or getattr(self, "_refreshing_favorites", False):
            return
        if self._favorite_refresh_job is not None:
            try:
                self.after_cancel(self._favorite_refresh_job)
            except (tk.TclError, ValueError):
                pass
            self._favorite_refresh_job = None
        self._refreshing_favorites = True
        scroll_position = self._current_scroll_position()
        try:
            favorites = cfg.config.medication_favorites()
            categories = sorted(
                {favorite.get("category", "").strip() for favorite in favorites
                 if favorite.get("category", "").strip()}, key=str.casefold)
            selected_category = self.favorite_category_var.get()
            allowed_categories = [I.t("all_categories"), I.t("favorite_starred")] + categories
            if selected_category not in allowed_categories:
                selected_category = I.t("all_categories")
                self.favorite_category_var.set(selected_category)
            self._rebuild_favorite_category_chips(favorites, categories)

            valid_ids = {favorite["id"] for favorite in favorites}
            self._favorite_selected_ids.intersection_update(valid_ids)

            query = self.favorite_search_var.get().casefold().strip()
            indexed = []
            for index, favorite in enumerate(favorites):
                category = favorite.get("category", "").strip()
                searchable = " ".join(str(favorite.get(key, "")) for key in (
                    "generic_name", "brand_name", "category", "regimen_name",
                    "dosage", "frequency", "duration", "notes")).casefold()
                if query and query not in searchable:
                    continue
                if (selected_category == I.t("favorite_starred")
                        and not favorite.get("pinned")):
                    continue
                if (selected_category not in (
                        I.t("all_categories"), I.t("favorite_starred"))
                        and category != selected_category):
                    continue
                indexed.append((index, favorite))

            sort_mode = self.favorite_sort_var.get()
            if sort_mode == I.t("favorite_sort_used"):
                indexed.sort(key=lambda pair: (
                    not bool(pair[1].get("pinned")),
                    -int(pair[1].get("use_count", 0)),
                    self._favorite_sort_name(pair[1])))
            elif sort_mode == I.t("favorite_sort_recent"):
                indexed.sort(key=lambda pair: (
                    not bool(pair[1].get("pinned")),
                    -self._favorite_used_timestamp(pair[1]),
                    self._favorite_sort_name(pair[1])))
            elif sort_mode == I.t("favorite_sort_name_reverse"):
                indexed.sort(
                    key=lambda pair: self._favorite_sort_name(pair[1]), reverse=True)
                indexed.sort(key=lambda pair: not bool(pair[1].get("pinned")))
            else:
                indexed.sort(key=lambda pair: (
                    not bool(pair[1].get("pinned")),
                    self._favorite_sort_name(pair[1])))

            filter_signature = (query, selected_category, sort_mode)
            if filter_signature != self._favorite_filter_signature:
                self._favorite_filter_signature = filter_signature
                self._favorite_render_limit = 60
            visible_indexed = indexed[:self._favorite_render_limit]

            for favorite_id in list(self._favorite_card_widgets):
                if favorite_id not in valid_ids:
                    self._favorite_card_widgets.pop(favorite_id)["card"].destroy()
            for widgets in self._favorite_card_widgets.values():
                widgets["card"].grid_forget()
            empty_label = getattr(self, "favorite_empty_label", None)
            if empty_label is not None:
                empty_label.destroy()
                self.favorite_empty_label = None
            column_count = self._favorite_columns_for_width(
                self._logical_widget_width(
                    self.favorite_cards, self.favorite_cards.winfo_width()))
            self._favorite_card_columns = column_count
            for column in range(3):
                self.favorite_cards.grid_columnconfigure(
                    column, weight=1 if column < column_count else 0,
                    uniform="favorite" if column < column_count else "")
            more_frame = getattr(self, "_favorite_more_frame", None)
            if more_frame is not None and more_frame.winfo_exists():
                more_frame.destroy()
            self._favorite_more_frame = None
            for display_index, (source_index, favorite) in enumerate(visible_indexed):
                self._render_favorite_card(display_index, source_index, favorite)
            if not indexed:
                self.favorite_empty_label = ctk.CTkLabel(
                    self.favorite_cards,
                    text=I.t("favorite_empty") if not favorites else I.t("favorite_no_matches"),
                    text_color=MUTED, font=ctk.CTkFont(size=15),
                    justify="center")
                self.favorite_empty_label.grid(
                    row=0, column=0, columnspan=column_count, sticky="ew", pady=28)
            elif len(visible_indexed) < len(indexed):
                remaining = len(indexed) - len(visible_indexed)
                self._favorite_more_frame = ctk.CTkFrame(
                    self.favorite_cards, fg_color="transparent")
                self._favorite_more_frame.grid(
                    row=(len(visible_indexed) + column_count - 1) // column_count,
                    column=0, columnspan=column_count, sticky="ew", pady=8)
                ctk.CTkLabel(
                    self._favorite_more_frame,
                    text=I.t("more_class_medicines", n=remaining),
                    text_color=MUTED).pack(side="left", padx=(4, 10))
                VisualButton(
                    self._favorite_more_frame, text=I.t("load_more"),
                    width=110, height=32, fg_color=PRIMARY,
                    hover_color=ACCENT_HOVER,
                    command=self._load_more_favorites).pack(side="left")
            self._refresh_favorite_selection_bar()
            self._show_favorite_notice()
        finally:
            self._refreshing_favorites = False
            self._restore_scroll_position(scroll_position)

    def _load_more_favorites(self):
        self._favorite_render_limit += 60
        self.refresh_favorites_page()

    def _schedule_favorites_refresh(self, delay=150):
        if not hasattr(self, "favorite_cards"):
            return
        if self._favorite_refresh_job is not None:
            try:
                self.after_cancel(self._favorite_refresh_job)
            except (tk.TclError, ValueError):
                pass
        self._favorite_refresh_job = self.after(delay, self.refresh_favorites_page)

    def _current_scroll_position(self):
        canvas = getattr(getattr(self, "scroll", None), "_parent_canvas", None)
        if canvas is None:
            return None
        try:
            return canvas.yview()[0]
        except tk.TclError:
            return None

    def _restore_scroll_position(self, position):
        if position is None:
            return

        def restore():
            canvas = getattr(getattr(self, "scroll", None), "_parent_canvas", None)
            if canvas is not None:
                try:
                    canvas.yview_moveto(position)
                except tk.TclError:
                    pass

        self.after_idle(restore)

    @staticmethod
    def _favorite_used_timestamp(favorite):
        value = favorite.get("last_used", "")
        if not value:
            return 0.0
        try:
            return datetime.datetime.fromisoformat(value).timestamp()
        except (TypeError, ValueError):
            return 0.0

    @staticmethod
    def _favorite_sort_name(favorite):
        return (favorite.get("brand_name", "").strip()
                or favorite.get("generic_name", "").strip()).casefold()

    def _favorite_matches_database(self, favorite):
        return self.db.contains_name(
            favorite.get("generic_name", ""), favorite.get("brand_name", ""))

    def _rebuild_favorite_category_chips(self, favorites, categories):
        for child in self.favorite_category_chips.winfo_children():
            child.destroy()
        selected = self.favorite_category_var.get()
        all_label = I.t("all_categories")
        starred_label = I.t("favorite_starred")
        category_values = list(categories)
        category_values.sort(key=lambda category: (
            -sum(int(favorite.get("use_count", 0)) for favorite in favorites
                 if favorite.get("category", "").strip() == category),
            -sum(1 for favorite in favorites
                 if favorite.get("category", "").strip() == category),
            category.casefold()))

        available_width = self._logical_widget_width(
            self.favorite_category_chips,
            self.favorite_category_chips.winfo_width())
        if available_width < 300:
            available_width = 900
        candidates = [all_label, starred_label] + category_values
        rows = []
        current_row = []
        used_width = 0
        for category in candidates:
            button_width = max(58, min(155, 22 + len(category) * 7))
            if current_row and used_width + button_width + 5 > available_width:
                rows.append(current_row)
                current_row = []
                used_width = 0
            current_row.append((category, button_width))
            used_width += button_width + 5
        if current_row:
            rows.append(current_row)

        for row_values in rows:
            row_frame = ctk.CTkFrame(
                self.favorite_category_chips, fg_color="transparent")
            row_frame.pack(fill="x", pady=1)
            pack_side = "right" if I.LANG == "ar" else "left"
            for category, button_width in row_values:
                active = category == selected
                VisualButton(
                    row_frame, text=category, width=button_width,
                    height=26, corner_radius=13,
                    fg_color=PRIMARY if active else ACCENT_SOFT,
                    text_color="white" if active else ACCENT,
                    hover_color=ACCENT_HOVER if active else LINE,
                    font=ctk.CTkFont(size=11, weight="bold"),
                    command=lambda value=category:
                        self._choose_favorite_category(value)).pack(
                            side=pack_side, padx=(0, 5))

    def _on_favorite_category_resize(self, event):
        if abs(event.width - self._favorite_category_width) < 60:
            return
        self._favorite_category_width = event.width
        self._schedule_favorites_refresh(120)

    def _choose_favorite_category(self, category):
        self.favorite_category_var.set(category)
        self.refresh_favorites_page()

    @staticmethod
    def _favorite_columns_for_width(width):
        if width and width >= 820:
            return 3
        if width and width >= 540:
            return 2
        return 1 if width and width >= 300 else 3

    @staticmethod
    def _logical_widget_width(widget, raw_width):
        try:
            scaling = ctk.ScalingTracker.get_widget_scaling(widget)
            return raw_width / scaling if scaling else raw_width
        except (KeyError, tk.TclError):
            return raw_width

    def _on_favorite_cards_resize(self, event):
        columns = self._favorite_columns_for_width(
            self._logical_widget_width(self.favorite_cards, event.width))
        if columns != getattr(self, "_favorite_card_columns", 3):
            self._favorite_card_columns = columns
            self._schedule_favorites_refresh(80)

    def _toggle_favorite_selection(self, favorite_id, variable, card):
        if variable.get():
            self._favorite_selected_ids.add(favorite_id)
            card.configure(border_color=ACCENT, border_width=2)
        else:
            self._favorite_selected_ids.discard(favorite_id)
            card.configure(border_color=LINE, border_width=1)
        self._refresh_favorite_selection_bar()

    def _refresh_favorite_selection_bar(self):
        count = len(self._favorite_selected_ids)
        if count:
            self.favorite_selection_label.configure(
                text=I.t("favorite_selected_count", n=count))
            if not self.favorite_selection_bar.winfo_manager():
                self.favorite_selection_bar.pack(
                    fill="x", padx=PAD, pady=(0, 6), before=self.favorite_cards)
        elif self.favorite_selection_bar.winfo_manager():
            self.favorite_selection_bar.pack_forget()

    def clear_favorite_selection(self):
        self._favorite_selected_ids.clear()
        self.refresh_favorites_page()

    def use_selected_favorites(self):
        favorites = cfg.config.medication_favorites()
        selected = [
            (index, favorites[index])
            for index in range(len(favorites))
            if favorites[index]["id"] in self._favorite_selected_ids
        ]
        if not selected:
            return
        for index, _favorite in selected:
            cfg.config.record_medication_favorite_use(index)
        self._favorite_selected_ids.clear()
        self.show_page("medications")
        for _index, favorite in selected:
            drug_fields = {key: favorite.get(key, "") for key in
                           ("generic_name", "brand_name", "dosage", "frequency",
                            "duration", "notes")}
            self.add_row(data=qu.DrugItem(**drug_fields))

    @staticmethod
    def _favorite_autocomplete_navigation_key(event):
        return bool(event and event.keysym in {
            "Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R",
            "Control_L", "Control_R", "Alt_L", "Alt_R",
        })

    def _on_favorite_scientific_type(self, event=None):
        if self._favorite_autocomplete_navigation_key(event):
            return
        query = self.favorite_vars["generic_name"].get().strip()
        self._schedule_favorite_autocomplete(query, "scientific")

    def _on_favorite_brand_type(self, event=None):
        if self._favorite_autocomplete_navigation_key(event):
            return
        query = self.favorite_vars["brand_name"].get().strip()
        self._schedule_favorite_autocomplete(query, "brand")

    def _schedule_favorite_autocomplete(self, query, mode):
        if self._favorite_ac_job is not None:
            try:
                self.after_cancel(self._favorite_ac_job)
            except (tk.TclError, ValueError):
                pass
        self._favorite_ac_token += 1
        token = self._favorite_ac_token
        if not query:
            self._hide_favorite_autocomplete()
            return

        def launch():
            self._favorite_ac_job = None
            search = (self.db.search_scientific if mode == "scientific"
                      else self.db.search_prescribable)

            def finished(matches):
                current = self.favorite_vars[
                    "generic_name" if mode == "scientific" else "brand_name"
                ].get().strip()
                if token == self._favorite_ac_token and current == query:
                    self._show_favorite_autocomplete(matches, mode)

            self.submit_background(
                lambda: search(query, 20), finished, silent=True)

        self._favorite_ac_job = self.after(180, launch)

    def _show_favorite_autocomplete(self, matches, mode):
        self._hide_favorite_autocomplete()
        if not matches:
            return
        anchor = (self.favorite_generic_entry if mode == "scientific"
                  else self.favorite_brand_entry)
        anchor.update_idletasks()
        top = tk.Toplevel(self)
        top.wm_overrideredirect(True)
        top.attributes("-topmost", True)
        top.geometry(
            f"+{anchor.winfo_rootx()}+{anchor.winfo_rooty() + anchor.winfo_height()}")
        self._favorite_ac_top = top
        self._favorite_ac_matches = matches
        self._favorite_ac_mode = mode
        self._favorite_ac_index = 0
        listbox = PopupListbox(
            top, height=min(10, len(matches)), font=LIST_FONT,
            bg=CARD, fg=TEXT, relief="flat", borderwidth=0,
            activestyle="none", highlightthickness=1,
            highlightbackground=LINE, selectbackground=PRIMARY,
            selectforeground="white", exportselection=False)
        max_chars = max(
            len(drug.generic_name) + len(drug.brand_name or "")
            + len(drug.strength or "") + 8 for drug in matches)
        listbox.configure(width=max(38, min(72, max_chars + 5)))
        for drug in matches:
            primary = (drug.generic_name if mode == "scientific"
                       else (drug.brand_name or drug.generic_name))
            linked = (drug.brand_name if mode == "scientific"
                      else (drug.generic_name if drug.brand_name else ""))
            label = primary
            if linked:
                label += f"  ({linked})"
            if drug.strength:
                label += f"  [{drug.strength}]"
            listbox.insert(tk.END, label)
        listbox.selection_set(0)
        listbox.see(0)
        listbox.bind("<ButtonRelease-1>", self._favorite_ac_choose)
        listbox.bind("<Return>", self._favorite_ac_choose)
        listbox.bind("<Escape>", lambda _event: self._hide_favorite_autocomplete())
        fit_autocomplete_popup(top, anchor, listbox, len(matches))
        self._favorite_ac_listbox = listbox

    def _hide_favorite_autocomplete(self):
        top = getattr(self, "_favorite_ac_top", None)
        if top is not None:
            try:
                top.destroy()
            except tk.TclError:
                pass
        self._favorite_ac_top = None
        self._favorite_ac_listbox = None
        self._favorite_ac_matches = []
        self._favorite_ac_index = -1

    def _move_favorite_autocomplete(self, step):
        listbox = getattr(self, "_favorite_ac_listbox", None)
        if listbox is None or not listbox.size():
            return None
        self._favorite_ac_index = max(
            0, min(listbox.size() - 1, self._favorite_ac_index + step))
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(self._favorite_ac_index)
        listbox.see(self._favorite_ac_index)
        return "break"

    def _favorite_ac_down(self, _event=None):
        return self._move_favorite_autocomplete(1)

    def _favorite_ac_up(self, _event=None):
        return self._move_favorite_autocomplete(-1)

    def _favorite_ac_choose_current(self, _event=None):
        if getattr(self, "_favorite_ac_listbox", None) is not None:
            self._favorite_ac_choose()
            return "break"
        return None

    def _favorite_ac_choose(self, _event=None):
        listbox = getattr(self, "_favorite_ac_listbox", None)
        if listbox is None:
            return
        selection = listbox.curselection()
        if not selection:
            return
        drug = self._favorite_ac_matches[selection[0]]
        if self._favorite_ac_mode == "brand":
            self.favorite_vars["brand_name"].set(
                drug.brand_name or drug.generic_name)
            if drug.brand_name:
                self.favorite_vars["generic_name"].set(drug.generic_name)
        else:
            self.favorite_vars["generic_name"].set(drug.generic_name)
            self.favorite_vars["brand_name"].set(drug.brand_name or "")
        if drug.category and not self.favorite_vars["category"].get().strip():
            self.favorite_vars["category"].set(drug.category)
        if drug.strength and not self.favorite_vars["dosage"].get().strip():
            self.favorite_vars["dosage"].set(drug.strength)
        self._hide_favorite_autocomplete()
        return "break"

    @staticmethod
    def _favorite_identity(favorite):
        return tuple(str(favorite.get(key, "")).strip().casefold() for key in (
            "generic_name", "brand_name"))

    @staticmethod
    def _favorite_prescription_line(favorite):
        name = favorite.get("generic_name", "").strip()
        brand = favorite.get("brand_name", "").strip()
        if brand:
            name = f"{name} ({brand})" if name else brand
        pieces = [name] + [favorite.get(key, "").strip() for key in (
            "dosage", "frequency", "duration", "notes")]
        return "     ".join(piece for piece in pieces if piece)

    @staticmethod
    def _favorite_regimen_line(favorite):
        """Compact card preview: instructions only; medicine identity is in the header."""
        return "     ".join(
            favorite.get(key, "").strip()
            for key in ("dosage", "frequency", "duration", "notes")
            if favorite.get(key, "").strip())

    def _render_favorite_card(self, display_index, source_index, favorite):
        favorite_id = favorite["id"]
        signature = tuple(str(favorite.get(key, "")) for key in (
            "brand_name", "generic_name", "category", "dosage",
            "frequency", "duration", "notes"))
        widgets = self._favorite_card_widgets.get(favorite_id)
        if widgets is not None and widgets["signature"] != signature:
            widgets["card"].destroy()
            widgets = None

        selected = favorite_id in self._favorite_selected_ids
        if widgets is None:
            card = GlassFrame(
                self.favorite_cards, fg_color=CARD,
                border_color=ACCENT if selected else LINE,
                border_width=2 if selected else 1, corner_radius=10)
            card.grid_columnconfigure(0, weight=1)

            heading = ctk.CTkFrame(card, fg_color="transparent")
            heading.grid(row=0, column=0, sticky="ew", padx=8, pady=(5, 0))
            heading.grid_columnconfigure(0, weight=1)
            name_line = ctk.CTkFrame(heading, fg_color="transparent")
            name_line.grid(row=0, column=0, sticky="w")
            brand_name = favorite.get("brand_name", "").strip()
            scientific_name = favorite.get("generic_name", "").strip()
            if brand_name:
                ctk.CTkLabel(
                    name_line, text=brand_name, text_color=TEXT,
                    anchor="w", justify="left",
                    font=ctk.CTkFont(size=15, weight="bold")).pack(side="left")
            if scientific_name and scientific_name.casefold() != brand_name.casefold():
                separator = "  ·  " if brand_name else ""
                ctk.CTkLabel(
                    name_line, text=f"{separator}{scientific_name}", text_color=MUTED,
                    anchor="w", justify="left",
                    font=ctk.CTkFont(size=15)).pack(side="left")
            if favorite.get("category", "").strip():
                ctk.CTkLabel(
                    name_line, text=favorite["category"].strip(), height=20,
                    corner_radius=10, fg_color=ACCENT_SOFT, text_color=ACCENT,
                    font=ctk.CTkFont(size=9, weight="bold")).pack(
                        side="left", padx=(7, 0))
            regimen_line = self._favorite_regimen_line(favorite)
            if regimen_line:
                ctk.CTkLabel(
                    card, text=regimen_line, text_color=TEXT_SECONDARY,
                    anchor="w", justify="left", wraplength=350,
                    font=ctk.CTkFont(size=12)).grid(
                        row=1, column=0, sticky="ew", padx=8, pady=(1, 3))
            actions = ctk.CTkFrame(card, fg_color="transparent")
            actions.grid(row=2, column=0, sticky="ew", padx=6, pady=(0, 5))
            add_button = VisualButton(
                actions, text="+", width=28, height=28,
                fg_color="transparent", text_color=ACCENT, hover_color=ACCENT_SOFT,
                border_width=0, font=_ui_font(22),
                command=lambda item_id=favorite_id: self.use_favorite(item_id))
            add_button.pack(side="left", padx=2)
            selected_var = tk.BooleanVar(value=selected)
            ctk.CTkCheckBox(
                actions, text="", variable=selected_var, width=24, height=24,
                checkbox_width=20, checkbox_height=20, corner_radius=5,
                border_width=2, fg_color=PRIMARY, hover_color=ACCENT_HOVER,
                command=lambda item_id=favorite_id, variable=selected_var,
                target=card: self._toggle_favorite_selection(
                    item_id, variable, target)).pack(side="left", padx=(5, 2))
            edit_button = VisualButton(
                actions, text="", image=_edit_icon(18), width=28, height=28,
                fg_color="transparent", text_color=ACCENT, border_width=0,
                hover_color=ACCENT_SOFT, font=ctk.CTkFont(size=15, weight="bold"),
                command=lambda item_id=favorite_id: self.edit_favorite(item_id))
            edit_button.pack(side="left", padx=2)
            delete_button = VisualButton(
                actions, text="🗑", width=28, height=28,
                fg_color="transparent", text_color=DANGER, border_width=0,
                hover_color=DANGER_SOFT, font=ctk.CTkFont(size=15),
                command=lambda item_id=favorite_id: self.delete_favorite(item_id))
            delete_button.pack(side="left", padx=2)
            star_button = VisualButton(
                actions, text="★" if favorite.get("pinned") else "☆",
                width=28, height=28, fg_color="transparent",
                text_color=ICON_BLUE, border_width=0, hover_color=ACCENT_SOFT,
                font=ctk.CTkFont(size=17),
                command=lambda item_id=favorite_id: self.toggle_favorite_pin(item_id))
            star_button.pack(side="left", padx=2)
            widgets = {
                "card": card, "star": star_button, "selected": selected_var,
                "signature": signature,
            }
            self._favorite_card_widgets[favorite_id] = widgets
        else:
            card = widgets["card"]
            widgets["star"].configure(text="★" if favorite.get("pinned") else "☆")
            widgets["selected"].set(selected)

        highlighted = favorite_id == self._favorite_highlight_id
        card.configure(
            fg_color=ACCENT_SOFT if highlighted else CARD,
            border_color=ACCENT if selected or highlighted else LINE,
            border_width=2 if selected or highlighted else 1)
        columns = getattr(self, "_favorite_card_columns", 3)
        card.grid(row=display_index // columns, column=display_index % columns,
                  sticky="new", padx=CARD_GAP, pady=CARD_GAP)
        if highlighted:
            self.after(900, lambda item_id=favorite_id:
                       self._finish_favorite_highlight(item_id))

    def _finish_favorite_highlight(self, favorite_id):
        if self._favorite_highlight_id != favorite_id:
            return
        self._favorite_highlight_id = None
        widgets = self._favorite_card_widgets.get(favorite_id)
        if widgets is None:
            return
        selected = favorite_id in self._favorite_selected_ids
        widgets["card"].configure(
            fg_color=CARD, border_color=ACCENT if selected else LINE,
            border_width=2 if selected else 1)

    @staticmethod
    def _favorite_index(favorites, identifier):
        if isinstance(identifier, int):
            return identifier if 0 <= identifier < len(favorites) else None
        return next((index for index, favorite in enumerate(favorites)
                     if favorite.get("id") == identifier), None)

    def _show_favorite_notice(self, mismatch_count=0):
        if self._deleted_favorite:
            return
        if self.favorite_notice.winfo_manager():
            self.favorite_notice.pack_forget()

    def new_favorite(self):
        self.favorite_edit_index = None
        for variable in self.favorite_vars.values():
            variable.set("")
        self.favorite_editor_title.configure(text=I.t("new_favorite"))
        self._show_favorite_editor()

    def edit_favorite(self, identifier):
        favorites = cfg.config.medication_favorites()
        index = self._favorite_index(favorites, identifier)
        if index is None:
            return
        self.favorite_edit_index = index
        favorite = favorites[index]
        for key, variable in self.favorite_vars.items():
            variable.set(favorite.get(key, ""))
        self.favorite_editor_title.configure(text=I.t("edit_favorite"))
        self._show_favorite_editor()

    def _show_favorite_editor(self):
        categories = sorted({
            favorite.get("category", "").strip()
            for favorite in cfg.config.medication_favorites()
            if favorite.get("category", "").strip()
        }, key=str.casefold)
        self.favorite_category_combo.configure(values=categories or [""])
        if not self.favorite_editor.winfo_manager():
            self.favorite_editor.pack(
                fill="x", padx=PAD, pady=(0, 10), before=self.favorite_cards)
        self.update_favorite_preview()

    def close_favorite_editor(self):
        self._hide_favorite_autocomplete()
        self.favorite_editor.pack_forget()
        self.favorite_edit_index = None

    def update_favorite_preview(self):
        if not hasattr(self, "favorite_preview_label"):
            return
        item = {key: variable.get() for key, variable in self.favorite_vars.items()}
        line = self._favorite_prescription_line(item)
        self.favorite_preview_label.configure(
            text=f"1.     {line}" if line else I.t("favorite_preview_empty"))

    def save_favorite_changes(self):
        item = {key: variable.get() for key, variable in self.favorite_vars.items()}
        if not (item["generic_name"].strip() or item["brand_name"].strip()):
            messagebox.showinfo(APP_TITLE, I.t("msg_no_drugs"))
            return
        favorites = cfg.config.medication_favorites()
        identity = self._favorite_identity(item)
        duplicate_index = next((
            index for index, favorite in enumerate(favorites)
            if index != self.favorite_edit_index and self._favorite_identity(favorite) == identity
        ), None)
        if duplicate_index is not None:
            if not messagebox.askyesno(
                    I.t("favorite_duplicate_title"), I.t("favorite_duplicate_confirm")):
                return
            saved = cfg.config.update_medication_favorite(duplicate_index, item)
        elif self.favorite_edit_index is None:
            saved = cfg.config.add_medication_favorite(item)
        else:
            saved = cfg.config.update_medication_favorite(self.favorite_edit_index, item)
        if not saved:
            messagebox.showinfo(APP_TITLE, I.t("msg_no_drugs"))
            return
        self.refresh_favorites_page()
        self.close_favorite_editor()

    def delete_favorite(self, identifier=None):
        if identifier is None:
            identifier = self.favorite_edit_index
        favorites = cfg.config.medication_favorites()
        index = self._favorite_index(favorites, identifier)
        if index is None:
            return
        deleted = favorites[index].copy()
        display_name = (deleted.get("brand_name", "").strip()
                        or deleted.get("generic_name", "").strip())
        if not messagebox.askyesno(
                I.t("delete_favorite"),
                I.t("favorite_delete_confirm", name=display_name)):
            return
        if cfg.config.remove_medication_favorite(index):
            self._deleted_favorite_recovery_id = cfg.config.add_recovery_item(
                "favorite", display_name, {"index": index, "favorite": deleted})
            self._favorite_selected_ids.discard(deleted.get("id", ""))
            self._deleted_favorite = (index, deleted)
            if self.favorite_undo_timer:
                self.after_cancel(self.favorite_undo_timer)
            self.favorite_undo_timer = self.after(8000, self._expire_favorite_undo)
            self.refresh_favorites_page()
            self.close_favorite_editor()
            self.favorite_notice_label.configure(
                text=I.t("favorite_deleted", name=display_name))
            if not self.favorite_undo_button.winfo_manager():
                self.favorite_undo_button.pack(side="right", padx=8, pady=5)
            if not self.favorite_notice.winfo_manager():
                self.favorite_notice.pack(
                    fill="x", padx=PAD, pady=(0, 8), before=self.favorite_cards)

    def undo_delete_favorite(self):
        if not self._deleted_favorite:
            return
        index, favorite = self._deleted_favorite
        cfg.config.insert_medication_favorite(index, favorite)
        recovery_id = getattr(self, "_deleted_favorite_recovery_id", "")
        if recovery_id:
            cfg.config.discard_recovery_item(recovery_id)
            self._deleted_favorite_recovery_id = ""
        self._expire_favorite_undo()
        self.refresh_favorites_page()

    def _expire_favorite_undo(self):
        self._deleted_favorite = None
        self.favorite_undo_timer = None
        if self.favorite_notice.winfo_manager():
            self.favorite_notice.pack_forget()
        self._show_favorite_notice()

    def toggle_favorite_pin(self, identifier):
        favorites = cfg.config.medication_favorites()
        index = self._favorite_index(favorites, identifier)
        if index is None:
            return
        favorite_id = favorites[index]["id"]
        cfg.config.toggle_medication_favorite_pin(index)
        self._favorite_highlight_id = favorite_id
        self.refresh_favorites_page()

    def use_favorite(self, identifier):
        favorites = cfg.config.medication_favorites()
        index = self._favorite_index(favorites, identifier)
        if index is None:
            return
        favorite = favorites[index]
        cfg.config.record_medication_favorite_use(index)
        self.show_page("medications")
        drug_fields = {key: favorite.get(key, "") for key in
                       ("generic_name", "brand_name", "dosage", "frequency", "duration", "notes", "quantity")}
        self.add_row(data=qu.DrugItem(**drug_fields))

    def use_selected_favorite(self):
        if self.favorite_edit_index is None:
            return
        self.use_favorite(self.favorite_edit_index)

    # -- therapeutic drug classes -----------------------------------------
    def ordered_therapeutic_groups(self):
        pinned = [code for code in cfg.config.favorite_therapeutic_groups()
                  if code in classes.GROUPS]
        return pinned + [code for code in classes.GROUPS if code not in pinned]

    def _invalidate_database_caches(self):
        self._classification_cache = None
        self._class_browser_render_signature = None

    def _ensure_classification_cache(self):
        """Classify each medicine once and reuse counts/lists across the browser."""
        if self._classification_cache is not None:
            return self._classification_cache
        by_group = {code: [] for code in classes.GROUPS}
        by_pair = {}
        unclassified = []
        suggested = []
        medicines_by_name = {}
        mappings_by_name = {}
        medicines_by_identity = {}
        mappings_by_identity = {}
        for drug in self.db.drugs:
            mappings = classes.groups_for(drug)
            normalized_name = drug.generic_name.strip().casefold()
            identity = self.db.drug_identity(drug)
            mapping_set = frozenset(
                (mapping.code, mapping.detail) for mapping in mappings)
            medicines_by_identity.setdefault(identity, []).append(drug)
            mappings_by_identity.setdefault(identity, set()).add(mapping_set)
            if normalized_name:
                medicines_by_name.setdefault(normalized_name, []).append(drug)
                mappings_by_name.setdefault(normalized_name, set()).add(mapping_set)
            if not mappings:
                unclassified.append(drug)
                continue
            if any(mapping.confidence == "suggested" for mapping in mappings):
                suggested.append(drug)
            seen_groups = set()
            for mapping in mappings:
                key = (mapping.code, mapping.detail.casefold())
                by_pair.setdefault(key, []).append(drug)
                if mapping.code not in seen_groups:
                    by_group.setdefault(mapping.code, []).append(drug)
                    seen_groups.add(mapping.code)
        sort_key = lambda item: (item.brand_name or item.generic_name).casefold()
        for medicines in by_group.values():
            medicines.sort(key=sort_key)
        for medicines in by_pair.values():
            medicines.sort(key=sort_key)
        unclassified.sort(key=sort_key)
        suggested.sort(key=sort_key)
        variations = [drug for name, medicines in medicines_by_name.items()
                      if len(mappings_by_name.get(name, ())) > 1
                      for drug in medicines]
        conflicting = [drug for identity, medicines in medicines_by_identity.items()
                       if len(mappings_by_identity.get(identity, ())) > 1
                       for drug in medicines]
        variations.sort(key=sort_key)
        conflicting.sort(key=sort_key)
        self._classification_cache = {
            "by_group": by_group, "by_pair": by_pair,
            "unclassified": unclassified, "suggested": suggested,
            "variations": variations, "conflicting": conflicting,
        }
        return self._classification_cache

    def unclassified_medicine_count(self):
        return len(self._ensure_classification_cache()["unclassified"])

    def therapeutic_group_medicine_count(self, code):
        return len(self._ensure_classification_cache()["by_group"].get(code, ()))

    def refresh_class_overview(self):
        if not hasattr(self, "class_tiles"):
            return
        ordered = self.ordered_therapeutic_groups()
        for code in ordered:
            tile = self.class_tiles[code]
            tile.pack_forget()
            tile.pack(fill="x", pady=3)
            self.class_count_badges[code].configure(
                text=str(self.therapeutic_group_medicine_count(code)))
        self.refresh_class_browser()

    def _schedule_class_browser_refresh(self, delay=140):
        if self._class_search_job is not None:
            try:
                self.after_cancel(self._class_search_job)
            except (tk.TclError, ValueError):
                pass
        self._class_search_job = self.after(delay, self.refresh_class_browser)

    def _set_class_browser_mode(self, unclassified):
        """Switch between the classification browser and focused unclassified search."""
        if unclassified:
            self.class_left_panel.pack_forget()
            self.class_right_panel.pack_forget()
            if not self.class_unclassified_panel.winfo_manager():
                self.class_unclassified_panel.pack(fill="both", expand=True)
            return

        self.class_unclassified_panel.pack_forget()
        if not self.class_left_panel.winfo_manager():
            self.class_left_panel.pack(side="left", fill="y", padx=(0, 8))
        if not self.class_right_panel.winfo_manager():
            self.class_right_panel.pack(side="left", fill="both", expand=True)

    def render_unclassified_class_search(self, query):
        """Prepare unclassified results and render a small responsive first page."""
        for child in self.class_unclassified_results.winfo_children():
            child.destroy()
        medicines = [drug for drug in self._ensure_classification_cache()["unclassified"]
                     if self._class_drug_matches_query(drug, query)]
        medicines.sort(
            key=lambda drug: (drug.brand_name or drug.generic_name).casefold())
        self.class_visible_drugs = medicines
        self._unclassified_rendered_count = 0
        self._unclassified_page_size = 24
        self.class_unclassified_more_frame = None
        self.class_unclassified_heading.configure(
            text=I.t("unclassified_medicines", n=len(medicines)))
        if not medicines:
            ctk.CTkLabel(
                self.class_unclassified_results,
                text=I.t("class_search_no_results"), text_color=MUTED,
                anchor="w").grid(row=0, column=0, columnspan=2, sticky="ew",
                                 padx=6, pady=10)
            return
        self._append_unclassified_class_search_page()

    def _append_unclassified_class_search_page(self):
        """Append one two-column page without rebuilding already-visible cards."""
        footer = getattr(self, "class_unclassified_more_frame", None)
        if footer is not None and footer.winfo_exists():
            footer.destroy()
        start = self._unclassified_rendered_count
        end = min(start + self._unclassified_page_size, len(self.class_visible_drugs))
        for index in range(start, end):
            self._add_unclassified_class_card(index, self.class_visible_drugs[index])
        self._unclassified_rendered_count = end
        remaining = len(self.class_visible_drugs) - end
        if remaining <= 0:
            self.class_unclassified_more_frame = None
            return
        self.class_unclassified_more_frame = ctk.CTkFrame(
            self.class_unclassified_results, fg_color="transparent")
        self.class_unclassified_more_frame.grid(
            row=(end + 1) // 2, column=0, columnspan=2,
            sticky="ew", padx=4, pady=8)
        ctk.CTkLabel(
            self.class_unclassified_more_frame,
            text=I.t("more_class_medicines", n=remaining),
            text_color=MUTED, anchor="w").pack(side="left", padx=(4, 10))
        VisualButton(
            self.class_unclassified_more_frame, text=I.t("load_more"),
            width=110, height=32, fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            command=self._append_unclassified_class_search_page).pack(side="left")

    def _add_unclassified_class_card(self, index, drug):
        """Add one compact unclassified medicine card to the results grid."""
        card = GlassFrame(
            self.class_unclassified_results, fg_color=CARD,
            border_color=LINE, border_width=1, corner_radius=9)
        card.grid(
            row=index // 2, column=index % 2, sticky="ew", padx=4, pady=4)
        card.grid_columnconfigure(0, weight=1)
        names = ctk.CTkFrame(card, fg_color="transparent")
        names.grid(row=0, column=0, sticky="ew", padx=10, pady=7)
        brand = drug.brand_name.strip() or drug.generic_name.strip()
        scientific = drug.generic_name.strip() if drug.brand_name.strip() else ""
        ctk.CTkLabel(
            names, text=brand, text_color=TEXT, anchor="w",
            font=_ui_font(13, "bold")).pack(side="left")
        if scientific and scientific.casefold() != brand.casefold():
            ctk.CTkLabel(
                names, text="  " + scientific, text_color=MUTED,
                anchor="w", font=_ui_font(11)).pack(side="left")
        ctk.CTkLabel(
            names, text=I.t(
                "classification_unrecognized"
                if drug.mapping_status == "unrecognized"
                else "classification_unclassified"), height=22,
            corner_radius=11, fg_color=WARNING_SOFT, text_color=WARNING,
            font=_ui_font(9, "bold")).pack(
                side="left", padx=(8, 0))
        VisualButton(
            card, text=I.t("favorite_use_rx"), width=82, height=30,
            fg_color=CARD, text_color=ACCENT, border_width=1,
            border_color=LINE, hover_color=ACCENT_SOFT,
            command=lambda item=drug: self.add_drug_database_item(item)).grid(
                row=0, column=1, padx=(2, 4), pady=5)
        VisualButton(
            card, text="★" if self._class_drug_is_starred(drug) else "☆",
            width=34, height=30, fg_color="transparent", text_color=ACCENT,
            hover_color=ACCENT_SOFT,
            command=lambda item=drug: self.toggle_class_drug_star(item)).grid(
                row=0, column=2, padx=(0, 6), pady=5)
        self._bind_class_medicine_context(card, drug)

    def _class_drug_matches_query(self, drug, query):
        if not query:
            return True
        mode = self.class_name_filter_var.get() if hasattr(self, "class_name_filter_var") else ""
        if mode == I.t("filter_brand_name"):
            return query in drug.brand_name.casefold()
        if mode == I.t("filter_scientific_name"):
            return query in drug.generic_name.casefold()
        return query in self._drug_search_text(drug)

    def _drugs_in_class(self, code, detail=None):
        cache = self._ensure_classification_cache()
        if detail is None:
            return list(cache["by_group"].get(code, ()))
        return list(cache["by_pair"].get((code, detail.casefold()), ()))

    @staticmethod
    def _classification_status_for(drug, code=None, detail=None):
        for mapping in classes.groups_for(drug):
            if code and mapping.code != code:
                continue
            if detail and mapping.detail.casefold() != detail.casefold():
                continue
            return mapping.confidence
        return "unclassified"

    @staticmethod
    def _classification_status_colors(status):
        return {
            "confirmed": (ACCENT_SOFT, ACCENT),
            "suggested": (WARNING_SOFT, WARNING),
            "unclassified": (BG, MUTED),
        }.get(status, (BG, MUTED))

    @staticmethod
    def _drug_search_text(drug):
        return " ".join((
            drug.brand_name, drug.generic_name, drug.strength, drug.form,
            drug.category, drug.therapeutic_group, drug.detailed_class,
        )).casefold()

    def _favorite_index_for_class_drug(self, drug):
        names = {value.strip().casefold() for value in (
            drug.brand_name, drug.generic_name) if value.strip()}
        for index, favorite in enumerate(cfg.config.medication_favorites()):
            favorite_names = {str(favorite.get(key, "")).strip().casefold()
                              for key in ("brand_name", "generic_name")
                              if str(favorite.get(key, "")).strip()}
            if names.intersection(favorite_names):
                return index
        return None

    def _class_drug_is_starred(self, drug):
        index = self._favorite_index_for_class_drug(drug)
        if index is None:
            return False
        return bool(cfg.config.medication_favorites()[index].get("pinned"))

    def _class_drug_was_used(self, drug):
        index = self._favorite_index_for_class_drug(drug)
        if index is None:
            return False
        return bool(cfg.config.medication_favorites()[index].get("last_used"))

    def select_class_browser_group(self, code):
        self._selected_therapeutic_group = code
        details = classes.subclasses_for(code)
        if self._selected_detailed_class not in details:
            self._selected_detailed_class = details[0] if details else None
        self.refresh_class_browser()

    def schedule_class_group_selection(self, code):
        """Delay the single-click action so a double click can open the group."""
        job = getattr(self, "_class_group_click_job", None)
        if job is not None:
            try:
                self.after_cancel(job)
            except (tk.TclError, ValueError):
                pass
        self._class_group_click_job = self.after(
            220, lambda selected=code: self._finish_class_group_selection(selected))

    def _finish_class_group_selection(self, code):
        self._class_group_click_job = None
        self.select_class_browser_group(code)

    def select_class_browser_detail(self, detail):
        self._selected_detailed_class = detail
        self.refresh_class_browser()

    def schedule_class_detail_selection(self, detail):
        """Select on one click while reserving double click for page navigation."""
        job = getattr(self, "_class_detail_click_job", None)
        if job is not None:
            try:
                self.after_cancel(job)
            except (tk.TclError, ValueError):
                pass
        self._class_detail_click_job = self.after(
            220, lambda picked=detail: self._finish_class_detail_selection(picked))

    def _finish_class_detail_selection(self, detail):
        self._class_detail_click_job = None
        self.select_class_browser_detail(detail)

    def _set_class_breadcrumb(self, code=None, detail=None):
        """Show only the selected group and class as clickable breadcrumbs."""
        if not hasattr(self, "class_breadcrumb_group"):
            return
        if not code:
            self.class_breadcrumb_group.configure(text=I.t("major_therapeutic_groups"),
                                                   command=lambda: None)
            self.class_breadcrumb_detail.pack_forget()
            self.class_breadcrumb_separator.pack_forget()
            return
        self.class_breadcrumb_group.configure(
            text=I.t("class_" + code),
            command=lambda selected=code: self.show_subclass_page(selected))
        if detail:
            self.class_breadcrumb_separator.pack(side="left")
            self.class_breadcrumb_detail.configure(
                text=detail,
                command=lambda group=code, picked=detail:
                    self.show_detail_medicines_page(group, picked))
            self.class_breadcrumb_detail.pack(side="left")
        else:
            self.class_breadcrumb_detail.pack_forget()
            self.class_breadcrumb_separator.pack_forget()

    def refresh_class_browser(self):
        """Render the major group, detailed class, and medicine panes in place."""
        if not hasattr(self, "class_detail_list"):
            return
        self._class_search_job = None
        query = self.class_search_var.get().casefold().strip()
        only_unclassified = bool(self.class_unclassified_filter_var.get()) if hasattr(
            self, "class_unclassified_filter_var") else False
        render_signature = (
            query, only_unclassified,
            self._selected_therapeutic_group, self._selected_detailed_class,
            id(self._ensure_classification_cache()),
        )
        if render_signature == getattr(self, "_class_browser_render_signature", None):
            return
        self._class_browser_render_signature = render_signature
        self._set_class_browser_mode(only_unclassified)
        if only_unclassified:
            self._selected_therapeutic_group = None
            self._selected_detailed_class = None
            self.render_unclassified_class_search(query)
            return
        name_mode = self.class_name_filter_var.get() if hasattr(
            self, "class_name_filter_var") else I.t("filter_all_names")
        search_taxonomy = name_mode == I.t("filter_all_names")

        def medicine_matches(drug):
            return self._class_drug_matches_query(drug, query)

        matching_groups = []
        for code in (() if only_unclassified else self.ordered_therapeutic_groups()):
            group_text = I.t("class_" + code).casefold()
            details = classes.subclasses_for(code)
            detail_match = (any(query in detail.casefold() for detail in details)
                            if query and search_taxonomy
                            else not query)
            group_match = bool(query and search_taxonomy and query in group_text)
            medicine_match = any(
                medicine_matches(drug)
                for detail in details for drug in self._drugs_in_class(code, detail))
            if (not query or group_match
                    or detail_match or medicine_match):
                matching_groups.append(code)
        for code in self.ordered_therapeutic_groups():
            tile = self.class_tiles[code]
            tile.pack_forget()
            if code in matching_groups:
                tile.pack(fill="x", pady=3)
        self.class_group_empty_label.pack_forget()
        if not matching_groups and not only_unclassified:
            self.class_group_empty_label.pack(fill="x", padx=6, pady=8)
            self._selected_therapeutic_group = None
            self._selected_detailed_class = None
        if only_unclassified:
            self._selected_therapeutic_group = None
            self._selected_detailed_class = None
        if matching_groups and self._selected_therapeutic_group not in matching_groups:
            self._selected_therapeutic_group = matching_groups[0]
            details = classes.subclasses_for(matching_groups[0])
            self._selected_detailed_class = details[0] if details else None
        code = self._selected_therapeutic_group
        for group_code, button in self.class_buttons.items():
            selected = group_code == code
            button.configure(
                fg_color=PRIMARY if selected else SURFACE,
                text_color="white" if selected else TEXT,
                border_color=ACCENT if selected else LINE)

        for child in self.class_detail_list.winfo_children():
            child.destroy()
        all_details = list(classes.subclasses_for(code)) if code else []
        group_query_match = bool(
            query and search_taxonomy and code
            and query in I.t("class_" + code).casefold())
        visible_details = []
        for detail in all_details:
            drugs = self._drugs_in_class(code, detail)
            detail_query_match = bool(
                query and search_taxonomy and query in detail.casefold())
            matching_drugs = [drug for drug in drugs if medicine_matches(drug)]
            if (not query
                    or group_query_match or detail_query_match or matching_drugs):
                visible_details.append(detail)
        if visible_details and self._selected_detailed_class not in visible_details:
            self._selected_detailed_class = visible_details[0]
        if not visible_details and not only_unclassified:
            self._selected_detailed_class = None
            ctk.CTkLabel(
                self.class_detail_list, text=I.t("no_detailed_classes_found"),
                text_color=MUTED, anchor="w").pack(fill="x", padx=5, pady=8)
        if only_unclassified:
            ctk.CTkLabel(
                self.class_detail_list, text=I.t("unclassified"),
                text_color=WARNING, anchor="w",
                font=ctk.CTkFont(size=13, weight="bold")).pack(
                    fill="x", padx=5, pady=8)
        self.class_detail_buttons = {}
        for detail in visible_details:
            count = len(self._drugs_in_class(code, detail))
            selected = detail == self._selected_detailed_class
            button = VisualButton(
                self.class_detail_list,
                text=I.t("detailed_class_with_count", detail=detail, n=count),
                height=34, corner_radius=8, anchor="w",
                fg_color=PRIMARY if selected else CARD,
                text_color="white" if selected else TEXT,
                border_width=1, border_color=ACCENT if selected else LINE,
                hover_color=ACCENT_SOFT,
                command=lambda picked=detail:
                    self.schedule_class_detail_selection(picked))
            button.pack(fill="x", pady=2)
            button.bind(
                "<Double-Button-1>",
                lambda _event, group=code, picked=detail:
                    self.show_detail_medicines_page(group, picked))
            self.class_detail_buttons[detail] = button

        self._set_class_breadcrumb(code, self._selected_detailed_class)
        self.class_visible_drugs = []

    def toggle_class_drug_star(self, drug):
        index = self._favorite_index_for_class_drug(drug)
        if index is None:
            cfg.config.add_medication_favorite({
                "brand_name": drug.brand_name,
                "generic_name": drug.generic_name,
                "category": drug.category,
                "dosage": drug.strength,
                "pinned": True,
            })
        else:
            cfg.config.toggle_medication_favorite_pin(index)
        self.refresh_class_browser()
        self.refresh_favorites_page()

    def _bind_class_medicine_context(self, widget, drug):
        handler = lambda event, item=drug: self.show_class_medicine_context_menu(event, item)
        def bind_tree(node):
            node.bind("<Button-3>", handler)
            for child in node.winfo_children():
                bind_tree(child)
        bind_tree(widget)

    def show_class_medicine_context_menu(self, event, drug):
        menu = VisualMenu(self, tearoff=False, font=("Segoe UI", 44))
        menu.add_command(
            label=I.t("context_use_rx"),
            command=lambda: self.add_drug_database_item(drug))
        menu.add_command(
            label=(I.t("context_unstar_drug") if self._class_drug_is_starred(drug)
                   else I.t("context_star_drug")),
            command=lambda: self.toggle_class_context_star(drug))
        menu.add_separator()
        menu.add_command(
            label=I.t("context_edit_mapping"),
            command=lambda: self.edit_class_drug_mapping(drug))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def toggle_class_context_star(self, drug):
        group = self._selected_therapeutic_group
        detail = self._selected_detailed_class
        detail_page_open = bool(self.class_subpage.winfo_manager() and group and detail)
        self.toggle_class_drug_star(drug)
        if detail_page_open:
            self.show_detail_medicines_page(group, detail)

    def edit_class_drug_mapping(self, drug):
        self.show_class_mapping_editor(target_drug=drug)
        if self.mapping_visible_drugs:
            self.mapping_list.selection_set(0)
            self.mapping_list.activate(0)
            self.select_mapping_drug()

    def view_class_drug_reference(self, drug):
        self.show_page("reference")
        medicine = drug.generic_name.strip() or drug.brand_name.strip()
        if not medicine:
            return
        self.reference_lookup_button.configure(state="disabled")
        self.reference_status.configure(text=I.t("openfda_searching"))
        self._clear_reference_cards()
        threading.Thread(
            target=self._lookup_openfda_worker, args=([medicine],), daemon=True).start()

    def class_mapping_integrity(self):
        invalid_groups = []
        invalid_details = []
        mapping_by_name = {}
        mapping_by_identity = {}
        for drug in self.db.drugs:
            key = drug.generic_name.strip().casefold()
            mappings = classes.groups_for(drug)
            mapping_set = frozenset((mapping.code, mapping.detail) for mapping in mappings)
            if key:
                mapping_by_name.setdefault(key, []).append(mapping_set)
            mapping_by_identity.setdefault(
                self.db.drug_identity(drug), []).append(mapping_set)
            if drug.mapping_status == "unrecognized":
                invalid_details.append(drug.detailed_class or drug.generic_name or drug.brand_name)
            for mapping in mappings:
                if mapping.code not in classes.GROUPS:
                    invalid_groups.append(drug.generic_name)
                elif mapping.detail not in classes.subclasses_for(mapping.code):
                    invalid_details.append(drug.generic_name)
        variations = sum(
            1 for values in mapping_by_name.values()
            if len(set(values)) > 1)
        conflicts = sum(
            1 for values in mapping_by_identity.values()
            if len(set(values)) > 1)
        missing_favorites = sum(
            1 for favorite in cfg.config.medication_favorites()
            if not self._favorite_matches_database(favorite))
        return {
            "unclassified": self.unclassified_medicine_count(),
            "invalid_groups": len(invalid_groups),
            "invalid_details": len(invalid_details),
            "variations": variations,
            "conflicts": conflicts,
            "missing_favorites": missing_favorites,
        }

    def show_mapping_integrity_report(self):
        report = self.class_mapping_integrity()
        messagebox.showinfo(
            I.t("mapping_integrity"), I.t("mapping_integrity_report", **report),
            parent=self)

    def toggle_therapeutic_group_favorite(self, code):
        cfg.config.toggle_favorite_therapeutic_group(code)
        self.refresh_class_overview()

    def show_class_mapping_editor(self, review_unclassified=False, target_drug=None):
        """Edit local class metadata without requiring CSV editing."""
        self.class_overview.pack_forget()
        for child in self.class_subpage.winfo_children():
            child.destroy()
        self.class_subpage.pack(fill="both", expand=True)
        card = self.section(self.class_subpage, "")
        bar = ctk.CTkFrame(card, fg_color="transparent")
        bar.pack(fill="x", padx=10, pady=(7, 5))
        ctk.CTkLabel(bar, text=I.t("class_mapping_editor"), text_color=ACCENT,
                     font=ctk.CTkFont(size=16, weight="bold"), anchor="w").pack(
                         side="left", padx=(0, 8))
        VisualButton(bar, text="← " + I.t("back"), height=ACTION_HEIGHT,
                      width=86, fg_color=CARD, text_color=ACCENT,
                      border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
                      command=self.back_to_major_groups).pack(side="left")
        body = ctk.CTkFrame(card, height=720, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        body.pack_propagate(False)
        left = ctk.CTkFrame(body, fg_color="transparent")
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))
        right = GlassFrame(body, width=360, fg_color=SURFACE,
                             border_color=LINE, border_width=1, corner_radius=12)
        right.pack(side="left", fill="both")
        self.mapping_search_var = tk.StringVar()
        self._mapping_target_identity = (
            self.db.drug_identity(target_drug) if target_drug else None)
        self._mapping_search_job = None
        self.mapping_search_var.trace_add("write", lambda *_: self._schedule_mapping_refresh())
        mapping_control_font = ctk.CTkFont(size=13)
        VisualEntry(
            left, textvariable=self.mapping_search_var, height=FIELD_HEIGHT,
            placeholder_text=I.t("search_medicines"), border_width=2,
            border_color=ACCENT, corner_radius=9,
            font=mapping_control_font).pack(fill="x", pady=(0, 6))
        self.mapping_unclassified_only = tk.BooleanVar(value=False)
        self.mapping_unclassified_only.set(bool(review_unclassified))
        self.mapping_suggested_only = tk.BooleanVar(value=False)
        self.mapping_conflicting_only = tk.BooleanVar(value=False)
        mapping_filters = ctk.CTkFrame(left, fg_color="transparent")
        mapping_filters.pack(fill="x", pady=(0, 6))
        ctk.CTkCheckBox(
            mapping_filters, text=I.t("show_unclassified_only"),
            variable=self.mapping_unclassified_only, text_color=MUTED,
            fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            command=self.refresh_mapping_list).pack(side="left", padx=(0, 16))
        ctk.CTkCheckBox(
            mapping_filters, text=I.t("show_suggested"),
            variable=self.mapping_suggested_only, text_color=MUTED,
            fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            command=self.refresh_mapping_list).pack(side="left", padx=(0, 16))
        ctk.CTkCheckBox(
            mapping_filters, text=I.t("show_mapping_variations"),
            variable=self.mapping_conflicting_only, text_color=MUTED,
            fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            command=self.refresh_mapping_list).pack(side="left")
        self.mapping_list = tk.Listbox(left, height=5, font=LIST_FONT,
                                       bg=SURFACE, fg=TEXT, relief="flat", borderwidth=0,
                                       highlightthickness=1, highlightbackground=LINE,
                                       selectbackground=PRIMARY, selectforeground="white", activestyle="none")
        self.mapping_list.configure(selectmode=tk.EXTENDED, exportselection=False)
        self.mapping_list.pack(fill="both", expand=True)
        self.mapping_list.bind("<<ListboxSelect>>", self.select_mapping_drug)
        review_actions = ctk.CTkFrame(left, fg_color="transparent")
        review_actions.pack(fill="x", pady=(6, 0))
        self.mapping_result_label = ctk.CTkLabel(
            review_actions, text="", text_color=MUTED,
            font=ctk.CTkFont(size=10), anchor="w")
        self.mapping_result_label.pack(side="left", padx=(0, 8))
        VisualButton(
            review_actions, text=I.t("previous"), width=92, height=32,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=lambda: self.mapping_select_relative(-1)).pack(
                side="left", padx=(0, 4))
        VisualButton(
            review_actions, text=I.t("skip"), width=78, height=32,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=lambda: self.mapping_select_relative(1)).pack(
                side="left", padx=4)
        VisualButton(
            review_actions, text=I.t("save_and_next"), width=130, height=32,
            fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            command=self.mapping_save_and_next).pack(side="right")
        self.mapping_group_labels = {I.t("class_" + code): code for code in classes.GROUPS
                                     if classes.subclasses_for(code)}
        self.mapping_group_var = tk.StringVar(value=I.t("choose_major_group"))
        self.mapping_detail_var = tk.StringVar(value=I.t("choose_detailed_class"))
        VisualButton(
            right, text=I.t("mapping_integrity"), width=150, height=34,
            fg_color=CARD, text_color=ACCENT, border_width=1,
            border_color=LINE, hover_color=ACCENT_SOFT,
            command=self.show_mapping_integrity_report).pack(
                anchor="center", pady=(10, 5))
        self.mapping_selected_label = ctk.CTkLabel(right, text=I.t("no_medicine_selected"),
                                                    text_color=TEXT, anchor="w",
                                                    font=ctk.CTkFont(
                                                        size=round(SELECTED_MEDICINE_FONT_SIZE / 2),
                                                        weight="bold"), wraplength=246)
        self.mapping_selected_label.pack(fill="x", padx=12, pady=(10, 8))
        ctk.CTkLabel(right, text=I.t("major_therapeutic_group"), text_color=MUTED,
                     font=ctk.CTkFont(size=10, weight="bold"), anchor="w").pack(fill="x", padx=12)
        self.mapping_group_menu = VisualOptionMenu(
            right, values=[I.t("choose_major_group")] + list(self.mapping_group_labels),
            variable=self.mapping_group_var, height=FIELD_HEIGHT,
            corner_radius=9, font=mapping_control_font,
            dropdown_font=mapping_control_font,
            fg_color=ACCENT_SOFT, text_color=ACCENT,
            button_color=PRIMARY, button_hover_color=ACCENT_HOVER,
            command=self.mapping_group_changed)
        self.mapping_group_menu.pack(fill="x", padx=12, pady=(3, 8))
        self.mapping_detail_shell = ctk.CTkFrame(
            right, height=76, fg_color="transparent")
        self.mapping_detail_shell.pack_propagate(False)
        self.mapping_detail_label = ctk.CTkLabel(
            self.mapping_detail_shell, text=I.t("detailed_drug_class"), text_color=MUTED,
            font=ctk.CTkFont(size=10, weight="bold"), anchor="w")
        self.mapping_detail_label.pack(fill="x", pady=(0, 3))
        self.mapping_detail_menu = VisualOptionMenu(
            self.mapping_detail_shell, values=[I.t("choose_detailed_class")],
            variable=self.mapping_detail_var, height=FIELD_HEIGHT,
            corner_radius=9, font=mapping_control_font,
            dropdown_font=mapping_control_font,
            fg_color=ACCENT_SOFT, text_color=ACCENT, button_color=PRIMARY,
            button_hover_color=ACCENT_HOVER, command=lambda _: self.mapping_detail_changed())
        self.mapping_detail_menu.pack(fill="x")
        self.mapping_keep_existing_var = tk.BooleanVar(value=True)
        self.mapping_keep_existing_checkbox = ctk.CTkCheckBox(
            right, text=I.t("keep_existing_classes"),
            variable=self.mapping_keep_existing_var, fg_color=PRIMARY,
            hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=11))
        self.mapping_keep_existing_checkbox.pack(fill="x", padx=12, pady=(0, 6))
        mapping_actions = ctk.CTkFrame(right, fg_color="transparent")
        mapping_actions.pack(fill="x", padx=12)
        self.add_mapping_drug_button = VisualButton(
            mapping_actions, text=I.t("add_drug_to_class"), height=34,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, state="disabled",
                                                      command=self.add_drug_to_class)
        self.add_mapping_drug_button.pack(
            side="left", fill="x", expand=True, padx=(0, 3))
        self.save_mapping_button = VisualButton(
            mapping_actions, text=I.t("save_changes"), height=34,
            fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            state="disabled", command=self.save_class_mapping)
        self.save_mapping_button.pack(
            side="left", fill="x", expand=True, padx=(3, 0))
        self.mapping_status = ctk.CTkLabel(right, text="", text_color=MUTED,
                                           font=ctk.CTkFont(size=10), wraplength=270, justify="left")
        self.mapping_status.pack(fill="x", padx=12, pady=(8, 0))
        self.mapping_selected_drug = None
        self.mapping_selected_drugs = []
        self.pending_class_mapping = None
        self.mapping_visible_drugs = []
        self.refresh_mapping_list()
        if (review_unclassified or target_drug) and self.mapping_visible_drugs:
            self.mapping_list.selection_set(0)
            self.mapping_list.activate(0)
            self.select_mapping_drug()

    def _schedule_mapping_refresh(self, delay=180):
        if self._mapping_search_job is not None:
            try:
                self.after_cancel(self._mapping_search_job)
            except (tk.TclError, ValueError):
                pass
        self._mapping_search_job = self.after(delay, self.refresh_mapping_list)

    def refresh_mapping_list(self):
        if not hasattr(self, "mapping_list"):
            return
        self._mapping_search_job = None
        query = self.mapping_search_var.get().casefold().strip()
        target_identity = getattr(self, "_mapping_target_identity", None)
        if target_identity:
            drugs = [drug for drug in self.db.drugs
                     if self.db.drug_identity(drug) == target_identity]
        elif (self.mapping_unclassified_only.get()
              or self.mapping_suggested_only.get()
              or self.mapping_conflicting_only.get()):
            cache = self._ensure_classification_cache()
            candidates = []
            if self.mapping_unclassified_only.get():
                candidates.extend(cache["unclassified"])
            if self.mapping_suggested_only.get():
                candidates.extend(cache["suggested"])
            if self.mapping_conflicting_only.get():
                candidates.extend(cache["variations"])
            candidates = list({self.db.drug_identity(drug): drug
                               for drug in candidates}.values())
            drugs = [drug for drug in candidates
                     if not query or query in (
                         drug.generic_name + " " + drug.brand_name).casefold()]
        elif query:
            drugs = self.db.search(query, limit=501)
        else:
            drugs = self.db.fetch_page(0, 501)
        total_visible = len(drugs)
        drugs = drugs[:500]
        self.mapping_visible_drugs = sorted(
            drugs, key=lambda drug: (drug.generic_name or drug.brand_name).casefold())
        if hasattr(self, "mapping_result_label"):
            suffix = "+" if total_visible > 500 else ""
            self.mapping_result_label.configure(text=f"{len(self.mapping_visible_drugs)}{suffix}")
        self.mapping_list.delete(0, tk.END)
        for drug in self.mapping_visible_drugs:
            found = classes.groups_for(drug)
            if drug.mapping_status == "unrecognized":
                suffix = I.t("unrecognized_class_detail", detail=drug.detailed_class)
                status = "unrecognized"
            else:
                suffix = (", ".join(mapping.detail for mapping in found)
                          if found else I.t("unclassified"))
                status = found[0].confidence if found else "unclassified"
            medicine_name = drug.generic_name.strip() or drug.brand_name.strip()
            self.mapping_list.insert(
                tk.END,
                f"{medicine_name}  —  {suffix}  [{I.t('classification_' + status)}]")
        self.mapping_selected_drugs = []
        self.mapping_selected_drug = None
        self.pending_class_mapping = None
        if hasattr(self, "mapping_selected_label"):
            self.mapping_selected_label.configure(text=I.t("no_medicine_selected"))
        if hasattr(self, "add_mapping_drug_button"):
            self.add_mapping_drug_button.configure(state="disabled")
        if hasattr(self, "save_mapping_button"):
            self.save_mapping_button.configure(state="disabled")

    def select_mapping_drug(self, event=None):
        selected_indices = self.mapping_list.curselection()
        if not selected_indices:
            return
        selected = [self.mapping_visible_drugs[index] for index in selected_indices]
        self.mapping_selected_drugs = selected
        drug = selected[0]
        self.mapping_selected_drug = drug
        if len(selected) == 1:
            selected_name = drug.brand_name.strip() or drug.generic_name
            if drug.brand_name.strip() and drug.generic_name.strip():
                selected_name += "\n" + drug.generic_name
            self.mapping_selected_label.configure(text=selected_name)
        else:
            self.mapping_selected_label.configure(
                text=I.t("mapping_selected_count", n=len(selected)))
        found = classes.group_for(drug)
        if len(selected) == 1 and found and found.code in self.mapping_group_labels.values():
            label = next(name for name, code in self.mapping_group_labels.items() if code == found.code)
            self.mapping_group_var.set(label)
            self.mapping_group_changed(label, preferred_detail=found.detail)
        else:
            self.mapping_group_var.set(I.t("choose_major_group"))
            self.mapping_detail_menu.configure(values=[I.t("choose_detailed_class")])
            self.mapping_detail_var.set(I.t("choose_detailed_class"))
            self.mapping_detail_shell.pack_forget()
        self.pending_class_mapping = None
        self.add_mapping_drug_button.configure(state="normal")
        self.save_mapping_button.configure(state="disabled")
        self.mapping_status.configure(text="", text_color=MUTED)

    def mapping_group_changed(self, label, preferred_detail=None):
        code = self.mapping_group_labels.get(label)
        details = list(classes.subclasses_for(code)) if code else []
        if not code:
            self.mapping_detail_shell.pack_forget()
            self.mapping_detail_menu.configure(values=[I.t("choose_detailed_class")])
            self.mapping_detail_var.set(I.t("choose_detailed_class"))
            return
        if not self.mapping_detail_shell.winfo_manager():
            self.mapping_detail_shell.pack(fill="x", padx=12, pady=(0, 8),
                                           before=self.mapping_keep_existing_checkbox)
        values = details or [I.t("choose_detailed_class")]
        self.mapping_detail_menu.configure(values=values)
        self.mapping_detail_var.set(preferred_detail if preferred_detail in details else values[0])

    def mapping_detail_changed(self):
        self.pending_class_mapping = None
        self.save_mapping_button.configure(state="disabled")
        self.mapping_status.configure(text="", text_color=MUTED)

    def add_drug_to_class(self):
        if not self.mapping_selected_drugs:
            return
        code = self.mapping_group_labels.get(self.mapping_group_var.get())
        detail = self.mapping_detail_var.get()
        if not code or detail == I.t("choose_detailed_class"):
            self.mapping_status.configure(text=I.t("choose_group_and_class"), text_color=DANGER)
            return
        medicines = tuple(self.mapping_selected_drugs)
        self.pending_class_mapping = (medicines, code, detail)
        self.mapping_status.configure(
            text=I.t("mapping_batch_ready", n=len(medicines), detail=detail),
            text_color=GOOD)
        self.save_mapping_button.configure(state="normal")

    def save_class_mapping(self, advance=False):
        if not self.pending_class_mapping:
            return False
        medicines, code, detail = self.pending_class_mapping
        prior_states = self.db.classification_states_for_drugs(medicines)
        current_indices = self.mapping_list.curselection()
        next_index = current_indices[0] if current_indices else 0
        updated = self.db.update_drug_classifications(
            medicines, code, detail, append=self.mapping_keep_existing_var.get())
        if updated:
            if prior_states:
                cfg.config.add_recovery_item(
                    "mapping", I.t("mapping_recovery_label", n=updated), prior_states)
            self._invalidate_database_caches()
            self.mapping_status.configure(
                text=I.t("class_mapping_batch_saved", n=updated,
                         group=I.t("class_" + code), detail=detail), text_color=GOOD)
            self.refresh_mapping_list()
            self.refresh_class_overview()
            self.pending_class_mapping = None
            self.mapping_selected_drugs = []
            self.mapping_selected_drug = None
            self.save_mapping_button.configure(state="disabled")
            if advance and self.mapping_visible_drugs:
                next_index = min(next_index, len(self.mapping_visible_drugs) - 1)
                self.mapping_list.selection_set(next_index)
                self.mapping_list.activate(next_index)
                self.mapping_list.see(next_index)
                self.select_mapping_drug()
            return True
        return False

    def mapping_select_relative(self, step):
        if not self.mapping_visible_drugs:
            return
        selected = self.mapping_list.curselection()
        current = selected[0] if selected else (0 if step > 0 else len(self.mapping_visible_drugs) - 1)
        target = max(0, min(len(self.mapping_visible_drugs) - 1, current + step))
        self.mapping_list.selection_clear(0, tk.END)
        self.mapping_list.selection_set(target)
        self.mapping_list.activate(target)
        self.mapping_list.see(target)
        self.select_mapping_drug()

    def mapping_save_and_next(self):
        if not self.mapping_selected_drugs:
            self.mapping_select_relative(1)
            return
        if not self.pending_class_mapping:
            self.add_drug_to_class()
        if self.pending_class_mapping:
            self.save_class_mapping(advance=True)

    def open_therapeutic_group(self, code):
        """Open a focused subpage containing this group's detailed classes."""
        job = getattr(self, "_class_group_click_job", None)
        if job is not None:
            try:
                self.after_cancel(job)
            except (tk.TclError, ValueError):
                pass
            self._class_group_click_job = None
        self._selected_therapeutic_group = code
        self.show_subclass_page(code)

    def show_all_detailed_classes(self):
        """Show every configured detailed class, grouped on one scrollable page."""
        self.class_page_header.pack_forget()
        self.class_overview.pack_forget()
        for child in self.class_subpage.winfo_children():
            child.destroy()
        self.class_subpage.pack(fill="both", expand=True)
        card = ctk.CTkFrame(self.class_subpage, fg_color="transparent")
        card.pack(fill="both", expand=True)
        toolbar = ctk.CTkFrame(card, fg_color="transparent")
        toolbar.pack(fill="x", padx=2, pady=(0, 5))
        toolbar.grid_columnconfigure((1, 2), weight=1, uniform="all_class_toolbar")
        VisualButton(
            toolbar, text="← " + I.t("back"), height=ACTION_HEIGHT,
            width=86, fg_color=CARD, text_color=ACCENT,
            border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
            command=self.back_to_major_groups).grid(row=0, column=0, sticky="w")
        self.all_class_search_var = tk.StringVar()
        self.all_class_search_var.trace_add("write", lambda *_: self.render_all_detailed_classes())
        VisualEntry(
            toolbar, textvariable=self.all_class_search_var, height=FIELD_HEIGHT,
            placeholder_text=I.t("search_detailed_classes"),
            border_color=LINE).grid(row=0, column=1, sticky="ew", padx=(6, 3))
        self.all_classes_body = ctk.CTkScrollableFrame(
            card, height=520, fg_color="transparent")
        self.all_classes_body.pack(
            fill="both", expand=True, padx=2, pady=(0, 2))
        self.render_all_detailed_classes()

    def render_all_detailed_classes(self):
        if not hasattr(self, "all_classes_body"):
            return
        for child in self.all_classes_body.winfo_children():
            child.destroy()
        query = self.all_class_search_var.get().casefold().strip()
        found_any = False
        for group_code in classes.GROUPS:
            details = classes.subclasses_for(group_code)
            if query:
                group_name = I.t("class_" + group_code).casefold()
                details = tuple(detail for detail in details
                                if query in detail.casefold() or query in group_name)
            if not details:
                continue
            found_any = True
            group = ctk.CTkFrame(self.all_classes_body, fg_color="transparent")
            group.pack(fill="x", pady=(4, 2))
            ctk.CTkLabel(
                group, text=I.t("class_" + group_code), text_color=ACCENT,
                font=ctk.CTkFont(size=13, weight="bold"), anchor="w").pack(
                    fill="x", pady=(2, 3))
            grid = ctk.CTkFrame(group, fg_color="transparent")
            grid.pack(fill="x")
            grid.grid_columnconfigure((0, 1), weight=1, uniform="detailed_classes")
            for index, detail in enumerate(details):
                VisualButton(
                    grid, text=detail, height=33, corner_radius=8, anchor="w",
                    fg_color=SURFACE, text_color=TEXT, border_width=1,
                    border_color=LINE, hover_color=ACCENT_SOFT,
                    command=lambda group=group_code, picked=detail:
                        self.show_subclass_page(group, picked, return_to_all=True)).grid(
                            row=index // 2, column=index % 2,
                            sticky="ew", padx=2, pady=2)
        if not found_any:
            ctk.CTkLabel(self.all_classes_body, text=I.t("no_detailed_classes_found"),
                         text_color=MUTED, anchor="w").pack(fill="x", pady=12)

    def show_subclass_page(self, group_code, selected_detail=None, return_to_all=False):
        self._active_subclass_group = group_code
        self._class_detail_return = "all" if return_to_all else "overview"
        self.class_overview.pack_forget()
        for child in self.class_subpage.winfo_children():
            child.destroy()
        self.class_subpage.pack(fill="both", expand=True)
        card = self.section(self.class_subpage, "")
        bar = ctk.CTkFrame(card, fg_color="transparent")
        bar.pack(fill="x", padx=PAD, pady=(0, 8))
        VisualButton(bar, text="← " + I.t("back"), height=ACTION_HEIGHT,
                      width=86, fg_color=CARD, text_color=ACCENT,
                      border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
                      command=self.back_from_subclass_page).pack(side="left")
        trail = ctk.CTkFrame(bar, fg_color="transparent")
        trail.pack(side="left", padx=(8, 0))
        VisualButton(
            trail, text=I.t("major_therapeutic_groups"), height=34, width=40,
            fg_color="transparent", text_color=ACCENT, hover_color=ACCENT_SOFT,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.back_to_major_groups).pack(side="left")
        ctk.CTkLabel(
            trail, text="›", width=24, text_color=MUTED,
            font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")
        ctk.CTkLabel(
            trail, text=I.t("class_" + group_code), text_color=ACCENT,
            font=ctk.CTkFont(size=14, weight="bold"), anchor="w").pack(side="left")
        list_frame = ctk.CTkScrollableFrame(card, height=470, fg_color="transparent")
        list_frame.pack(fill="both", expand=True, padx=PAD, pady=(0, PAD))
        self.subclass_buttons = {}
        subclass_counts = {}
        for detail in classes.subclasses_for(group_code):
            subclass_counts[detail] = len(self._drugs_in_class(group_code, detail))
        for detail in classes.subclasses_for(group_code):
            button = VisualButton(
                list_frame, text=I.t("detailed_class_with_count", detail=detail,
                                     n=subclass_counts[detail]), height=36, corner_radius=9, anchor="w",
                fg_color=SURFACE, text_color=TEXT, border_width=1,
                border_color=LINE, hover_color=ACCENT_SOFT,
                command=lambda picked=detail:
                    self.schedule_subclass_selection(picked))
            button.pack(fill="x", pady=3)
            button.bind(
                "<Double-Button-1>",
                lambda _event, group=group_code, picked=detail:
                    self.show_detail_medicines_page(group, picked))
            self.subclass_buttons[detail] = button
        if selected_detail:
            self.show_detail_medicines_page(group_code, selected_detail)

    def schedule_subclass_selection(self, detail):
        job = getattr(self, "_subclass_click_job", None)
        if job is not None:
            try:
                self.after_cancel(job)
            except (tk.TclError, ValueError):
                pass
        self._subclass_click_job = self.after(
            220, lambda picked=detail: self._finish_subclass_selection(picked))

    def _finish_subclass_selection(self, detail):
        self._subclass_click_job = None
        self._selected_subpage_detail = detail
        for name, button in self.subclass_buttons.items():
            selected = name == detail
            button.configure(
                fg_color=PRIMARY if selected else SURFACE,
                text_color="white" if selected else TEXT,
                border_color=ACCENT if selected else LINE)

    def show_detail_medicines_page(self, group_code, detail):
        """Open one dedicated page for the medicines mapped to a detail class."""
        job = getattr(self, "_class_detail_click_job", None)
        if job is not None:
            try:
                self.after_cancel(job)
            except (tk.TclError, ValueError):
                pass
            self._class_detail_click_job = None
        job = getattr(self, "_subclass_click_job", None)
        if job is not None:
            try:
                self.after_cancel(job)
            except (tk.TclError, ValueError):
                pass
            self._subclass_click_job = None
        self._active_subclass_group = group_code
        self._selected_therapeutic_group = group_code
        self._selected_detailed_class = detail
        self.class_overview.pack_forget()
        for child in self.class_subpage.winfo_children():
            child.destroy()
        self.class_subpage.pack(fill="both", expand=True)
        card = self.section(self.class_subpage, "")
        bar = ctk.CTkFrame(card, fg_color="transparent")
        bar.pack(fill="x", padx=PAD, pady=(0, 10))
        self._detail_page_bar = bar
        VisualButton(
            bar, text="← " + I.t("back"),
            height=ACTION_HEIGHT, width=86, fg_color=CARD, text_color=ACCENT,
            border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
            command=self.back_to_major_groups).pack(side="left")
        trail = ctk.CTkFrame(bar, fg_color="transparent")
        trail.pack(side="left", padx=(8, 0))
        VisualButton(
            trail, text=I.t("class_" + group_code), height=34, width=40,
            fg_color="transparent", text_color=ACCENT, hover_color=ACCENT_SOFT,
            font=ctk.CTkFont(size=16, weight="bold"),
            command=lambda: self.show_subclass_page(group_code)).pack(side="left")
        ctk.CTkLabel(
            trail, text="›", width=24, text_color=MUTED,
            font=ctk.CTkFont(size=17, weight="bold")).pack(side="left")
        ctk.CTkLabel(
            trail, text=detail, text_color=ACCENT, anchor="e",
            font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")
        add_button = VisualButton(
            bar, text="+", width=34, height=34, corner_radius=17,
            fg_color=SURFACE, text_color=ACCENT, hover_color=ACCENT_SOFT,
            border_width=1, border_color=LINE,
            font=ctk.CTkFont(size=20, weight="bold"),
            command=self.toggle_detail_drug_creator)
        add_button.pack(side="right")

        self.detail_drug_creator = ctk.CTkFrame(
            card, fg_color=ACCENT_SOFT, border_color=LINE,
            border_width=1, corner_radius=9)
        self.detail_drug_name_var = tk.StringVar()
        ctk.CTkLabel(
            self.detail_drug_creator, text=I.t("drug_name"),
            text_color=MUTED, font=ctk.CTkFont(size=11, weight="bold")).pack(
                side="left", padx=(10, 5), pady=7)
        self.detail_drug_name_entry = VisualEntry(
            self.detail_drug_creator, textvariable=self.detail_drug_name_var,
            width=360, height=36, border_color=ACCENT,
            placeholder_text=I.t("class_add_drug_placeholder"))
        self.detail_drug_name_entry.pack(
            side="left", fill="x", expand=True, padx=5, pady=7)
        self.detail_drug_name_entry.bind(
            "<Return>", lambda _event: self.save_detail_class_medicine(
                group_code, detail))
        VisualButton(
            self.detail_drug_creator, text=I.t("save"), width=80, height=34,
            fg_color=GOOD, hover_color=GOOD_HOVER,
            command=lambda: self.save_detail_class_medicine(
                group_code, detail)).pack(side="right", padx=(5, 10), pady=7)

        medicines = self._drugs_in_class(group_code, detail)
        ctk.CTkLabel(
            card, text=I.t("mapped_medicine_count", detail=detail, n=len(medicines)),
            text_color=MUTED, anchor="w",
            font=ctk.CTkFont(size=11)).pack(fill="x", padx=PAD, pady=(0, 7))
        body = ctk.CTkScrollableFrame(card, height=450, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=PAD, pady=(0, PAD))
        if not medicines:
            ctk.CTkLabel(body, text=I.t("no_mapped_medicines"),
                         text_color=MUTED, anchor="w").pack(fill="x", pady=10)
            return
        self._detail_page_body = body
        self._detail_page_medicines = medicines
        self._detail_page_group = group_code
        self._detail_page_class = detail
        self._detail_page_rendered = 0
        self._detail_page_size = 40
        self._detail_page_footer = None
        self._append_detail_medicine_page()

    def toggle_detail_drug_creator(self):
        panel = getattr(self, "detail_drug_creator", None)
        if panel is None:
            return
        if panel.winfo_manager():
            panel.pack_forget()
            return
        panel.pack(fill="x", padx=PAD, pady=(0, 7), after=self._detail_page_bar)
        self.detail_drug_name_entry.focus_set()

    def save_detail_class_medicine(self, group_code, detail):
        name = self.detail_drug_name_var.get().strip()
        if not name:
            messagebox.showinfo(
                I.t("add_drug_tooltip"), I.t("class_add_drug_required"), parent=self)
            self.detail_drug_name_entry.focus_set()
            return
        try:
            _drug, created = self.db.add_confirmed_medicine(
                name, group_code, detail)
        except (OSError, ValueError) as exc:
            messagebox.showerror(I.t("add_drug_tooltip"), str(exc), parent=self)
            return
        self._invalidate_database_caches()
        self.db_label.configure(text=I.t("db_count", n=len(self.db.drugs)))
        messagebox.showinfo(
            I.t("add_drug_tooltip"),
            I.t("class_drug_added" if created else "class_drug_mapped",
                medicine=name, detail=detail), parent=self)
        self.show_detail_medicines_page(group_code, detail)

    def _append_detail_medicine_page(self):
        footer = getattr(self, "_detail_page_footer", None)
        if footer is not None and footer.winfo_exists():
            footer.destroy()
        start = self._detail_page_rendered
        end = min(start + self._detail_page_size, len(self._detail_page_medicines))
        for drug in self._detail_page_medicines[start:end]:
            self._add_detail_medicine_row(
                self._detail_page_body, drug,
                self._detail_page_group, self._detail_page_class)
        self._detail_page_rendered = end
        remaining = len(self._detail_page_medicines) - end
        if remaining <= 0:
            self._detail_page_footer = None
            return
        footer = ctk.CTkFrame(self._detail_page_body, fg_color="transparent")
        footer.pack(fill="x", pady=8)
        self._detail_page_footer = footer
        ctk.CTkLabel(
            footer, text=I.t("more_class_medicines", n=remaining),
            text_color=MUTED).pack(side="left", padx=(4, 10))
        VisualButton(
            footer, text=I.t("load_more"), width=110, height=32,
            fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            command=self._append_detail_medicine_page).pack(side="left")

    def _add_detail_medicine_row(self, body, drug, group_code, detail):
        row = ctk.CTkFrame(body, fg_color=CARD, border_color=LINE,
                           border_width=1, corner_radius=9)
        row.pack(fill="x", pady=3)
        names = ctk.CTkFrame(row, fg_color="transparent")
        names.pack(side="left", fill="x", expand=True, padx=10, pady=7)
        brand = drug.brand_name.strip() or drug.generic_name.strip()
        scientific = drug.generic_name.strip() if drug.brand_name.strip() else ""
        ctk.CTkLabel(
            names, text=brand, text_color=TEXT, anchor="w",
            font=_ui_font(13, "bold")).pack(side="left")
        if scientific and scientific.casefold() != brand.casefold():
            ctk.CTkLabel(
                names, text="  " + scientific, text_color=MUTED, anchor="w",
                font=_ui_font(12)).pack(side="left")
        status = self._classification_status_for(drug, group_code, detail)
        status_background, status_foreground = self._classification_status_colors(status)
        ctk.CTkLabel(
            names, text=I.t("classification_" + status), height=22,
            corner_radius=11, fg_color=status_background,
            text_color=status_foreground,
            font=_ui_font(9, "bold")).pack(
                side="left", padx=(8, 0))
        delete_button = VisualButton(
            row, text="🗑", width=32, height=32,
            fg_color="transparent", text_color=DANGER, border_width=0,
            hover_color=DANGER_SOFT, font=_ui_font(14),
            command=lambda item=drug, group=group_code, picked=detail:
                self.delete_detail_class_drug(item, group, picked))
        delete_button.pack(side="right", padx=(2, 7), pady=5)
        VisualButton(
            row, text="★" if self._class_drug_is_starred(drug) else "☆",
            width=36, height=32, fg_color="transparent", text_color=ACCENT,
            hover_color=ACCENT_SOFT,
            command=lambda item=drug, group=group_code, picked=detail:
                self.toggle_detail_drug_star(item, group, picked)).pack(
                    side="right", padx=(2, 7), pady=5)
        VisualButton(
            row, text="", image=_edit_icon(20), width=32, height=32, corner_radius=8,
            fg_color="transparent", text_color=ACCENT,
            border_width=0, hover_color=ACCENT_SOFT,
            font=_ui_font(14, "bold"),
            command=lambda item=drug: self.edit_class_drug_mapping(item)).pack(
                side="right", padx=2, pady=5)
        add_button = VisualButton(
            row, text="+", width=32, height=32,
            fg_color="transparent", text_color=ACCENT, border_width=0,
            hover_color=ACCENT_SOFT, font=_ui_font(22),
            command=lambda item=drug: self.add_drug_database_item(item))
        add_button.pack(side="right", padx=2, pady=5)
        self._bind_class_medicine_context(row, drug)

    def delete_detail_class_drug(self, drug, group_code, detail):
        """Delete one database product only after explicit confirmation."""
        name = drug.brand_name.strip() or drug.generic_name.strip()
        if not messagebox.askyesno(
                I.t("class_delete_drug"),
                I.t("class_delete_drug_confirm", medicine=name), parent=self):
            return
        try:
            deleted = self.db.delete_drug(drug)
        except OSError as exc:
            logging.exception("Could not delete medicine from the local database")
            messagebox.showerror(
                I.t("class_delete_drug"), str(exc), parent=self)
            return
        if not deleted:
            messagebox.showerror(
                I.t("class_delete_drug"), I.t("class_delete_drug_failed"), parent=self)
            return
        self._invalidate_database_caches()
        self.db_label.configure(text=I.t("db_count", n=len(self.db.drugs)))
        self.show_detail_medicines_page(group_code, detail)

    def toggle_detail_drug_star(self, drug, group_code, detail):
        self.toggle_class_drug_star(drug)
        self.show_detail_medicines_page(group_code, detail)

    def back_to_major_groups(self):
        self.class_subpage.pack_forget()
        if not self.class_page_header.winfo_manager():
            self.class_page_header.pack(fill="x", padx=4, pady=(0, 3))
        self.class_overview.pack(fill="both", expand=True)
        code = self._active_subclass_group
        if code not in classes.GROUPS:
            code = classes.GROUPS[0]
        self.select_therapeutic_group(code)

    def back_from_subclass_page(self):
        if getattr(self, "_class_detail_return", "overview") == "all":
            self.show_all_detailed_classes()
            return
        self.back_to_major_groups()

    def add_drug_database_item(self, drug):
        """Place a browsed medicine into the first blank prescription row."""
        target = next((row for row in self.rows
                       if not row.name_var.get().strip()
                       and not row.trade_var.get().strip()), None)
        if target is None:
            target = self.add_row()
        target.name_var.set(drug.generic_name)
        target.trade_var.set(drug.brand_name)
        target._brand = drug.brand_name
        target._category = drug.category
        target._form = drug.form
        target._auto_quantity()
        target._update_class_badge(drug)
        self.show_page("medications")
        self.on_any_change()

    def select_therapeutic_group(self, code):
        """Backward-compatible entry point for selecting the browser group."""
        self.select_class_browser_group(code)

    def move_row(self, row, direction):
        index = self.rows.index(row)
        target = index + direction
        if not 0 <= target < len(self.rows):
            return
        self.rows[index], self.rows[target] = self.rows[target], self.rows[index]
        # Repack only the moved row. Rebuilding every medication widget made
        # long prescriptions visibly stutter during navigation and scrolling.
        row.pack_forget()
        if target + 1 < len(self.rows):
            row.pack(fill="x", padx=2, pady=4, before=self.rows[target + 1])
        else:
            row.pack(fill="x", padx=2, pady=4)
        self._number_drug_rows()
        self.on_any_change()

    def drag_row(self, row, phase, y_root):
        """Reorder a medication when its compact handle crosses a neighbour."""
        if phase == "start":
            row.configure(border_color=ACCENT, border_width=2)
            return
        if phase == "end":
            row.configure(border_color=LINE, border_width=1)
            return
        index = self.rows.index(row)
        if index > 0:
            previous = self.rows[index - 1]
            midpoint = previous.winfo_rooty() + previous.winfo_height() / 2
            if y_root < midpoint:
                self.move_row(row, -1)
                return
        if index + 1 < len(self.rows):
            following = self.rows[index + 1]
            midpoint = following.winfo_rooty() + following.winfo_height() / 2
            if y_root > midpoint:
                self.move_row(row, 1)

    def remove_row(self, row):
        if len(self.rows) <= 1:
            messagebox.showinfo("Info", I.t("msg_at_least_one"))
            return
        row.destroy()
        self.rows.remove(row)
        self._number_drug_rows()
        self.on_any_change()

    def clear_all(self):
        for v in self.doctor_vars.values():
            v.set("")
        self.clear_patient_details()
        for r in list(self.rows):
            r.destroy()
        self.rows = []
        self.add_row()
        self.on_any_change()

    # -- encrypted patient history -----------------------------------------
    def _schedule_patient_history_refresh(self, delay=180):
        if self._patient_search_job is not None:
            try:
                self.after_cancel(self._patient_search_job)
            except (tk.TclError, ValueError):
                pass
        self._patient_search_job = self.after(delay, self.refresh_patient_history)

    def refresh_patient_history(self):
        if not hasattr(self, "patient_history_list"):
            return
        self._patient_search_job = None
        query = self.patient_search_var.get()
        token = self._latest_query_tokens.get("patient_history", 0) + 1
        self._latest_query_tokens["patient_history"] = token
        self.submit_background(
            lambda: self.patient_history.search(query),
            lambda records: self._render_patient_history(query, token, records),
            silent=True)

    def _render_patient_history(self, query, token, records):
        if (token != self._latest_query_tokens.get("patient_history")
                or self.patient_search_var.get() != query):
            return
        self._patient_history_records = records
        self.patient_history_list.delete(0, tk.END)
        for record in self._patient_history_records:
            prescriptions = record.get("prescriptions", [])
            last_saved = max((str(item.get("saved_at", ""))[:10] for item in prescriptions), default="")
            detail = I.t("last_prescription", date=last_saved or I.t("none_short"))
            display = f"{record.get('name', '')} — {detail}"
            self.patient_history_list.insert(tk.END, directional_display_text(display))
        self._select_patient_record(self._loaded_patient_id)

    def _select_patient_record(self, record_id):
        if not record_id or not hasattr(self, "_patient_history_records"):
            return
        for index, record in enumerate(self._patient_history_records):
            if record.get("id") == record_id:
                self.patient_history_list.selection_clear(0, tk.END)
                self.patient_history_list.selection_set(index)
                self.patient_history_list.see(index)
                break

    def _set_patient_sex(self, label):
        self.patient_vars["sex"].set(self._patient_sex_codes.get(label, ""))

    def _patient_form_snapshot(self):
        return tuple(
            strip_bidi_display_controls(self.patient_vars[key].get()).strip()
            for key in ("name", "age", "sex"))

    def _set_patient_status(self, state):
        if not hasattr(self, "patient_status_label"):
            return
        styles = {
            "new": (I.t("patient_status_new"), ACCENT_SOFT, ACCENT),
            "saved": (I.t("patient_status_saved"), "#e5f6ef", "#176b51"),
            "modified": (I.t("patient_status_modified"), WARNING_SOFT, WARNING),
        }
        text, background, foreground = styles.get(state, styles["new"])
        self._patient_status_state = state
        self.patient_status_label.configure(
            text=text, fg_color=background, text_color=foreground)

    def _on_patient_form_change(self, *_):
        if self._patient_status_suspend:
            return
        if getattr(self, "_loaded_patient_id", ""):
            state = ("saved" if self._patient_form_snapshot()
                     == self._patient_saved_snapshot else "modified")
        else:
            state = "new"
        self._set_patient_status(state)

    def _set_patient_form(self, record):
        self._patient_status_suspend = True
        try:
            for key, variable in self.patient_vars.items():
                variable.set(str(record.get(key, "")))
            sex_code = str(record.get("sex", ""))
            sex_label = next(
                (label for label, code in self._patient_sex_codes.items() if code == sex_code),
                I.t("sex_m"))
            self.patient_sex_menu.set(sex_label)
            self._loaded_patient_id = str(record.get("id", ""))
            self._patient_saved_snapshot = self._patient_form_snapshot()
            self.patient_delete_button.configure(state="normal")
            self._set_patient_status("saved")
        finally:
            self._patient_status_suspend = False

    def _selected_or_loaded_patient(self):
        selected = self.patient_history_list.curselection()
        if selected and selected[0] < len(self._patient_history_records):
            return self._patient_history_records[selected[0]]
        if self._loaded_patient_id:
            return self.patient_history.get(self._loaded_patient_id)
        return None

    def _warn_similar_patient(self):
        if self._loaded_patient_id:
            return True
        name = self.patient_vars["name"].get().strip()
        age = self.patient_vars["age"].get().strip()
        sex = self.patient_vars["sex"].get().strip()
        matches = self.patient_history.find_similar(name, age, sex=sex)
        if not matches:
            return True
        match = matches[0]
        load_existing = messagebox.askyesno(
            I.t("similar_patient_title"),
            I.t("similar_patient_message", name=match.get("name", ""),
                age=match.get("age", "") or I.t("none_short")),
            parent=self)
        if load_existing:
            self._load_patient_record(match)
            return False
        return True

    def _show_empty_prescriptions(self):
        if not hasattr(self, "patient_prescriptions_body"):
            return
        for child in self.patient_prescriptions_body.winfo_children():
            child.destroy()
        ctk.CTkLabel(self.patient_prescriptions_body, text=I.t("previous_prescriptions_empty"),
                     text_color=MUTED, anchor="w", font=ctk.CTkFont(size=10)).pack(fill="x", pady=(2, 4))

    @staticmethod
    def _history_drug_name(drug):
        return str(drug.get("brand_name", "") or drug.get("generic_name", "") or "—")

    def refresh_prescription_comparison(self):
        """Show a live comparison with the selected patient's latest saved Rx."""
        if not hasattr(self, "patient_comparison_body"):
            return
        for child in self.patient_comparison_body.winfo_children():
            child.destroy()
        loaded_id = str(getattr(self, "_loaded_patient_id", ""))
        cached = getattr(self, "_current_history_record", None)
        visible_id = str((cached or {}).get("id", ""))
        # Never compare the current medicine workspace with another patient's
        # history. Browsing a different patient hides this panel until that
        # patient is explicitly loaded and owns the current medicine set.
        if (not loaded_id or (visible_id and visible_id != loaded_id)
                or self._medication_patient_id != loaded_id):
            self.patient_comparison_body.pack_forget()
            return
        record = self.patient_history.get(loaded_id)
        current = [row.get_data().__dict__ for row in self.rows
                   if row.get_data().generic_name or row.get_data().brand_name]
        prescriptions = list((record or {}).get("prescriptions", []))
        if not record or not prescriptions or not current:
            self.patient_comparison_body.pack_forget()
            return
        if not self.patient_comparison_body.winfo_manager():
            self.patient_comparison_body.pack(
                fill="x", pady=(0, 3), before=self.patient_prescriptions_body)
        previous = max(prescriptions, key=lambda item: str(item.get("saved_at", "")))
        result = compare_prescriptions(current, previous.get("drugs", []))
        ctk.CTkLabel(self.patient_comparison_body, text=I.t("comparison_title"),
                     text_color=ACCENT, anchor="w",
                     font=ctk.CTkFont(size=11, weight="bold")).pack(
                         fill="x", padx=9, pady=(7, 3))
        styles = {
            "added": (I.t("comparison_added"), "#e5f6ef", "#176b51"),
            "removed": (I.t("comparison_removed"), DANGER_SOFT, DANGER),
            "changed": (I.t("comparison_changed"), WARNING_SOFT, WARNING),
        }
        shown = False
        for kind in ("added", "removed", "changed"):
            label, bg, fg = styles[kind]
            for item in result[kind]:
                drug = item.get("after", {}) if kind == "changed" else item
                suffix = (" · " + ", ".join(item.get("fields", []))) if kind == "changed" else ""
                ctk.CTkLabel(
                    self.patient_comparison_body,
                    text=f"{label}: {self._history_drug_name(drug)}{suffix}",
                    fg_color=bg, text_color=fg, corner_radius=7, anchor="w",
                    font=ctk.CTkFont(size=10, weight="bold")).pack(
                        fill="x", padx=8, pady=2)
                shown = True
        if not shown:
            ctk.CTkLabel(self.patient_comparison_body, text=I.t("comparison_unchanged"),
                         text_color=ACCENT, anchor="w",
                         font=ctk.CTkFont(size=10)).pack(fill="x", padx=9, pady=(0, 7))

    def select_patient_history(self, event=None):
        selected = self.patient_history_list.curselection()
        if not selected:
            return
        record = self._patient_history_records[selected[0]]
        self._current_history_record = record
        self.patient_delete_button.configure(state="normal")
        self.show_patient_prescriptions(record)
        self.refresh_prescription_comparison()

    def show_patient_prescriptions(self, record):
        prescriptions = list(record.get("prescriptions", []))
        if not prescriptions:
            self._show_empty_prescriptions()
            return
        for child in self.patient_prescriptions_body.winfo_children():
            child.destroy()
        prescriptions.sort(key=lambda item: str(item.get("saved_at", "")), reverse=True)
        for prescription in prescriptions:
            saved_at = str(prescription.get("saved_at", ""))[:10]
            drugs = prescription.get("drugs", [])
            card = GlassFrame(self.patient_prescriptions_body, fg_color=SURFACE,
                                border_color=LINE, border_width=1, corner_radius=9)
            card.pack(fill="x", pady=(0, 3))
            prescription_id = str(prescription.get("id", ""))
            expanded = prescription_id in self._expanded_prescription_ids
            arrow = "▾" if expanded else "▸"
            VisualButton(
                card,
                text=(f"{arrow}  {I.t('last_prescription', date=saved_at or '—')}"
                      f"  ·  {I.t('medicine_count', n=len(drugs))}"),
                height=30, anchor="w", fg_color="transparent", text_color=ACCENT,
                hover_color=ACCENT_SOFT, font=ctk.CTkFont(size=12, weight="bold"),
                command=lambda item_id=prescription_id, patient=record:
                    self.toggle_patient_prescription(patient, item_id)).pack(
                        fill="x", padx=4, pady=1)
            if not expanded:
                continue
            details = ctk.CTkFrame(card, fg_color="transparent")
            details.pack(fill="x", padx=9, pady=(0, 4))
            for number, drug in enumerate(drugs, 1):
                name = str(drug.get("brand_name", "") or drug.get("generic_name", ""))
                scientific = str(drug.get("generic_name", ""))
                if drug.get("brand_name") and scientific:
                    name = f"{name} ({scientific})"
                regimen = "  ·  ".join(
                    str(drug.get(key, "")).strip()
                    for key in ("dosage", "frequency", "duration", "notes", "quantity")
                    if str(drug.get(key, "")).strip())
                line = f"{number}. {name}" + (f" — {regimen}" if regimen else "")
                ctk.CTkLabel(
                    details, text=directional_display_text(line), text_color=TEXT,
                    font=ctk.CTkFont(size=11), anchor="w", justify="left",
                    wraplength=720).pack(fill="x", pady=2)
            actions = ctk.CTkFrame(details, fg_color="transparent")
            actions.pack(fill="x", pady=(5, 0))
            VisualButton(
                actions, text=I.t("load_rx"), width=100, height=32,
                fg_color=CARD, text_color=ACCENT, border_width=1,
                border_color=LINE, hover_color=ACCENT_SOFT,
                command=lambda item=prescription, patient=record:
                    self.load_saved_prescription(patient, item)).pack(side="left")
            VisualButton(
                actions, text=I.t("delete_rx"), width=90, height=32,
                fg_color="transparent", text_color=DANGER, hover_color=DANGER_SOFT,
                command=lambda item=prescription, patient=record:
                    self.delete_saved_prescription(patient, item)).pack(side="right")

    def toggle_patient_prescription(self, record, prescription_id):
        if prescription_id in self._expanded_prescription_ids:
            self._expanded_prescription_ids.remove(prescription_id)
        else:
            self._expanded_prescription_ids.add(prescription_id)
        self.show_patient_prescriptions(record)

    def save_patient_history(self):
        if not self._warn_similar_patient():
            return
        if (self._loaded_patient_id
                and self._patient_form_snapshot() != self._patient_saved_snapshot
                and not messagebox.askyesno(
                    I.t("patient_overwrite_title"),
                    I.t("patient_overwrite_message"), parent=self)):
            return
        try:
            record = self.patient_history.save_patient(
                {key: variable.get() for key, variable in self.patient_vars.items()},
                self._loaded_patient_id)
        except ValueError as exc:
            messagebox.showinfo(APP_TITLE, str(exc))
            return
        self._loaded_patient_id = str(record.get("id", ""))
        self._current_history_record = record
        self._patient_saved_snapshot = self._patient_form_snapshot()
        self._set_patient_status("saved")
        self.patient_delete_button.configure(state="normal")
        self.patient_search_var.set("")
        self.refresh_patient_history()

    def load_selected_patient(self, event=None):
        selected = self.patient_history_list.curselection()
        if not selected:
            return
        self._load_patient_record(self._patient_history_records[selected[0]])
        return "break" if event else None

    def _load_patient_record(self, record):
        self._set_patient_form(record)
        self._current_history_record = record
        self._select_patient_record(record.get("id", ""))
        self.show_patient_prescriptions(record)
        self.refresh_prescription_comparison()

    def delete_saved_prescription(self, patient, prescription):
        if not messagebox.askyesno(I.t("delete_rx"), I.t("delete_rx_confirm"), parent=self):
            return
        deleted = self.patient_history.delete_prescription(
            str(patient.get("id", "")), str(prescription.get("id", "")))
        if not deleted:
            return
        cfg.config.add_recovery_item(
            "prescription", self._history_drug_name((deleted.get("drugs") or [{}])[0]),
            {"patient_id": patient.get("id", ""), "prescription": deleted})
        record = self.patient_history.get(str(patient.get("id", ""))) or patient
        self._current_history_record = record
        self.show_patient_prescriptions(record)
        self.refresh_prescription_comparison()

    def delete_selected_patient(self):
        record = self._selected_or_loaded_patient()
        if not record:
            return
        if not messagebox.askyesno(I.t("delete_patient"), I.t("delete_patient_confirm", name=record.get("name", ""))):
            return
        if self.patient_history.delete(record.get("id", "")):
            cfg.config.add_recovery_item("patient", record.get("name", ""), record)
            if record.get("id") == self._loaded_patient_id:
                self.clear_patient_details()
            self.refresh_patient_history()
            self._show_empty_prescriptions()

    def new_patient(self):
        self.clear_patient_details()
        self.patient_search_var.set("")
        self.patient_name_entry.focus_set()

    def clear_patient_details(self):
        self._patient_status_suspend = True
        try:
            for variable in self.patient_vars.values():
                variable.set("")
            self.patient_sex_menu.set(I.t("sex_m"))
            self._loaded_patient_id = ""
            self._medication_patient_id = ""
            self._current_history_record = None
            self._patient_saved_snapshot = self._patient_form_snapshot()
            self.patient_history_list.selection_clear(0, tk.END)
            self.patient_delete_button.configure(state="disabled")
            self._set_patient_status("new")
        finally:
            self._patient_status_suspend = False
        self._show_empty_prescriptions()
        self.refresh_prescription_comparison()

    def clear_patient_search(self, event=None):
        self.patient_search_var.set("")
        return "break" if event else None

    def load_saved_prescription(self, patient, prescription):
        if self.rows and any(
                row.get_data().generic_name or row.get_data().brand_name
                for row in self.rows):
            if not messagebox.askyesno(I.t("load_previous_prescription"), I.t("replace_current_medicines")):
                return
        self._set_patient_form(patient)
        for row in list(self.rows):
            row.destroy()
        self.rows = []
        for saved_drug in prescription.get("drugs", []):
            self.add_row(data=qu.DrugItem(**{
                key: saved_drug.get(key, "") for key in
                ("generic_name", "brand_name", "dosage", "frequency",
                 "duration", "notes", "quantity")
            }))
        if not self.rows:
            self.add_row()
        self._medication_patient_id = str(patient.get("id", ""))
        self.refresh_prescription_comparison()
        self.show_page("medications")

    def save_prescription_for_patient(self):
        if not self._warn_similar_patient():
            return
        patient = {key: variable.get() for key, variable in self.patient_vars.items()}
        drugs = [row.get_data() for row in self.rows
                 if row.get_data().generic_name or row.get_data().brand_name]
        try:
            record = self.patient_history.save_prescription(
                patient, [drug.__dict__ for drug in drugs], self._loaded_patient_id)
        except ValueError as exc:
            messagebox.showinfo(I.t("save_patient_prescription"), str(exc))
            return
        self._loaded_patient_id = str(record.get("id", ""))
        self._medication_patient_id = self._loaded_patient_id
        self._current_history_record = record
        self._patient_saved_snapshot = self._patient_form_snapshot()
        self._set_patient_status("saved")
        self.patient_search_var.set("")
        self.refresh_patient_history()
        self.show_patient_prescriptions(record)
        messagebox.showinfo(I.t("save_patient_prescription"), I.t("patient_prescription_saved"))

    # -- data ----------------------------------------------------------------
    def collect(self):
        doctor = qu.Doctor(**{k: v.get().strip() for k, v in self.doctor_vars.items()})
        patient = qu.Patient(**{k: v.get().strip() for k, v in self.patient_vars.items()})
        drugs = [r.get_data() for r in self.rows
                 if r.get_data().generic_name or r.get_data().brand_name]
        rx_id = "RX-" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        clinic = qu.Clinic(**{k: v for k, v in cfg.config.get_clinic().items()
                              if k in {"name", "address", "phone", "logo_path"}})
        return qu.Prescription(clinic=clinic, doctor=doctor, patient=patient, drugs=drugs,
                               date=datetime.datetime.now().strftime("%Y-%m-%d"), rx_id=rx_id)

    def on_any_change(self, *a):
        # Text-entry widgets can fire several events for one visible edit.
        # Coalesce them so the Word preview is rebuilt once after typing pauses.
        if self._word_preview_job is not None:
            try:
                self.after_cancel(self._word_preview_job)
            except (tk.TclError, ValueError):
                pass
        self._word_preview_job = self.after(180, self._render_word_preview_now)
        if self._loaded_patient_id:
            self._medication_patient_id = self._loaded_patient_id
            if self._comparison_job is not None:
                try:
                    self.after_cancel(self._comparison_job)
                except (tk.TclError, ValueError):
                    pass
            self._comparison_job = self.after(250, self._render_prescription_comparison_now)

    def _render_prescription_comparison_now(self):
        self._comparison_job = None
        self.refresh_prescription_comparison()

    def _render_word_preview_now(self):
        self._word_preview_job = None
        if self.word_preview_visible:
            self.render_word_preview()

    def toggle_word_preview(self):
        if self.word_preview_visible:
            self.close_medication_subpage()
            return
        self.open_medication_subpage("preview")
        self.render_word_preview()

    def open_medication_subpage(self, kind):
        """Show secondary medication tools without expanding the entry form."""
        self._hide_quick_results()
        for row in self.rows:
            row._hide_ac()
        self.medication_main.pack_forget()
        self.medication_favorite_panel.pack_forget()
        self.word_preview_section.pack_forget()
        self.word_preview_visible = kind == "preview"
        self.medication_subpage_title.configure(
            text=I.t("word_preview" if self.word_preview_visible else "starred_drugs"))
        self.medication_subpage.pack(fill="both", expand=True)
        panel = self.word_preview_section if self.word_preview_visible else self.medication_favorite_panel
        panel.pack(fill="x", padx=2, pady=CARD_GAP)
        if self.word_preview_visible:
            self.word_preview_body.pack(fill="x", padx=PAD, pady=PAD)
        self.scroll._parent_canvas.yview_moveto(0)

    def close_medication_subpage(self):
        self.word_preview_visible = False
        self.medication_subpage.pack_forget()
        self.medication_favorite_panel.pack_forget()
        self.word_preview_section.pack_forget()
        self.medication_main.pack(fill="both", expand=True)
        self.scroll._parent_canvas.yview_moveto(0)

    # -- openFDA online drug reference -------------------------------------
    def _clear_reference_cards(self):
        for child in self.reference_cards.winfo_children():
            child.destroy()

    def _show_reference_placeholder(self):
        if not hasattr(self, "reference_cards"):
            return
        self._clear_reference_cards()
        ctk.CTkLabel(self.reference_cards, text=I.t("openfda_reference_note"),
                     text_color=MUTED, justify="left", wraplength=680, anchor="w").pack(fill="x")

    def lookup_openfda_labels(self):
        medicines = [row.get_data().generic_name.strip() for row in self.rows
                     if row.get_data().generic_name.strip()]
        if not medicines:
            messagebox.showinfo(I.t("online_drug_reference"), I.t("openfda_no_drugs"))
            return
        self._start_openfda_lookup(medicines)

    def lookup_openfda_search(self, _event=None):
        medicine = self.reference_search_var.get().strip()
        if not medicine:
            self.reference_search_entry.focus_set()
            return "break" if _event else None
        self._start_openfda_lookup([medicine])
        return "break" if _event else None

    def _set_reference_lookup_state(self, busy):
        state = "disabled" if busy else "normal"
        self.reference_lookup_button.configure(state=state)
        self.reference_search_button.configure(state=state)
        self.reference_refresh_button.configure(state=state)

    def _start_openfda_lookup(self, medicines, force=False):
        medicines = list(dict.fromkeys(" ".join(name.split()) for name in medicines if name.strip()))
        if not medicines:
            return
        self._reference_last_medicines = medicines
        self._set_reference_lookup_state(True)
        self.reference_status.configure(text=I.t("openfda_searching"))
        self._clear_reference_cards()
        threading.Thread(target=self._lookup_openfda_worker,
                         args=(medicines, force), daemon=True).start()

    def refresh_openfda_labels(self):
        if self._reference_last_medicines:
            self._start_openfda_lookup(self._reference_last_medicines, force=True)
        else:
            self.lookup_openfda_search()

    def clear_openfda_cache(self):
        cfg.config.set("openfda_cache", {})
        self.reference_status.configure(text=I.t("reference_cache_cleared"))

    @staticmethod
    def _openfda_cache_key(medicine):
        return openfda.scientific_name_candidate(medicine).strip().casefold()

    def _lookup_openfda_with_cache(self, medicines, force=False):
        now = datetime.datetime.now(datetime.timezone.utc)
        cache = cfg.config.get("openfda_cache", {})
        cache = dict(cache) if isinstance(cache, dict) else {}
        references, missing, cached_names, checked = [], [], set(), {}
        for medicine in medicines:
            key = self._openfda_cache_key(medicine)
            item = cache.get(key, {}) if not force else {}
            try:
                timestamp = datetime.datetime.fromisoformat(str(item.get("checked_at", "")))
                if timestamp.tzinfo is None:
                    timestamp = timestamp.replace(tzinfo=datetime.timezone.utc)
                valid = now - timestamp <= datetime.timedelta(days=7)
                reference = openfda.reference_from_dict(item["reference"]) if valid else None
            except (KeyError, TypeError, ValueError):
                reference = None
            if reference is not None:
                references.append(reference)
                cached_names.add(reference.medicine.casefold())
                checked[reference.medicine.casefold()] = timestamp.isoformat(timespec="seconds")
            else:
                missing.append(medicine)
        fresh = openfda.lookup_labels(missing) if missing else []
        stamp = now.isoformat(timespec="seconds")
        fresh_keys = set()
        for reference in fresh:
            key = self._openfda_cache_key(reference.medicine)
            fresh_keys.add(key)
            cache[key] = {
                "checked_at": stamp, "reference": openfda.reference_to_dict(reference)}
            checked[reference.medicine.casefold()] = stamp
        for medicine in missing:
            key = self._openfda_cache_key(medicine)
            if key not in fresh_keys:
                cache.pop(key, None)
        if fresh or force:
            cfg.config.set("openfda_cache", cache)
        by_name = {reference.medicine.casefold(): reference for reference in references + fresh}
        ordered = [by_name[medicine.casefold()] for medicine in medicines
                   if medicine.casefold() in by_name]
        return ordered, cached_names, checked

    def _lookup_openfda_worker(self, medicines, force=False):
        try:
            references, cached_names, checked = self._lookup_openfda_with_cache(medicines, force)
            self.after(0, lambda: self._show_openfda_results(
                references, medicines, cached_names, checked))
        except openfda.OpenFDALookupError as exc:
            self.after(0, lambda detail=str(exc): self._show_openfda_error(detail))
        except Exception as exc:
            logging.exception("openFDA lookup failed")
            self.after(0, lambda detail=str(exc): self._show_openfda_error(detail))

    def _show_openfda_error(self, detail):
        self._set_reference_lookup_state(False)
        self.reference_status.configure(text=I.t("openfda_error"))
        self._clear_reference_cards()
        ctk.CTkLabel(self.reference_cards, text=I.t("openfda_error_detail", detail=detail),
                     text_color=DANGER, justify="left", wraplength=680, anchor="w").pack(fill="x")

    def _reference_line(self, parent, title, paragraphs, kind, source_sections=(), status=""):
        background, foreground = REFERENCE_TAGS[kind]
        line = ctk.CTkFrame(parent, fg_color=background, border_color=foreground,
                            border_width=1, corner_radius=10)
        ctk.CTkLabel(line, text=title, text_color=foreground,
                     font=_ui_font(13, "bold"), anchor="w").pack(
                         fill="x", padx=10, pady=(8, 3))
        if status:
            ctk.CTkLabel(line, text=status, height=22, corner_radius=11,
                         fg_color=CARD, text_color=foreground,
                         font=_ui_font(10, "bold"), anchor="w").pack(
                             anchor="w", padx=10, pady=(0, 4))
        value = "\n\n".join(paragraphs) or I.t("openfda_not_stated")
        preview = value if len(value) <= 280 else value[:277].rsplit(" ", 1)[0] + "…"
        content = ctk.CTkLabel(line, text=preview, text_color=TEXT,
                              font=_ui_font(12), justify="left", wraplength=300,
                              anchor="nw")
        content.pack(fill="x", padx=10, pady=(0, 8))
        line.bind("<Configure>", lambda event: content.configure(
            wraplength=max(100, event.width - 24)))
        if preview != value:
            expanded = False
            def toggle():
                nonlocal expanded
                expanded = not expanded
                content.configure(text=value if expanded else preview)
                expand_button.configure(text=I.t("reference_collapse" if expanded else "reference_expand"))
            expand_button = VisualButton(
                line, text=I.t("reference_expand"), height=26, width=120,
                fg_color="transparent", text_color=foreground,
                hover_color=background, font=_ui_font(11), command=toggle)
            expand_button.pack(anchor="w", padx=6, pady=(0, 5))
        if source_sections:
            ctk.CTkLabel(line, text=I.t("reference_section_source", sections=", ".join(source_sections)),
                         text_color=foreground, font=_ui_font(10), anchor="w",
                         wraplength=300).pack(fill="x", padx=10, pady=(0, 7))
        return line

    @staticmethod
    def _reference_sources(reference):
        return {key: tuple(values) for key, values in reference.field_sources}

    def _dose_reference_values(self, reference):
        parts = list(reference.dosage)
        details = (("reference_adult_dose", reference.adult_dose),
                   ("reference_maximum_dose", reference.maximum_dose),
                   ("reference_route", reference.route),
                   ("reference_hepatic_adjustment", reference.hepatic_adjustment))
        for label, values in details:
            if values:
                parts.append(I.t(label) + ": " + " ".join(values))
        return tuple(parts)

    @staticmethod
    def _format_label_date(value):
        digits = "".join(character for character in str(value) if character.isdigit())
        if len(digits) < 8:
            return "", "undated"
        try:
            date = datetime.datetime.strptime(digits[:8], "%Y%m%d").date()
        except ValueError:
            return "", "undated"
        age = (datetime.datetime.now(datetime.timezone.utc).date() - date).days
        return date.isoformat(), "old" if age > 365 * 5 else "current"

    def _layout_reference_sections(self, parent, panels, width):
        columns = 2 if width >= 720 else 1
        if getattr(parent, "_reference_columns", None) == columns:
            return
        parent._reference_columns = columns
        if not hasattr(parent, "_reference_stacks"):
            parent._reference_stacks = tuple(
                ctk.CTkFrame(parent, fg_color="transparent", width=1, height=1)
                for _ in range(2))
        parent.grid_columnconfigure(0, weight=1, uniform="label_sections")
        parent.grid_columnconfigure(1, weight=1 if columns == 2 else 0,
                                     uniform="label_sections" if columns == 2 else "")
        for index, stack in enumerate(parent._reference_stacks):
            if index < columns:
                stack.grid(row=0, column=index, sticky="new")
            else:
                stack.grid_remove()
        # Independent vertical stacks avoid row-height gaps beneath shorter cards.
        for index, panel in enumerate(panels):
            panel.pack_forget()
            panel.pack(in_=parent._reference_stacks[index % columns],
                       fill="x", padx=3, pady=3)
            # pack(in_=...) changes the geometry manager, not the Tk parent.
            # The later-created CTk stack canvases otherwise cover these siblings.
            panel.lift()

    def _show_openfda_results(self, references, medicines, cached_names=None, checked=None):
        self._set_reference_lookup_state(False)
        self.reference_status.configure(text=I.t("openfda_found", n=len(references)))
        self._clear_reference_cards()
        found = {reference.medicine.casefold() for reference in references}
        cached_names = cached_names or set()
        checked = checked or {}
        for number, reference in enumerate(references, 1):
            card = GlassFrame(self.reference_cards, fg_color=SURFACE, border_color=LINE,
                                border_width=1, corner_radius=14)
            card.pack(fill="x", pady=5)
            heading = I.t("openfda_card_title", number=number, name=reference.scientific_name)
            date_text, date_status = self._format_label_date(reference.effective_date)
            metadata = []
            if date_text:
                metadata.append(I.t("reference_label_date", date=date_text))
            else:
                metadata.append(I.t("reference_label_undated"))
            if date_status == "old":
                metadata.append(I.t("reference_label_old"))
            checked_at = checked.get(reference.medicine.casefold(), "")
            if reference.medicine.casefold() in cached_names:
                metadata.append(I.t("reference_cached", date=checked_at[:10]))
            elif checked_at:
                metadata.append(I.t("reference_checked", date=checked_at[:10]))
            header = ctk.CTkFrame(card, fg_color="transparent")
            header.pack(fill="x", padx=12, pady=(8, 5))
            VisualButton(header, text=I.t("reference_full_label"), height=28, width=130,
                          fg_color="transparent", text_color=ACCENT,
                          hover_color=ACCENT_SOFT,
                          command=lambda item=reference: self.open_full_drug_label(item)).pack(side="right", padx=(8, 0))
            header_text = "  ·  ".join((heading,
                I.t("openfda_label_name", name=reference.label_name), *metadata))
            identity = ctk.CTkLabel(header, text=header_text, text_color=TEXT,
                                    font=_ui_font(12), anchor="w", justify="left")
            identity.pack(side="left", fill="x", expand=True)
            # Keep all metadata readable if a long label cannot fit on one line.
            identity.bind("<Configure>", lambda event, label=identity:
                          label.configure(wraplength=max(100, event.width)))
            sources = self._reference_sources(reference)
            renal_status = I.t("reference_renal_" + reference.renal_status)
            sections = ctk.CTkFrame(card, fg_color="transparent")
            sections.pack(fill="x", padx=9, pady=(0, 8))
            panels = [self._reference_line(sections, I.t(title), values, kind,
                                            sources.get(source_key, ()), status)
                      for title, values, kind, source_key, status in (
                          ("label_indication", reference.indications, "indication", "indication", ""),
                          ("label_dose", self._dose_reference_values(reference), "dose", "dose", ""),
                          ("label_contraindications", reference.contraindications,
                           "contraindication", "contraindication", ""),
                          ("label_pregnancy", reference.pregnancy, "pregnancy", "pregnancy", ""),
                          ("label_renal_adjustment", reference.renal_adjustment,
                           "renal", "renal", renal_status))]
            self._layout_reference_sections(sections, panels, self.reference_cards.winfo_width())
            sections.bind("<Configure>", lambda event, parent=sections, items=panels:
                          self._layout_reference_sections(parent, items, event.width))
        missing = [medicine for medicine in medicines if medicine.casefold() not in found]
        if missing:
            ctk.CTkLabel(self.reference_cards, text=I.t("openfda_not_found", names=", ".join(missing)),
                         text_color=MUTED, justify="left", wraplength=680, anchor="w").pack(fill="x", pady=5)
        if not references:
            ctk.CTkLabel(self.reference_cards, text=I.t("openfda_no_matches"), text_color=MUTED,
                         justify="left", wraplength=680, anchor="w").pack(fill="x")

    def open_full_drug_label(self, reference):
        window = ctk.CTkToplevel(self)
        self.reference_full_label_window = window
        window.title(I.t("reference_full_label"))
        window.geometry("900x680")
        window.minsize(650, 460)
        window.transient(self)
        ctk.CTkLabel(window, text=reference.label_name, text_color=ACCENT,
                     font=_ui_font(22, "bold"), anchor="w").pack(
                         fill="x", padx=18, pady=(16, 7))
        search_var = tk.StringVar()
        self.reference_full_label_search_var = search_var
        search = VisualEntry(window, textvariable=search_var,
                              placeholder_text=I.t("reference_search_full_label"),
                              height=FIELD_HEIGHT, border_color=LINE)
        search.pack(fill="x", padx=18, pady=(0, 8))
        text = ctk.CTkTextbox(window, fg_color=CARD, border_color=LINE, border_width=1,
                              text_color=TEXT, font=_ui_font(12), wrap="word")
        self.reference_full_label_text = text
        text.pack(fill="both", expand=True, padx=18, pady=(0, 12))
        sections = reference.full_sections or (
            (I.t("label_indication"), reference.indications),
            (I.t("label_dose"), reference.dosage),
            (I.t("label_contraindications"), reference.contraindications),
            (I.t("label_pregnancy"), reference.pregnancy),
            (I.t("label_renal_adjustment"), reference.renal_adjustment))
        # Filter at display time so previously cached labels follow the same policy.
        sections = tuple((heading, values) for heading, values in sections
                         if heading.replace("_", " ").strip().casefold()
                         not in {"pediatric use", "how supplied", "warnings and cautions",
                                 "warning and caution", "use in specific populations",
                                 "clinical studies", "clinical pharmacology"})

        def render(*_args):
            query = search_var.get().strip().casefold()
            text.configure(state="normal")
            text.delete("1.0", "end")
            for heading, values in sections:
                content = "\n\n".join(values)
                if query and query not in (heading + " " + content).casefold():
                    continue
                text.insert("end", heading.upper() + "\n", "heading")
                text.insert("end", (content or I.t("openfda_not_stated")) + "\n\n")
            widget = text._textbox
            widget.tag_configure("heading", foreground=ACCENT, font=("Segoe UI", 24, "bold"))
            widget.tag_remove("match", "1.0", "end")
            widget.tag_configure("match", background=WARNING_SOFT, foreground=TEXT)
            if query:
                position = "1.0"
                while True:
                    position = widget.search(query, position, stopindex="end", nocase=True)
                    if not position:
                        break
                    end = f"{position}+{len(query)}c"
                    widget.tag_add("match", position, end)
                    position = end
            text.configure(state="disabled")
        search_var.trace_add("write", render)
        render()
        search.focus_set()

    # -- Medscape handoff (manual clinical review, no scraping) ------------
    def check_interactions_in_medscape(self):
        medicines = [row.get_data().generic_name.strip() for row in self.rows
                     if row.get_data().generic_name.strip()]
        if not medicines:
            messagebox.showinfo(I.t("interaction_review"), I.t("openfda_no_drugs"))
            return
        names = []
        for medicine in medicines:
            normalized = openfda.scientific_name_candidate(medicine)
            if normalized and normalized.casefold() not in {name.casefold() for name in names}:
                names.append(normalized)
        self.clipboard_clear()
        self.clipboard_append("\n".join(names))
        self.update()
        messagebox.showinfo(I.t("interaction_review"), I.t("medscape_copied", names=", ".join(names)))
        webbrowser.open("https://reference.medscape.com/drug-interactionchecker")

    def render_word_preview(self):
        if not hasattr(self, "word_preview_body"):
            return
        for child in self.word_preview_body.winfo_children():
            child.destroy()
        drugs = [row.get_data() for row in self.rows
                 if row.get_data().generic_name or row.get_data().brand_name]
        if not drugs:
            ctk.CTkLabel(self.word_preview_body, text=I.t("word_preview_empty"),
                         text_color=MUTED, anchor="w").pack(fill="x")
            return
        for row_number, drug in enumerate(drugs, 1):
            if drug.generic_name and drug.brand_name:
                name = f"{drug.brand_name} ({drug.generic_name})"
            else:
                name = drug.generic_name or drug.brand_name
            parts = [name] + [value for value in (
                drug.dosage, drug.frequency, drug.duration, drug.notes) if value]
            line = "     ".join(parts)
            line_row = ctk.CTkFrame(self.word_preview_body, fg_color=CARD)
            line_row.pack(fill="x", pady=2)
            ctk.CTkLabel(
                line_row, text=f"{row_number}.", width=34, text_color=ACCENT,
                font=ctk.CTkFont(size=12, weight="bold"), anchor="w").pack(
                    side="left", padx=(0, 5))
            ctk.CTkLabel(
                line_row, text=line, text_color=TEXT,
                font=ctk.CTkFont(size=12), anchor="w", justify="left",
                wraplength=900).pack(side="left", fill="x", expand=True)

    def on_paper_change(self, value):
        cfg.config.paper_size = value

    def on_lang_change(self, value):
        cfg.config.language = value
        I.set_lang(value)
        self.reload_texts()

    def create_backup(self):
        initial_dir = cfg.config.get("document_defaults", {}).get("export_folder", "")
        path = filedialog.asksaveasfilename(
            title=I.t("backup"), defaultextension=".rxbackup",
            initialdir=initial_dir if Path(initial_dir).is_dir() else None,
            filetypes=[("Prescription backup", "*.rxbackup")])
        if not path:
            return
        try:
            cfg.config.create_backup(path)
        except Exception as exc:
            messagebox.showerror(I.t("backup"), str(exc))
            return
        messagebox.showinfo(I.t("backup"), I.t("msg_backup_created", path=path))

    def restore_backup(self):
        path = filedialog.askopenfilename(
            title=I.t("restore"), filetypes=[("Prescription backup", "*.rxbackup")])
        if not path:
            return
        if not messagebox.askyesno(I.t("restore"), "Restore settings and replace the current drug database?"):
            return
        try:
            cfg.config.restore_backup(path)
        except Exception as exc:
            messagebox.showerror(I.t("restore"), str(exc))
            return
        I.set_lang(cfg.config.language)
        self.db = dbmod.DrugDatabase(cfg.config.drug_db_path)
        for child in list(self.children.values()):
            child.destroy()
        self.rows = []
        self._build_ui()
        messagebox.showinfo(I.t("restore"), I.t("msg_backup_restored"))

    def reload_texts(self):
        for child in list(self.children.values()):
            child.destroy()
        self.rows = []
        self._build_ui()

    # -- profile -------------------------------------------------------------
    def save_profile(self):
        cfg.config.save_profile(self.profile_var.get(),
                                {k: v.get().strip() for k, v in self.doctor_vars.items()})
        messagebox.showinfo(I.t("app_title"), I.t("msg_saved_profile"))

    def load_profile_into_ui(self):
        for k, v in cfg.config.get_doctor().items():
            if k in self.doctor_vars:
                self.doctor_vars[k].set(v)
        self._refresh_specialty_values(self.doctor_vars["specialty"].get())

    # -- DB ops --------------------------------------------------------------
    @staticmethod
    def _drug_import_identity(drug):
        return "|".join((drug.generic_name.strip().casefold(),
                         drug.brand_name.strip().casefold()))

    def _classification_snapshot(self):
        return {
            self._drug_import_identity(drug): frozenset(
                (mapping.code, mapping.detail) for mapping in classes.groups_for(drug))
            for drug in self.db.drugs
        }

    def _import_classification_report(self, before):
        after = self._classification_snapshot()
        identities = {self._drug_import_identity(drug): drug for drug in self.db.drugs}
        new_keys = set(after) - set(before)
        confirmed = [
            drug for key, drug in identities.items()
            if after.get(key) and drug.mapping_status != "suggested"
        ]
        suggested = [
            drug for key, drug in identities.items()
            if after.get(key) and drug.mapping_status == "suggested"
        ]
        unclassified = [drug for key, drug in identities.items() if not after.get(key)]
        changed_keys = {
            key for key in set(before).intersection(after)
            if before.get(key) != after.get(key)
        }
        missing_keys = {key for key in set(before) - set(after) if before.get(key)}
        return {
            "new": [identities[key] for key in new_keys if key in identities],
            "recognized": confirmed + suggested,
            "confirmed": confirmed,
            "suggested": suggested,
            "unclassified": unclassified,
            "changed_missing": list(changed_keys | missing_keys),
            "import_stats": dict(self.db.last_import_report),
        }

    def show_import_classification_assistant(self, report):
        window = ctk.CTkToplevel(self)
        window.title(I.t("import_classification_assistant"))
        window.geometry("760x500")
        window.minsize(680, 450)
        window.transient(self)
        window.grab_set()
        ctk.CTkLabel(
            window, text=I.t("import_classification_assistant"),
            text_color=ACCENT, anchor="w",
            font=ctk.CTkFont(size=24, weight="bold")).pack(
                fill="x", padx=22, pady=(20, 4))
        stats = report.get("import_stats", {})
        ctk.CTkLabel(
            window, text=I.t(
                "import_classification_summary",
                source=stats.get("source_rows", len(self.db.drugs)),
                imported=stats.get("imported_rows", len(self.db.drugs)),
                brand_only=stats.get("brand_only_rows", 0),
                skipped=stats.get("skipped_blank_names", 0),
                unrecognized=stats.get("unrecognized_class_rows", 0)),
            text_color=MUTED, anchor="w", justify="left",
            wraplength=700, font=ctk.CTkFont(size=12)).pack(
                fill="x", padx=22, pady=(0, 15))
        grid = ctk.CTkFrame(window, fg_color="transparent")
        grid.pack(fill="both", expand=True, padx=22)
        grid.grid_columnconfigure((0, 1), weight=1, uniform="import-report")
        items = (
            ("confirmed", "import_confirmed_mappings", ACCENT_SOFT, ACCENT),
            ("suggested", "import_suggested_mappings", WARNING_SOFT, WARNING),
            ("unclassified", "import_unclassified_medicines", WARNING_SOFT, WARNING),
            ("changed_missing", "import_changed_missing", BG, MUTED),
        )
        for index, (key, label_key, background, foreground) in enumerate(items):
            values = report.get(key, [])
            card = ctk.CTkFrame(
                grid, fg_color=background, border_color=LINE,
                border_width=1, corner_radius=12)
            card.grid(row=index // 2, column=index % 2, sticky="nsew", padx=5, pady=5)
            ctk.CTkLabel(
                card, text=str(len(values)), text_color=foreground,
                font=ctk.CTkFont(size=26, weight="bold"), anchor="w").pack(
                    fill="x", padx=14, pady=(12, 0))
            ctk.CTkLabel(
                card, text=I.t(label_key), text_color=TEXT,
                font=ctk.CTkFont(size=12, weight="bold"), anchor="w").pack(
                    fill="x", padx=14, pady=(0, 4))
            if values and key != "changed_missing":
                names = ", ".join(
                    (drug.brand_name.strip() or drug.generic_name.strip())
                    for drug in values[:3])
                ctk.CTkLabel(
                    card, text=names, text_color=MUTED, anchor="w",
                    font=ctk.CTkFont(size=10), wraplength=270).pack(
                        fill="x", padx=14, pady=(0, 10))
        actions = ctk.CTkFrame(window, fg_color="transparent")
        actions.pack(fill="x", padx=22, pady=18)
        if report.get("unclassified"):
            VisualButton(
                actions, text=I.t("review_unclassified"), height=ACTION_HEIGHT,
                fg_color=PRIMARY, hover_color=ACCENT_HOVER,
                command=lambda: (window.destroy(),
                                 self.show_page("drug_classes"),
                                 self.show_class_mapping_editor(True))).pack(side="left")
        VisualButton(
            actions, text=I.t("manage_class_mappings"), height=ACTION_HEIGHT,
            fg_color=CARD, text_color=ACCENT, border_width=1,
            border_color=LINE, hover_color=ACCENT_SOFT,
            command=lambda: (window.destroy(), self.show_page("drug_classes"),
                             self.show_class_mapping_editor())).pack(side="left", padx=8)
        VisualButton(
            actions, text=I.t("close"), height=ACTION_HEIGHT, width=90,
            fg_color=ACCENT_SOFT, text_color=ACCENT, hover_color=LINE,
            command=window.destroy).pack(side="right")

    def import_db(self, on_done=None):
        path = filedialog.askopenfilename(
            title=I.t("import_db"),
            filetypes=[("Drug database", "*.csv *.xlsx *.xls"),
                       ("CSV files", "*.csv"),
                       ("Excel files", "*.xlsx"),
                       ("Legacy Excel files", "*.xls"),
                       ("All files", "*.*")])
        if not path:
            return
        if not messagebox.askyesno(I.t("replace_database_title"),
                                   I.t("replace_database_confirm")):
            return
        def worker():
            before = self._classification_snapshot()
            total = self.db.import_file(path, replace=True)
            with cfg.config.batch_save():
                cfg.config.drug_db_path = str(self.db.path)
                cfg.config.set("drug_db_imported_at", datetime.datetime.now(
                    datetime.timezone.utc).isoformat(timespec="seconds"))
            return total, self._import_classification_report(before)

        def finished(result):
            total, report = result
            self._invalidate_database_caches()
            self.db_label.configure(text=I.t("db_count", n=len(self.db.drugs)))
            self.refresh_class_overview()
            self.refresh_favorites_page()
            self.show_import_classification_assistant(report)
            if on_done:
                on_done()

        self.submit_background(
            worker, finished,
            on_error=lambda exc: messagebox.showerror(I.t("import_db"), str(exc), parent=self),
            label=I.t("import_db") + "…")

    def remove_db(self, on_done=None):
        if not messagebox.askyesno(I.t("remove_db"), I.t("remove_db_confirm")):
            return
        try:
            self.db.clear()
            self._invalidate_database_caches()
            self.db_label.configure(text=I.t("db_count", n=0))
            self.refresh_class_overview()
            self.refresh_favorites_page()
            messagebox.showinfo(I.t("remove_db"), I.t("msg_db_removed"))
            if on_done:
                on_done()
        except Exception as exc:
            logging.exception("Could not remove local drug database")
            messagebox.showerror(I.t("remove_db"), str(exc))

    def export_db(self, on_done=None):
        initial_dir = cfg.config.get("document_defaults", {}).get("export_folder", "")
        path = filedialog.asksaveasfilename(
            title=I.t("export_db"), defaultextension=".xlsx",
            initialdir=initial_dir if Path(initial_dir).is_dir() else None,
            filetypes=[("Excel workbook", "*.xlsx"),
                       ("Legacy Excel workbook", "*.xls"),
                       ("CSV files", "*.csv")])
        if not path:
            return
        def finished(_result):
            messagebox.showinfo(I.t("export_db"), I.t("msg_exported_db", path=path))
            if on_done:
                on_done()

        self.submit_background(
            lambda: self.db.export_file(path), finished,
            on_error=lambda exc: messagebox.showerror(I.t("export_db"), str(exc), parent=self),
            label=I.t("export_db") + "…")

    # -- output --------------------------------------------------------------
    def _qr(self, rx):
        """Deprecated synchronous entry point; interactive exports use workers."""
        raise RuntimeError("Use the asynchronous cloud export workflow")

    def _validate(self, rx, action):
        errors, warnings = qu.validate_prescription(rx)
        if errors:
            messagebox.showwarning(action, "Please correct the following:\n\n" + "\n".join(f"• {x}" for x in errors))
            return False
        if warnings and not messagebox.askyesno(action, "Warnings:\n\n" + "\n".join(
                f"• {x}" for x in warnings) + "\n\nContinue?"):
            return False
        return True

    def _build_full(self, path_pdf=None, path_docx=None):
        """Legacy callers now enter the same asynchronous cloud workflow."""
        return self._start_cloud_export(I.t("preview"), path_pdf=path_pdf, path_docx=path_docx)

    def _generate_full_document(self, rx, qr, document, path_pdf=None, path_docx=None):
        """Generate a prepared document; safe to execute on a worker thread."""
        with self._document_lock:
            document_language = document.get("language", "interface")
            with I.document_language(document_language):
                options = {
                    "show_header": bool(document.get("show_header", True)),
                    "logo_size": document.get("logo_size", "medium"),
                    "margin_mm_value": int(document.get("margin_mm", 16)),
                }
                if path_pdf:
                    pdfgen.generate_prescription_pdf(
                        rx, path_pdf, paper_size=document.get("_paper_size", "A4"),
                        qr_pil_image=qr,
                        **options)
                if path_docx:
                    pdfgen.generate_prescription_docx(
                        rx, path_docx, qr_pil_image=qr,
                        paper_size=document.get("_paper_size", "A4"), **options)

    def _prepare_full_document(self, action):
        rx = copy.deepcopy(self.collect())
        if not self._validate(rx, action):
            return None
        document = copy.deepcopy(cfg.config.get("document_defaults", {}))
        document["_paper_size"] = self.paper_var.get()
        if document.get("language", "interface") == "interface":
            document["language"] = I.get_lang()
        return rx, document

    def _set_export_busy(self, busy):
        self._export_busy = busy
        for widget in self.action.winfo_children():
            if isinstance(widget, VisualButton):
                widget.configure(state="disabled" if busy else "normal")

    def _start_cloud_export(self, action, *, path_pdf=None, path_docx=None,
                            compact=False, on_success=None):
        if self._export_busy or self._closing:
            return None
        prepared = self._prepare_full_document(action)
        if prepared is None:
            return None
        rx, document = prepared
        operation = {"rx": rx, "document": document, "payload": copy.deepcopy(rx.to_cloud_payload()),
                     "api_key": cfg.config.cloud_rx_api_key, "action": action,
                     "path_pdf": path_pdf, "path_docx": path_docx,
                     "compact": compact, "on_success": on_success}
        self._set_export_busy(True)
        self._upload_export_link(operation)
        return operation

    def _upload_export_link(self, operation):
        def ready(link):
            self._write_cloud_export(operation, link.url)

        self.submit_background(
            lambda: cloud_rx.upload_prescription(operation["payload"], operation["api_key"]),
            ready, on_error=lambda error: self._cloud_export_failed(operation, error),
            label=I.t("cloud_creating"))

    def _write_cloud_export(self, operation, url):
        def worker():
            qr = qu.make_qr_image(url) if url is not None else None
            if operation["compact"]:
                # Use the same document lock/language snapshot as the headed path.
                with self._document_lock:
                    with I.document_language(operation["document"]["language"]):
                        return pdfgen.generate_medication_label_docx(
                            operation["rx"], operation["path_docx"], qr_pil_image=qr,
                            paper_size=operation["document"]["_paper_size"])
            return self._generate_full_document(
                operation["rx"], qr, operation["document"],
                operation["path_pdf"], operation["path_docx"])

        def finished(result):
            self._set_export_busy(False)
            if operation["on_success"]:
                operation["on_success"](result)

        def failed(_error):
            self._set_export_busy(False)
            logging.error("Prescription document generation failed")
            messagebox.showerror(operation["action"], I.t("cloud_document_failed"), parent=self)

        self.submit_background(worker, finished, on_error=failed, label=operation["action"] + "…")

    def _cloud_export_failed(self, operation, error):
        code = error.code if isinstance(error, cloud_rx.CloudRxError) else "unexpected"
        logging.warning("Cloud prescription upload failed (%s)", code)
        dialog = ctk.CTkToplevel(self)
        dialog.title(I.t("cloud_failed"))
        dialog.geometry("600x240")
        dialog.transient(self)
        ctk.CTkLabel(dialog, text=I.t("cloud_error_" + code), font=_ui_font(13),
                     wraplength=550, justify="left").pack(fill="x", padx=20, pady=(24, 10))
        ctk.CTkLabel(dialog, text=I.t("cloud_failure_choices"), font=_ui_font(12),
                     wraplength=550, justify="left").pack(fill="x", padx=20, pady=(0, 16))
        actions = ctk.CTkFrame(dialog, fg_color="transparent")
        actions.pack(padx=12, pady=8)

        def choose(choice):
            dialog.destroy()
            self._resolve_cloud_export_failure(operation, choice)

        for label, choice in (("cloud_retry", "retry"), ("cloud_without_qr", "without"), ("cloud_cancel", "cancel")):
            VisualButton(actions, text=I.t(label), width=_ui_font(12).measure(I.t(label)) + 22,
                         height=32, font=_ui_font(12), fg_color=SURFACE, text_color=ACCENT,
                         border_width=1, border_color=LINE,
                         command=lambda value=choice: choose(value)).pack(side="left", padx=4)
        dialog.protocol("WM_DELETE_WINDOW", lambda: choose("cancel"))
        dialog.bind("<Escape>", lambda _event: choose("cancel"))
        dialog.grab_set()

    def _resolve_cloud_export_failure(self, operation, choice):
        if self._closing:
            return
        if choice == "retry":
            self._upload_export_link(operation)
        elif choice == "without":
            self._write_cloud_export(operation, None)
        else:
            self._set_export_busy(False)

    def preview(self):
        path = os.path.join(tempfile.gettempdir(), f"rx_preview_{uuid.uuid4().hex}.pdf")
        def finished(_result):
            try:
                os.startfile(path)
            except Exception:
                webbrowser.open(path)

        self._start_cloud_export(I.t("preview"), path_pdf=path, on_success=finished)

    def print_pdf(self):
        path = os.path.join(tempfile.gettempdir(), f"rx_print_{uuid.uuid4().hex}.pdf")
        def finished(_result):
            try:
                os.startfile(path, "print")
            except Exception:
                os.startfile(path)

        self._start_cloud_export(I.t("print"), path_pdf=path, on_success=finished)

    def export_word(self):
        if self._export_busy:
            return
        initial_dir = cfg.config.get("document_defaults", {}).get("export_folder", "")
        path = filedialog.asksaveasfilename(
            title=I.t("export_word"), defaultextension=".docx",
            initialdir=initial_dir if Path(initial_dir).is_dir() else None,
            filetypes=[("Word documents", "*.docx")])
        if not path:
            return
        def finished(_result):
            messagebox.showinfo(I.t("export_word"), I.t("msg_exported_word", path=path))
            try:
                os.startfile(path)
            except Exception:
                webbrowser.open(path)

        self._start_cloud_export(I.t("export_word"), path_docx=path, on_success=finished)

    def export_label(self):
        if self._export_busy:
            return
        initial_dir = cfg.config.get("document_defaults", {}).get("export_folder", "")
        path = filedialog.asksaveasfilename(
            title=I.t("export_compact"), defaultextension=".docx",
            initialdir=initial_dir if Path(initial_dir).is_dir() else None,
            filetypes=[("Word documents", "*.docx")])
        if not path:
            return
        def finished(_result):
            messagebox.showinfo(
                I.t("export_compact"), I.t("msg_exported_compact", path=path))
            try:
                os.startfile(path)
            except Exception:
                webbrowser.open(path)

        self._start_cloud_export(I.t("export_compact"), path_docx=path, compact=True, on_success=finished)

    def open_settings(self):
        SettingsWindow(self)

    def apply_ui_font_preferences(self):
        """Update existing controls without rebuilding pages or losing edits."""
        pending = [self]
        while pending:
            widget = pending.pop()
            if isinstance(widget, (VisualComboBox, VisualOptionMenu, PopupListbox, VisualMenu)):
                widget.apply_preferences()
            pending.extend(widget.winfo_children())
        patient_size = cfg.config.ui_font_size("patient_name_font_size")
        self.patient_name_entry.configure(font=_ui_font(patient_size), height=max(FIELD_HEIGHT, patient_size + 12))
        self._hide_quick_results()
        self._hide_treatment_template_suggestions()
        for row in self.rows:
            row._hide_ac()


class SettingsWindow(ctk.CTkToplevel):
    """Organized application settings with category navigation and one save flow."""

    _fixed_dropdown_fonts = True

    SECTIONS = (
        ("general", "⚒", "General"),
        ("clinic", "✚", "Clinic identity"),
        ("documents", "▤", "Documents"),
        ("qr", "▦", "QR verification"),
        ("database", "⌬", "Database"),
        ("gemini", "✦", "Gemini reference"),
        ("security", "▣", "Backup & security"),
        ("recovery", "↶", "Recovery centre"),
        ("about", "ⓘ", "About"),
    )

    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.title(I.t("set_title"))
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = min(920, max(780, screen_width - 200))
        height = min(620, max(540, screen_height - 220))
        x = max(20, (screen_width - width) // 2)
        y = max(20, (screen_height - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(780, 540)
        self.configure(fg_color=BG)
        self.transient(master)
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.bind("<Escape>", lambda _event: self.cancel())
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        clinic = cfg.config.get_clinic()
        self.cloud_key_var = tk.StringVar(value=cfg.config.cloud_rx_api_key)
        self.clinic_name_var = tk.StringVar(value=clinic.get("name", ""))
        self.clinic_address_var = tk.StringVar(value=clinic.get("address", ""))
        self.clinic_phone_var = tk.StringVar(value=clinic.get("phone", ""))
        self.logo_var = tk.StringVar(value=clinic.get("logo_path", ""))
        self.paper_var = tk.StringVar(value=cfg.config.paper_size)
        self.language_var = tk.StringVar(
            value="العربية" if cfg.config.language == "ar" else "English")
        dropdown_size = cfg.config.ui_font_size("dropdown_font_size")
        self.dropdown_font_var = tk.StringVar(value=str(dropdown_size) if dropdown_size else I.t("settings_font_default"))
        self.patient_font_var = tk.StringVar(value=str(cfg.config.ui_font_size("patient_name_font_size")))
        self.gemini_enabled_var = tk.BooleanVar(value=cfg.config.gemini_enabled)
        self.gemini_key_var = tk.StringVar(value=cfg.config.gemini_api_key)
        document = cfg.config.get("document_defaults", {})
        language_labels = {"interface": I.t("settings_same_as_interface"),
                           "en": "English", "ar": "العربية"}
        self.document_language_var = tk.StringVar(
            value=language_labels.get(document.get("language", "interface"),
                                      I.t("settings_same_as_interface")))
        self.document_header_var = tk.BooleanVar(
            value=bool(document.get("show_header", True)))
        self.export_folder_var = tk.StringVar(value=document.get("export_folder", ""))
        self.auto_backup_var = tk.BooleanVar(
            value=bool(cfg.config.get("auto_backup_enabled", False)))
        self.settings_search_var = tk.StringVar()
        self._active_section = "general"
        self._recovery_page_index = 0
        self._nav_buttons = {}
        self._nav_rows = {}
        self._nav_indicators = {}
        self._settings_nav_icons = {}
        self._pages = {}
        self._build_workspace()
        self._build_footer()
        self._build_pages()
        self._show_section("general")
        self._tracked_variables = (
            self.cloud_key_var, self.clinic_name_var, self.clinic_address_var,
            self.clinic_phone_var, self.logo_var, self.paper_var,
            self.language_var, self.gemini_enabled_var, self.gemini_key_var,
            self.document_language_var, self.document_header_var,
            self.export_folder_var, self.auto_backup_var,
            self.dropdown_font_var, self.patient_font_var,
        )
        self._saved_snapshot = self._snapshot()
        for variable in self._tracked_variables:
            variable.trace_add("write", self._mark_dirty)
        for variable in (self.clinic_name_var, self.clinic_address_var,
                         self.clinic_phone_var, self.logo_var,
                         self.document_header_var):
            variable.trace_add("write", lambda *_: self.update_clinic_preview())
        self.update_clinic_preview()
        self.after(100, self.grab_set)

    def _build_workspace(self):
        workspace = GlassFrame(self, glass_surface="workspace", fg_color=BG, corner_radius=0)
        workspace.grid(row=0, column=0, sticky="nsew")
        workspace.grid_rowconfigure(0, weight=1)
        workspace.grid_columnconfigure(1, weight=1)

        self.sidebar = GlassFrame(
            workspace, glass_surface="sidebar", width=218, fg_color=SIDEBAR_SURFACE,
            border_color=LINE, border_width=1, corner_radius=12)
        self.sidebar.grid(row=0, column=0, sticky="nsw", padx=(0, 1))
        self.sidebar.grid_propagate(False)
        self.settings_search_entry = VisualEntry(
            self.sidebar, textvariable=self.settings_search_var, height=FIELD_HEIGHT,
            placeholder_text=I.t("settings_search"), border_color=LINE,
            corner_radius=9, font=ctk.CTkFont(size=13))
        self.settings_search_entry.pack(fill="x", padx=12, pady=(10, 5))
        attach_search_hint(self.settings_search_entry, self.settings_search_var,
                           I.t("settings_search"))
        self.settings_search_var.trace_add(
            "write", lambda *_: self._filter_settings_sections())
        for key, _symbol, _label in self.SECTIONS:
            icons = (_dashboard_icon(key), _dashboard_icon(key, NAV_ACTIVE))
            self._settings_nav_icons[key] = icons
            row = ctk.CTkFrame(self.sidebar, fg_color="transparent", corner_radius=8)
            row.pack(fill="x", padx=8, pady=1)
            indicator = ctk.CTkFrame(row, width=3, height=22, corner_radius=1,
                                     fg_color=NAV_ACTIVE)
            button = VisualButton(
                row,
                text=I.t('settings_' + key) if key != 'security' else I.t('settings_backup_security'),
                image=icons[0], compound="left",
                anchor="w", height=38, corner_radius=8, fg_color="transparent",
                hover_color=ACCENT_SOFT, text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(size=16),
                command=lambda section=key: self._show_section(section))
            button._image_label_spacing = 10
            button.configure(anchor="w")
            button.pack(fill="x", padx=(5, 0))
            self._nav_buttons[key] = button
            self._nav_rows[key] = row
            self._nav_indicators[key] = indicator
        ctk.CTkLabel(
            self.sidebar, text=I.t("local_encrypted", version=cfg.APP_VERSION),
            text_color=MUTED, justify="left", anchor="w",
            font=ctk.CTkFont(size=11)).pack(
                side="bottom", fill="x", padx=22, pady=12)

        self.content_host = GlassFrame(
            workspace, glass_surface="sheet", fg_color=SURFACE, corner_radius=10, border_width=0,
            border_color=LINE)
        self.content_host.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        self.content_host.grid_rowconfigure(0, weight=1)
        self.content_host.grid_columnconfigure(0, weight=1)

    def _build_footer(self):
        footer = GlassFrame(self, glass_surface="sheet", height=66, fg_color=SURFACE, corner_radius=0,
                              border_width=0)
        footer.grid(row=1, column=0, sticky="ew")
        footer.grid_propagate(False)
        footer.pack_propagate(False)
        self._settings_footer = footer
        row = ctk.CTkFrame(footer, fg_color="transparent", corner_radius=0)
        row.pack(fill="x", pady=(10, 16))
        self._settings_footer_row = row
        self.footer_status = ctk.CTkLabel(
            row, text=I.t("settings_no_changes"), text_color=SUCCESS,
            font=ctk.CTkFont(size=13), anchor="w")
        self.footer_status.pack(side="left", padx=12)
        VisualButton(
            row, text=I.t("settings_restore_defaults"),
            width=_ui_font(13).measure(I.t("settings_restore_defaults")) + 22, font=_ui_font(13),
            height=ACTION_HEIGHT, fg_color=SURFACE, hover_color=ACCENT_SOFT,
            text_color=MUTED, border_width=1, border_color=LINE,
            command=self.restore_defaults).pack(side="left", padx=4)
        VisualButton(
            row, text=I.t("settings_reset_all"),
            width=_ui_font(13).measure(I.t("settings_reset_all")) + 22, font=_ui_font(13),
            height=ACTION_HEIGHT, fg_color=SURFACE, hover_color=DANGER_SOFT,
            text_color=DANGER, border_width=1, border_color=LINE,
            command=self.reset_all_settings).pack(side="left", padx=4)
        VisualButton(
            row, text=I.t("settings_save_changes"),
            width=_ui_font(13).measure(I.t("settings_save_changes")) + 22,
            font=_ui_font(13), height=ACTION_HEIGHT,
            fg_color=GOOD, hover_color=GOOD_HOVER, command=self.save).pack(
                side="right", padx=(6, 12))

    def _new_page(self, key, title, subtitle):
        page = GlassFrame(
            self.content_host, glass_surface="sheet", fg_color=SURFACE, corner_radius=10)
        page.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        page.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            page, text=title, text_color=TEXT, anchor="w",
            font=ctk.CTkFont(size=24, weight="bold")).grid(
                row=0, column=0, sticky="ew", padx=14, pady=(8, 12))
        self._pages[key] = page
        return page

    def _card(self, parent, row, title=None):
        card = GlassFrame(
            parent, fg_color=CARD, corner_radius=10, border_width=0,
            border_color=LINE)
        card.grid(row=row, column=0, sticky="ew", padx=10, pady=(0, 8))
        card.grid_columnconfigure(0, weight=1)
        if title:
            ctk.CTkLabel(
                card, text=title, text_color=TEXT_SECONDARY, anchor="w",
                font=ctk.CTkFont(size=14, weight="bold")).grid(
                    row=0, column=0, sticky="ew", padx=PAD, pady=(6, CARD_GAP))
        return card

    @staticmethod
    def _entry(parent, variable, row, label, show=None):
        ctk.CTkLabel(
            parent, text=label, text_color=MUTED, anchor="w",
            font=_ui_font(12)).grid(
                row=row, column=0, sticky="ew", padx=14, pady=(5, 3))
        entry = VisualEntry(
            parent, textvariable=variable, height=36, corner_radius=8,
            border_color=LINE, show=show,
            font=ctk.CTkFont(size=13))
        entry.grid(row=row + 1, column=0, sticky="ew", padx=14, pady=(0, 6))
        return entry

    def _build_pages(self):
        self._build_general_page()
        self._build_clinic_page()
        self._build_documents_page()
        self._build_qr_page()
        self._build_database_page()
        self._build_gemini_page()
        self._build_security_page()
        self._build_recovery_page()
        self._build_about_page()

    def _build_general_page(self):
        page = self._new_page(
            "general", I.t("settings_general"), I.t("settings_general_tip"))
        card = self._card(page, 2, I.t("settings_app_preferences"))
        card.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkLabel(card, text=I.t("language"), text_color=MUTED,
                     anchor="w", font=ctk.CTkFont(size=14, weight="bold")).grid(
                         row=1, column=0, sticky="ew", padx=(20, 8), pady=(6, 4))
        ctk.CTkLabel(card, text=I.t("settings_app_version"), text_color=MUTED,
                     anchor="w", font=ctk.CTkFont(size=14, weight="bold")).grid(
                         row=1, column=1, sticky="ew", padx=(8, 20), pady=(6, 4))
        VisualOptionMenu(
            card, values=["English", "العربية"], variable=self.language_var,
            height=44, corner_radius=9, fg_color=ACCENT_SOFT,
            button_color=PRIMARY, button_hover_color=ACCENT_HOVER,
            text_color=TEXT_SECONDARY, font=ctk.CTkFont(size=15)).grid(
                row=2, column=0, sticky="ew", padx=(20, 8), pady=(0, 20))
        ctk.CTkLabel(
            card, text=f"Rx Prescription Printer  ·  {cfg.APP_VERSION}",
            height=44, corner_radius=9, fg_color=SURFACE,
            text_color=ACCENT, anchor="w", padx=14,
            font=ctk.CTkFont(size=15, weight="bold")).grid(
                row=2, column=1, sticky="ew", padx=(8, 20), pady=(0, 20))
        display = self._card(page, 3, I.t("settings_display_fonts"))
        display.grid_columnconfigure((0, 1), weight=1)
        display.winfo_children()[-1].grid_configure(columnspan=2)
        sizes = [str(size) for size in range(10, 57, 2)]
        for column, (label, variable, values) in enumerate((
                (I.t("settings_dropdown_font"), self.dropdown_font_var, [I.t("settings_font_default")] + sizes),
                (I.t("settings_patient_font"), self.patient_font_var, sizes))):
            ctk.CTkLabel(display, text=label, text_color=MUTED, anchor="w", font=_ui_font(12)).grid(
                row=1, column=column, sticky="ew", padx=14, pady=(2, 4))
            VisualOptionMenu(display, variable=variable, values=values, height=FIELD_HEIGHT,
                             font=_ui_font(13), fg_color=SURFACE, text_color=TEXT,
                             button_color=ACCENT_SOFT, button_hover_color=LINE).grid(
                row=2, column=column, sticky="ew", padx=14, pady=(0, 10))

    def _build_clinic_page(self):
        page = self._new_page(
            "clinic", I.t("settings_clinic"), I.t("settings_clinic_tip"))
        card = self._card(page, 2, I.t("settings_clinic_contact"))
        card.grid_columnconfigure((0, 1, 2), weight=1, uniform="clinic_contact_fields")
        card.winfo_children()[-1].grid_configure(columnspan=3)
        contact_fields = (
            (self.clinic_name_var, I.t("settings_clinic_name")),
            (self.clinic_address_var, I.t("settings_clinic_address")),
            (self.clinic_phone_var, I.t("settings_clinic_phone")),
        )
        for column, (variable, label) in enumerate(contact_fields):
            left_pad = 14 if column == 0 else 5
            right_pad = 14 if column == 2 else 5
            ctk.CTkLabel(
                card, text=label, text_color=MUTED, anchor="w",
                font=ctk.CTkFont(size=12, weight="bold")).grid(
                    row=1, column=column, sticky="ew",
                    padx=(left_pad, right_pad), pady=(2, 3))
            VisualEntry(
                card, textvariable=variable, height=36, corner_radius=8,
                border_color=LINE, font=ctk.CTkFont(size=13)).grid(
                    row=2, column=column, sticky="ew",
                    padx=(left_pad, right_pad), pady=(0, 10))
        logo_card = self._card(page, 3, I.t("settings_logo"))
        logo_card.grid_columnconfigure(0, weight=0)
        logo_card.grid_columnconfigure(1, weight=1)
        logo_card.winfo_children()[-1].grid_configure(columnspan=4)
        self.logo_thumbnail = ctk.CTkLabel(
            logo_card, text="Rx", width=54, height=46, corner_radius=8,
            fg_color=SURFACE, text_color=ACCENT, font=_ui_font(18, "bold"))
        self.logo_thumbnail.grid(row=1, column=0, padx=(14, 8), pady=(0, 10))
        self.logo_entry = VisualEntry(
            logo_card, textvariable=self.logo_var, height=36, corner_radius=8,
            border_color=LINE, font=ctk.CTkFont(size=13))
        self.logo_entry.grid(row=1, column=1, sticky="ew", padx=(0, 6), pady=(0, 10))
        VisualButton(
            logo_card, text=I.t("settings_browse"), width=88, height=36,
            fg_color=SURFACE, hover_color=ACCENT_SOFT, text_color=ACCENT,
            border_width=1, border_color=LINE,
            command=self.choose_logo).grid(row=1, column=2, padx=(0, 6), pady=(0, 10))
        VisualButton(
            logo_card, text=I.t("settings_remove"), width=88, height=36,
            fg_color=SURFACE, hover_color=DANGER_SOFT, text_color=DANGER,
            border_width=1, border_color=LINE,
            command=self.remove_logo).grid(
                row=1, column=3, padx=(0, 14), pady=(0, 10))
        preview = self._card(page, 4, I.t("settings_header_preview"))
        self.clinic_preview_card = preview
        preview.grid_columnconfigure(0, weight=0)
        preview.grid_columnconfigure(1, weight=1)
        preview.winfo_children()[-1].grid_configure(columnspan=2)
        placeholder = Image.new("RGBA", (1, 1), (0, 0, 0, 0))
        self._clinic_preview_placeholder = ctk.CTkImage(
            light_image=placeholder, dark_image=placeholder, size=(1, 1))
        self.clinic_preview_logo = ctk.CTkLabel(
            preview, text="Rx", width=60, height=50, corner_radius=8,
            fg_color=ACCENT_SOFT, text_color=ACCENT,
            image=self._clinic_preview_placeholder, compound="left",
            font=ctk.CTkFont(size=18, weight="bold"))
        self.clinic_preview_logo.grid(row=1, column=0, padx=(14, 10), pady=(0, 10))
        self.clinic_preview_text = ctk.CTkLabel(
            preview, text="", text_color=TEXT_SECONDARY, anchor="w",
            justify="left", font=ctk.CTkFont(size=12, weight="bold"))
        self.clinic_preview_text.grid(
            row=1, column=1, sticky="ew", padx=(0, 14), pady=(0, 10))

    def _build_documents_page(self):
        page = self._new_page(
            "documents", I.t("settings_documents"), I.t("settings_documents_tip"))
        card = self._card(page, 2, I.t("settings_page_format"))
        card.grid_columnconfigure((0, 1), weight=1)
        fields = ((I.t("paper"), 0), (I.t("settings_document_language"), 1))
        for label, column in fields:
            ctk.CTkLabel(
                card, text=label, text_color=MUTED, anchor="w",
                font=ctk.CTkFont(size=13, weight="bold")).grid(
                    row=1, column=column, sticky="ew", padx=20, pady=(4, 4))
        VisualOptionMenu(
            card, values=list(cfg.PAPER_SIZES.keys()), variable=self.paper_var,
            height=44, corner_radius=9, fg_color=ACCENT_SOFT,
            button_color=PRIMARY, button_hover_color=ACCENT_HOVER,
            text_color=TEXT_SECONDARY, font=ctk.CTkFont(size=15)).grid(
                row=2, column=0, sticky="ew", padx=20, pady=(0, 10))
        VisualOptionMenu(
            card, values=[I.t("settings_same_as_interface"), "English", "العربية"],
            variable=self.document_language_var, height=44, corner_radius=9,
            fg_color=ACCENT_SOFT, button_color=PRIMARY,
            button_hover_color=ACCENT_HOVER, text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(size=15)).grid(
                row=2, column=1, sticky="ew", padx=20, pady=(0, 10))
        ctk.CTkCheckBox(
            card, text=I.t("settings_show_header"),
            variable=self.document_header_var, fg_color=PRIMARY,
            hover_color=ACCENT_HOVER).grid(
                row=3, column=0, columnspan=2, sticky="w",
                padx=20, pady=(4, 14))
        folder = self._card(page, 3, I.t("settings_export_folder"))
        folder.grid_columnconfigure(0, weight=1)
        VisualEntry(
            folder, textvariable=self.export_folder_var, height=42,
            border_color=LINE).grid(
                row=1, column=0, sticky="ew", padx=(20, 8), pady=(0, 16))
        VisualButton(
            folder, text=I.t("settings_browse"), width=100, height=42,
            fg_color=ACCENT_SOFT, text_color=ACCENT,
            command=self.choose_export_folder).grid(
                row=1, column=1, padx=(0, 20), pady=(0, 16))

    def _build_qr_page(self):
        page = self._new_page(
            "qr", I.t("settings_qr"), I.t("settings_qr_tip"))
        card = self._card(page, 2, I.t("cloud_settings_title"))
        ctk.CTkLabel(card, text=cloud_rx.API_URL, text_color=TEXT,
                     font=_ui_font(12), anchor="w").grid(
            row=1, column=0, sticky="ew", padx=20, pady=(4, 8))
        self.cloud_key_entry = self._entry(card, self.cloud_key_var, 2, I.t("cloud_api_key"), show="•")
        cloud_notice = ctk.CTkLabel(
            card, text=I.t("cloud_privacy"), text_color=MUTED, justify="left",
            anchor="w", wraplength=500, font=ctk.CTkFont(size=13))
        cloud_notice.grid(
                row=4, column=0, sticky="ew", padx=20, pady=(0, 12))
        card.bind("<Configure>", lambda event: cloud_notice.configure(
            wraplength=max(200, event.width - 40)), add="+")
        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=5, column=0, sticky="ew", padx=20, pady=(0, 20))
        VisualButton(
            actions, text=I.t("settings_open_viewer"), height=ACTION_HEIGHT,
            fg_color=ACCENT_SOFT, hover_color=LINE, text_color=ACCENT,
            command=lambda: webbrowser.open(cloud_rx.VIEWER_URL)).pack(side="left", padx=(0, 8))
        VisualButton(
            actions, text=I.t("cloud_remove_key"), height=ACTION_HEIGHT,
            fg_color=CARD, hover_color=ACCENT_SOFT, text_color=ACCENT,
            border_width=1, border_color=LINE,
            command=lambda: self.cloud_key_var.set("")).pack(side="left")

    def _build_database_page(self):
        page = self._new_page(
            "database", I.t("settings_database"), I.t("settings_database_tip"))
        card = self._card(page, 2, I.t("settings_current_database"))
        status = ctk.CTkFrame(card, fg_color="transparent")
        status.grid(row=1, column=0, sticky="ew", padx=14, pady=(1, 3))
        status.grid_columnconfigure(0, weight=1)
        self.database_path_label = ctk.CTkLabel(
            status, text=str(self.master.db.path), text_color=TEXT_SECONDARY,
            anchor="w", justify="left", wraplength=600,
            font=ctk.CTkFont(size=12))
        self.database_path_label.grid(row=0, column=0, sticky="ew")
        self.database_count_label = ctk.CTkLabel(
            status, text=str(len(self.master.db.drugs)), fg_color=ACCENT_SOFT,
            text_color=ACCENT, corner_radius=12, width=48, height=24,
            font=ctk.CTkFont(size=12, weight="bold"))
        self.database_count_label.grid(row=0, column=1, padx=(8, 0))
        self.database_meta_label = ctk.CTkLabel(
            status, text="", text_color=MUTED, anchor="w", justify="left",
            font=ctk.CTkFont(size=10))
        self.database_meta_label.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(2, 0))
        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=14, pady=(3, 10))
        self._database_action_button(
            actions, I.t("import_db"), self.import_database, primary=True)
        self._database_action_button(
            actions, I.t("export_db"), self.master.export_db)
        self._database_action_button(
            actions, I.t("remove_db"), self.remove_database, danger=True)
        self._database_action_button(
            actions, I.t("settings_validate_database"), self.validate_database)

        treatment_card = self._card(page, 3, I.t("settings_treatment_database"))
        treatment_status = ctk.CTkFrame(treatment_card, fg_color="transparent")
        treatment_status.grid(row=1, column=0, sticky="ew", padx=14, pady=(1, 3))
        treatment_status.grid_columnconfigure(0, weight=1)
        self.treatment_database_meta_label = ctk.CTkLabel(
            treatment_status, text="", text_color=TEXT_SECONDARY, anchor="w",
            justify="left", font=ctk.CTkFont(size=12))
        self.treatment_database_meta_label.grid(row=0, column=0, sticky="ew")
        self.treatment_database_count_label = ctk.CTkLabel(
            treatment_status, text="0", fg_color=ACCENT_SOFT,
            text_color=ACCENT, corner_radius=12, width=48, height=24,
            font=ctk.CTkFont(size=12, weight="bold"))
        self.treatment_database_count_label.grid(row=0, column=1, padx=(8, 0))
        treatment_actions = ctk.CTkFrame(treatment_card, fg_color="transparent")
        treatment_actions.grid(row=2, column=0, sticky="ew", padx=14, pady=(3, 10))
        self._database_action_button(
            treatment_actions, I.t("settings_import_treatment_database"),
            self.import_treatment_database, primary=True)
        self._database_action_button(
            treatment_actions, I.t("settings_export_treatment_database"),
            self.export_treatment_database)
        self._database_action_button(
            treatment_actions, I.t("settings_remove_treatment_database"),
            self.remove_treatment_database, danger=True)
        self._database_action_button(
            treatment_actions, I.t("settings_validate_treatment_database"),
            self.validate_treatment_database)
        self._sync_database_status()
        self._sync_treatment_database_status()

    @staticmethod
    def _database_action_button(parent, text, command, primary=False, danger=False):
        foreground = PRIMARY if primary else CARD
        text_color = "white" if primary else DANGER if danger else ACCENT
        hover = ACCENT_HOVER if primary else DANGER_SOFT if danger else ACCENT_SOFT
        button = VisualButton(
            parent, text=text, width=max(76, len(text) * 7 + 24), height=32,
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=foreground, hover_color=hover, text_color=text_color,
            border_width=0 if primary else 1, border_color=LINE,
            command=command)
        button.pack(side="left", padx=(0, 5))
        return button

    def _build_gemini_page(self):
        page = self._new_page(
            "gemini", I.t("settings_gemini"), I.t("settings_gemini_tip"))
        card = self._card(page, 2)
        ctk.CTkCheckBox(
            card, text=I.t("gemini_enable"), variable=self.gemini_enabled_var,
            fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=13, weight="bold")).grid(
                row=1, column=0, sticky="w", padx=14, pady=(3, 6))
        self.gemini_key_entry = self._entry(
            card, self.gemini_key_var, 2, I.t("gemini_api_key"), show="•")
        ctk.CTkLabel(
            card, text=I.t("gemini_api_key_tip"), text_color=MUTED,
            justify="left", anchor="w", wraplength=720,
            font=ctk.CTkFont(size=11)).grid(
                row=4, column=0, sticky="ew", padx=14, pady=(0, 5))
        self.gemini_status_label = ctk.CTkLabel(
            card, text="", text_color=MUTED, justify="left", anchor="w",
            wraplength=720, font=ctk.CTkFont(size=12))
        self.gemini_status_label.grid(row=5, column=0, sticky="ew", padx=14, pady=(0, 5))
        self._set_initial_gemini_status()
        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=6, column=0, sticky="ew", padx=14, pady=(0, 10))
        self.gemini_test_button = VisualButton(
            actions, text=I.t("gemini_test"), height=34,
            fg_color=ACCENT_SOFT, hover_color=LINE, text_color=ACCENT,
            command=self.test_gemini_connection)
        self.gemini_test_button.pack(side="left", padx=(0, 8))
        VisualButton(
            actions, text=I.t("gemini_remove_key"), height=34,
            fg_color=CARD, hover_color=DANGER_SOFT, text_color=DANGER,
            border_width=1, border_color=LINE,
            command=self.remove_gemini_key).pack(side="left")
        help_card = self._card(page, 3)
        help_header = ctk.CTkFrame(help_card, fg_color="transparent")
        help_header.grid(row=0, column=0, sticky="ew", padx=14, pady=(9, 3))
        ctk.CTkLabel(
            help_header, text=I.t("gemini_api_key_help"), text_color=TEXT_SECONDARY,
            anchor="w", font=ctk.CTkFont(size=13, weight="bold")).pack(
                side="left", fill="x", expand=True)
        VisualButton(
            help_header, text="aistudio.google.com/app/apikey ↗", width=225,
            height=30, fg_color="transparent", text_color=ICON_BLUE,
            border_width=0, hover_color=ACCENT_SOFT,
            font=ctk.CTkFont(size=11, underline=True),
            command=lambda: webbrowser.open(
                "https://aistudio.google.com/app/apikey")).pack(side="right")
        for row, key in enumerate((
                "gemini_api_key_step_1", "gemini_api_key_step_2",
                "gemini_api_key_step_3", "gemini_api_key_step_4"), 1):
            ctk.CTkLabel(
                help_card, text=f"{row}. {I.t(key)}", text_color=MUTED,
                justify="left", anchor="w", wraplength=760,
                font=ctk.CTkFont(size=11)).grid(
                    row=row, column=0, sticky="ew", padx=14,
                    pady=(2, 7 if row == 4 else 2))

    def _build_security_page(self):
        page = self._new_page(
            "security", I.t("settings_backup_security"),
            I.t("settings_backup_tip"))
        card = self._card(page, 2, I.t("settings_backup_restore"))
        ctk.CTkLabel(
            card, text=I.t("settings_backup_note"), text_color=MUTED,
            justify="left", anchor="w", wraplength=720,
            font=ctk.CTkFont(size=12)).grid(
                row=1, column=0, sticky="ew", padx=14, pady=(3, 8))
        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 10))
        VisualButton(
            actions, text=I.t("settings_create_backup"), height=34,
            fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            command=self.create_backup).pack(side="left", padx=(0, 8))
        VisualButton(
            actions, text=I.t("settings_restore_backup"), height=34,
            fg_color=ACCENT_SOFT, hover_color=LINE, text_color=ACCENT,
            command=self.master.restore_backup).pack(side="left")
        VisualButton(
            actions, text=I.t("settings_open_backup_folder"), height=34,
            fg_color=CARD, hover_color=ACCENT_SOFT, text_color=ACCENT,
            border_width=1, border_color=LINE,
            command=self.open_backup_folder).pack(side="left", padx=(8, 0))
        self.backup_status_label = ctk.CTkLabel(
            card, text=self._last_backup_text(), text_color=MUTED,
            anchor="w", font=ctk.CTkFont(size=11))
        self.backup_status_label.grid(
            row=3, column=0, sticky="ew", padx=14, pady=(0, 6))
        ctk.CTkCheckBox(
            card, text=I.t("settings_automatic_backup"),
            variable=self.auto_backup_var, fg_color=PRIMARY,
            hover_color=ACCENT_HOVER, font=ctk.CTkFont(size=12)).grid(
                row=4, column=0, sticky="w", padx=14, pady=(0, 9))
        protection = self._card(page, 3, I.t("settings_local_protection"))
        self.security_protection_card = protection
        ctk.CTkLabel(
            protection, text=I.t("settings_protection_note"), text_color=MUTED,
            justify="left", anchor="w", wraplength=720,
            font=ctk.CTkFont(size=12)).grid(
                row=1, column=0, sticky="ew", padx=14, pady=(2, 10))

    def _build_about_page(self):
        page = self._new_page(
            "about", I.t("settings_about"), I.t("settings_about_tip"))
        card = self._card(page, 2)
        ctk.CTkLabel(
            card, text=I.t("settings_safe_credit"), text_color=ACCENT,
            justify="center", anchor="center",
            font=ctk.CTkFont(size=18, weight="bold")).grid(
                row=0, column=0, sticky="ew", padx=18, pady=18)

    def _build_recovery_page(self):
        page = self._new_page("recovery", I.t("settings_recovery"), "")
        card = self._card(page, 2, I.t("recovery_recently_deleted"))
        ctk.CTkLabel(
            card, text=I.t("recovery_retention"), text_color=MUTED,
            anchor="w", font=ctk.CTkFont(size=11)).grid(
                row=1, column=0, sticky="ew", padx=14, pady=(0, 7))
        self.recovery_items_body = ctk.CTkFrame(
            card, height=330, fg_color="transparent")
        self.recovery_items_body.grid(row=2, column=0, sticky="nsew", padx=8, pady=(0, 10))
        recovery_nav = ctk.CTkFrame(card, fg_color="transparent")
        recovery_nav.grid(row=3, column=0, sticky="e", padx=10, pady=(0, 8))
        VisualButton(recovery_nav, text=I.t("previous"), width=86, height=28,
                      fg_color=CARD, text_color=ACCENT, border_width=1,
                      border_color=LINE, command=lambda: self.change_recovery_page(-1)).pack(
                          side="left", padx=3)
        self.recovery_page_label = ctk.CTkLabel(
            recovery_nav, text="1 / 1", width=60, text_color=MUTED)
        self.recovery_page_label.pack(side="left", padx=3)
        VisualButton(recovery_nav, text=I.t("next_page"), width=86, height=28,
                      fg_color=CARD, text_color=ACCENT, border_width=1,
                      border_color=LINE, command=lambda: self.change_recovery_page(1)).pack(
                          side="left", padx=3)
        self.refresh_recovery_page()

    def refresh_recovery_page(self):
        if not hasattr(self, "recovery_items_body"):
            return
        for child in self.recovery_items_body.winfo_children():
            child.destroy()
        all_items = cfg.config.recovery_items()
        page_count = max(1, (len(all_items) + 6) // 7)
        self._recovery_page_index = max(0, min(self._recovery_page_index, page_count - 1))
        start = self._recovery_page_index * 7
        items = all_items[start:start + 7]
        if hasattr(self, "recovery_page_label"):
            self.recovery_page_label.configure(
                text=f"{self._recovery_page_index + 1} / {page_count}")
        if not items:
            ctk.CTkLabel(self.recovery_items_body, text=I.t("recovery_empty"),
                         text_color=MUTED, anchor="w").pack(fill="x", padx=8, pady=10)
            return
        kind_names = {
            "favorite": I.t("favorite_drugs"), "template": I.t("treatment_templates"),
            "patient": I.t("patient_details"), "prescription": I.t("previous_prescriptions"),
            "mapping": I.t("class_mapping_editor"),
            "template_database": I.t("settings_treatment_database"),
        }
        for item in items:
            row = ctk.CTkFrame(self.recovery_items_body, fg_color=SURFACE,
                               border_color=LINE, border_width=1, corner_radius=8)
            row.pack(fill="x", pady=3)
            text = (f"{kind_names.get(item.get('kind'), item.get('kind', ''))} · "
                    f"{item.get('label', '')}\n{str(item.get('deleted_at', ''))[:16].replace('T', ' ')}")
            ctk.CTkLabel(row, text=text, text_color=TEXT_SECONDARY, anchor="w",
                         justify="left", font=ctk.CTkFont(size=11, weight="bold")).pack(
                             side="left", fill="x", expand=True, padx=10, pady=7)
            VisualButton(
                row, text=I.t("restore_item"), width=82, height=30,
                fg_color=PRIMARY, hover_color=ACCENT_HOVER,
                command=lambda item_id=item.get("id", ""): self.restore_recovery_item(item_id)
            ).pack(side="left", padx=3)
            VisualButton(
                row, text=I.t("delete_permanently"), width=108, height=30,
                fg_color="transparent", text_color=DANGER, hover_color=DANGER_SOFT,
                command=lambda item_id=item.get("id", ""): self.delete_recovery_item(item_id)
            ).pack(side="left", padx=(3, 8))

    def change_recovery_page(self, step):
        self._recovery_page_index = max(0, self._recovery_page_index + step)
        self.refresh_recovery_page()

    def restore_recovery_item(self, item_id):
        item = next((entry for entry in cfg.config.recovery_items()
                     if entry.get("id") == item_id), None)
        if not item:
            return
        payload, kind = item.get("payload"), item.get("kind")
        restored = False
        if kind == "favorite" and isinstance(payload, dict):
            favorite = payload.get("favorite", payload)
            restored = cfg.config.insert_medication_favorite(
                int(payload.get("index", 0)), favorite)
            self.master.refresh_favorites_page()
        elif kind == "template" and isinstance(payload, dict):
            restored = bool(cfg.config.save_treatment_template(payload))
            self.master.refresh_treatment_template_menu()
        elif kind == "template_database" and isinstance(payload, list):
            restored = bool(cfg.config.merge_treatment_templates(payload))
            self.master.refresh_treatment_template_menu()
        elif kind == "patient" and isinstance(payload, dict):
            restored = self.master.patient_history.restore_record(payload)
            self.master.refresh_patient_history()
        elif kind == "prescription" and isinstance(payload, dict):
            restored = self.master.patient_history.restore_prescription(
                str(payload.get("patient_id", "")), payload.get("prescription", {}))
            self.master.refresh_patient_history()
        elif kind == "mapping" and isinstance(payload, list):
            restored = bool(self.master.db.restore_classification_states(payload))
            self.master._invalidate_database_caches()
            self.master.refresh_class_overview()
        if restored:
            cfg.config.discard_recovery_item(item_id)
            self.refresh_recovery_page()

    def delete_recovery_item(self, item_id):
        if not messagebox.askyesno(
                I.t("delete_permanently"), I.t("delete_permanently_confirm"), parent=self):
            return
        cfg.config.discard_recovery_item(item_id)
        self.refresh_recovery_page()

    def _show_section(self, section):
        self._active_section = section
        if section == "recovery":
            self.refresh_recovery_page()
        for key, page in self._pages.items():
            if key == section:
                page.grid()
            else:
                page.grid_remove()
        for key, button in self._nav_buttons.items():
            active = key == section
            button.configure(
                fg_color=NAV_ACTIVE_SOFT if active else "transparent",
                text_color=NAV_ACTIVE if active else TEXT_SECONDARY,
                image=self._settings_nav_icons[key][1 if active else 0],
                font=_ui_font(16))
            self._nav_rows[key].configure(fg_color=NAV_ACTIVE_SOFT if active else "transparent")
            if active:
                self._nav_indicators[key].place(x=0, rely=0.5, anchor="w")
            else:
                self._nav_indicators[key].place_forget()

    def _filter_settings_sections(self):
        query = self.settings_search_var.get().casefold().strip()
        keywords = {
            "general": "language version application",
            "clinic": "clinic identity name address phone logo header preview",
            "documents": "document pdf word paper margin export folder header logo language",
            "qr": "qr verification viewer key link",
            "database": "drug database medicine import replace remove validate classification",
            "gemini": "gemini api key ai connection quota offline reference",
            "security": "backup restore automatic security encryption local patient",
            "recovery": "recovery restore deleted favorites templates patients prescriptions mappings",
            "about": "about version diagnostic log data folder privacy",
        }
        matches = []
        for key, button in self._nav_buttons.items():
            haystack = f"{button.cget('text')} {keywords.get(key, '')}".casefold()
            row = self._nav_rows[key]
            row.pack_forget()
            if not query or query in haystack:
                row.pack(fill="x", padx=8, pady=1)
                matches.append(key)
        if matches and self._active_section not in matches:
            self._show_section(matches[0])

    def _snapshot(self):
        return tuple(variable.get() for variable in self._tracked_variables) if hasattr(
            self, "_tracked_variables") else ()

    def _mark_dirty(self, *_args):
        if not hasattr(self, "footer_status"):
            return
        dirty = self._snapshot() != self._saved_snapshot
        self.footer_status.configure(
            text=I.t("settings_unsaved") if dirty else I.t("settings_no_changes"),
            text_color=WARNING if dirty else SUCCESS)

    def cancel(self):
        if hasattr(self, "_saved_snapshot") and self._snapshot() != self._saved_snapshot:
            if not messagebox.askyesno(
                    I.t("set_title"), I.t("settings_discard_confirm"), parent=self):
                return
        self.destroy()

    def save(self):
        logo_path = self.logo_var.get().strip()
        if logo_path and not Path(logo_path).is_file():
            self._show_section("clinic")
            messagebox.showerror(
                I.t("set_title"), I.t("settings_invalid_logo"), parent=self)
            return
        export_folder = self.export_folder_var.get().strip()
        if export_folder and not Path(export_folder).is_dir():
            self._show_section("documents")
            messagebox.showerror(
                I.t("set_title"), I.t("settings_invalid_export_folder"), parent=self)
            return
        gemini_key = self.gemini_key_var.get().strip()
        if self.gemini_enabled_var.get() and not gemini_key:
            self._show_section("gemini")
            self.gemini_status_label.configure(
                text=I.t("gemini_not_configured"), text_color=DANGER)
            return
        language = "ar" if self.language_var.get() == "العربية" else "en"
        language_changed = language != cfg.config.language
        try:
            with cfg.config.batch_save():
                cfg.config.set_clinic(name=self.clinic_name_var.get().strip(),
                                      address=self.clinic_address_var.get().strip(),
                                      phone=self.clinic_phone_var.get().strip(),
                                      logo_path=logo_path)
                cfg.config.cloud_rx_api_key = self.cloud_key_var.get()
                cfg.config.paper_size = self.paper_var.get()
                cfg.config.language = language
                cfg.config.set_ui_font_sizes(
                    0 if self.dropdown_font_var.get() == I.t("settings_font_default") else int(self.dropdown_font_var.get()),
                    int(self.patient_font_var.get()))
                cfg.config.set_gemini(gemini_key, self.gemini_enabled_var.get())
                document_language = {
                    "English": "en", "العربية": "ar",
                    I.t("settings_same_as_interface"): "interface",
                }.get(self.document_language_var.get(), "interface")
                cfg.config.data["document_defaults"] = {
                    "language": document_language,
                    "show_header": bool(self.document_header_var.get()),
                    "logo_size": "medium",
                    "margin_mm": 16,
                    "export_folder": export_folder,
                }
                cfg.config.data["auto_backup_enabled"] = bool(self.auto_backup_var.get())
                cfg.config.save()
            cfg.config.maybe_create_automatic_backup()
        except Exception as exc:
            logging.exception("Could not save application settings")
            messagebox.showerror(
                I.t("set_title"), f"Settings could not be saved:\n{exc}", parent=self)
            return
        if hasattr(self.master, "paper_var"):
            self.master.paper_var.set(self.paper_var.get())
        if hasattr(self.master, "lang_var"):
            self.master.lang_var.set(language)
        self._saved_snapshot = self._snapshot()
        self.destroy()
        if language_changed:
            I.set_lang(language)
            self.master.reload_texts()
        if hasattr(self.master, "apply_ui_font_preferences"):
            self.master.apply_ui_font_preferences()
        messagebox.showinfo(I.t("set_title"), I.t("settings_saved"), parent=self.master)

    def restore_defaults(self):
        if not messagebox.askyesno(
                I.t("settings_restore_defaults"), I.t("settings_defaults_confirm"),
                parent=self):
            return
        if self._active_section == "general":
            self.language_var.set("English")
            self.dropdown_font_var.set(I.t("settings_font_default"))
            self.patient_font_var.set("14")
        elif self._active_section == "clinic":
            self.clinic_name_var.set("")
            self.clinic_address_var.set("")
            self.clinic_phone_var.set("")
            self.logo_var.set("")
        elif self._active_section == "documents":
            self.paper_var.set(cfg.DEFAULT_PAPER)
            self.document_language_var.set(I.t("settings_same_as_interface"))
            self.document_header_var.set(True)
            self.export_folder_var.set("")
        elif self._active_section == "qr":
            self.cloud_key_var.set("")
        elif self._active_section == "gemini":
            self.gemini_enabled_var.set(False)
            self.gemini_key_var.set("")
        elif self._active_section == "security":
            self.auto_backup_var.set(False)
        else:
            messagebox.showinfo(
                I.t("settings_restore_defaults"), I.t("settings_no_defaults"), parent=self)

    def reset_all_settings(self):
        if not messagebox.askyesno(
                I.t("settings_reset_all"), I.t("settings_reset_all_confirm"), parent=self):
            return
        self.language_var.set("English")
        self.clinic_name_var.set("")
        self.dropdown_font_var.set(I.t("settings_font_default"))
        self.patient_font_var.set("14")
        self.clinic_address_var.set("")
        self.clinic_phone_var.set("")
        self.logo_var.set("")
        self.paper_var.set(cfg.DEFAULT_PAPER)
        self.document_language_var.set(I.t("settings_same_as_interface"))
        self.document_header_var.set(True)
        self.export_folder_var.set("")
        self.cloud_key_var.set("")
        self.gemini_enabled_var.set(False)
        self.gemini_key_var.set("")
        self.auto_backup_var.set(False)

    def open_viewer(self):
        webbrowser.open(cloud_rx.VIEWER_URL)

    def choose_logo(self):
        path = filedialog.askopenfilename(
            title=I.t("settings_logo"), parent=self,
            filetypes=[("Images", "*.png *.jpg *.jpeg *.webp *.bmp"),
                       ("All files", "*.*")])
        if not path:
            return
        try:
            with Image.open(path) as image:
                image.verify()
        except Exception:
            logging.exception("Selected clinic logo is not a valid image")
            messagebox.showerror(
                I.t("settings_logo"), I.t("settings_invalid_logo_image"), parent=self)
            return
        self.logo_var.set(path)

    def remove_logo(self):
        if self.logo_var.get().strip() and not messagebox.askyesno(
                I.t("settings_remove"), I.t("settings_remove_logo_confirm"), parent=self):
            return
        self.logo_var.set("")

    def choose_export_folder(self):
        path = filedialog.askdirectory(
            title=I.t("settings_export_folder"),
            initialdir=self.export_folder_var.get().strip() or None,
            parent=self)
        if path:
            self.export_folder_var.set(path)

    def update_clinic_preview(self):
        if not hasattr(self, "clinic_preview_text"):
            return
        header_visible = self.document_header_var.get()
        name = self.clinic_name_var.get().strip() or I.t("settings_clinic_name")
        details = "  |  ".join(
            item for item in (self.clinic_address_var.get().strip(),
                              self.clinic_phone_var.get().strip()) if item)
        self.clinic_preview_text.configure(text=(name + ("\n" + details if details else ""))
                                           if header_visible else I.t("settings_header_hidden"))
        logo_path = self.logo_var.get().strip()
        try:
            if logo_path and Path(logo_path).is_file():
                pixels = 54
                image = Image.open(logo_path).convert("RGBA")
                image.thumbnail((pixels, pixels), Image.Resampling.LANCZOS)
                self._clinic_preview_image = ctk.CTkImage(
                    light_image=image, dark_image=image, size=image.size)
                self.clinic_preview_logo.configure(
                    text="" if header_visible else "—",
                    image=self._clinic_preview_image if header_visible else self._clinic_preview_placeholder)
                self.logo_thumbnail.configure(text="", image=self._clinic_preview_image)
                return
        except Exception:
            logging.exception("Clinic logo preview could not be rendered")
        self._clinic_preview_image = None
        self.clinic_preview_logo.configure(
            text="Rx" if header_visible else "—", image=self._clinic_preview_placeholder)
        self.logo_thumbnail.configure(text="Rx", image=self._clinic_preview_placeholder)

    def import_database(self):
        self.master.import_db(on_done=self._sync_database_status)

    def remove_database(self):
        self.master.remove_db(on_done=self._sync_database_status)

    def _sync_database_status(self):
        try:
            self.database_path_label.configure(text=str(self.master.db.path))
            self.database_count_label.configure(text=str(len(self.master.db.drugs)))
            report = self.master.class_mapping_integrity()
            imported = cfg.config.get("drug_db_imported_at", "") or I.t("settings_not_available")
            self.database_meta_label.configure(text=I.t(
                "settings_database_metadata",
                name=Path(self.master.db.path).name,
                imported=imported,
                classified=max(0, len(self.master.db.drugs) - report["unclassified"]),
                unclassified=report["unclassified"]))
        except Exception:
            pass

    def validate_database(self):
        def worker():
            drugs = self.master.db.drugs
            blank = 0
            seen = set()
            duplicates = 0
            for drug in drugs:
                blank += not (drug.brand_name.strip() or drug.generic_name.strip())
                key = (drug.brand_name.strip().casefold(), drug.generic_name.strip().casefold())
                duplicates += key in seen
                seen.add(key)
            report = self.master.class_mapping_integrity()
            return len(drugs), blank, duplicates, report

        def finished(result):
            if not self.winfo_exists():
                return
            total, blank, duplicates, report = result
            messagebox.showinfo(
                I.t("settings_validate_database"), I.t(
                    "settings_database_validation", total=total, blank=blank,
                    duplicates=duplicates, unclassified=report["unclassified"],
                    invalid=report["invalid_groups"] + report["invalid_details"]),
                parent=self)

        self.master.submit_background(
            worker, finished, label=I.t("settings_validate_database") + "…")

    def _sync_treatment_database_status(self):
        try:
            templates = cfg.config.treatment_templates()
            medicine_count = sum(
                len(template.get("medications", [])) for template in templates)
            self.treatment_database_count_label.configure(text=str(len(templates)))
            self.treatment_database_meta_label.configure(text=I.t(
                "settings_treatment_database_metadata",
                templates=len(templates), medicines=medicine_count))
        except Exception:
            pass

    def import_treatment_database(self):
        self.master.import_treatment_templates_xlsx(
            parent=self, on_done=self._sync_treatment_database_status)

    def export_treatment_database(self):
        self.master.export_treatment_templates_review(
            parent=self, on_done=self._sync_treatment_database_status)

    def remove_treatment_database(self):
        templates = cfg.config.treatment_templates()
        if not templates:
            messagebox.showinfo(
                I.t("settings_remove_treatment_database"),
                I.t("settings_treatment_database_empty"), parent=self)
            return
        if not messagebox.askyesno(
                I.t("settings_remove_treatment_database"),
                I.t("settings_remove_treatment_confirm"), parent=self):
            return
        cfg.config.add_recovery_item(
            "template_database", I.t("settings_treatment_database"), templates)
        cfg.config.data["treatment_templates"] = []
        cfg.config.save()
        self.master.new_treatment_template()
        self._sync_treatment_database_status()
        messagebox.showinfo(
            I.t("settings_remove_treatment_database"),
            I.t("settings_treatment_database_removed"), parent=self)

    def validate_treatment_database(self):
        def worker():
            templates = cfg.config.treatment_templates()
            medicines = blank_medicines = duplicate_diseases = 0
            seen_diseases = set()
            for template in templates:
                disease_key = template.get("disease", "").strip().casefold()
                duplicate_diseases += disease_key in seen_diseases
                seen_diseases.add(disease_key)
                for medicine in template.get("medications", []):
                    medicines += 1
                    if not (str(medicine.get("brand_name", "")).strip()
                            or str(medicine.get("generic_name", "")).strip()):
                        blank_medicines += 1
            return len(templates), medicines, blank_medicines, duplicate_diseases

        def finished(result):
            if not self.winfo_exists():
                return
            template_count, medicines, blank_medicines, duplicates = result
            messagebox.showinfo(
                I.t("settings_validate_treatment_database"), I.t(
                    "settings_treatment_database_validation",
                    templates=template_count, medicines=medicines,
                    blank=blank_medicines, duplicates=duplicates), parent=self)

        self.master.submit_background(
            worker, finished,
            label=I.t("settings_validate_treatment_database") + "…")

    def remove_gemini_key(self):
        if self.gemini_key_var.get().strip() and not messagebox.askyesno(
                I.t("gemini_remove_key"), I.t("settings_remove_key_confirm"), parent=self):
            return
        self.gemini_key_var.set("")
        self.gemini_enabled_var.set(False)
        gemini_drug.clear_cache()
        self.gemini_status_label.configure(
            text=I.t("gemini_key_removed"), text_color=ACCENT)

    def _set_initial_gemini_status(self):
        key = self.gemini_key_var.get().strip()
        last_test = cfg.config.get("gemini_last_test", "")
        if not key:
            text, color = I.t("gemini_not_configured"), MUTED
        elif not self.gemini_enabled_var.get():
            text, color = I.t("settings_connection_disabled"), MUTED
        elif last_test:
            text, color = I.t("settings_connection_last_ok", date=last_test), ACCENT
        else:
            text, color = I.t("settings_connection_not_tested"), WARNING
        self.gemini_status_label.configure(text=text, text_color=color)

    def test_gemini_connection(self):
        key = self.gemini_key_var.get().strip()
        if not key:
            self.gemini_status_label.configure(
                text=I.t("gemini_not_configured"), text_color=DANGER)
            return
        self.gemini_test_button.configure(state="disabled")
        self.gemini_status_label.configure(
            text=I.t("gemini_testing"), text_color=WARNING)

        def worker():
            try:
                gemini_drug.GeminiDrugClient(key).test_connection()
            except Exception as exc:
                logging.exception("Gemini connection test failed")
                self.after(0, lambda: self._gemini_test_finished(False, str(exc)))
                return
            self.after(0, lambda: self._gemini_test_finished(True, ""))

        threading.Thread(target=worker, daemon=True).start()

    def _gemini_test_finished(self, success, detail):
        try:
            if not self.winfo_exists():
                return
            self.gemini_test_button.configure(state="normal")
            if success:
                checked = datetime.datetime.now(datetime.timezone.utc).isoformat(
                    timespec="seconds")
                cfg.config.set("gemini_last_test", checked)
                text = I.t("settings_connection_last_ok", date=checked)
            else:
                lowered = detail.casefold()
                prefix = (I.t("settings_connection_quota") if "quota" in lowered
                          else I.t("settings_connection_offline")
                          if any(term in lowered for term in ("network", "connect", "offline"))
                          else I.t("settings_connection_failed"))
                text = f"{prefix}: {detail}"
            self.gemini_status_label.configure(
                text=text, text_color=ACCENT if success else DANGER)
        except Exception:
            pass

    def _last_backup_text(self):
        last = cfg.config.get("last_backup_at", "")
        return (I.t("settings_last_backup", date=last) if last
                else I.t("settings_no_backup"))

    def create_backup(self):
        self.master.create_backup()
        if hasattr(self, "backup_status_label"):
            self.backup_status_label.configure(text=self._last_backup_text())

    def open_backup_folder(self):
        folder = cfg.APP_DIR / "backups"
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder))

class GeminiSettingsWindow(ctk.CTkToplevel):
    """Manage the per-user Gemini key stored in the encrypted app config."""

    def __init__(self, master):
        super().__init__(master)
        self.title(I.t("gemini_settings"))
        self.geometry("600x390")
        self.transient(master)

        ctk.CTkLabel(
            self, text=I.t("gemini_settings"), text_color=ACCENT,
            font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(
                fill="x", padx=18, pady=(18, 10))
        self.enabled_var = tk.BooleanVar(value=cfg.config.gemini_enabled)
        ctk.CTkCheckBox(
            self, text=I.t("gemini_enable"), variable=self.enabled_var,
            fg_color=PRIMARY, hover_color=ACCENT_HOVER).pack(
                anchor="w", padx=18, pady=(0, 10))
        ctk.CTkLabel(self, text=I.t("gemini_api_key"), text_color=MUTED,
                     anchor="w").pack(fill="x", padx=18, pady=(0, 3))
        self.key_var = tk.StringVar(value=cfg.config.gemini_api_key)
        self.key_entry = VisualEntry(
            self, textvariable=self.key_var, show="•", height=FIELD_HEIGHT,
            corner_radius=9, border_color=LINE)
        self.key_entry.pack(fill="x", padx=18)
        ctk.CTkLabel(
            self, text=I.t("gemini_api_key_tip"), text_color=MUTED,
            anchor="w", justify="left", wraplength=550).pack(
                fill="x", padx=18, pady=(6, 12))
        self.status_label = ctk.CTkLabel(
            self, text="", text_color=MUTED, anchor="w", wraplength=550,
            justify="left")
        self.status_label.pack(fill="x", padx=18, pady=(0, 10))
        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.pack(fill="x", padx=18, pady=(4, 0))
        self.test_button = VisualButton(
            actions, text=I.t("gemini_test"), height=ACTION_HEIGHT,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=self.test_connection)
        self.test_button.pack(side="left", padx=(0, 6))
        VisualButton(
            actions, text=I.t("gemini_remove_key"), height=ACTION_HEIGHT,
            fg_color=DANGER_FILL, hover_color=DANGER_HOVER, command=self.remove_key).pack(
                side="left", padx=6)
        VisualButton(
            actions, text=I.t("set_save"), height=ACTION_HEIGHT,
            fg_color=GOOD, hover_color=GOOD_HOVER, command=self.save).pack(side="right")

    def save(self):
        key = self.key_var.get().strip()
        if self.enabled_var.get() and not key:
            self.status_label.configure(text=I.t("gemini_not_configured"), text_color=DANGER)
            return
        try:
            cfg.config.set_gemini(key, self.enabled_var.get())
        except Exception as exc:
            logging.exception("Could not save Gemini settings")
            self.status_label.configure(text=str(exc), text_color=DANGER)
            return
        messagebox.showinfo(I.t("gemini_settings"), I.t("gemini_settings_saved"))
        self.destroy()

    def remove_key(self):
        try:
            cfg.config.remove_gemini_key()
            gemini_drug.clear_cache()
        except Exception as exc:
            logging.exception("Could not remove Gemini API key")
            self.status_label.configure(text=str(exc), text_color=DANGER)
            return
        self.key_var.set("")
        self.enabled_var.set(False)
        self.status_label.configure(text=I.t("gemini_key_removed"), text_color=ACCENT)

    def test_connection(self):
        key = self.key_var.get().strip()
        if not key:
            self.status_label.configure(text=I.t("gemini_not_configured"), text_color=DANGER)
            return
        self.test_button.configure(state="disabled")
        self.status_label.configure(text=I.t("gemini_testing"), text_color=WARNING)

        def worker():
            try:
                gemini_drug.GeminiDrugClient(key).test_connection()
            except Exception as exc:
                logging.exception("Gemini connection test failed")
                self.after(0, lambda: self._test_finished(False, str(exc)))
                return
            self.after(0, lambda: self._test_finished(True, ""))

        threading.Thread(target=worker, daemon=True).start()

    def _test_finished(self, success, detail):
        try:
            if not self.winfo_exists():
                return
            self.test_button.configure(state="normal")
            self.status_label.configure(
                text=I.t("gemini_test_ok") if success else detail,
                text_color=ACCENT if success else DANGER)
        except Exception:
            pass


if __name__ == "__main__":
    app = App()
    app.mainloop()
