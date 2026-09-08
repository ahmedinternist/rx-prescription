"""Prescription desktop app — main GUI (CustomTkinter).

Features
--------
* Bilingual (English / Arabic) UI + documents.
* Prescriber (doctor) and Patient boxes shown SIDE BY SIDE and kept compact.
* Repeatable drug rows: drug NAME on top, with Dosage / Frequency / Duration /
  Notes boxes directly beneath it (in that order), and autocomplete.
* Importable drug database (replace / merge) + export.
* Paper size: A5 / A4 / Letter.
* Outputs:
    - Preview/Print : full prescription PDF (doctor, patient, drug table, QR).
    - Medication Label : compact Word doc with the full medication content + QR.
    - Export Word : editable .docx of the full prescription.

Run:  python main.py
"""
from __future__ import annotations

import datetime
import logging
import os
import tempfile
import threading
import tkinter as tk
import unicodedata
import uuid
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog
from typing import List

import customtkinter as ctk
from PIL import Image, ImageDraw, ImageFont

import config as cfg
import drug_db as dbmod
import pdf_generator as pdfgen
import qr_utils as qu
import i18n as I
import openfda
import drug_classes as classes
import gemini_drug
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

# --- Calm clinical visual system --------------------------------------------
BG = "#f3f8f7"          # soft mint-white workspace
CARD = "#ffffff"        # paper-like card surface
ACCENT = "#167d78"      # single clinical teal accent
ACCENT_HOVER = "#106863"
ACCENT_SOFT = "#e4f4f1" # restrained teal tint for selections and chips
ICON_BLUE = "#3568b2"   # reference-screen icon blue
LINE = "#d9e7e4"        # subtle cool border
MUTED = "#647775"       # secondary text
DANGER = "#c94747"      # red is reserved for destructive/warning actions
DANGER_HOVER = "#ad3636"
WARNING = "#b7791f"
WARNING_SOFT = "#fff3d6"
GOOD = ACCENT            # successful actions share the clinical accent
GOOD_HOVER = ACCENT_HOVER

PAD = 16                 # generous card breathing room
FIELD_HEIGHT = 40        # consistent text-entry and dropdown height
ACTION_HEIGHT = 38       # consistent action-button height
LIST_FONT = ("Segoe UI", 30)
PAGE_TITLE_FONT_SIZE = 35
SELECTED_MEDICINE_FONT_SIZE = 35
DASHBOARD_WIDTH = 190
PRESCRIBER_FIELD_WIDTH = 440
NAV_ICONS = {
    "prescriber": "✚", "patient": "⚕", "medications": "◒",
    "favorites": "★", "drug_classes": "⌬", "interaction_review": "⚯",
    "reference": "▤", "settings": "⚒",
}
REFERENCE_TAGS = {
    "contraindication": ("#fbe9e8", "#ae3d3d"),
    "pregnancy": ("#fff2d8", "#956014"),
    "interaction": ("#e6f0fb", "#2765a0"),
}

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


class DrugRow(ctk.CTkFrame):
    """One medication row with linked generic/trade and scientific-name inputs."""

    def __init__(self, master, db, on_change, on_remove,
                 on_move_up, on_move_down, on_reference, **kwargs):
        super().__init__(master, **kwargs)
        self.db = db
        self.on_change = on_change
        self.on_remove = on_remove
        self.on_move_up = on_move_up
        self.on_move_down = on_move_down
        self.on_reference = on_reference
        self._ac_top = None
        self._ac_listbox = None
        self._brand = ""
        self._category = ""
        self._bidi_bindings = []

        self.name_var = tk.StringVar()
        self.trade_var = tk.StringVar()
        name_row = ctk.CTkFrame(self, fg_color="transparent")
        name_row.pack(fill="x", padx=PAD, pady=(8, 2))
        self.number_badge = ctk.CTkLabel(
            name_row, text="1", width=28, height=28, corner_radius=14,
            fg_color=ACCENT_SOFT, text_color=ACCENT,
            font=ctk.CTkFont(weight="bold", size=12))
        self.number_badge.pack(side="left", padx=(0, 8))
        ctk.CTkLabel(name_row, text=I.t("drug"), width=48, anchor="w",
                     text_color=ACCENT, font=ctk.CTkFont(weight="bold", size=11)
                     ).pack(side="left")
        header_actions = ctk.CTkFrame(name_row, fg_color="transparent")
        header_actions.pack(side="right")
        self.up_button = ctk.CTkButton(
            header_actions, text="↑", width=30, height=28, corner_radius=7,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=lambda: self.on_move_up(self))
        self.up_button.pack(side="left", padx=2)
        self.down_button = ctk.CTkButton(
            header_actions, text="↓", width=30, height=28, corner_radius=7,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=lambda: self.on_move_down(self))
        self.down_button.pack(side="left", padx=2)
        ctk.CTkButton(
            header_actions, text=I.t("delete"), width=62, height=28, corner_radius=7,
            fg_color=DANGER, hover_color=DANGER_HOVER, font=ctk.CTkFont(size=10),
            command=self.on_remove).pack(side="left", padx=(2, 0))
        names = ctk.CTkFrame(self, fg_color="transparent")
        names.pack(fill="x", padx=PAD, pady=(0, 6))
        names.grid_columnconfigure(0, weight=1)
        names.grid_columnconfigure(1, weight=1)
        trade_col = ctk.CTkFrame(names, fg_color="transparent")
        trade_col.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        science_col = ctk.CTkFrame(names, fg_color="transparent")
        science_col.grid(row=0, column=1, sticky="ew", padx=(4, 0))
        ctk.CTkLabel(trade_col, text=I.t("generic_trade_name"), anchor="w", text_color=MUTED,
                     font=ctk.CTkFont(size=10, weight="bold")).pack(fill="x", pady=(0, 2))
        self.trade_entry = ctk.CTkEntry(trade_col, textvariable=self.trade_var,
                                        placeholder_text=I.t("generic_trade_name"), height=FIELD_HEIGHT,
                                        corner_radius=8, border_color=LINE,
                                        font=ctk.CTkFont(weight="bold", size=13))
        self.trade_entry.pack(fill="x")
        self.trade_entry.bind("<KeyRelease>", self._on_trade_type)
        science_heading = ctk.CTkFrame(science_col, fg_color="transparent")
        science_heading.pack(fill="x", pady=(0, 2))
        ctk.CTkLabel(science_heading, text=I.t("scientific_name"), anchor="w",
                     text_color=MUTED,
                     font=ctk.CTkFont(size=10, weight="bold")).pack(side="left")
        self.reference_button = ctk.CTkButton(
            science_heading, text="!", width=30, height=26, corner_radius=13,
            fg_color=WARNING_SOFT, text_color=WARNING, border_width=1,
            border_color=WARNING, hover_color="#f8dfaa",
            font=ctk.CTkFont(size=15, weight="bold"),
            command=lambda: self.on_reference(self, False))
        self.name_entry = ctk.CTkEntry(science_col, textvariable=self.name_var,
                                       placeholder_text=I.t("scientific_name"), height=FIELD_HEIGHT,
                                       corner_radius=8, border_color=LINE,
                                       font=ctk.CTkFont(weight="bold", size=13))
        self.name_entry.pack(fill="x")
        # The imported database supplies the first Generic / trade field.
        # Scientific name remains editable but deliberately has no DB popup.
        self.name_entry.bind("<KeyRelease>", lambda _event: self.on_change())
        self.name_entry.bind("<FocusIn>", lambda _event: self._hide_ac())
        # Sticky database dropdown belongs only to the Generic / trade field.
        self.trade_entry.bind("<Down>", self._ac_down)
        self.trade_entry.bind("<Up>", self._ac_up)
        self.trade_entry.bind("<Return>", self._ac_choose_current)
        self.trade_entry.bind("<Escape>", lambda _event: self._hide_ac())
        self.trade_picker = ctk.CTkOptionMenu(
            trade_col, values=[I.t("choose_linked_trade_name")], height=28,
            fg_color=ACCENT_SOFT, text_color=ACCENT, button_color=ACCENT,
            button_hover_color=ACCENT_HOVER, command=self._choose_linked_trade)

        self.class_badge = ctk.CTkLabel(
            self, text="", height=22, corner_radius=11, fg_color=ACCENT_SOFT,
            text_color=ACCENT, font=ctk.CTkFont(size=10, weight="bold"), anchor="w")

        details = ctk.CTkFrame(self, fg_color="transparent")
        self.details = details
        details.pack(fill="x", padx=PAD, pady=(0, 8))
        # Balanced 20% / 20% / 15% / 45% proportions for dosage, frequency,
        # duration and notes.
        for column, weight in enumerate((4, 4, 3, 9)):
            details.grid_columnconfigure(column, weight=weight)

        self.dosage_var = tk.StringVar()
        self.freq_var = tk.StringVar()
        self.dur_var = tk.StringVar()
        self.notes_var = tk.StringVar()
        self.dosage_entry = self._box(details, 0, I.t("dosage"),
                                      I.t("ph_dosage"), self.dosage_var)
        self.freq_entry = self._frequency_box(details, 1, I.t("frequency"),
                                              self.freq_var)
        self.dur_entry = self._box(details, 2, I.t("duration"),
                                   I.t("ph_duration"), self.dur_var)
        self.notes_entry = self._notes_box(details, 3, I.t("notes"),
                                           self.notes_var)

        self._reference_drug = ""
        self._reference_display = ""
        self.reference_panel = ctk.CTkFrame(
            self, fg_color="#fffdf7", border_color="#ead7a7",
            border_width=1, corner_radius=10)
        reference_head = ctk.CTkFrame(self.reference_panel, fg_color="transparent")
        reference_head.pack(fill="x", padx=10, pady=(8, 4))
        ctk.CTkLabel(
            reference_head, text=I.t("gemini_reference_title"), text_color=WARNING,
            font=ctk.CTkFont(size=12, weight="bold"), anchor="w").pack(
                side="left", fill="x", expand=True)
        ctk.CTkButton(
            reference_head, text=I.t("gemini_refresh"), width=72, height=28,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT,
            command=lambda: self.on_reference(self, True)).pack(side="right", padx=2)
        ctk.CTkButton(
            reference_head, text=I.t("gemini_copy"), width=58, height=28,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=self.copy_reference).pack(side="right", padx=2)
        ctk.CTkButton(
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

    def _scientific_name_changed(self, *_args):
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
                     font=ctk.CTkFont(size=10, weight="bold")).pack(fill="x", pady=(0, 2))
        binding = DirectionalTextBinding(self, var)
        self._bidi_bindings.append(binding)
        e = ctk.CTkEntry(col, textvariable=binding.display_var,
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
                     font=ctk.CTkFont(size=10, weight="bold")).pack(
                         fill="x", pady=(0, 2))
        binding = DirectionalTextBinding(self, var)
        self._bidi_bindings.append(binding)
        picker = ctk.CTkComboBox(
            col, values=[directional_display_text(value) for value in FREQUENCY_OPTIONS],
            variable=binding.display_var, height=FIELD_HEIGHT,
            corner_radius=9, border_width=2, border_color=ACCENT, fg_color=CARD,
            button_color=ACCENT, button_hover_color=ACCENT_HOVER,
            dropdown_fg_color=CARD, dropdown_hover_color=ACCENT_SOFT,
            font=ctk.CTkFont(size=13), dropdown_font=ctk.CTkFont(size=13),
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
                     font=ctk.CTkFont(size=10, weight="bold")).pack(
                         fill="x", pady=(0, 2))
        binding = DirectionalTextBinding(self, var)
        self._bidi_bindings.append(binding)
        picker = ctk.CTkComboBox(
            col, values=[directional_display_text(value) for value in NOTE_OPTIONS],
            variable=binding.display_var, height=FIELD_HEIGHT,
            corner_radius=9, border_width=2, border_color=ACCENT, fg_color=CARD,
            button_color=ACCENT, button_hover_color=ACCENT_HOVER,
            dropdown_fg_color=CARD, dropdown_hover_color=ACCENT_SOFT,
            font=ctk.CTkFont(size=13), dropdown_font=ctk.CTkFont(size=13),
            justify="left",
            command=lambda _value: self.on_change())
        binding.attach(picker)
        picker.pack(fill="x")
        picker.bind("<KeyRelease>", lambda ev: self.on_change())
        return picker

    def set_position(self, number):
        """Keep the number badge accurate after additions and reordering."""
        self.number_badge.configure(text=str(number))

    # autocomplete (sticky, large)
    def _on_scientific_type(self, event=None):
        self.on_change()
        self._update_class_badge()
        self._hide_trade_picker()
        q = self.name_var.get().strip()
        if len(q) < 1:
            self._hide_ac()
            return
        matches = self.db.search_scientific(q, limit=20)
        if not matches:
            self._hide_ac()
            return
        self._show_ac(matches, mode="scientific")

    def _on_trade_type(self, event=None):
        if event and event.keysym in {
                "Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R",
                "Control_L", "Control_R", "Alt_L", "Alt_R"}:
            return
        self.on_change()
        self._hide_trade_picker()
        q = self.trade_var.get().strip()
        if len(q) < 1:
            self._hide_ac()
            return
        matches = self.db.search_prescribable(q, limit=20)
        if not matches:
            self._hide_ac()
            return
        self._show_ac(matches, mode="trade")

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
        lb = tk.Listbox(top,
                        height=min(36, len(matches) * 2),
                        font=LIST_FONT,
                        bg="white", fg="#111", relief="solid", borderwidth=3,
                        activestyle="dotbox",
                        highlightthickness=3, highlightbackground=ACCENT,
                        selectbackground=ACCENT, selectforeground="white")
        max_chars = max((len(d.generic_name) + len(d.brand_name or "")
                         + len(d.strength or "") + 8) for d in matches)
        width = max(110, min(190, max_chars * 2))
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
        lb.pack(padx=2, pady=2)
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
        self._real_hide()
        self._update_class_badge(d)
        if mode == "scientific":
            self._show_trade_picker(d.generic_name, selected=d.brand_name)
        else:
            self._hide_trade_picker()
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
        self.patient_history = PatientHistory()
        self.rows: List[DrugRow] = []
        self._word_preview_job = None
        self._favorite_refresh_job = None
        self._build_ui()

    def _build_ui(self):
        self.paper_var = tk.StringVar(value=cfg.config.paper_size)
        self.lang_var = tk.StringVar(value=cfg.config.language)
        self.profile_var = tk.StringVar(value=cfg.config.get("active_profile", "Default"))

        # Dashboard: persistent navigation on the left; one data-entry page at a time.
        self.workspace = ctk.CTkFrame(self, fg_color="transparent")
        self.workspace.pack(fill="both", expand=True, padx=4, pady=(6, 4))
        self.dashboard = ctk.CTkFrame(self.workspace, width=DASHBOARD_WIDTH, fg_color="white",
                                      border_color=LINE, border_width=1, corner_radius=12)
        self.dashboard.pack(side="left", fill="y", padx=(0, 4))
        self.dashboard.pack_propagate(False)
        ctk.CTkLabel(self.dashboard, text=I.t("dashboard"), text_color=ACCENT,
                     font=ctk.CTkFont(weight="bold", size=16), anchor="w").pack(
                         fill="x", padx=16, pady=(18, 14))
        self.page_buttons = {}
        self._page_button_icons = {}
        self._dashboard_action_icons = []
        self._add_page_button("prescriber", I.t("prescriber_details"))
        self._add_page_button("patient", I.t("patient_details"))
        self._add_page_button("medications", I.t("medication_entry"))
        self._add_page_button("favorites", I.t("favorite_drugs"))
        self._add_page_button("drug_classes", I.t("drug_classes"))
        self._add_page_button("interaction_review", I.t("interaction_review"))
        self._add_page_button("reference", I.t("online_drug_reference"))
        self._add_dashboard_action("settings", I.t("settings"), self.open_settings)
        ctk.CTkLabel(self.dashboard,
                     text=I.t("local_encrypted", version=cfg.APP_VERSION),
                     text_color=MUTED, font=ctk.CTkFont(size=10),
                     justify="left", anchor="w", wraplength=155).pack(
                         side="bottom", fill="x", padx=14, pady=(6, 10))
        database_card = ctk.CTkFrame(
            self.dashboard, fg_color=ACCENT_SOFT, border_color=LINE,
            border_width=1, corner_radius=10)
        database_card.pack(side="bottom", fill="x", padx=10, pady=(4, 6))
        ctk.CTkLabel(
            database_card, text=I.t("settings_database"), text_color=MUTED,
            font=ctk.CTkFont(size=10, weight="bold"), anchor="w").pack(
                fill="x", padx=10, pady=(7, 0))
        self.db_label = ctk.CTkLabel(
            database_card, text=I.t("db_count", n=len(self.db.drugs)),
            text_color=ACCENT, font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w")
        self.db_label.pack(fill="x", padx=10, pady=(0, 7))

        self.content = ctk.CTkFrame(self.workspace, fg_color="transparent")
        self.content.pack(side="left", fill="both", expand=True)
        self.scroll = ctk.CTkScrollableFrame(self.content, fg_color=BG, corner_radius=8)
        self.scroll.pack(fill="both", expand=True)
        # CustomTkinter defaults to 30 px per scroll unit on Windows and then
        # requests many units per wheel event.  A smaller increment keeps fast
        # wheel/touchpad movement smooth when many medication cards are visible.
        self.scroll._parent_canvas.configure(yscrollincrement=12)

        self.build_forms()

        # action bar (fixed footer)
        self.action = ctk.CTkFrame(self, height=62, fg_color="white",
                                   border_color=LINE, border_width=1,
                                   corner_radius=12)
        self.action.pack(side="bottom", fill="x", padx=4, pady=(0, 4))
        # left cluster
        ctk.CTkButton(self.action, text=I.t("add_drug"), width=120, height=ACTION_HEIGHT,
                      corner_radius=9, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.add_drug_and_show).pack(side="left", padx=6)
        ctk.CTkButton(self.action, text=I.t("preview"), width=120, height=ACTION_HEIGHT,
                      corner_radius=9, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.preview).pack(side="left", padx=6)
        ctk.CTkButton(self.action, text=I.t("clear"), width=100, height=ACTION_HEIGHT,
                      corner_radius=9, fg_color=CARD, text_color=MUTED,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.clear_all).pack(side="left", padx=6)
        # right cluster (primary actions)
        ctk.CTkButton(self.action, text=I.t("export_compact"), width=160, height=ACTION_HEIGHT,
                      corner_radius=9, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.export_label).pack(side="right", padx=6)
        ctk.CTkButton(self.action, text=I.t("export_word"), width=120, height=ACTION_HEIGHT,
                      corner_radius=9, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.export_word).pack(side="right", padx=6)
        ctk.CTkButton(self.action, text=I.t("save_profile"), width=130, height=ACTION_HEIGHT,
                      corner_radius=9, fg_color=GOOD, hover_color=GOOD_HOVER,
                      command=self.save_profile).pack(side="right", padx=6)
        ctk.CTkButton(self.action, text=I.t("print"), width=110, height=ACTION_HEIGHT,
                      corner_radius=9, fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      font=ctk.CTkFont(weight="bold", size=13),
                      command=self.print_pdf).pack(side="right", padx=6)

        self.add_row()
        self.refresh_favorites_page()
        self.refresh_patient_history()
        self.load_profile_into_ui()
        self.show_page("prescriber")

    def confirm_close(self):
        """Require an explicit confirmation before closing the desktop app."""
        if messagebox.askyesno(
                I.t("confirm_close_title"), I.t("confirm_close_message"), parent=self):
            self.destroy()

    # -- forms ---------------------------------------------------------------
    def section(self, parent, title, accent=True):
        f = ctk.CTkFrame(parent, fg_color=CARD, border_color=LINE,
                         border_width=1, corner_radius=16)
        f.pack(fill="x", padx=2, pady=8)
        if accent:
            bar = ctk.CTkFrame(f, height=4, fg_color=ACCENT, corner_radius=0)
            bar.pack(fill="x", side="top")
        ctk.CTkLabel(f, text=title, font=ctk.CTkFont(weight="bold", size=14),
                     text_color=ACCENT, anchor="w").pack(anchor="w", padx=PAD, pady=(10, 4))
        return f

    def page_header(self, parent, title, subtitle):
        """A consistent title block makes every dashboard page easy to scan."""
        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(fill="x", padx=4, pady=(10, 5))
        ctk.CTkLabel(header, text=title, text_color=ACCENT,
                     font=ctk.CTkFont(weight="bold", size=PAGE_TITLE_FONT_SIZE),
                     anchor="w").pack(fill="x")
        if subtitle:
            ctk.CTkLabel(header, text=subtitle, text_color=MUTED,
                         font=ctk.CTkFont(size=11), anchor="w", justify="left",
                         wraplength=720).pack(fill="x", pady=(2, 0))
        return header

    def _add_page_button(self, key, text):
        symbol = NAV_ICONS.get(key, "•")
        normal_icon = _glyph_icon(symbol, ICON_BLUE)
        selected_icon = _glyph_icon(symbol, "white")
        button = ctk.CTkButton(
            self.dashboard, text=text, image=normal_icon, compound="left",
            anchor="w", height=42, corner_radius=9,
            fg_color="transparent", text_color=ACCENT, hover_color=ACCENT_SOFT,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda page=key: self.show_page(page))
        button.pack(fill="x", padx=10, pady=3)
        self.page_buttons[key] = button
        self._page_button_icons[key] = (normal_icon, selected_icon)

    def _add_dashboard_action(self, key, text, command):
        icon = _glyph_icon(NAV_ICONS.get(key, "•"), ICON_BLUE)
        self._dashboard_action_icons.append(icon)
        ctk.CTkButton(
            self.dashboard, text=text, image=icon, compound="left",
            anchor="w", height=42,
            corner_radius=9, fg_color="transparent", text_color=ACCENT,
            hover_color=ACCENT_SOFT, font=ctk.CTkFont(size=12, weight="bold"),
            command=command).pack(fill="x", padx=10, pady=3)

    def show_page(self, key):
        for name, page in self.pages.items():
            page.pack_forget()
            normal_icon, selected_icon = self._page_button_icons[name]
            self.page_buttons[name].configure(
                fg_color=ACCENT if name == key else "transparent",
                text_color="white" if name == key else ACCENT,
                image=selected_icon if name == key else normal_icon,
            )
        self.pages[key].pack(fill="both", expand=True, padx=4, pady=4)
        self.active_page = key

    def field(self, parent, label, var, width=260, placeholder=None, expand=True,
              directional=False):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=PAD, pady=7)
        ctk.CTkLabel(row, text=label, width=130, anchor="w",
                     text_color=MUTED, font=ctk.CTkFont(size=11),
                     height=FIELD_HEIGHT).pack(side="left")
        display_var = var
        binding = None
        if directional:
            binding = DirectionalTextBinding(self, var)
            self._bidi_bindings.append(binding)
            display_var = binding.display_var
        entry = ctk.CTkEntry(
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
        row.pack(fill="x", padx=PAD, pady=7)
        ctk.CTkLabel(
            row, text=label, width=130, anchor="w", text_color=MUTED,
            font=ctk.CTkFont(size=11), height=FIELD_HEIGHT,
        ).pack(side="left")
        binding = DirectionalTextBinding(self, var)
        self._bidi_bindings.append(binding)
        self.specialty_menu = ctk.CTkComboBox(
            row, values=self._specialty_values(var.get()),
            variable=binding.display_var,
            width=width, height=FIELD_HEIGHT, corner_radius=9,
            border_width=2, border_color=ACCENT, fg_color=CARD,
            button_color=ACCENT, button_hover_color=ACCENT_HOVER,
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
        row.pack(fill="x", padx=PAD, pady=7)
        ctk.CTkLabel(row, text=label, width=130, anchor="w",
                     text_color=MUTED, font=ctk.CTkFont(size=11),
                     height=FIELD_HEIGHT).pack(side="left")
        opts = [I.t("sex_m"), I.t("sex_f")]
        code_of = {I.t("sex_m"): "M", I.t("sex_f"): "F"}
        cur = var.get()
        sel = next((lab for lab, c in code_of.items() if c == cur), opts[0])
        om = ctk.CTkOptionMenu(row, values=opts, width=120, height=FIELD_HEIGHT,
                               corner_radius=9, fg_color=ACCENT_SOFT,
                               text_color=ACCENT, button_color=ACCENT,
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
        self.pages = {
            "prescriber": ctk.CTkFrame(self.scroll, fg_color="transparent"),
            "patient": ctk.CTkFrame(self.scroll, fg_color="transparent"),
            "medications": ctk.CTkFrame(self.scroll, fg_color="transparent"),
            "favorites": ctk.CTkFrame(self.scroll, fg_color="transparent"),
            "drug_classes": ctk.CTkFrame(self.scroll, fg_color="transparent"),
            "interaction_review": ctk.CTkFrame(self.scroll, fg_color="transparent"),
            "reference": ctk.CTkFrame(self.scroll, fg_color="transparent"),
        }

        self.page_header(self.pages["prescriber"], I.t("prescriber_details"), "")
        d = self.section(self.pages["prescriber"], I.t("prescriber"))
        self.field(
            d, I.t("f_name"), self.doctor_vars["name"],
            PRESCRIBER_FIELD_WIDTH, expand=False, directional=True)
        self.field(
            d, I.t("license"), self.doctor_vars["license_no"],
            PRESCRIBER_FIELD_WIDTH, expand=False)
        self.specialty_field(
            d, I.t("specialty"), self.doctor_vars["specialty"],
            PRESCRIBER_FIELD_WIDTH)

        self.page_header(self.pages["patient"], I.t("patient_details"), "")
        p = self.section(self.pages["patient"], I.t("patient"))
        self.field(
            p, I.t("f_name"), self.patient_vars["name"], 220,
            directional=True)
        self.field(p, I.t("age"), self.patient_vars["age"], 90)
        self.sex_field(p, I.t("sex"), self.patient_vars["sex"])
        patient_actions = ctk.CTkFrame(p, fg_color="transparent")
        patient_actions.pack(fill="x", padx=PAD, pady=(2, PAD))
        ctk.CTkButton(patient_actions, text=I.t("new_patient"), height=ACTION_HEIGHT,
                      command=self.new_patient).pack(side="left", padx=(0, 4))
        ctk.CTkButton(patient_actions, text=I.t("clear_patient"), height=ACTION_HEIGHT,
                      fg_color=CARD, text_color=ACCENT, border_width=1,
                      border_color=LINE, hover_color=ACCENT_SOFT,
                      command=self.clear_patient_details).pack(side="left")

        history = self.section(self.pages["patient"], I.t("patient_history"))
        self.patient_search_var = tk.StringVar()
        self.patient_search_var.trace_add("write", lambda *_: self.refresh_patient_history())
        # Compact, readable patient finder: a half-width search field paired
        # with a shorter result list rather than two full-width boxes.
        patient_finder = ctk.CTkFrame(history, fg_color="transparent")
        patient_finder.pack(anchor="w", padx=PAD, pady=(0, 8))
        ctk.CTkEntry(patient_finder, textvariable=self.patient_search_var, width=360,
                     height=FIELD_HEIGHT, placeholder_text=I.t("search_patients"), border_color=ACCENT,
                     corner_radius=9, font=ctk.CTkFont(size=16)).pack(anchor="w", pady=(0, 6))
        result_shell = ctk.CTkFrame(patient_finder, width=720, fg_color="#f8fcfb",
                                    border_color=LINE, border_width=1, corner_radius=9)
        result_shell.pack(anchor="w")
        self.patient_history_list = tk.Listbox(result_shell, height=5, width=39,
                                               font=LIST_FONT, bg="#f8fcfb", fg="#1a302e",
                                               relief="flat", borderwidth=0, highlightthickness=0,
                                               selectbackground=ACCENT, selectforeground="white",
                                               activestyle="none")
        self.patient_history_list.pack(fill="both", expand=True, padx=4, pady=4)
        self.patient_history_list.bind("<<ListboxSelect>>", self.select_patient_history)
        history_actions = ctk.CTkFrame(history, fg_color="transparent")
        history_actions.pack(fill="x", padx=PAD, pady=(0, PAD))
        ctk.CTkButton(history_actions, text=I.t("save_patient"), height=ACTION_HEIGHT,
                      command=self.save_patient_history).pack(side="left", padx=(0, 4))
        ctk.CTkButton(history_actions, text=I.t("load_patient"), height=ACTION_HEIGHT,
                      fg_color=CARD, text_color=ACCENT, border_width=1,
                      border_color=LINE, hover_color=ACCENT_SOFT,
                      command=self.load_selected_patient).pack(side="left", padx=4)
        ctk.CTkButton(history_actions, text=I.t("delete_patient"), height=ACTION_HEIGHT,
                      fg_color=DANGER, hover_color=DANGER_HOVER,
                      command=self.delete_selected_patient).pack(side="right")
        ctk.CTkLabel(history, text=I.t("previous_prescriptions"), text_color=ACCENT,
                     font=ctk.CTkFont(size=11, weight="bold"), anchor="w").pack(
                         fill="x", padx=PAD, pady=(8, 3))
        self.patient_prescriptions_body = ctk.CTkFrame(history, fg_color="transparent")
        self.patient_prescriptions_body.pack(fill="x", padx=PAD, pady=(0, PAD))
        self._show_empty_prescriptions()

        self.page_header(self.pages["medications"], I.t("medication_entry"), "")
        dr = self.section(self.pages["medications"], I.t("medications"))
        self.drugs_frame = ctk.CTkFrame(dr, fg_color="transparent")
        self.drugs_frame.pack(fill="x", padx=6, pady=(0, 2))

        self.page_header(self.pages["interaction_review"], I.t("interaction_review"),
                         I.t("page_interaction_help"))
        review = self.section(self.pages["interaction_review"], I.t("interaction_review"))
        ctk.CTkLabel(review, text=I.t("interaction_review_tip"), text_color=MUTED,
                     justify="left", wraplength=720, anchor="w").pack(fill="x", padx=PAD, pady=(0, 6))
        self.medscape_button = ctk.CTkButton(
            review, text=I.t("check_medscape"), height=ACTION_HEIGHT, width=250,
            fg_color=CARD, text_color=ACCENT, border_width=1,
            border_color=LINE, hover_color=ACCENT_SOFT,
            command=self.check_interactions_in_medscape)
        self.medscape_button.pack(anchor="w", padx=PAD, pady=(0, 8))

        preview = self.section(self.pages["medications"], I.t("word_preview"))
        self.word_preview_body = ctk.CTkFrame(preview, fg_color="transparent")
        self.word_preview_body.pack(fill="x", padx=PAD, pady=(0, PAD))
        ctk.CTkButton(self.pages["medications"], text=I.t("save_patient_prescription"),
                      height=ACTION_HEIGHT, width=230, fg_color=GOOD, hover_color=GOOD_HOVER,
                      command=self.save_prescription_for_patient).pack(
                          anchor="w", padx=18, pady=(4, 10))

        self.page_header(self.pages["reference"], I.t("online_drug_reference"),
                         I.t("page_reference_help"))
        safety = self.section(self.pages["reference"], I.t("online_drug_reference"))
        ctk.CTkLabel(safety, text=I.t("openfda_connected"), text_color=MUTED,
                     justify="left", wraplength=720, anchor="w").pack(fill="x", padx=PAD, pady=(0, 6))
        self.reference_lookup_button = ctk.CTkButton(
            safety, text=I.t("lookup_drug_labels"), height=ACTION_HEIGHT, width=230,
            fg_color=CARD, text_color=ACCENT, border_width=1,
            border_color=LINE, hover_color=ACCENT_SOFT,
            command=self.lookup_openfda_labels)
        self.reference_lookup_button.pack(anchor="w", padx=PAD, pady=(0, 6))
        self.reference_status = ctk.CTkLabel(safety, text=I.t("openfda_ready"),
                                             text_color=MUTED, anchor="w")
        self.reference_status.pack(fill="x", padx=PAD, pady=(0, 6))
        self.reference_cards = ctk.CTkFrame(safety, fg_color="transparent")
        self.reference_cards.pack(fill="x", padx=PAD, pady=(0, PAD))
        self._show_reference_placeholder()

        self.page_header(self.pages["favorites"], I.t("favorite_drugs"), "")
        favorites = self.section(self.pages["favorites"], I.t("favorite_drugs"))
        favorite_toolbar = ctk.CTkFrame(favorites, fg_color="transparent")
        favorite_toolbar.pack(fill="x", padx=PAD, pady=(0, 8))
        self.favorite_search_var = tk.StringVar()
        self.favorite_search_var.trace_add(
            "write", lambda *_: self._schedule_favorites_refresh())
        ctk.CTkEntry(
            favorite_toolbar, textvariable=self.favorite_search_var,
            height=FIELD_HEIGHT, placeholder_text=I.t("favorite_search"),
            border_color=LINE, font=ctk.CTkFont(size=15)).pack(
                side="left", fill="x", expand=True, padx=(0, 8))
        self.favorite_sort_var = tk.StringVar(value=I.t("favorite_sort_used"))
        self.favorite_sort_menu = ctk.CTkOptionMenu(
            favorite_toolbar,
            values=[I.t("favorite_sort_used"), I.t("favorite_sort_recent"),
                    I.t("favorite_sort_name"), I.t("favorite_sort_name_reverse")],
            variable=self.favorite_sort_var, width=145, height=38,
            fg_color=ACCENT_SOFT, text_color=ACCENT, button_color=ACCENT,
            button_hover_color=ACCENT_HOVER, font=ctk.CTkFont(size=13),
            command=lambda _value: self.refresh_favorites_page())
        self.favorite_sort_menu.pack(side="left", padx=(0, 5))
        ctk.CTkButton(
            favorite_toolbar, text=I.t("new_favorite"), width=125,
            height=38, fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=13),
            command=self.new_favorite).pack(side="left")

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
            border_color="#ead7a7", corner_radius=9)
        self.favorite_notice_label = ctk.CTkLabel(
            self.favorite_notice, text="", text_color=WARNING, anchor="w",
            font=ctk.CTkFont(size=12, weight="bold"))
        self.favorite_notice_label.pack(side="left", fill="x", expand=True, padx=12, pady=8)
        self.favorite_undo_button = ctk.CTkButton(
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
            favorites, fg_color="#f8fcfb", border_color=LINE,
            border_width=1, corner_radius=12)
        self.favorite_editor.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.favorite_editor_title = ctk.CTkLabel(
            self.favorite_editor, text=I.t("new_favorite"), text_color=ACCENT,
            font=ctk.CTkFont(size=18, weight="bold"), anchor="w")
        self.favorite_editor_title.grid(
            row=0, column=0, columnspan=3, sticky="ew", padx=14, pady=(12, 8))
        ctk.CTkButton(
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
                font=ctk.CTkFont(size=11, weight="bold")).grid(
                    row=row, column=column, sticky="ew", padx=6, pady=(2, 3))
            if values == "categories":
                widget = ctk.CTkComboBox(
                    self.favorite_editor, values=[""],
                    variable=self.favorite_vars[key], height=38,
                    border_color=LINE, fg_color=CARD,
                    button_color=ACCENT, button_hover_color=ACCENT_HOVER,
                    font=ctk.CTkFont(size=12))
                self.favorite_category_combo = widget
            elif values:
                widget = ctk.CTkComboBox(
                    self.favorite_editor, values=list(values),
                    variable=self.favorite_vars[key], height=38,
                    border_color=LINE, fg_color=CARD,
                    button_color=ACCENT, button_hover_color=ACCENT_HOVER,
                    font=ctk.CTkFont(size=12))
            else:
                widget = ctk.CTkEntry(
                    self.favorite_editor, textvariable=self.favorite_vars[key],
                    height=38, border_color=LINE,
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
            fg_color=CARD, text_color="#163d3a", anchor="w", padx=12,
            font=ctk.CTkFont(size=14, weight="bold"))
        self.favorite_preview_label.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(
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
        ctk.CTkButton(
            self.favorite_selection_bar, text=I.t("favorite_clear_selection"),
            width=72, height=28, fg_color="transparent", text_color=ACCENT,
            hover_color=LINE, command=self.clear_favorite_selection).pack(
                side="right", padx=(3, 7), pady=4)
        ctk.CTkButton(
            self.favorite_selection_bar, text=I.t("favorite_add_selected"),
            width=130, height=28, fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self.use_selected_favorites).pack(side="right", padx=3, pady=4)
        self.favorite_cards = ctk.CTkFrame(favorites, fg_color="transparent")
        self.favorite_cards.pack(fill="x", expand=False, padx=PAD, pady=(0, PAD))
        self.favorite_cards.grid_columnconfigure((0, 1, 2), weight=1, uniform="favorite")
        self._favorite_card_columns = 3
        self.favorite_cards.bind("<Configure>", self._on_favorite_cards_resize)

        self.page_header(self.pages["drug_classes"], I.t("drug_classes"), "")
        self.class_overview = ctk.CTkFrame(self.pages["drug_classes"], fg_color="transparent")
        self.class_overview.pack(fill="both", expand=True)
        class_page = self.section(self.class_overview, I.t("major_therapeutic_groups"))
        top = ctk.CTkFrame(class_page, fg_color="transparent")
        top.pack(fill="x", padx=PAD, pady=(0, 8))
        ctk.CTkLabel(top, text=I.t("drug_classes_hint"), text_color=MUTED,
                     font=ctk.CTkFont(size=11), anchor="w", justify="left",
                     wraplength=560).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(top, text=I.t("manage_class_mappings"), height=ACTION_HEIGHT, width=180,
                      fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
                      hover_color=ACCENT_SOFT, command=self.show_class_mapping_editor).pack(
                          side="right", padx=(10, 0))
        ctk.CTkButton(top, text=I.t("show_all_detailed_classes"), height=ACTION_HEIGHT, width=210,
                      fg_color=CARD, text_color=ACCENT, border_width=1,
                      border_color=LINE, hover_color=ACCENT_SOFT,
                      command=self.show_all_detailed_classes).pack(side="right", padx=(10, 0))
        self.unclassified_label = ctk.CTkLabel(class_page, text="", text_color=MUTED,
                                               font=ctk.CTkFont(size=11), anchor="w")
        self.unclassified_label.pack(fill="x", padx=PAD, pady=(0, 6))
        class_grid = ctk.CTkFrame(class_page, fg_color="transparent")
        class_grid.pack(fill="x", padx=PAD, pady=(0, 12))
        class_grid.grid_columnconfigure(0, weight=1)
        class_grid.grid_columnconfigure(1, weight=1)
        self.class_buttons = {}
        self.class_tiles = {}
        self.class_count_badges = {}
        self.class_favorite_buttons = {}
        for index, code in enumerate(self.ordered_therapeutic_groups()):
            tile = ctk.CTkFrame(class_grid, fg_color="transparent")
            button = ctk.CTkButton(
                tile, text=I.t("class_" + code), height=42, corner_radius=11,
                anchor="w", fg_color="#f8fcfb", text_color="#1a302e",
                border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
                command=lambda selected=code: self.open_therapeutic_group(selected))
            grid_args = {"row": index // 2, "column": index % 2, "sticky": "ew", "padx": 4, "pady": 4}
            if index == len(classes.GROUPS) - 1 and len(classes.GROUPS) % 2:
                grid_args["columnspan"] = 2
            tile.grid(**grid_args)
            tile.grid_columnconfigure(0, weight=1)
            button.grid(row=0, column=0, sticky="ew")
            count_badge = ctk.CTkLabel(
                tile, text="0", width=30, height=30, corner_radius=15,
                fg_color=ACCENT_SOFT, text_color=ACCENT,
                font=ctk.CTkFont(size=11, weight="bold"))
            count_badge.grid(row=0, column=1, padx=(5, 0))
            favorite_button = ctk.CTkButton(
                tile, text="☆", width=34, height=34, corner_radius=9,
                fg_color="transparent", text_color=ACCENT, hover_color=ACCENT_SOFT,
                command=lambda selected=code: self.toggle_therapeutic_group_favorite(selected))
            favorite_button.grid(row=0, column=2, padx=(4, 0))
            self.class_buttons[code] = button
            self.class_tiles[code] = tile
            self.class_count_badges[code] = count_badge
            self.class_favorite_buttons[code] = favorite_button
        self._selected_therapeutic_group = None
        self._class_visible_drugs = []
        self.refresh_class_overview()
        self.class_subpage = ctk.CTkFrame(self.pages["drug_classes"], fg_color="transparent")
        self._active_subclass_group = None

    # -- drug rows ----------------------------------------------------------
    def add_drug_and_show(self):
        self.show_page("medications")
        self.add_row()

    def add_row(self, data=None):
        row = DrugRow(self.drugs_frame, self.db, self.on_any_change,
                      lambda: self.remove_row(row),
                      lambda r: self.move_row(r, -1), lambda r: self.move_row(r, 1),
                      lambda r, force=False: self.query_gemini_drug(r, force),
                      fg_color=CARD, border_color=LINE, border_width=1, corner_radius=12)
        if data:
            row.name_var.set(data.generic_name)
            row.trade_var.set(data.brand_name)
            row._brand = data.brand_name
            row.dosage_var.set(data.dosage)
            row.freq_var.set(data.frequency)
            row.dur_var.set(data.duration)
            row.notes_var.set(data.notes)
        row.pack(fill="x", padx=2, pady=4)
        self.rows.append(row)
        self._number_drug_rows()
        if data:
            row._update_class_badge(self.db.find_exact(data.generic_name))
        self.on_any_change()
        return row

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
            for display_index, (source_index, favorite) in enumerate(indexed):
                self._render_favorite_card(display_index, source_index, favorite)
            if not indexed:
                self.favorite_empty_label = ctk.CTkLabel(
                    self.favorite_cards,
                    text=I.t("favorite_empty") if not favorites else I.t("favorite_no_matches"),
                    text_color=MUTED, font=ctk.CTkFont(size=15),
                    justify="center")
                self.favorite_empty_label.grid(
                    row=0, column=0, columnspan=column_count, sticky="ew", pady=28)
            self._refresh_favorite_selection_bar()
            self._show_favorite_notice()
        finally:
            self._refreshing_favorites = False
            self._restore_scroll_position(scroll_position)

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
        generic = favorite.get("generic_name", "").strip().casefold()
        brand = favorite.get("brand_name", "").strip().casefold()
        return any(
            drug.generic_name.strip().casefold() == generic
            or (brand and drug.brand_name.strip().casefold() == brand)
            for drug in self.db.drugs)

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
                ctk.CTkButton(
                    row_frame, text=category, width=button_width,
                    height=26, corner_radius=13,
                    fg_color=ACCENT if active else ACCENT_SOFT,
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
        matches = self.db.search_scientific(query, 20) if query else []
        self._show_favorite_autocomplete(matches, "scientific")

    def _on_favorite_brand_type(self, event=None):
        if self._favorite_autocomplete_navigation_key(event):
            return
        query = self.favorite_vars["brand_name"].get().strip()
        matches = []
        seen = set()
        if query:
            candidates = self.db.search_prescribable(query, 20)
            for drug in candidates:
                identity = (drug.generic_name.casefold(), drug.brand_name.casefold(),
                            drug.strength.casefold(), drug.form.casefold())
                if identity not in seen:
                    seen.add(identity)
                    matches.append(drug)
                if len(matches) == 20:
                    break
        self._show_favorite_autocomplete(matches, "brand")

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
        listbox = tk.Listbox(
            top, height=min(10, len(matches)), font=("Segoe UI", 30, "bold"),
            bg="white", fg="#111", relief="solid", borderwidth=2,
            activestyle="dotbox", highlightthickness=2,
            highlightbackground=ACCENT, selectbackground=ACCENT,
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
        listbox.pack(padx=2, pady=2)
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
            card = ctk.CTkFrame(
                self.favorite_cards, fg_color="#fbfdfd",
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
                    name_line, text=brand_name, text_color="#163d3a",
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
            star_button = ctk.CTkButton(
                heading, text="★" if favorite.get("pinned") else "☆",
                width=26, height=26, corner_radius=8, fg_color="transparent",
                text_color=ICON_BLUE, hover_color=ACCENT_SOFT,
                font=ctk.CTkFont(size=17),
                command=lambda item_id=favorite_id: self.toggle_favorite_pin(item_id))
            star_button.grid(row=0, column=1, padx=(6, 0))

            regimen_line = self._favorite_regimen_line(favorite)
            if regimen_line:
                ctk.CTkLabel(
                    card, text=regimen_line, text_color="#365a57",
                    anchor="w", justify="left", wraplength=350,
                    font=ctk.CTkFont(size=12)).grid(
                        row=1, column=0, sticky="ew", padx=8, pady=(1, 3))
            actions = ctk.CTkFrame(card, fg_color="transparent")
            actions.grid(row=2, column=0, sticky="ew", padx=6, pady=(0, 5))
            ctk.CTkButton(
                actions, text=I.t("favorite_use_rx"), width=78, height=28,
                fg_color=ACCENT, hover_color=ACCENT_HOVER,
                command=lambda item_id=favorite_id: self.use_favorite(item_id)).pack(
                    side="left", padx=2)
            selected_var = tk.BooleanVar(value=selected)
            ctk.CTkCheckBox(
                actions, text="", variable=selected_var, width=24, height=24,
                checkbox_width=20, checkbox_height=20, corner_radius=5,
                border_width=2, fg_color=ACCENT, hover_color=ACCENT_HOVER,
                command=lambda item_id=favorite_id, variable=selected_var,
                target=card: self._toggle_favorite_selection(
                    item_id, variable, target)).pack(side="left", padx=(5, 2))
            ctk.CTkButton(
                actions, text=I.t("favorite_edit_short"), width=52, height=28,
                fg_color=CARD, text_color=ACCENT, border_width=1,
                border_color=LINE, hover_color=ACCENT_SOFT,
                command=lambda item_id=favorite_id: self.edit_favorite(item_id)).pack(
                    side="left", padx=2)
            ctk.CTkButton(
                actions, text=I.t("favorite_delete_short"), width=56, height=28,
                fg_color="transparent", text_color=DANGER, hover_color="#fae9e9",
                command=lambda item_id=favorite_id: self.delete_favorite(item_id)).pack(
                    side="right", padx=2)
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
            fg_color="#eaf8f5" if highlighted else "#fbfdfd",
            border_color=ACCENT if selected or highlighted else LINE,
            border_width=2 if selected or highlighted else 1)
        columns = getattr(self, "_favorite_card_columns", 3)
        card.grid(row=display_index // columns, column=display_index % columns,
                  sticky="nsew", padx=3, pady=3)
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
            fg_color="#fbfdfd", border_color=ACCENT if selected else LINE,
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
                       ("generic_name", "brand_name", "dosage", "frequency", "duration", "notes")}
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

    def unclassified_medicine_count(self):
        return sum(1 for drug in self.db.drugs if classes.group_for(drug) is None)

    def therapeutic_group_medicine_count(self, code):
        return sum(1 for drug in self.db.drugs
                   if (found := classes.group_for(drug)) and found.code == code)

    def refresh_class_overview(self):
        if not hasattr(self, "class_tiles"):
            return
        pinned = set(cfg.config.favorite_therapeutic_groups())
        ordered = self.ordered_therapeutic_groups()
        for index, code in enumerate(ordered):
            tile = self.class_tiles[code]
            tile.grid_forget()
            grid_args = {"row": index // 2, "column": index % 2, "sticky": "ew", "padx": 4, "pady": 4}
            if index == len(ordered) - 1 and len(ordered) % 2:
                grid_args["columnspan"] = 2
            tile.grid(**grid_args)
            self.class_favorite_buttons[code].configure(text="★" if code in pinned else "☆")
            self.class_count_badges[code].configure(
                text=str(self.therapeutic_group_medicine_count(code)))
        if hasattr(self, "unclassified_label"):
            self.unclassified_label.configure(
                text=I.t("unclassified_medicines", n=self.unclassified_medicine_count()))

    def toggle_therapeutic_group_favorite(self, code):
        cfg.config.toggle_favorite_therapeutic_group(code)
        self.refresh_class_overview()

    def show_class_mapping_editor(self):
        """Edit local class metadata without requiring CSV editing."""
        self.class_overview.pack_forget()
        for child in self.class_subpage.winfo_children():
            child.destroy()
        self.class_subpage.pack(fill="both", expand=True)
        card = self.section(self.class_subpage, I.t("class_mapping_editor"))
        bar = ctk.CTkFrame(card, fg_color="transparent")
        bar.pack(fill="x", padx=PAD, pady=(0, 8))
        ctk.CTkButton(bar, text="← " + I.t("back_to_major_groups"), height=ACTION_HEIGHT,
                      width=170, fg_color=CARD, text_color=ACCENT,
                      border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
                      command=self.back_to_major_groups).pack(side="left")
        ctk.CTkLabel(bar, text=I.t("class_mapping_editor"), text_color=ACCENT,
                     font=ctk.CTkFont(size=16, weight="bold"), anchor="e").pack(
                         side="right", fill="x", expand=True)
        ctk.CTkLabel(card, text=I.t("class_mapping_hint"), text_color=MUTED,
                     font=ctk.CTkFont(size=11), anchor="w", justify="left",
                     wraplength=650).pack(fill="x", padx=PAD, pady=(0, 10))
        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=PAD, pady=(0, PAD))
        left = ctk.CTkFrame(body, fg_color="transparent")
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))
        right = ctk.CTkScrollableFrame(body, width=340, height=460, fg_color="#f8fcfb",
                                       border_color=LINE, border_width=1, corner_radius=12)
        right.pack(side="left", fill="y")
        self.mapping_search_var = tk.StringVar()
        self.mapping_search_var.trace_add("write", lambda *_: self.refresh_mapping_list())
        ctk.CTkEntry(left, textvariable=self.mapping_search_var, height=FIELD_HEIGHT,
                     placeholder_text=I.t("search_medicines"), border_color=LINE).pack(fill="x", pady=(0, 6))
        self.mapping_unclassified_only = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(left, text=I.t("show_unclassified_only"), variable=self.mapping_unclassified_only,
                         text_color=MUTED, fg_color=ACCENT, hover_color=ACCENT_HOVER,
                         command=self.refresh_mapping_list).pack(anchor="w", pady=(0, 6))
        self.mapping_list = tk.Listbox(left, height=5, font=LIST_FONT,
                                       bg="#f8fcfb", fg="#1a302e", relief="flat", borderwidth=0,
                                       highlightthickness=1, highlightbackground=LINE,
                                       selectbackground=ACCENT, selectforeground="white", activestyle="none")
        self.mapping_list.pack(fill="both", expand=True)
        self.mapping_list.bind("<<ListboxSelect>>", self.select_mapping_drug)
        self.mapping_group_labels = {I.t("class_" + code): code for code in classes.GROUPS
                                     if classes.subclasses_for(code)}
        self.mapping_group_var = tk.StringVar(value=I.t("choose_major_group"))
        self.mapping_detail_var = tk.StringVar(value=I.t("choose_detailed_class"))
        ctk.CTkLabel(right, text=I.t("mapping_selected_medicine"), text_color=MUTED,
                     font=ctk.CTkFont(size=10, weight="bold"), anchor="w").pack(
                         fill="x", padx=12, pady=(14, 2))
        self.mapping_selected_label = ctk.CTkLabel(right, text=I.t("no_medicine_selected"),
                                                    text_color="#1a302e", anchor="w",
                                                    font=ctk.CTkFont(
                                                        size=SELECTED_MEDICINE_FONT_SIZE,
                                                        weight="bold"), wraplength=270)
        self.mapping_selected_label.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkLabel(right, text=I.t("major_therapeutic_group"), text_color=MUTED,
                     font=ctk.CTkFont(size=10, weight="bold"), anchor="w").pack(fill="x", padx=12)
        self.mapping_group_menu = ctk.CTkOptionMenu(
            right, values=[I.t("choose_major_group")] + list(self.mapping_group_labels),
            variable=self.mapping_group_var, height=FIELD_HEIGHT, fg_color=ACCENT_SOFT, text_color=ACCENT,
            button_color=ACCENT, button_hover_color=ACCENT_HOVER,
            command=self.mapping_group_changed)
        self.mapping_group_menu.pack(fill="x", padx=12, pady=(3, 12))
        self.mapping_detail_shell = ctk.CTkFrame(
            right, height=108, fg_color=ACCENT_SOFT, border_color=ACCENT, border_width=1,
            corner_radius=10)
        self.mapping_detail_shell.pack_propagate(False)
        self.mapping_detail_label = ctk.CTkLabel(
            self.mapping_detail_shell, text=I.t("detailed_drug_class"), text_color=ACCENT,
            font=ctk.CTkFont(size=13, weight="bold"), anchor="w")
        self.mapping_detail_label.pack(fill="x", padx=10, pady=(8, 3))
        self.mapping_detail_menu = ctk.CTkOptionMenu(
            self.mapping_detail_shell, values=[I.t("choose_detailed_class")],
            variable=self.mapping_detail_var, height=FIELD_HEIGHT, font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=CARD, text_color=ACCENT, button_color=ACCENT,
            button_hover_color=ACCENT_HOVER, command=lambda _: self.mapping_detail_changed())
        self.mapping_detail_menu.pack(fill="x", padx=10, pady=(0, 8))
        self.add_mapping_drug_button = ctk.CTkButton(
            right, text=I.t("add_drug_to_class"), height=ACTION_HEIGHT,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, state="disabled",
                                                      command=self.add_drug_to_class)
        self.add_mapping_drug_button.pack(fill="x", padx=12, pady=(0, 6))
        self.save_mapping_button = ctk.CTkButton(right, text=I.t("save_changes"), height=ACTION_HEIGHT,
                                                  fg_color=ACCENT, hover_color=ACCENT_HOVER,
                                                  state="disabled", command=self.save_class_mapping)
        self.save_mapping_button.pack(fill="x", padx=12)
        self.mapping_status = ctk.CTkLabel(right, text="", text_color=MUTED,
                                           font=ctk.CTkFont(size=10), wraplength=270, justify="left")
        self.mapping_status.pack(fill="x", padx=12, pady=(8, 0))
        self.mapping_selected_drug = None
        self.pending_class_mapping = None
        self.mapping_visible_drugs = []
        self.refresh_mapping_list()

    def refresh_mapping_list(self):
        if not hasattr(self, "mapping_list"):
            return
        query = self.mapping_search_var.get().casefold().strip()
        drugs = []
        for drug in self.db.drugs:
            if self.mapping_unclassified_only.get() and classes.group_for(drug) is not None:
                continue
            if query and query not in (drug.generic_name + " " + drug.brand_name).casefold():
                continue
            drugs.append(drug)
        self.mapping_visible_drugs = sorted(drugs, key=lambda drug: drug.generic_name.casefold())
        self.mapping_list.delete(0, tk.END)
        for drug in self.mapping_visible_drugs:
            found = classes.group_for(drug)
            suffix = found.detail if found else I.t("unclassified")
            self.mapping_list.insert(tk.END, f"{drug.generic_name}  —  {suffix}")

    def select_mapping_drug(self, event=None):
        selected = self.mapping_list.curselection()
        if not selected:
            return
        drug = self.mapping_visible_drugs[selected[0]]
        self.mapping_selected_drug = drug
        self.mapping_selected_label.configure(text=drug.generic_name)
        found = classes.group_for(drug)
        if found and found.code in self.mapping_group_labels.values():
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
            self.mapping_detail_shell.pack(fill="x", padx=12, pady=(0, 14),
                                           before=self.add_mapping_drug_button)
        values = details or [I.t("choose_detailed_class")]
        self.mapping_detail_menu.configure(values=values)
        self.mapping_detail_var.set(preferred_detail if preferred_detail in details else values[0])

    def mapping_detail_changed(self):
        self.pending_class_mapping = None
        self.save_mapping_button.configure(state="disabled")
        self.mapping_status.configure(text="", text_color=MUTED)

    def add_drug_to_class(self):
        if not self.mapping_selected_drug:
            return
        code = self.mapping_group_labels.get(self.mapping_group_var.get())
        detail = self.mapping_detail_var.get()
        if not code or detail == I.t("choose_detailed_class"):
            self.mapping_status.configure(text=I.t("choose_group_and_class"), text_color=DANGER)
            return
        self.pending_class_mapping = (self.mapping_selected_drug.generic_name, code, detail)
        self.mapping_status.configure(
            text=I.t("mapping_ready_to_save", medicine=self.mapping_selected_drug.generic_name,
                     detail=detail), text_color=GOOD)
        self.save_mapping_button.configure(state="normal")

    def save_class_mapping(self):
        if not self.pending_class_mapping:
            return
        medicine, code, detail = self.pending_class_mapping
        if self.db.update_classification(medicine, code, detail):
            self.db.load()
            self.mapping_status.configure(
                text=I.t("class_mapping_saved_visible", group=I.t("class_" + code),
                         detail=detail), text_color=GOOD)
            self.refresh_mapping_list()
            self.refresh_class_overview()
            self.pending_class_mapping = None
            self.save_mapping_button.configure(state="disabled")

    def open_therapeutic_group(self, code):
        """Open detailed classes where a major group has a defined taxonomy."""
        if classes.subclasses_for(code):
            self.show_subclass_page(code)
            return
        self.select_therapeutic_group(code)

    def show_all_detailed_classes(self):
        """Show every configured detailed class, grouped on one scrollable page."""
        self.class_overview.pack_forget()
        for child in self.class_subpage.winfo_children():
            child.destroy()
        self.class_subpage.pack(fill="both", expand=True)
        card = self.section(self.class_subpage, I.t("all_detailed_drug_classes"))
        bar = ctk.CTkFrame(card, fg_color="transparent")
        bar.pack(fill="x", padx=PAD, pady=(0, 8))
        ctk.CTkButton(bar, text="← " + I.t("back_to_major_groups"), height=ACTION_HEIGHT,
                      width=170, fg_color=CARD, text_color=ACCENT,
                      border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
                      command=self.back_to_major_groups).pack(side="left")
        ctk.CTkLabel(bar, text=I.t("all_detailed_drug_classes"), text_color=ACCENT,
                     font=ctk.CTkFont(size=16, weight="bold"), anchor="e").pack(
                         side="right", fill="x", expand=True)
        ctk.CTkLabel(card, text=I.t("all_detailed_classes_hint"), text_color=MUTED,
                     font=ctk.CTkFont(size=11), anchor="w", justify="left",
                     wraplength=650).pack(fill="x", padx=PAD, pady=(0, 10))
        self.all_class_search_var = tk.StringVar()
        self.all_class_search_var.trace_add("write", lambda *_: self.render_all_detailed_classes())
        ctk.CTkEntry(card, textvariable=self.all_class_search_var, height=FIELD_HEIGHT,
                     placeholder_text=I.t("search_detailed_classes"), border_color=LINE).pack(
                         fill="x", padx=PAD, pady=(0, 8))
        self.all_classes_body = ctk.CTkScrollableFrame(card, height=440, fg_color="transparent")
        self.all_classes_body.pack(fill="x", padx=PAD, pady=(0, PAD))
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
            ctk.CTkLabel(self.all_classes_body, text=I.t("class_" + group_code), text_color=ACCENT,
                         font=ctk.CTkFont(size=13, weight="bold"), anchor="w").pack(
                             fill="x", pady=(10, 4))
            for detail in details:
                ctk.CTkButton(
                    self.all_classes_body, text=detail, height=33, corner_radius=8, anchor="w",
                    fg_color="#f8fcfb", text_color="#1a302e", border_width=1,
                    border_color=LINE, hover_color=ACCENT_SOFT,
                    command=lambda group=group_code, picked=detail:
                        self.show_subclass_page(group, picked, return_to_all=True)).pack(
                            fill="x", pady=2)
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
        card = self.section(self.class_subpage, I.t("detailed_drug_classes"))
        bar = ctk.CTkFrame(card, fg_color="transparent")
        bar.pack(fill="x", padx=PAD, pady=(0, 8))
        ctk.CTkButton(bar, text="← " + I.t("back_to_major_groups"), height=ACTION_HEIGHT,
                      width=170, fg_color=CARD, text_color=ACCENT,
                      border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
                      command=self.back_from_subclass_page).pack(side="left")
        ctk.CTkLabel(bar, text=I.t("class_" + group_code), text_color=ACCENT,
                     font=ctk.CTkFont(size=16, weight="bold"), anchor="e").pack(
                         side="right", fill="x", expand=True)
        ctk.CTkLabel(card, text=I.t("subclass_hint"), text_color=MUTED,
                     font=ctk.CTkFont(size=11), anchor="w", justify="left",
                     wraplength=650).pack(fill="x", padx=PAD, pady=(0, 10))
        # Keep the selection result on screen beneath the detailed-class list.
        list_frame = ctk.CTkScrollableFrame(card, height=230, fg_color="transparent")
        list_frame.pack(fill="x", padx=PAD, pady=(0, 10))
        self.subclass_buttons = {}
        subclass_counts = {}
        for detail in classes.subclasses_for(group_code):
            subclass_counts[detail] = sum(
                1 for drug in self.db.drugs
                if (found := classes.group_for(drug)) and found.code == group_code
                and found.detail.casefold() == detail.casefold())
        for detail in classes.subclasses_for(group_code):
            button = ctk.CTkButton(
                list_frame, text=I.t("detailed_class_with_count", detail=detail,
                                     n=subclass_counts[detail]), height=36, corner_radius=9, anchor="w",
                fg_color="#f8fcfb", text_color="#1a302e", border_width=1,
                border_color=LINE, hover_color=ACCENT_SOFT,
                command=lambda picked=detail: self.select_drug_subclass(group_code, picked))
            button.pack(fill="x", pady=3)
            self.subclass_buttons[detail] = button
        self.subclass_result_label = ctk.CTkLabel(card, text=I.t("choose_drug_class"),
                                                   text_color=MUTED, anchor="w")
        self.subclass_result_label.pack(fill="x", padx=PAD, pady=(0, 5))
        self.subclass_drug_list = tk.Listbox(card, height=3, font=LIST_FONT,
                                              bg="#f8fcfb", fg="#1a302e", relief="flat",
                                              borderwidth=0, highlightthickness=1,
                                              highlightbackground=LINE, selectbackground=ACCENT,
                                              selectforeground="white", activestyle="none")
        self.subclass_drug_list.pack(fill="x", padx=PAD, pady=(0, 8))
        self.subclass_drug_list.bind("<<ListboxSelect>>", self.select_subclass_drug)
        self.subclass_visible_drugs = []
        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.pack(fill="x", padx=PAD, pady=(0, PAD))
        self.add_subclass_drug_button = ctk.CTkButton(
            actions, text=I.t("add_selected_medicine"), height=ACTION_HEIGHT,
            fg_color=ACCENT, hover_color=ACCENT_HOVER, state="disabled",
            command=self.add_selected_subclass_drug)
        self.add_subclass_drug_button.pack(side="right")
        if selected_detail:
            self.select_drug_subclass(group_code, selected_detail)

    def back_to_major_groups(self):
        self.class_subpage.pack_forget()
        self.class_overview.pack(fill="both", expand=True)
        self.select_therapeutic_group(self._active_subclass_group)

    def back_from_subclass_page(self):
        if getattr(self, "_class_detail_return", "overview") == "all":
            self.show_all_detailed_classes()
            return
        self.back_to_major_groups()

    def select_drug_subclass(self, group_code, detail):
        for subclass, button in self.subclass_buttons.items():
            selected = subclass == detail
            button.configure(fg_color=ACCENT if selected else "#f8fcfb",
                             text_color="white" if selected else "#1a302e",
                             border_color=ACCENT if selected else LINE)
        visible = [(drug, found) for drug in self.db.drugs
                   if (found := classes.group_for(drug))
                   and found.code == group_code and found.detail.casefold() == detail.casefold()]
        visible.sort(key=lambda item: item[0].generic_name.casefold())
        self.subclass_visible_drugs = visible
        self.subclass_result_label.configure(
            text=I.t("mapped_medicine_count", detail=detail, n=len(visible)))
        self.subclass_drug_list.delete(0, tk.END)
        if visible:
            for drug, _ in visible:
                self.subclass_drug_list.insert(tk.END, drug.generic_name)
        else:
            self.subclass_drug_list.insert(tk.END, I.t("no_mapped_medicines"))
        self.add_subclass_drug_button.configure(state="disabled")

    def select_subclass_drug(self, event=None):
        selected = self.subclass_drug_list.curselection()
        self.add_subclass_drug_button.configure(
            state="normal" if selected and self.subclass_visible_drugs else "disabled")

    def add_selected_subclass_drug(self):
        selected = self.subclass_drug_list.curselection()
        if not selected:
            return
        drug, _ = self.subclass_visible_drugs[selected[0]]
        self.add_drug_database_item(drug)

    def add_drug_database_item(self, drug):
        """Place a browsed medicine into the first blank prescription row."""
        target = next((row for row in self.rows if not row.name_var.get().strip()), None)
        if target is None:
            target = self.add_row()
        target.name_var.set(drug.generic_name)
        target.trade_var.set(drug.brand_name)
        target._brand = drug.brand_name
        target._category = drug.category
        target._update_class_badge(drug)
        self.show_page("medications")
        self.on_any_change()

    def select_therapeutic_group(self, code):
        """Filter the local database by one major therapeutic group."""
        self._selected_therapeutic_group = code
        for group_code, button in self.class_buttons.items():
            selected = group_code == code
            button.configure(
                fg_color=ACCENT if selected else "#f8fcfb",
                text_color="white" if selected else "#1a302e",
                border_color=ACCENT if selected else LINE,
            )
        visible = []
        for drug in self.db.drugs:
            found = classes.group_for(drug)
            if found and (code is None or found.code == code):
                visible.append((drug, found))
        visible.sort(key=lambda item: item[0].generic_name.casefold())
        self._class_visible_drugs = visible
        # Major-group tiles are navigation only.  Their small count badges are
        # the complete summary; detailed pages show the medicine names.

    def select_class_drug(self, event=None):
        selected = self.class_drug_list.curselection()
        self.add_class_drug_button.configure(state="normal" if selected else "disabled")

    def add_selected_class_drug(self):
        selected = self.class_drug_list.curselection()
        if not selected:
            return
        drug, _ = self._class_visible_drugs[selected[0]]
        self.add_drug_database_item(drug)

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
        for v in self.patient_vars.values():
            v.set("")
        for r in list(self.rows):
            r.destroy()
        self.rows = []
        self.add_row()
        self.on_any_change()

    # -- encrypted patient history -----------------------------------------
    def refresh_patient_history(self):
        if not hasattr(self, "patient_history_list"):
            return
        self._patient_history_records = self.patient_history.search(self.patient_search_var.get())
        self.patient_history_list.delete(0, tk.END)
        for record in self._patient_history_records:
            prescriptions = record.get("prescriptions", [])
            last_saved = max((str(item.get("saved_at", ""))[:10] for item in prescriptions), default="")
            detail = I.t("last_prescription", date=last_saved) if last_saved else " · ".join(
                value for value in [record.get("age", ""), record.get("sex", "")] if value)
            self.patient_history_list.insert(tk.END, record.get("name", "") + (f" — {detail}" if detail else ""))

    def _show_empty_prescriptions(self):
        if not hasattr(self, "patient_prescriptions_body"):
            return
        for child in self.patient_prescriptions_body.winfo_children():
            child.destroy()
        ctk.CTkLabel(self.patient_prescriptions_body, text=I.t("previous_prescriptions_empty"),
                     text_color=MUTED, anchor="w", font=ctk.CTkFont(size=10)).pack(fill="x", pady=(2, 4))

    def select_patient_history(self, event=None):
        selected = self.patient_history_list.curselection()
        if not selected:
            return
        self.show_patient_prescriptions(self._patient_history_records[selected[0]])

    def show_patient_prescriptions(self, record):
        prescriptions = list(record.get("prescriptions", []))
        if not prescriptions:
            self._show_empty_prescriptions()
            return
        for child in self.patient_prescriptions_body.winfo_children():
            child.destroy()
        for number, prescription in enumerate(reversed(prescriptions), 1):
            saved_at = str(prescription.get("saved_at", ""))[:10]
            drugs = prescription.get("drugs", [])
            card = ctk.CTkFrame(self.patient_prescriptions_body, fg_color="#f8faff",
                                border_color=LINE, border_width=1, corner_radius=9)
            card.pack(fill="x", pady=3)
            summary = I.t("previous_prescription_heading", number=number, date=saved_at or "—")
            ctk.CTkLabel(card, text=summary, text_color=ACCENT,
                         font=ctk.CTkFont(size=11, weight="bold"), anchor="w").pack(
                             fill="x", padx=10, pady=(7, 1))
            names = ", ".join(str(drug.get("generic_name", "")) for drug in drugs[:3])
            if len(drugs) > 3:
                names += I.t("more_medicines", n=len(drugs) - 3)
            ctk.CTkLabel(card, text=I.t("previous_prescription_summary", count=len(drugs), names=names),
                         text_color=MUTED, font=ctk.CTkFont(size=10), anchor="w").pack(
                             fill="x", padx=10, pady=(0, 5))
            ctk.CTkButton(card, text=I.t("load_previous_prescription"), height=ACTION_HEIGHT, width=178,
                          fg_color=CARD, text_color=ACCENT, border_width=1,
                          border_color=LINE, hover_color=ACCENT_SOFT,
                          command=lambda item=prescription, patient=record: self.load_saved_prescription(patient, item)).pack(
                              anchor="e", padx=8, pady=(0, 7))

    def save_patient_history(self):
        try:
            self.patient_history.save_patient({key: variable.get() for key, variable in self.patient_vars.items()})
        except ValueError as exc:
            messagebox.showinfo(APP_TITLE, str(exc))
            return
        self.refresh_patient_history()

    def load_selected_patient(self):
        selected = self.patient_history_list.curselection()
        if not selected:
            return
        record = self._patient_history_records[selected[0]]
        for key, variable in self.patient_vars.items():
            variable.set(record.get(key, ""))
        self.show_patient_prescriptions(record)

    def delete_selected_patient(self):
        selected = self.patient_history_list.curselection()
        if not selected:
            return
        record = self._patient_history_records[selected[0]]
        if not messagebox.askyesno(I.t("delete_patient"), I.t("delete_patient_confirm", name=record.get("name", ""))):
            return
        if self.patient_history.delete(record.get("id", "")):
            self.refresh_patient_history()
            self._show_empty_prescriptions()

    def new_patient(self):
        self.clear_patient_details()
        self.patient_search_var.set("")

    def clear_patient_details(self):
        for variable in self.patient_vars.values():
            variable.set("")
        self.patient_history_list.selection_clear(0, tk.END)
        self._show_empty_prescriptions()

    def load_saved_prescription(self, patient, prescription):
        if self.rows and any(
                row.get_data().generic_name or row.get_data().brand_name
                for row in self.rows):
            if not messagebox.askyesno(I.t("load_previous_prescription"), I.t("replace_current_medicines")):
                return
        for key, variable in self.patient_vars.items():
            variable.set(patient.get(key, ""))
        for row in list(self.rows):
            row.destroy()
        self.rows = []
        for saved_drug in prescription.get("drugs", []):
            self.add_row(data=qu.DrugItem(**{key: saved_drug.get(key, "") for key in
                                              ("generic_name", "brand_name", "dosage", "frequency", "duration", "notes")}))
        if not self.rows:
            self.add_row()
        self.show_page("medications")

    def save_prescription_for_patient(self):
        patient = {key: variable.get() for key, variable in self.patient_vars.items()}
        drugs = [row.get_data() for row in self.rows
                 if row.get_data().generic_name or row.get_data().brand_name]
        try:
            record = self.patient_history.save_prescription(
                patient, [drug.__dict__ for drug in drugs])
        except ValueError as exc:
            messagebox.showinfo(I.t("save_patient_prescription"), str(exc))
            return
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

    def _render_word_preview_now(self):
        self._word_preview_job = None
        self.render_word_preview()

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
        self.reference_lookup_button.configure(state="disabled")
        self.reference_status.configure(text=I.t("openfda_searching"))
        self._clear_reference_cards()
        threading.Thread(target=self._lookup_openfda_worker, args=(medicines,), daemon=True).start()

    def _lookup_openfda_worker(self, medicines):
        try:
            references = openfda.lookup_labels(medicines)
            self.after(0, lambda: self._show_openfda_results(references, medicines))
        except openfda.OpenFDALookupError as exc:
            self.after(0, lambda: self._show_openfda_error(str(exc)))
        except Exception as exc:
            logging.exception("openFDA lookup failed")
            self.after(0, lambda: self._show_openfda_error(str(exc)))

    def _show_openfda_error(self, detail):
        self.reference_lookup_button.configure(state="normal")
        self.reference_status.configure(text=I.t("openfda_error"))
        self._clear_reference_cards()
        ctk.CTkLabel(self.reference_cards, text=I.t("openfda_error_detail", detail=detail),
                     text_color=DANGER, justify="left", wraplength=680, anchor="w").pack(fill="x")

    def _reference_line(self, parent, title, value, kind):
        line = ctk.CTkFrame(parent, fg_color="transparent")
        line.pack(fill="x", padx=12, pady=4)
        background, foreground = REFERENCE_TAGS[kind]
        ctk.CTkLabel(line, text=title, width=122, height=24, corner_radius=12,
                     fg_color=background, text_color=foreground,
                     font=ctk.CTkFont(size=10, weight="bold"), anchor="center").pack(
                         side="left", padx=(0, 9), anchor="n")
        ctk.CTkLabel(line, text=value or I.t("openfda_not_stated"), text_color="#20312f",
                     font=ctk.CTkFont(size=11), justify="left", wraplength=510,
                     anchor="nw").pack(side="left", fill="x", expand=True)

    def _show_openfda_results(self, references, medicines):
        self.reference_lookup_button.configure(state="normal")
        self.reference_status.configure(text=I.t("openfda_found", n=len(references)))
        self._clear_reference_cards()
        found = {reference.medicine.casefold() for reference in references}
        for number, reference in enumerate(references, 1):
            card = ctk.CTkFrame(self.reference_cards, fg_color="#f8fcfb", border_color=LINE,
                                border_width=1, corner_radius=14)
            card.pack(fill="x", pady=5)
            heading = I.t("openfda_card_title", number=number, name=reference.scientific_name)
            ctk.CTkLabel(card, text=heading, text_color=ACCENT,
                         font=ctk.CTkFont(size=14, weight="bold"), anchor="w").pack(
                             fill="x", padx=12, pady=(10, 2))
            ctk.CTkLabel(card, text=I.t("openfda_label_name", name=reference.label_name),
                         text_color=MUTED, font=ctk.CTkFont(size=11), anchor="w").pack(
                             fill="x", padx=12, pady=(0, 5))
            self._reference_line(card, I.t("label_contraindications"),
                                 openfda.concise(reference.contraindications), "contraindication")
            self._reference_line(card, I.t("label_pregnancy"),
                                 openfda.concise(reference.pregnancy), "pregnancy")
            self._reference_line(card, I.t("label_interactions"),
                                 openfda.concise(reference.interactions), "interaction")
        missing = [medicine for medicine in medicines if medicine.casefold() not in found]
        if missing:
            ctk.CTkLabel(self.reference_cards, text=I.t("openfda_not_found", names=", ".join(missing)),
                         text_color=MUTED, justify="left", wraplength=680, anchor="w").pack(fill="x", pady=5)
        if not references:
            ctk.CTkLabel(self.reference_cards, text=I.t("openfda_no_matches"), text_color=MUTED,
                         justify="left", wraplength=680, anchor="w").pack(fill="x")

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
        fields = [
            ("dosage", I.t("dosage")), ("frequency", I.t("frequency")),
            ("duration", I.t("duration")), ("notes", I.t("notes")),
        ]
        fields = [(name, label) for name, label in fields
                  if any(getattr(drug, name) for drug in drugs)]
        headers = ["#", I.t("drug")] + [label for _, label in fields]
        for column in range(len(headers)):
            self.word_preview_body.grid_columnconfigure(column, weight=2 if column == 1 else 1)
        for column, header in enumerate(headers):
            ctk.CTkLabel(self.word_preview_body, text=header, text_color=ACCENT,
                         fg_color=ACCENT_SOFT, font=ctk.CTkFont(size=11, weight="bold"),
                         anchor="w").grid(row=0, column=column, sticky="ew", padx=1, pady=1)
        for row_number, drug in enumerate(drugs, 1):
            if drug.generic_name and drug.brand_name:
                name = f"{drug.generic_name}\n({drug.brand_name})"
            else:
                name = drug.generic_name or drug.brand_name
            values = [str(row_number), name] + [getattr(drug, field) for field, _ in fields]
            for column, value in enumerate(values):
                ctk.CTkLabel(self.word_preview_body, text=value, text_color="#1a1a1a",
                             fg_color="white", font=ctk.CTkFont(size=12), anchor="w",
                             justify="left", wraplength=240).grid(
                                 row=row_number, column=column, sticky="nsew", padx=1, pady=1)

    def on_paper_change(self, value):
        cfg.config.paper_size = value

    def on_lang_change(self, value):
        cfg.config.language = value
        I.set_lang(value)
        self.reload_texts()

    def create_backup(self):
        path = filedialog.asksaveasfilename(
            title=I.t("backup"), defaultextension=".rxbackup",
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

    def on_profile_change(self, value):
        if value == "+ New profile…":
            name = simpledialog.askstring(APP_TITLE, "New profile name:", parent=self)
            if not name:
                self.profile_var.set(cfg.config.get("active_profile", "Default"))
                return
            cfg.config.save_profile(name, {k: v.get().strip() for k, v in self.doctor_vars.items()})
            self.profile_menu.configure(values=cfg.config.profile_names() + ["+ New profile…"])
            self.profile_var.set(name)
            return
        doctor = cfg.config.use_profile(value)
        for key, variable in self.doctor_vars.items():
            variable.set(doctor.get(key, ""))
        self._refresh_specialty_values(doctor.get("specialty", ""))

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
    def import_db(self):
        path = filedialog.askopenfilename(
            title=I.t("import_db"),
            filetypes=[("Drug database", "*.csv;*.xlsx;*.xls"),
                       ("CSV files", "*.csv"),
                       ("Excel files", "*.xlsx"),
                       ("Legacy Excel files", "*.xls"),
                       ("All files", "*.*")])
        if not path:
            return
        if not messagebox.askyesno(I.t("replace_database_title"),
                                   I.t("replace_database_confirm")):
            return
        try:
            total = self.db.import_file(path, replace=True)
            cfg.config.drug_db_path = str(self.db.path)
            self.db_label.configure(text=I.t("db_count", n=len(self.db.drugs)))
            self.refresh_class_overview()
            self.refresh_favorites_page()
            messagebox.showinfo(I.t("import_db"), I.t("msg_imported", n=total))
        except Exception as e:
            messagebox.showerror(I.t("import_db"), str(e))

    def remove_db(self):
        if not messagebox.askyesno(I.t("remove_db"), I.t("remove_db_confirm")):
            return
        try:
            self.db.clear()
            self.db_label.configure(text=I.t("db_count", n=0))
            self.refresh_class_overview()
            self.refresh_favorites_page()
            messagebox.showinfo(I.t("remove_db"), I.t("msg_db_removed"))
        except Exception as exc:
            logging.exception("Could not remove local drug database")
            messagebox.showerror(I.t("remove_db"), str(exc))

    def export_db(self):
        path = filedialog.asksaveasfilename(
            title=I.t("export_db"), defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        try:
            self.db.export_csv(path)
            messagebox.showinfo(I.t("export_db"), I.t("msg_exported_db", path=path))
        except Exception as e:
            messagebox.showerror(I.t("export_db"), str(e))

    # -- output --------------------------------------------------------------
    def _qr(self, rx):
        try:
            url = qu.build_qr_url(rx)
            info = qu.qr_info(url)
        except Exception as exc:
            logging.exception("Could not create QR code")
            messagebox.showerror(APP_TITLE, f"Could not create signed QR code:\n{exc}")
            return None
        if info["qr_version"] > 25:
            answer = messagebox.askyesno(
                APP_TITLE,
                f"This QR is dense (version {info['qr_version']}). It may not scan reliably when printed. Continue?")
            if not answer:
                return None
        return qu.make_qr_image(url)

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
        rx = self.collect()
        if not self._validate(rx, I.t("preview")):
            return None
        qr = self._qr(rx)
        if qr is None:
            return None
        if path_pdf:
            pdfgen.generate_prescription_pdf(rx, path_pdf,
                                            paper_size=self.paper_var.get(), qr_pil_image=qr)
        if path_docx:
            pdfgen.generate_prescription_docx(rx, path_docx, qr_pil_image=qr)
        return rx

    def preview(self):
        path = os.path.join(tempfile.gettempdir(), f"rx_preview_{uuid.uuid4().hex}.pdf")
        if not self._build_full(path_pdf=path):
            return
        try:
            os.startfile(path)
        except Exception:
            webbrowser.open(path)

    def print_pdf(self):
        path = os.path.join(tempfile.gettempdir(), f"rx_print_{uuid.uuid4().hex}.pdf")
        if not self._build_full(path_pdf=path):
            return
        try:
            os.startfile(path, "print")
        except Exception:
            os.startfile(path)

    def export_word(self):
        path = filedialog.asksaveasfilename(
            title=I.t("export_word"), defaultextension=".docx",
            filetypes=[("Word documents", "*.docx")])
        if not path:
            return
        if not self._build_full(path_docx=path):
            return
        messagebox.showinfo(I.t("export_word"), I.t("msg_exported_word", path=path))
        try:
            os.startfile(path)
        except Exception:
            webbrowser.open(path)

    def export_label(self):
        rx = self.collect()
        if not self._validate(rx, I.t("export_compact")):
            return
        path = filedialog.asksaveasfilename(
            title=I.t("export_compact"), defaultextension=".docx",
            filetypes=[("Word documents", "*.docx")])
        if not path:
            return
        qr = self._qr(rx)
        if qr is None:
            return
        pdfgen.generate_medication_label_docx(rx, path, qr_pil_image=qr)
        messagebox.showinfo(I.t("export_compact"), I.t("msg_exported_compact", path=path))
        try:
            os.startfile(path)
        except Exception:
            webbrowser.open(path)

    def open_settings(self):
        SettingsWindow(self)


class SettingsWindow(ctk.CTkToplevel):
    """Organized application settings with category navigation and one save flow."""

    SECTIONS = (
        ("general", "⚒", "General"),
        ("clinic", "✚", "Clinic identity"),
        ("documents", "▤", "Documents"),
        ("qr", "▦", "QR verification"),
        ("database", "⌬", "Drug database"),
        ("gemini", "✦", "Gemini reference"),
        ("security", "▣", "Backup & security"),
    )

    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.title(I.t("set_title"))
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = min(980, max(820, screen_width - 120))
        height = min(680, max(580, screen_height - 150))
        x = max(20, (screen_width - width) // 2)
        y = max(20, (screen_height - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(820, 580)
        self.configure(fg_color=BG)
        self.transient(master)
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.bind("<Escape>", lambda _event: self.cancel())
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        clinic = cfg.config.get_clinic()
        self.viewer_var = tk.StringVar(value=cfg.config.viewer_base_url)
        self.clinic_name_var = tk.StringVar(value=clinic.get("name", ""))
        self.clinic_address_var = tk.StringVar(value=clinic.get("address", ""))
        self.clinic_phone_var = tk.StringVar(value=clinic.get("phone", ""))
        self.logo_var = tk.StringVar(value=clinic.get("logo_path", ""))
        self.paper_var = tk.StringVar(value=cfg.config.paper_size)
        self.language_var = tk.StringVar(
            value="العربية" if cfg.config.language == "ar" else "English")
        self.gemini_enabled_var = tk.BooleanVar(value=cfg.config.gemini_enabled)
        self.gemini_key_var = tk.StringVar(value=cfg.config.gemini_api_key)
        self._active_section = "general"
        self._nav_buttons = {}
        self._settings_nav_icons = {}
        self._pages = {}
        self._build_header()
        self._build_workspace()
        self._build_footer()
        self._build_pages()
        self._show_section("general")
        self._tracked_variables = (
            self.viewer_var, self.clinic_name_var, self.clinic_address_var,
            self.clinic_phone_var, self.logo_var, self.paper_var,
            self.language_var, self.gemini_enabled_var, self.gemini_key_var,
        )
        self._saved_snapshot = self._snapshot()
        for variable in self._tracked_variables:
            variable.trace_add("write", self._mark_dirty)
        self.after(100, self.grab_set)

    def _build_header(self):
        header = ctk.CTkFrame(self, height=76, fg_color=CARD, corner_radius=0,
                              border_width=0)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        self._settings_header_icon = _glyph_icon("⚒", "white", 27)
        ctk.CTkLabel(
            header, text="", image=self._settings_header_icon,
            width=42, height=42, corner_radius=12, fg_color=ACCENT).pack(
                side="left", padx=(24, 12), pady=17)
        ctk.CTkLabel(
            header, text=I.t("set_title"), text_color="#153a38",
            font=ctk.CTkFont(size=26, weight="bold")).pack(side="left")
        ctk.CTkLabel(
            header, text=I.t("settings_encrypted"), fg_color=ACCENT_SOFT,
            text_color=ACCENT, corner_radius=14,
            font=ctk.CTkFont(size=13, weight="bold"), width=108,
            height=30).pack(side="right", padx=24)

    def _build_workspace(self):
        workspace = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        workspace.grid(row=1, column=0, sticky="nsew")
        workspace.grid_rowconfigure(0, weight=1)
        workspace.grid_columnconfigure(1, weight=1)

        self.sidebar = ctk.CTkFrame(
            workspace, width=240, fg_color="#edf5f3", corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsw", padx=(0, 1))
        self.sidebar.grid_propagate(False)
        ctk.CTkLabel(
            self.sidebar, text=I.t("settings_categories"), text_color=MUTED,
            font=ctk.CTkFont(size=12, weight="bold"), anchor="w").pack(
                fill="x", padx=22, pady=(24, 10))
        for key, symbol, _label in self.SECTIONS:
            icon = _glyph_icon(symbol, ICON_BLUE, 23)
            self._settings_nav_icons[key] = icon
            button = ctk.CTkButton(
                self.sidebar,
                text=I.t('settings_' + key) if key != 'security' else I.t('settings_backup_security'),
                image=icon, compound="left",
                anchor="w", height=48, corner_radius=10, fg_color="transparent",
                hover_color=ACCENT_SOFT, text_color=ACCENT,
                font=ctk.CTkFont(size=16),
                command=lambda section=key: self._show_section(section))
            button.pack(fill="x", padx=12, pady=3)
            self._nav_buttons[key] = button
        ctk.CTkLabel(
            self.sidebar, text=I.t("local_encrypted", version=cfg.APP_VERSION),
            text_color=MUTED, justify="left", anchor="w",
            font=ctk.CTkFont(size=11)).pack(
                side="bottom", fill="x", padx=22, pady=22)

        self.content_host = ctk.CTkFrame(
            workspace, fg_color=CARD, corner_radius=14, border_width=1,
            border_color=LINE)
        self.content_host.grid(row=0, column=1, sticky="nsew", padx=20, pady=18)
        self.content_host.grid_rowconfigure(0, weight=1)
        self.content_host.grid_columnconfigure(0, weight=1)

    def _build_footer(self):
        footer = ctk.CTkFrame(self, height=72, fg_color=CARD, corner_radius=0,
                              border_width=0)
        footer.grid(row=2, column=0, sticky="ew")
        footer.grid_propagate(False)
        self.footer_status = ctk.CTkLabel(
            footer, text=I.t("settings_no_changes"), text_color=MUTED,
            font=ctk.CTkFont(size=13), anchor="w")
        self.footer_status.pack(side="left", padx=24)
        ctk.CTkButton(
            footer, text=I.t("settings_restore_defaults"), width=150,
            height=40, fg_color="transparent", hover_color=ACCENT_SOFT,
            text_color=MUTED, border_width=1, border_color=LINE,
            command=self.restore_defaults).pack(side="left", padx=4)
        ctk.CTkButton(
            footer, text=I.t("settings_cancel"), width=100, height=40,
            fg_color=CARD, hover_color=ACCENT_SOFT, text_color=ACCENT,
            border_width=1, border_color=LINE, command=self.cancel).pack(
                side="right", padx=(6, 24))
        ctk.CTkButton(
            footer, text=I.t("settings_save_changes"), width=150, height=40,
            fg_color=GOOD, hover_color=GOOD_HOVER, command=self.save).pack(
                side="right", padx=6)

    def _new_page(self, key, title, subtitle):
        page = ctk.CTkFrame(
            self.content_host, fg_color=CARD, corner_radius=12)
        page.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        page.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            page, text=title, text_color="#153a38", anchor="w",
            font=ctk.CTkFont(size=PAGE_TITLE_FONT_SIZE, weight="bold")).grid(
                row=0, column=0, sticky="ew", padx=30, pady=(26, 4))
        if subtitle:
            ctk.CTkLabel(
                page, text=subtitle, text_color=MUTED, anchor="w",
                justify="left", wraplength=760,
                font=ctk.CTkFont(size=14)).grid(
                    row=1, column=0, sticky="ew", padx=30, pady=(0, 20))
        self._pages[key] = page
        return page

    def _card(self, parent, row, title=None):
        card = ctk.CTkFrame(
            parent, fg_color="#fbfdfd", corner_radius=14, border_width=1,
            border_color=LINE)
        card.grid(row=row, column=0, sticky="ew", padx=30, pady=(0, 16))
        card.grid_columnconfigure(0, weight=1)
        if title:
            ctk.CTkLabel(
                card, text=title, text_color="#244e4b", anchor="w",
                font=ctk.CTkFont(size=17, weight="bold")).grid(
                    row=0, column=0, sticky="ew", padx=20, pady=(18, 10))
        return card

    @staticmethod
    def _entry(parent, variable, row, label, show=None):
        ctk.CTkLabel(
            parent, text=label, text_color=MUTED, anchor="w",
            font=ctk.CTkFont(size=14, weight="bold")).grid(
                row=row, column=0, sticky="ew", padx=20, pady=(10, 4))
        entry = ctk.CTkEntry(
            parent, textvariable=variable, height=44, corner_radius=9,
            border_color=LINE, show=show,
            font=ctk.CTkFont(size=15))
        entry.grid(row=row + 1, column=0, sticky="ew", padx=20, pady=(0, 10))
        return entry

    def _build_pages(self):
        self._build_general_page()
        self._build_clinic_page()
        self._build_documents_page()
        self._build_qr_page()
        self._build_database_page()
        self._build_gemini_page()
        self._build_security_page()

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
        ctk.CTkOptionMenu(
            card, values=["English", "العربية"], variable=self.language_var,
            height=44, corner_radius=9, fg_color=ACCENT_SOFT,
            button_color=ACCENT, button_hover_color=ACCENT_HOVER,
            text_color="#244e4b", font=ctk.CTkFont(size=15)).grid(
                row=2, column=0, sticky="ew", padx=(20, 8), pady=(0, 20))
        ctk.CTkLabel(
            card, text=f"Rx Prescription Printer  ·  {cfg.APP_VERSION}",
            height=44, corner_radius=9, fg_color=ACCENT_SOFT,
            text_color=ACCENT, anchor="w", padx=14,
            font=ctk.CTkFont(size=15, weight="bold")).grid(
                row=2, column=1, sticky="ew", padx=(8, 20), pady=(0, 20))

    def _build_clinic_page(self):
        page = self._new_page(
            "clinic", I.t("settings_clinic"), I.t("settings_clinic_tip"))
        card = self._card(page, 2, I.t("settings_clinic_contact"))
        card.grid_columnconfigure((0, 1, 2), weight=1)
        contact_fields = (
            (self.clinic_name_var, I.t("settings_clinic_name")),
            (self.clinic_address_var, I.t("settings_clinic_address")),
            (self.clinic_phone_var, I.t("settings_clinic_phone")),
        )
        for column, (variable, label) in enumerate(contact_fields):
            left_pad = 20 if column == 0 else 6
            right_pad = 20 if column == 2 else 6
            ctk.CTkLabel(
                card, text=label, text_color=MUTED, anchor="w",
                font=ctk.CTkFont(size=13, weight="bold")).grid(
                    row=1, column=column, sticky="ew",
                    padx=(left_pad, right_pad), pady=(4, 4))
            ctk.CTkEntry(
                card, textvariable=variable, height=42, corner_radius=9,
                border_color=LINE, font=ctk.CTkFont(size=14)).grid(
                    row=2, column=column, sticky="ew",
                    padx=(left_pad, right_pad), pady=(0, 18))
        logo_card = self._card(page, 3, I.t("settings_logo"))
        logo_card.grid_columnconfigure(0, weight=1)
        self.logo_entry = ctk.CTkEntry(
            logo_card, textvariable=self.logo_var, height=44, corner_radius=9,
            border_color=LINE, font=ctk.CTkFont(size=14))
        self.logo_entry.grid(row=1, column=0, sticky="ew", padx=(20, 8), pady=(0, 18))
        ctk.CTkButton(
            logo_card, text=I.t("settings_browse"), width=100, height=44,
            fg_color=ACCENT_SOFT, hover_color=LINE, text_color=ACCENT,
            command=self.choose_logo).grid(row=1, column=1, padx=(0, 8), pady=(0, 18))
        ctk.CTkButton(
            logo_card, text=I.t("settings_remove"), width=100, height=44,
            fg_color=CARD, hover_color="#fae9e9", text_color=DANGER,
            border_width=1, border_color=LINE,
            command=lambda: self.logo_var.set("")).grid(
                row=1, column=2, padx=(0, 20), pady=(0, 18))

    def _build_documents_page(self):
        page = self._new_page(
            "documents", I.t("settings_documents"), I.t("settings_documents_tip"))
        card = self._card(page, 2, I.t("settings_page_format"))
        ctk.CTkLabel(
            card, text=I.t("paper"), text_color=MUTED, anchor="w",
            font=ctk.CTkFont(size=14, weight="bold")).grid(
                row=1, column=0, sticky="ew", padx=20, pady=(6, 4))
        ctk.CTkOptionMenu(
            card, values=list(cfg.PAPER_SIZES.keys()), variable=self.paper_var,
            height=44, corner_radius=9, fg_color=ACCENT_SOFT,
            button_color=ACCENT, button_hover_color=ACCENT_HOVER,
            text_color="#244e4b", font=ctk.CTkFont(size=15)).grid(
                row=2, column=0, sticky="ew", padx=20, pady=(0, 12))
        ctk.CTkLabel(
            card, text=I.t("settings_document_note"), text_color=MUTED,
            justify="left", anchor="w", wraplength=720,
            font=ctk.CTkFont(size=14)).grid(
                row=3, column=0, sticky="ew", padx=20, pady=(0, 20))

    def _build_qr_page(self):
        page = self._new_page(
            "qr", I.t("settings_qr"), I.t("settings_qr_tip"))
        card = self._card(page, 2, I.t("settings_verification_viewer"))
        self._entry(card, self.viewer_var, 1, I.t("set_viewer"))
        ctk.CTkLabel(
            card, text=I.t("set_tip"), text_color=MUTED, justify="left",
            anchor="w", wraplength=720, font=ctk.CTkFont(size=13)).grid(
                row=3, column=0, sticky="ew", padx=20, pady=(0, 12))
        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 20))
        ctk.CTkButton(
            actions, text=I.t("settings_open_viewer"), height=40,
            fg_color=ACCENT_SOFT, hover_color=LINE, text_color=ACCENT,
            command=self.open_viewer).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            actions, text=I.t("settings_copy_key"), height=40,
            fg_color=CARD, hover_color=ACCENT_SOFT, text_color=ACCENT,
            border_width=1, border_color=LINE,
            command=self.copy_verification_key).pack(side="left")

    def _build_database_page(self):
        page = self._new_page(
            "database", I.t("settings_database"), I.t("settings_database_tip"))
        card = self._card(page, 2, I.t("settings_current_database"))
        status = ctk.CTkFrame(card, fg_color="transparent")
        status.grid(row=1, column=0, sticky="ew", padx=20, pady=(4, 8))
        status.grid_columnconfigure(0, weight=1)
        self.database_path_label = ctk.CTkLabel(
            status, text=str(self.master.db.path), text_color="#244e4b",
            anchor="w", justify="left", wraplength=610,
            font=ctk.CTkFont(size=14))
        self.database_path_label.grid(row=0, column=0, sticky="ew")
        self.database_count_label = ctk.CTkLabel(
            status, text=str(len(self.master.db.drugs)), fg_color=ACCENT_SOFT,
            text_color=ACCENT, corner_radius=14, width=58, height=30,
            font=ctk.CTkFont(size=14, weight="bold"))
        self.database_count_label.grid(row=0, column=1, padx=(12, 0))
        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=20, pady=(8, 20))
        ctk.CTkButton(
            actions, text=I.t("import_db"), height=40,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self.import_database).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            actions, text=I.t("export_db"), height=40,
            fg_color=ACCENT_SOFT, hover_color=LINE, text_color=ACCENT,
            command=self.master.export_db).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            actions, text=I.t("remove_db"), height=40,
            fg_color=CARD, hover_color="#fae9e9", text_color=DANGER,
            border_width=1, border_color=LINE,
            command=self.remove_database).pack(side="left")

    def _build_gemini_page(self):
        page = self._new_page(
            "gemini", I.t("settings_gemini"), I.t("settings_gemini_tip"))
        card = self._card(page, 2, I.t("gemini_settings"))
        ctk.CTkCheckBox(
            card, text=I.t("gemini_enable"), variable=self.gemini_enabled_var,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=15, weight="bold")).grid(
                row=1, column=0, sticky="w", padx=20, pady=(6, 12))
        self.gemini_key_entry = self._entry(
            card, self.gemini_key_var, 2, I.t("gemini_api_key"), show="•")
        ctk.CTkLabel(
            card, text=I.t("gemini_api_key_tip"), text_color=MUTED,
            justify="left", anchor="w", wraplength=720,
            font=ctk.CTkFont(size=13)).grid(
                row=4, column=0, sticky="ew", padx=20, pady=(0, 8))
        self.gemini_status_label = ctk.CTkLabel(
            card, text="", text_color=MUTED, justify="left", anchor="w",
            wraplength=720, font=ctk.CTkFont(size=13))
        self.gemini_status_label.grid(row=5, column=0, sticky="ew", padx=20, pady=(0, 8))
        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=6, column=0, sticky="ew", padx=20, pady=(0, 20))
        self.gemini_test_button = ctk.CTkButton(
            actions, text=I.t("gemini_test"), height=40,
            fg_color=ACCENT_SOFT, hover_color=LINE, text_color=ACCENT,
            command=self.test_gemini_connection)
        self.gemini_test_button.pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            actions, text=I.t("gemini_remove_key"), height=40,
            fg_color=CARD, hover_color="#fae9e9", text_color=DANGER,
            border_width=1, border_color=LINE,
            command=self.remove_gemini_key).pack(side="left")

    def _build_security_page(self):
        page = self._new_page(
            "security", I.t("settings_backup_security"),
            I.t("settings_backup_tip"))
        card = self._card(page, 2, I.t("settings_backup_restore"))
        ctk.CTkLabel(
            card, text=I.t("settings_backup_note"), text_color=MUTED,
            justify="left", anchor="w", wraplength=720,
            font=ctk.CTkFont(size=14)).grid(
                row=1, column=0, sticky="ew", padx=20, pady=(6, 14))
        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 20))
        ctk.CTkButton(
            actions, text=I.t("settings_create_backup"), height=40,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self.master.create_backup).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            actions, text=I.t("settings_restore_backup"), height=40,
            fg_color=ACCENT_SOFT, hover_color=LINE, text_color=ACCENT,
            command=self.master.restore_backup).pack(side="left")
        protection = self._card(page, 3, I.t("settings_local_protection"))
        ctk.CTkLabel(
            protection, text=I.t("settings_protection_note"), text_color=MUTED,
            justify="left", anchor="w", wraplength=720,
            font=ctk.CTkFont(size=14)).grid(
                row=1, column=0, sticky="ew", padx=20, pady=(4, 20))

    def _show_section(self, section):
        self._active_section = section
        for key, page in self._pages.items():
            if key == section:
                page.grid()
            else:
                page.grid_remove()
        for key, button in self._nav_buttons.items():
            button.configure(
                fg_color=ACCENT_SOFT if key == section else "transparent",
                text_color=ACCENT if key == section else "#355b58",
                font=ctk.CTkFont(size=16, weight="bold" if key == section else "normal"))

    def _snapshot(self):
        return tuple(variable.get() for variable in self._tracked_variables) if hasattr(
            self, "_tracked_variables") else ()

    def _mark_dirty(self, *_args):
        if not hasattr(self, "footer_status"):
            return
        dirty = self._snapshot() != self._saved_snapshot
        self.footer_status.configure(
            text=I.t("settings_unsaved") if dirty else I.t("settings_no_changes"),
            text_color=WARNING if dirty else MUTED)

    def cancel(self):
        if hasattr(self, "_saved_snapshot") and self._snapshot() != self._saved_snapshot:
            if not messagebox.askyesno(
                    I.t("set_title"), I.t("settings_discard_confirm"), parent=self):
                return
        self.destroy()

    def save(self):
        viewer_url = self.viewer_var.get().strip().rstrip("/")
        logo_path = self.logo_var.get().strip()
        if not viewer_url.lower().startswith(("https://", "http://")):
            self._show_section("qr")
            messagebox.showerror(
                I.t("set_title"), I.t("settings_invalid_viewer"), parent=self)
            return
        if logo_path and not Path(logo_path).is_file():
            self._show_section("clinic")
            messagebox.showerror(
                I.t("set_title"), I.t("settings_invalid_logo"), parent=self)
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
            cfg.config.set_clinic(name=self.clinic_name_var.get().strip(),
                                  address=self.clinic_address_var.get().strip(),
                                  phone=self.clinic_phone_var.get().strip(),
                                  logo_path=logo_path)
            cfg.config.viewer_base_url = viewer_url
            cfg.config.paper_size = self.paper_var.get()
            cfg.config.language = language
            cfg.config.set_gemini(gemini_key, self.gemini_enabled_var.get())
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
        messagebox.showinfo(I.t("set_title"), I.t("settings_saved"), parent=self.master)

    def restore_defaults(self):
        if not messagebox.askyesno(
                I.t("settings_restore_defaults"), I.t("settings_defaults_confirm"),
                parent=self):
            return
        if self._active_section == "general":
            self.language_var.set("English")
        elif self._active_section == "clinic":
            self.clinic_name_var.set("")
            self.clinic_address_var.set("")
            self.clinic_phone_var.set("")
            self.logo_var.set("")
        elif self._active_section == "documents":
            self.paper_var.set(cfg.DEFAULT_PAPER)
        elif self._active_section == "qr":
            self.viewer_var.set(cfg.DEFAULT_VIEWER_BASE)
        elif self._active_section == "gemini":
            self.gemini_enabled_var.set(False)
            self.gemini_key_var.set("")
        else:
            messagebox.showinfo(
                I.t("settings_restore_defaults"), I.t("settings_no_defaults"), parent=self)

    def open_viewer(self):
        url = self.viewer_var.get().strip()
        if url.lower().startswith(("http://", "https://")):
            webbrowser.open(url)
        else:
            webbrowser.open(str(cfg.PROJECT_DIR / "viewer.html"))

    def copy_verification_key(self):
        import json
        try:
            key = json.dumps(qu.verification_key(), separators=(",", ":"))
            self.clipboard_clear()
            self.clipboard_append(key)
            messagebox.showinfo(
                APP_TITLE, I.t("settings_key_copied"), parent=self)
        except Exception as exc:
            logging.exception("Could not export verification key")
            messagebox.showerror(APP_TITLE, str(exc), parent=self)

    def choose_logo(self):
        path = filedialog.askopenfilename(title="Clinic logo",
                                          filetypes=[("Images", "*.png;*.jpg;*.jpeg"),
                                                     ("All files", "*.*")])
        if path:
            self.logo_var.set(path)

    def import_database(self):
        self.master.import_db()
        self._sync_database_status()

    def remove_database(self):
        self.master.remove_db()
        self._sync_database_status()

    def _sync_database_status(self):
        try:
            self.database_path_label.configure(text=str(self.master.db.path))
            self.database_count_label.configure(text=str(len(self.master.db.drugs)))
        except Exception:
            pass

    def remove_gemini_key(self):
        self.gemini_key_var.set("")
        self.gemini_enabled_var.set(False)
        gemini_drug.clear_cache()
        self.gemini_status_label.configure(
            text=I.t("gemini_key_removed"), text_color=ACCENT)

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
            self.gemini_status_label.configure(
                text=I.t("gemini_test_ok") if success else detail,
                text_color=ACCENT if success else DANGER)
        except Exception:
            pass


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
            fg_color=ACCENT, hover_color=ACCENT_HOVER).pack(
                anchor="w", padx=18, pady=(0, 10))
        ctk.CTkLabel(self, text=I.t("gemini_api_key"), text_color=MUTED,
                     anchor="w").pack(fill="x", padx=18, pady=(0, 3))
        self.key_var = tk.StringVar(value=cfg.config.gemini_api_key)
        self.key_entry = ctk.CTkEntry(
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
        self.test_button = ctk.CTkButton(
            actions, text=I.t("gemini_test"), height=ACTION_HEIGHT,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=self.test_connection)
        self.test_button.pack(side="left", padx=(0, 6))
        ctk.CTkButton(
            actions, text=I.t("gemini_remove_key"), height=ACTION_HEIGHT,
            fg_color=DANGER, hover_color=DANGER_HOVER, command=self.remove_key).pack(
                side="left", padx=6)
        ctk.CTkButton(
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
