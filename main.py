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
import csv
import concurrent.futures
import functools
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


@functools.lru_cache(maxsize=48)
def _ui_font(size=12, weight="normal", family="Segoe UI"):
    """Reuse immutable UI font objects instead of allocating them per result row."""
    return ctk.CTkFont(family=family, size=size, weight=weight)


class DrugRow(ctk.CTkFrame):
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
        self.drag_handle = ctk.CTkButton(
            header_actions, text="⠿", width=36, height=28, corner_radius=7,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, font=ctk.CTkFont(size=16, weight="bold"),
            command=self._show_move_menu)
        self.drag_handle.pack(side="left", padx=2)
        self.drag_handle.bind("<ButtonPress-1>", self._drag_start)
        self.drag_handle.bind("<B1-Motion>", self._drag_motion)
        self.drag_handle.bind("<ButtonRelease-1>", self._drag_end)
        self.drag_handle.bind("<Button-3>", self._show_move_menu)
        self.drag_handle.bind("<Shift-F10>", self._show_move_menu)
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
        menu = tk.Menu(
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
        self._schedule_autocomplete(q, "scientific")

    def _on_trade_type(self, event=None):
        # The deferred worker uses db.search_prescribable for this brand/trade field.
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
        try:
            cfg.config.maybe_create_automatic_backup()
        except Exception:
            logging.exception("Automatic backup could not be created")
        self.patient_history = PatientHistory()
        self.rows: List[DrugRow] = []
        self._executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=3, thread_name_prefix="rx-worker")
        self._background_tasks = 0
        self._document_lock = threading.Lock()
        self._loaded_pages = set()
        self._classification_cache = None
        self._latest_query_tokens = {}
        self._favorite_ac_job = None
        self._favorite_ac_token = 0
        self._favorite_render_limit = 60
        self._favorite_filter_signature = None
        self._word_preview_job = None
        self._favorite_refresh_job = None
        self._class_search_job = None
        self.word_preview_visible = False
        self._build_ui()

    def _build_ui(self):
        self.paper_var = tk.StringVar(value=cfg.config.paper_size)
        self.lang_var = tk.StringVar(value=cfg.config.language)
        self.profile_var = tk.StringVar(value=cfg.config.get("active_profile", "Default"))

        # Dashboard: persistent navigation on the left; one data-entry page at a time.
        self.workspace = ctk.CTkFrame(self, fg_color="transparent")
        self.workspace.pack(fill="both", expand=True, padx=4, pady=(2, 4))
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
        self._add_page_button("treatment_templates", I.t("treatment_templates"))
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
        ctk.CTkButton(self.action, text=I.t("preview"), width=120, height=ACTION_HEIGHT,
                      corner_radius=9, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.preview).pack(side="left", padx=6)
        ctk.CTkButton(self.action, text=I.t("clear"), width=100, height=ACTION_HEIGHT,
                      corner_radius=9, fg_color=CARD, text_color=MUTED,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.clear_all).pack(side="left", padx=6)
        self.busy_label = ctk.CTkLabel(
            self.action, text="", text_color=MUTED,
            font=ctk.CTkFont(size=10), anchor="w")
        self.busy_label.pack(side="left", padx=8)
        # right cluster (primary actions)
        ctk.CTkButton(self.action, text=I.t("export_compact"), width=160, height=ACTION_HEIGHT,
                      corner_radius=9, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.export_label).pack(side="right", padx=6)
        ctk.CTkButton(self.action, text=I.t("export_word"), width=120, height=ACTION_HEIGHT,
                      corner_radius=9, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.export_word).pack(side="right", padx=6)
        ctk.CTkButton(self.action, text=I.t("print"), width=110, height=ACTION_HEIGHT,
                      corner_radius=9, fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      font=ctk.CTkFont(weight="bold", size=13),
                      command=self.print_pdf).pack(side="right", padx=6)

        self.add_row()
        self.load_profile_into_ui()
        self.show_page("prescriber")

    def confirm_close(self):
        """Require an explicit confirmation before closing the desktop app."""
        if messagebox.askyesno(
                I.t("confirm_close_title"), I.t("confirm_close_message"), parent=self):
            self._executor.shutdown(wait=False, cancel_futures=True)
            self.destroy()

    def submit_background(self, worker, on_success=None, *, on_error=None,
                          label="Working…", silent=False):
        """Run blocking work without freezing Tk; callbacks always run on Tk."""
        if not silent:
            self._background_tasks += 1
            if hasattr(self, "busy_label"):
                self.busy_label.configure(text=label)
        future = self._executor.submit(worker)

        def completed(done):
            try:
                result = done.result()
                error = None
            except Exception as exc:
                result, error = None, exc

            def deliver():
                if not silent:
                    self._background_tasks = max(0, self._background_tasks - 1)
                    if hasattr(self, "busy_label") and self._background_tasks == 0:
                        self.busy_label.configure(text="")
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

            try:
                self.after(0, deliver)
            except (tk.TclError, RuntimeError):
                pass

        future.add_done_callback(completed)
        return future

    # -- forms ---------------------------------------------------------------
    def section(self, parent, title, accent=True):
        f = ctk.CTkFrame(parent, fg_color=CARD, border_color=LINE,
                         border_width=1, corner_radius=16)
        f.pack(fill="x", padx=2, pady=8)
        if accent:
            bar = ctk.CTkFrame(f, height=4, fg_color=ACCENT, corner_radius=0)
            bar.pack(fill="x", side="top")
        if title:
            ctk.CTkLabel(f, text=title, font=ctk.CTkFont(weight="bold", size=14),
                         text_color=ACCENT, anchor="w").pack(
                             anchor="w", padx=PAD, pady=(10, 4))
        return f

    def page_header(self, parent, title, subtitle):
        """A consistent title block makes every dashboard page easy to scan."""
        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(fill="x", padx=4, pady=(0, 3))
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
            "treatment_templates": ctk.CTkFrame(self.scroll, fg_color="transparent"),
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
        ctk.CTkButton(
            d, text=I.t("save_profile"), width=150, height=ACTION_HEIGHT,
            corner_radius=9, fg_color=GOOD, hover_color=GOOD_HOVER,
            command=self.save_profile).pack(
                anchor="w", padx=(PAD + 130, PAD), pady=(2, PAD))

        self.page_header(self.pages["patient"], I.t("patient_details"), "")
        p = self.section(self.pages["patient"], "")

        patient_fields = ctk.CTkFrame(p, fg_color="transparent")
        patient_fields.pack(fill="x", padx=PAD, pady=(10, 7))
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
        self.patient_name_entry = ctk.CTkEntry(
            name_col, textvariable=patient_name_binding.display_var,
            width=PATIENT_NAME_WIDTH, height=FIELD_HEIGHT,
            corner_radius=9, border_color=LINE,
            font=ctk.CTkFont(size=14), justify="left")
        patient_name_binding.attach(self.patient_name_entry)
        self.patient_name_entry.pack(fill="x")

        age_col = ctk.CTkFrame(patient_fields, fg_color="transparent")
        age_col.grid(row=0, column=1, sticky="ew", padx=6)
        ctk.CTkLabel(age_col, text=I.t("age"), text_color=MUTED,
                     font=ctk.CTkFont(size=11), anchor="w").pack(fill="x", pady=(0, 3))
        self.patient_age_entry = ctk.CTkEntry(
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
        self.patient_sex_menu = ctk.CTkOptionMenu(
            sex_col, values=sex_labels, width=PATIENT_SEX_WIDTH,
            height=FIELD_HEIGHT, corner_radius=9,
            fg_color=ACCENT_SOFT, text_color=ACCENT, button_color=ACCENT,
            button_hover_color=ACCENT_HOVER, dropdown_hover_color=ACCENT_SOFT,
            command=self._set_patient_sex)
        self.patient_sex_menu.set(I.t("sex_m"))
        self.patient_sex_menu.pack(anchor="w")

        patient_actions = ctk.CTkFrame(p, fg_color="transparent")
        patient_actions.pack(fill="x", padx=PAD, pady=(1, PAD))
        ctk.CTkButton(patient_actions, text=I.t("patient_action_new"), width=84,
                      height=ACTION_HEIGHT, command=self.new_patient).pack(
                          side="left", padx=(0, 4))
        ctk.CTkButton(patient_actions, text=I.t("patient_action_save"), width=84,
                      height=ACTION_HEIGHT, fg_color=GOOD, hover_color=GOOD_HOVER,
                      command=self.save_patient_history).pack(side="left", padx=4)
        ctk.CTkButton(patient_actions, text=I.t("patient_action_clear"), width=84,
                      height=ACTION_HEIGHT, fg_color=CARD, text_color=ACCENT,
                      border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
                      command=self.clear_patient_details).pack(side="left", padx=4)
        self.patient_delete_button = ctk.CTkButton(
            patient_actions, text=I.t("patient_action_delete"), width=84,
            height=ACTION_HEIGHT, fg_color=DANGER, hover_color=DANGER_HOVER,
            state="disabled", command=self.delete_selected_patient)
        self.patient_delete_button.pack(side="left", padx=4)

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
        self.patient_search_entry = ctk.CTkEntry(
            patient_finder, textvariable=patient_search_binding.display_var,
            width=PATIENT_SEARCH_WIDTH,
            height=FIELD_HEIGHT, placeholder_text=I.t("search_patients"),
            border_color=ACCENT, corner_radius=9, font=ctk.CTkFont(size=16),
            justify="left")
        patient_search_binding.attach(self.patient_search_entry)
        self.patient_search_entry.pack(anchor="w", pady=(0, 6))
        result_shell = ctk.CTkFrame(
            patient_finder, width=PATIENT_SEARCH_WIDTH,
            height=PATIENT_RESULTS_HEIGHT, fg_color="#f8fcfb",
            border_color=LINE, border_width=1, corner_radius=9)
        result_shell.pack_propagate(False)
        result_shell.pack(anchor="w")
        self.patient_history_list = tk.Listbox(result_shell, height=5,
                                               font=LIST_FONT, bg="#f8fcfb", fg="#1a302e",
                                               relief="flat", borderwidth=0, highlightthickness=0,
                                               selectbackground=ACCENT, selectforeground="white",
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
        self.patient_prescriptions_body = ctk.CTkFrame(
            prescriptions_panel, fg_color="transparent")
        self.patient_prescriptions_body.pack(fill="x")
        self._loaded_patient_id = ""
        self._expanded_prescription_ids = set()
        self._current_history_record = None
        self.patient_name_entry.bind("<Return>", lambda _event: self.patient_age_entry.focus_set())
        self.patient_age_entry.bind("<Return>", lambda _event: self.patient_sex_menu.focus_set())
        self.bind_all("<Control-s>", self._save_patient_shortcut)
        self._show_empty_prescriptions()

        self.page_header(self.pages["medications"], I.t("medication_entry"), "")
        self.word_preview_visible = False
        medication_toolbar = ctk.CTkFrame(
            self.pages["medications"], fg_color="transparent")
        medication_toolbar.pack(fill="x", padx=2, pady=(0, 4))
        ctk.CTkButton(
            medication_toolbar, text=I.t("add_drug"), width=120,
            height=ACTION_HEIGHT, corner_radius=9, fg_color=ACCENT,
            hover_color=ACCENT_HOVER, command=self.add_row).pack(side="left", padx=(0, 5))
        self.medication_favorites_button = ctk.CTkButton(
            medication_toolbar, text=I.t("starred_drugs"), width=140,
            height=ACTION_HEIGHT, corner_radius=9, fg_color=CARD,
            text_color=ACCENT, border_color=LINE, border_width=1,
            hover_color=ACCENT_SOFT,
            command=self.toggle_medication_favorite_picker)
        self.medication_favorites_button.pack(side="left")

        self.medication_favorite_panel = ctk.CTkFrame(
            self.pages["medications"], fg_color="#f8fcfb", border_color=LINE,
            border_width=1, corner_radius=12)
        favorite_picker_head = ctk.CTkFrame(
            self.medication_favorite_panel, fg_color="transparent")
        favorite_picker_head.pack(fill="x", padx=10, pady=(8, 5))
        self.medication_favorite_search_var = tk.StringVar()
        self.medication_favorite_search_var.trace_add(
            "write", lambda *_: self.refresh_medication_favorite_picker())
        self.medication_favorite_search_entry = ctk.CTkEntry(
            favorite_picker_head, textvariable=self.medication_favorite_search_var,
            placeholder_text=I.t("search_starred_drugs"), height=36,
            border_color=LINE, corner_radius=8)
        self.medication_favorite_search_entry.pack(
            side="left", fill="x", expand=True, padx=(0, 6))
        ctk.CTkButton(
            favorite_picker_head, text="×", width=34, height=32,
            fg_color="transparent", text_color=MUTED, hover_color=ACCENT_SOFT,
            command=self.toggle_medication_favorite_picker).pack(side="right")
        # Let the result container follow its cards instead of reserving a
        # fixed blank area when only one or two starred drugs are available.
        self.medication_favorite_results = ctk.CTkFrame(
            self.medication_favorite_panel, fg_color="transparent")
        self.medication_favorite_results.pack(fill="x", padx=8, pady=(0, 8))

        dr = self.section(self.pages["medications"], I.t("medications"))
        self.medications_section = dr
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
        self.word_preview_toggle = ctk.CTkButton(
            preview, text=I.t("show_word_preview"), width=140, height=32,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=self.toggle_word_preview)
        self.word_preview_toggle.pack(anchor="w", padx=PAD, pady=(0, 8))
        self.word_preview_body = ctk.CTkFrame(preview, fg_color="transparent")
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

        self.class_page_header = self.page_header(
            self.pages["drug_classes"], I.t("drug_classes"), "")
        self.class_overview = ctk.CTkFrame(self.pages["drug_classes"], fg_color="transparent")
        self.class_overview.pack(fill="both", expand=True)
        class_page = self.section(self.class_overview, "")
        top = ctk.CTkFrame(class_page, fg_color="transparent")
        top.pack(fill="x", padx=PAD, pady=(10, 6))
        self.class_breadcrumb = ctk.CTkFrame(top, fg_color="transparent")
        self.class_breadcrumb.pack(side="left", fill="x", expand=True)
        self.class_breadcrumb_group = ctk.CTkButton(
            self.class_breadcrumb, text="", height=30, width=40,
            fg_color="transparent", text_color=ACCENT,
            hover_color=ACCENT_SOFT, anchor="w",
            font=ctk.CTkFont(size=13, weight="bold"))
        self.class_breadcrumb_group.pack(side="left")
        self.class_breadcrumb_separator = ctk.CTkLabel(
            self.class_breadcrumb, text="›", width=22, text_color=MUTED,
            font=ctk.CTkFont(size=14, weight="bold"))
        self.class_breadcrumb_separator.pack(side="left")
        self.class_breadcrumb_detail = ctk.CTkButton(
            self.class_breadcrumb, text="", height=30, width=40,
            fg_color="transparent", text_color=ACCENT,
            hover_color=ACCENT_SOFT, anchor="w",
            font=ctk.CTkFont(size=13, weight="bold"))
        self.class_breadcrumb_detail.pack(side="left")
        ctk.CTkButton(top, text=I.t("manage_class_mappings"), height=ACTION_HEIGHT, width=180,
                      fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
                      hover_color=ACCENT_SOFT, command=self.show_class_mapping_editor).pack(
                          side="right", padx=(10, 0))
        ctk.CTkButton(top, text=I.t("show_all_detailed_classes"), height=ACTION_HEIGHT, width=210,
                      fg_color=CARD, text_color=ACCENT, border_width=1,
                      border_color=LINE, hover_color=ACCENT_SOFT,
                      command=self.show_all_detailed_classes).pack(side="right", padx=(10, 0))
        ctk.CTkButton(
            top, text=I.t("mapping_integrity"), height=ACTION_HEIGHT, width=150,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=self.show_mapping_integrity_report).pack(
                side="right", padx=(10, 0))

        self.class_search_var = tk.StringVar()
        self.class_search_var.trace_add(
            "write", lambda *_: self._schedule_class_browser_refresh())
        ctk.CTkEntry(
            class_page, textvariable=self.class_search_var, height=FIELD_HEIGHT,
            placeholder_text=I.t("search_classes_medicines"),
            border_color=LINE, corner_radius=9).pack(
                fill="x", padx=PAD, pady=(0, 7))
        self.class_filters = ctk.CTkFrame(class_page, fg_color="transparent")
        self.class_filters.pack(fill="x", padx=PAD, pady=(0, 7))
        self.class_filter_group_labels = {
            I.t("filter_all_groups"): "",
            **{I.t("class_" + code): code for code in classes.GROUPS},
        }
        self.class_group_filter_var = tk.StringVar(value=I.t("filter_all_groups"))
        self.class_detail_filter_var = tk.StringVar(value=I.t("filter_all_classes"))
        self.class_name_filter_var = tk.StringVar(value=I.t("filter_all_names"))
        self.class_unclassified_filter_var = tk.BooleanVar(value=False)
        self.class_group_filter_menu = ctk.CTkOptionMenu(
            self.class_filters, values=list(self.class_filter_group_labels),
            variable=self.class_group_filter_var, width=190, height=34,
            fg_color=ACCENT_SOFT, text_color=ACCENT, button_color=ACCENT,
            button_hover_color=ACCENT_HOVER,
            command=self.class_group_filter_changed)
        self.class_group_filter_menu.pack(side="left", padx=(0, 5))
        self.class_detail_filter_menu = ctk.CTkOptionMenu(
            self.class_filters, values=[I.t("filter_all_classes")],
            variable=self.class_detail_filter_var, width=190, height=34,
            fg_color=ACCENT_SOFT, text_color=ACCENT, button_color=ACCENT,
            button_hover_color=ACCENT_HOVER,
            command=lambda _value: self.refresh_class_browser())
        self.class_detail_filter_menu.pack(side="left", padx=5)
        self.class_name_filter_menu = ctk.CTkOptionMenu(
            self.class_filters,
            values=[I.t("filter_all_names"), I.t("filter_brand_name"),
                    I.t("filter_scientific_name")],
            variable=self.class_name_filter_var, width=145, height=34,
            fg_color=CARD, text_color=ACCENT, button_color=ACCENT,
            button_hover_color=ACCENT_HOVER,
            command=lambda _value: self.refresh_class_browser())
        self.class_name_filter_menu.pack(side="left", padx=5)
        self.class_unclassified_filter_checkbox = ctk.CTkCheckBox(
            self.class_filters, text=I.t("filter_unclassified"),
            variable=self.class_unclassified_filter_var, width=110,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self.class_unclassified_filter_changed)
        self.class_unclassified_filter_checkbox.pack(side="left", padx=6)
        self.unclassified_label = ctk.CTkLabel(
            self.class_filters, text="", text_color=MUTED, height=30,
            font=ctk.CTkFont(size=11, weight="bold"), anchor="e")
        self.unclassified_label.pack(side="right", padx=(10, 0))

        self.class_browser = ctk.CTkFrame(class_page, fg_color="transparent")
        self.class_browser.pack(fill="both", expand=True, padx=PAD, pady=(0, PAD))
        self.class_left_panel = ctk.CTkFrame(
            self.class_browser, width=320, fg_color="#f8fcfb", border_color=LINE,
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

        self.class_right_panel = ctk.CTkFrame(
            self.class_browser, fg_color="#f8fcfb", border_color=LINE,
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
            self.class_browser, fg_color="#f8fcfb", border_color=LINE,
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
        self.class_favorite_buttons = {}
        for code in classes.GROUPS:
            tile = ctk.CTkFrame(self.class_group_list, fg_color="transparent")
            button = ctk.CTkButton(
                tile, text=I.t("class_" + code), height=38, corner_radius=9,
                anchor="w", fg_color="#f8fcfb", text_color="#1a302e",
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
            favorite_button = ctk.CTkButton(
                tile, text="☆", width=30, height=30, corner_radius=8,
                fg_color="transparent", text_color=ACCENT, hover_color=ACCENT_SOFT,
                command=lambda selected=code: self.toggle_therapeutic_group_favorite(selected))
            favorite_button.grid(row=0, column=2, padx=(4, 0))
            self.class_buttons[code] = button
            self.class_tiles[code] = tile
            self.class_count_badges[code] = count_badge
            self.class_favorite_buttons[code] = favorite_button
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
        editor = self.section(page, I.t("treatment_template_editor"))

        toolbar = ctk.CTkFrame(editor, fg_color="transparent")
        toolbar.pack(fill="x", padx=PAD, pady=(4, 7))
        self.treatment_disease_var = tk.StringVar()
        self.treatment_template_selector_var = tk.StringVar(
            value=I.t("treatment_select_template"))
        self.treatment_template_selector = ctk.CTkComboBox(
            toolbar, values=[I.t("treatment_select_template")],
            variable=self.treatment_template_selector_var, width=280,
            height=ACTION_HEIGHT, fg_color=ACCENT_SOFT, text_color=ACCENT,
            border_color=LINE, dropdown_fg_color=CARD,
            dropdown_hover_color=ACCENT_SOFT,
            button_color=ACCENT, button_hover_color=ACCENT_HOVER,
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
        self._treatment_template_popup = None
        self._treatment_template_suggestion_buttons = []
        ctk.CTkButton(
            toolbar, text=I.t("treatment_new"), width=78, height=ACTION_HEIGHT,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=self.new_treatment_template).pack(
                side="left", padx=3)
        ctk.CTkButton(
            toolbar, text=I.t("treatment_save"), width=78, height=ACTION_HEIGHT,
            fg_color=GOOD, hover_color=GOOD_HOVER,
            command=self.save_treatment_template).pack(side="left", padx=3)
        ctk.CTkButton(
            toolbar, text=I.t("treatment_duplicate"), width=92,
            height=ACTION_HEIGHT, fg_color=CARD, text_color=ACCENT,
            border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
            command=self.duplicate_treatment_template).pack(side="left", padx=3)
        self.treatment_delete_button = ctk.CTkButton(
            toolbar, text=I.t("treatment_delete"), width=78,
            height=ACTION_HEIGHT, fg_color=CARD, text_color=DANGER,
            border_width=1, border_color=LINE, hover_color="#fae9e9",
            command=self.delete_treatment_template)
        self.treatment_delete_button.pack(side="left", padx=3)
        ctk.CTkButton(
            toolbar, text=I.t("treatment_use_rx"), width=105,
            height=ACTION_HEIGHT, fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self.use_treatment_template).pack(side="right")

        search_row = ctk.CTkFrame(editor, fg_color="transparent")
        search_row.pack(fill="x", padx=PAD, pady=(0, 5))
        ctk.CTkLabel(
            search_row, text=I.t("treatment_add_from_database"), width=125,
            text_color=MUTED, anchor="w",
            font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")
        self.treatment_drug_search_var = tk.StringVar()
        self.treatment_drug_search_var.trace_add(
            "write", lambda *_: self._schedule_treatment_drug_search())
        self.treatment_drug_search_entry = ctk.CTkEntry(
            search_row, textvariable=self.treatment_drug_search_var,
            height=FIELD_HEIGHT, border_color=LINE, corner_radius=9,
            placeholder_text=I.t("treatment_search_database"),
            font=ctk.CTkFont(size=14))
        self.treatment_drug_search_entry.pack(side="left", fill="x", expand=True)

        self.treatment_drug_results = ctk.CTkFrame(
            editor, fg_color="#f8fcfb", border_color=LINE,
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

    def _set_treatment_status(self, text, color=MUTED):
        self.treatment_template_status.configure(text=text, text_color=color)

    @staticmethod
    def _treatment_disease_values():
        diseases = {
            item.get("disease", "").strip()
            for item in cfg.config.treatment_templates()
            if item.get("disease", "").strip()
        }
        return sorted(diseases, key=str.casefold)

    def _restore_treatment_disease_values(self):
        if hasattr(self, "treatment_disease_entry"):
            values = self._treatment_disease_values()
            self.treatment_disease_entry.configure(
                values=[directional_display_text(value) for value in values] or [""])

    def _filter_treatment_disease_menu(self, event=None):
        """Autocomplete the disease field from locally saved treatment plans."""
        if event and event.keysym == "Escape":
            self._hide_treatment_disease_suggestions()
            return
        if event and event.keysym in {
                "Up", "Down", "Return", "Tab", "Shift_L", "Shift_R",
                "Control_L", "Control_R", "Alt_L", "Alt_R"}:
            return
        logical = strip_bidi_display_controls(self.treatment_disease_var.get()).strip()
        query = logical.casefold()
        matches = [value for value in self._treatment_disease_values()
                   if not query or query in value.casefold()]
        self.treatment_disease_entry.configure(
            values=[directional_display_text(value) for value in matches]
            or [directional_display_text(logical)])
        if query and matches:
            self._show_treatment_disease_suggestions(matches)
        else:
            self._hide_treatment_disease_suggestions()

    def _show_treatment_disease_suggestions(self, matches):
        """Open clickable saved-disease matches immediately below the input."""
        self._hide_treatment_disease_suggestions()
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
        width = max(340, self.treatment_disease_entry.winfo_width())
        x = self.treatment_disease_entry.winfo_rootx()
        anchor_top = self.treatment_disease_entry.winfo_rooty()
        anchor_bottom = anchor_top + self.treatment_disease_entry.winfo_height()
        shell = ctk.CTkFrame(
            popup, fg_color=CARD, border_color=ACCENT,
            border_width=2, corner_radius=8)
        shell.pack(fill="both", expand=True)
        self._treatment_disease_popup = popup
        self._treatment_disease_suggestion_buttons = []
        visible_rows = min(len(matches), 5)
        container = shell
        is_scrollable = len(matches) > 3
        if is_scrollable:
            container = ctk.CTkScrollableFrame(
                shell, fg_color=CARD, corner_radius=6,
                scrollbar_button_color=ACCENT,
                scrollbar_button_hover_color=ACCENT_HOVER)
            container.pack(fill="both", expand=True, padx=3, pady=3)
        for index, disease in enumerate(matches):
            button = ctk.CTkButton(
                container, text=directional_display_text(disease), height=36,
                corner_radius=5, fg_color="transparent", text_color="#173b38",
                hover_color=ACCENT_SOFT, anchor="w",
                font=ctk.CTkFont(size=13),
                command=lambda value=disease: self._choose_treatment_disease(value))
            bottom_inset = 4 if index == len(matches) - 1 else 0
            button.pack(fill="x", padx=4, pady=(4, bottom_inset))
            self._treatment_disease_suggestion_buttons.append(button)
        popup.update_idletasks()
        screen_height = popup.winfo_screenheight()
        max_height = max(96, screen_height - 24)
        if is_scrollable:
            row_height = max(
                button.winfo_reqheight()
                for button in self._treatment_disease_suggestion_buttons) + 8
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

    def _hide_treatment_disease_suggestions(self):
        popup = getattr(self, "_treatment_disease_popup", None)
        self._treatment_disease_popup = None
        self._treatment_disease_suggestion_buttons = []
        if popup is not None:
            try:
                if popup.winfo_exists():
                    popup.destroy()
            except tk.TclError:
                pass

    def _on_treatment_disease_focus_out(self, _event=None):
        self._restore_treatment_disease_values()
        self.after(160, self._hide_treatment_disease_suggestions)

    def _choose_treatment_disease(self, disease):
        self._hide_treatment_disease_suggestions()
        self.treatment_disease_var.set(disease)
        self._select_saved_treatment_disease(disease)

    def _select_saved_treatment_disease(self, value):
        self._hide_treatment_disease_suggestions()
        disease = strip_bidi_display_controls(value).strip()
        matches = [item for item in cfg.config.treatment_templates()
                   if item.get("disease", "").strip().casefold() == disease.casefold()]
        if not matches:
            return
        template = matches[-1]
        label = next((label for label, template_id in self._treatment_template_lookup.items()
                      if template_id == template["id"]), "")
        if label:
            self.treatment_template_selector_var.set(label)
        self._load_treatment_template_record(template)

    def _load_treatment_disease_from_search(self, _event=None):
        disease = strip_bidi_display_controls(self.treatment_disease_var.get()).strip()
        exact = [value for value in self._treatment_disease_values()
                 if value.casefold() == disease.casefold()]
        candidates = exact or [value for value in self._treatment_disease_values()
                               if disease.casefold() in value.casefold()]
        if len(candidates) == 1:
            self._select_saved_treatment_disease(candidates[0])
        return "break"

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
        width = max(360, self.treatment_template_selector.winfo_width())
        x = self.treatment_template_selector.winfo_rootx()
        anchor_top = self.treatment_template_selector.winfo_rooty()
        anchor_bottom = anchor_top + self.treatment_template_selector.winfo_height()
        shell = ctk.CTkFrame(
            popup, fg_color=CARD, border_color=ACCENT,
            border_width=2, corner_radius=8)
        shell.pack(fill="both", expand=True)
        self._treatment_template_popup = popup
        self._treatment_template_suggestion_buttons = []
        visible_rows = min(len(matches), 5)
        container = shell
        is_scrollable = len(matches) > 3
        if is_scrollable:
            container = ctk.CTkScrollableFrame(
                shell, fg_color=CARD, corner_radius=6,
                scrollbar_button_color=ACCENT,
                scrollbar_button_hover_color=ACCENT_HOVER)
            container.pack(fill="both", expand=True, padx=3, pady=3)
        for index, label in enumerate(matches):
            button = ctk.CTkButton(
                container, text=directional_display_text(label), height=36,
                corner_radius=5, fg_color="transparent", text_color="#173b38",
                hover_color=ACCENT_SOFT, anchor="w",
                font=ctk.CTkFont(size=13),
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
        self._hide_treatment_template_suggestions()
        self._treatment_template_id = ""
        self.treatment_disease_var.set("")
        self.treatment_drug_search_var.set("")
        self._treatment_template_drugs = []
        self.refresh_treatment_template_menu()
        self.treatment_template_selector_var.set("")
        self.render_treatment_template_drugs()
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
        template_id = template.get("id", "")
        self._treatment_template_id = template_id
        self.treatment_disease_var.set(template["disease"])
        self._treatment_template_drugs = [
            dict(item, _editor_open=False) for item in template["medications"]]
        self.render_treatment_template_drugs()
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
            ctk.CTkButton(
                self.treatment_drug_results, text="+  " + text, height=32,
                fg_color="transparent", text_color="#173b38", anchor="w",
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
            card = ctk.CTkFrame(
                self.treatment_medications_frame, fg_color="#fbfdfd",
                border_color=LINE, border_width=1, corner_radius=9)
            card.pack(fill="x", pady=3)
            header = ctk.CTkFrame(card, fg_color="transparent")
            header.pack(fill="x", padx=9, pady=(5, 3))
            ctk.CTkLabel(
                header, text=f"{index + 1}.", text_color="#173b38",
                font=ctk.CTkFont(size=13, weight="bold"), anchor="w").pack(side="left")

            actions = ctk.CTkFrame(header, fg_color="transparent")
            actions.pack(side="right")
            ctk.CTkButton(
                actions, text=I.t("treatment_remove"), width=68, height=27,
                fg_color="transparent", text_color=DANGER, hover_color="#fae9e9",
                command=lambda position=index: self.remove_treatment_drug(position)).pack(
                    side="right", padx=(4, 0))
            ctk.CTkButton(
                actions, text="↓", width=30, height=27, fg_color="transparent",
                text_color=ACCENT, hover_color=ACCENT_SOFT,
                command=lambda position=index: self.move_treatment_drug(position, 1)).pack(
                    side="right", padx=2)
            ctk.CTkButton(
                actions, text="↑", width=30, height=27, fg_color="transparent",
                text_color=ACCENT, hover_color=ACCENT_SOFT,
                command=lambda position=index: self.move_treatment_drug(position, -1)).pack(
                    side="right", padx=2)
            ctk.CTkButton(
                actions,
                text=(I.t("treatment_collapse") if medicine.get("_editor_open")
                      else I.t("treatment_edit")),
                width=64, height=27, fg_color=CARD, text_color=ACCENT,
                border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
                command=lambda position=index: self.toggle_treatment_drug_editor(position)).pack(
                    side="right", padx=2)

            if medicine.get("alternative_to_previous"):
                ctk.CTkLabel(
                    header, text=I.t("treatment_or"), width=32, height=22,
                    corner_radius=11, fg_color=WARNING_SOFT, text_color=WARNING,
                    font=ctk.CTkFont(size=10, weight="bold")).pack(
                        side="left", padx=(7, 3))
            primary, secondary = self._treatment_medicine_title(medicine)
            ctk.CTkLabel(
                header, text=primary or I.t("drug"), text_color="#173b38",
                font=ctk.CTkFont(size=14, weight="bold"), anchor="w").pack(
                    side="left", padx=(7, 0))
            if secondary:
                ctk.CTkLabel(
                    header, text=f"  ·  {secondary}", text_color=MUTED,
                    font=ctk.CTkFont(size=13), anchor="w").pack(side="left")

            summary = self._treatment_regimen_summary(medicine)
            if summary:
                ctk.CTkLabel(
                    card, text=summary, text_color=MUTED, anchor="w", justify="left",
                    wraplength=900, font=ctk.CTkFont(size=11)).pack(
                        fill="x", padx=38, pady=(0, 5))

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
                ctk.CTkEntry(
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
                    widget = ctk.CTkComboBox(
                        column_frame,
                        values=[directional_display_text(value) for value in options],
                        variable=binding.display_var, height=34, corner_radius=8,
                        border_width=1, border_color=ACCENT, fg_color=CARD,
                        button_color=ACCENT, button_hover_color=ACCENT_HOVER,
                        dropdown_fg_color=CARD, dropdown_hover_color=ACCENT_SOFT,
                        font=ctk.CTkFont(size=12),
                        dropdown_font=ctk.CTkFont(size=12), justify="left")
                else:
                    widget = ctk.CTkEntry(
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
                variable=alternative_var, height=26, fg_color=ACCENT,
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
        self._set_treatment_status(I.t("treatment_saved"), ACCENT)

    def duplicate_treatment_template(self):
        """Save the current form as a separate editable template."""
        disease = self._current_treatment_disease()
        if not disease:
            self._set_treatment_status(I.t("treatment_disease_required"), DANGER)
            self.treatment_template_selector.focus_set()
            return
        if not self._treatment_template_drugs:
            self._set_treatment_status(I.t("treatment_medicine_required"), DANGER)
            return
        copy_suffix = I.t("treatment_copy_suffix")
        copied_disease = f"{disease} {copy_suffix}".strip()
        self.treatment_disease_var.set(copied_disease)
        saved_id = cfg.config.save_treatment_template({
            "id": uuid.uuid4().hex, "disease": copied_disease,
            "variant": "",
            "medications": self._treatment_template_drugs,
        })
        if not saved_id:
            self._set_treatment_status(I.t("treatment_save_failed"), DANGER)
            return
        self._treatment_template_id = saved_id
        self.refresh_treatment_template_menu(saved_id)
        self._set_treatment_status(I.t("treatment_duplicated"), ACCENT)

    def delete_treatment_template(self):
        if not self._treatment_template_id:
            self._set_treatment_status(I.t("treatment_select_first"), WARNING)
            return
        if not messagebox.askyesno(
                I.t("treatment_delete"), I.t("treatment_delete_confirm"), parent=self):
            return
        cfg.config.remove_treatment_template(self._treatment_template_id)
        self.new_treatment_template()
        self._set_treatment_status(I.t("treatment_deleted"), ACCENT)

    def backup_treatment_templates(self):
        path = filedialog.asksaveasfilename(
            parent=self, title=I.t("treatment_backup"),
            defaultextension=".rxtemplates",
            filetypes=[(I.t("treatment_backup_file"), "*.rxtemplates")])
        if not path:
            return
        try:
            cfg.config.export_treatment_templates(path)
        except Exception as exc:
            logging.exception("Treatment-template backup failed")
            messagebox.showerror(I.t("treatment_backup"), str(exc), parent=self)
            return
        self._set_treatment_status(I.t("treatment_backup_done", path=path), ACCENT)

    def restore_treatment_templates(self):
        path = filedialog.askopenfilename(
            parent=self, title=I.t("treatment_restore"),
            filetypes=[(I.t("treatment_backup_file"), "*.rxtemplates")])
        if not path:
            return
        choice = messagebox.askyesnocancel(
            I.t("treatment_restore"), I.t("treatment_restore_choice"), parent=self)
        if choice is None:
            return
        try:
            count = cfg.config.import_treatment_templates(path, replace=bool(choice))
        except Exception as exc:
            logging.exception("Treatment-template restore failed")
            messagebox.showerror(I.t("treatment_restore"), str(exc), parent=self)
            return
        self.new_treatment_template()
        self._set_treatment_status(I.t("treatment_restore_done", n=count), ACCENT)

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
                self.new_treatment_template()
                self._set_treatment_status(I.t("treatment_import_done", n=count), ACCENT)
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

    def _show_treatment_apply_preview(self):
        """Preview mandatory steps and choose one medicine from every OR group."""
        groups = self._treatment_choice_groups(self._treatment_template_drugs)
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
            card = ctk.CTkFrame(
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
                        fg_color=ACCENT, hover_color=ACCENT_HOVER,
                        text_color="#173b38", font=ctk.CTkFont(size=12),
                        command=lambda: None).pack(
                            fill="x", padx=2, pady=2, anchor="w")
                else:
                    ctk.CTkLabel(
                        row, text=text, text_color="#173b38", anchor="w",
                        justify="left", font=ctk.CTkFont(size=12)).pack(
                            fill="x", padx=2, pady=2)
        actions = ctk.CTkFrame(dialog, fg_color="transparent")
        actions.pack(fill="x", padx=18, pady=(0, 14))
        ctk.CTkButton(
            actions, text=I.t("settings_cancel"), width=90, height=ACTION_HEIGHT,
            fg_color=CARD, text_color=MUTED, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=dialog.destroy).pack(side="left")
        ctk.CTkButton(
            actions, text=I.t("treatment_add_current"), width=150,
            height=ACTION_HEIGHT, fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=lambda: self._apply_treatment_selection(
                selections, False, dialog)).pack(side="right", padx=(6, 0))
        ctk.CTkButton(
            actions, text=I.t("treatment_replace_current"), width=165,
            height=ACTION_HEIGHT, fg_color=CARD, text_color=ACCENT,
            border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
            command=lambda: self._apply_treatment_selection(
                selections, True, dialog)).pack(side="right")
        dialog.grab_set()
        dialog.focus_set()

    def _apply_treatment_selection(self, selections, replace, dialog):
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

    def toggle_medication_favorite_picker(self):
        if self.medication_favorite_panel.winfo_manager():
            self.medication_favorite_panel.pack_forget()
            return
        self.medication_favorite_panel.pack(
            fill="x", padx=2, pady=(0, 6), before=self.medications_section)
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
            result_row.pack(fill="x", pady=2)
            brand = favorite.get("brand_name", "").strip()
            scientific = favorite.get("generic_name", "").strip()
            name = brand or scientific
            if brand and scientific and brand.casefold() != scientific.casefold():
                name = f"{brand}  ·  {scientific}"
            regimen = self._favorite_regimen_line(favorite)
            text = name if not regimen else f"{name}\n{regimen}"
            ctk.CTkLabel(
                result_row, text=text, text_color="#173b38", anchor="w", justify="left",
                font=ctk.CTkFont(size=12, weight="bold"), wraplength=680).pack(
                    side="left", fill="x", expand=True, padx=10, pady=6)
            ctk.CTkButton(
                result_row, text=I.t("favorite_use_rx"), width=82, height=30,
                fg_color=ACCENT, hover_color=ACCENT_HOVER,
                command=lambda item_id=favorite.get("id", index):
                    self.use_favorite_from_medication(item_id)).pack(
                        side="right", padx=6, pady=5)
            shown += 1
            if shown >= 12:
                break
        if not shown:
            ctk.CTkLabel(
                self.medication_favorite_results,
                text=(I.t("starred_drugs_empty") if not indexed and not query
                      else I.t("favorite_no_matches")),
                text_color=MUTED, anchor="w").pack(fill="x", padx=8, pady=8)

    def use_favorite_from_medication(self, identifier):
        self.use_favorite(identifier)
        self.medication_favorite_panel.pack_forget()
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
                ctk.CTkButton(
                    self._favorite_more_frame, text=I.t("load_more"),
                    width=110, height=32, fg_color=ACCENT,
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
        for drug in self.db.drugs:
            mappings = classes.groups_for(drug)
            if not mappings:
                unclassified.append(drug)
                continue
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
        self._classification_cache = {
            "by_group": by_group, "by_pair": by_pair,
            "unclassified": unclassified,
        }
        return self._classification_cache

    def unclassified_medicine_count(self):
        return len(self._ensure_classification_cache()["unclassified"])

    def therapeutic_group_medicine_count(self, code):
        return len(self._ensure_classification_cache()["by_group"].get(code, ()))

    def refresh_class_overview(self):
        if not hasattr(self, "class_tiles"):
            return
        pinned = set(cfg.config.favorite_therapeutic_groups())
        ordered = self.ordered_therapeutic_groups()
        for code in ordered:
            tile = self.class_tiles[code]
            tile.pack_forget()
            tile.pack(fill="x", pady=3)
            self.class_favorite_buttons[code].configure(text="★" if code in pinned else "☆")
            self.class_count_badges[code].configure(
                text=str(self.therapeutic_group_medicine_count(code)))
        if hasattr(self, "unclassified_label"):
            self.unclassified_label.configure(
                text=I.t("unclassified_medicines", n=self.unclassified_medicine_count()))
        self.refresh_class_browser()

    def _schedule_class_browser_refresh(self, delay=140):
        if self._class_search_job is not None:
            try:
                self.after_cancel(self._class_search_job)
            except (tk.TclError, ValueError):
                pass
        self._class_search_job = self.after(delay, self.refresh_class_browser)

    def class_group_filter_changed(self, label):
        code = self.class_filter_group_labels.get(label, "")
        details = list(classes.subclasses_for(code)) if code else []
        values = [I.t("filter_all_classes")] + details
        self.class_detail_filter_menu.configure(values=values)
        self.class_detail_filter_var.set(I.t("filter_all_classes"))
        if code:
            self._selected_therapeutic_group = code
            self.class_unclassified_filter_var.set(False)
        self.refresh_class_browser()

    def class_unclassified_filter_changed(self):
        if self.class_unclassified_filter_var.get():
            self.class_group_filter_var.set(I.t("filter_all_groups"))
            self.class_detail_filter_var.set(I.t("filter_all_classes"))
        self.refresh_class_browser()

    def _set_class_browser_mode(self, unclassified):
        """Switch between the classification browser and focused unclassified search."""
        if unclassified:
            self.class_left_panel.pack_forget()
            self.class_right_panel.pack_forget()
            if not self.class_unclassified_panel.winfo_manager():
                self.class_unclassified_panel.pack(fill="both", expand=True)
            self.class_group_filter_menu.pack_forget()
            self.class_detail_filter_menu.pack_forget()
            return

        self.class_unclassified_panel.pack_forget()
        if not self.class_left_panel.winfo_manager():
            self.class_left_panel.pack(side="left", fill="y", padx=(0, 8))
        if not self.class_right_panel.winfo_manager():
            self.class_right_panel.pack(side="left", fill="both", expand=True)
        if not self.class_group_filter_menu.winfo_manager():
            self.class_group_filter_menu.pack(
                side="left", padx=(0, 5), before=self.class_name_filter_menu)
        if not self.class_detail_filter_menu.winfo_manager():
            self.class_detail_filter_menu.pack(
                side="left", padx=5, before=self.class_name_filter_menu)

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
        ctk.CTkButton(
            self.class_unclassified_more_frame, text=I.t("load_more"),
            width=110, height=32, fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self._append_unclassified_class_search_page).pack(side="left")

    def _add_unclassified_class_card(self, index, drug):
        """Add one compact unclassified medicine card to the results grid."""
        card = ctk.CTkFrame(
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
            names, text=brand, text_color="#173b38", anchor="w",
            font=_ui_font(13, "bold")).pack(side="left")
        if scientific and scientific.casefold() != brand.casefold():
            ctk.CTkLabel(
                names, text="  " + scientific, text_color=MUTED,
                anchor="w", font=_ui_font(11)).pack(side="left")
        ctk.CTkLabel(
            names, text=I.t("classification_unclassified"), height=22,
            corner_radius=11, fg_color="#fff0cc", text_color=WARNING,
            font=_ui_font(9, "bold")).pack(
                side="left", padx=(8, 0))
        ctk.CTkButton(
            card, text=I.t("favorite_use_rx"), width=82, height=30,
            fg_color=CARD, text_color=ACCENT, border_width=1,
            border_color=LINE, hover_color=ACCENT_SOFT,
            command=lambda item=drug: self.add_drug_database_item(item)).grid(
                row=0, column=1, padx=(2, 4), pady=5)
        ctk.CTkButton(
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
            "unclassified": ("#f1f3f3", MUTED),
        }.get(status, ("#f1f3f3", MUTED))

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
        group_filter = self.class_filter_group_labels.get(
            self.class_group_filter_var.get(), "") if hasattr(
                self, "class_group_filter_var") else ""
        detail_filter = self.class_detail_filter_var.get() if hasattr(
            self, "class_detail_filter_var") else I.t("filter_all_classes")
        if detail_filter == I.t("filter_all_classes"):
            detail_filter = ""
        only_unclassified = bool(self.class_unclassified_filter_var.get()) if hasattr(
            self, "class_unclassified_filter_var") else False
        render_signature = (
            query, group_filter, detail_filter, only_unclassified,
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
            if group_filter and code != group_filter:
                continue
            group_text = I.t("class_" + code).casefold()
            details = classes.subclasses_for(code)
            if detail_filter:
                details = tuple(detail for detail in details if detail == detail_filter)
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
                fg_color=ACCENT if selected else "#f8fcfb",
                text_color="white" if selected else "#1a302e",
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
            if detail_filter and detail != detail_filter:
                continue
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
            button = ctk.CTkButton(
                self.class_detail_list,
                text=I.t("detailed_class_with_count", detail=detail, n=count),
                height=34, corner_radius=8, anchor="w",
                fg_color=ACCENT if selected else CARD,
                text_color="white" if selected else "#1a302e",
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
        menu = tk.Menu(self, tearoff=False, font=("Segoe UI", 44))
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
        self.show_class_mapping_editor()
        self.mapping_search_var.set(drug.brand_name.strip() or drug.generic_name.strip())
        self.refresh_mapping_list()
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
        for drug in self.db.drugs:
            key = drug.generic_name.strip().casefold()
            mappings = classes.groups_for(drug)
            mapping_set = frozenset((mapping.code, mapping.detail) for mapping in mappings)
            if key:
                mapping_by_name.setdefault(key, []).append(mapping_set)
            for mapping in mappings:
                if mapping.code not in classes.GROUPS:
                    invalid_groups.append(drug.generic_name)
                elif mapping.detail not in classes.subclasses_for(mapping.code):
                    invalid_details.append(drug.generic_name)
        conflicts = sum(
            1 for values in mapping_by_name.values()
            if len(set(values)) > 1)
        missing_favorites = sum(
            1 for favorite in cfg.config.medication_favorites()
            if not self._favorite_matches_database(favorite))
        return {
            "unclassified": self.unclassified_medicine_count(),
            "invalid_groups": len(invalid_groups),
            "invalid_details": len(invalid_details),
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

    def show_class_mapping_editor(self, review_unclassified=False):
        """Edit local class metadata without requiring CSV editing."""
        self.class_overview.pack_forget()
        for child in self.class_subpage.winfo_children():
            child.destroy()
        self.class_subpage.pack(fill="both", expand=True)
        card = self.section(self.class_subpage, I.t("class_mapping_editor"))
        bar = ctk.CTkFrame(card, fg_color="transparent")
        bar.pack(fill="x", padx=PAD, pady=(0, 8))
        ctk.CTkButton(bar, text="← " + I.t("back"), height=ACTION_HEIGHT,
                      width=86, fg_color=CARD, text_color=ACCENT,
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
        self._mapping_search_job = None
        self.mapping_search_var.trace_add("write", lambda *_: self._schedule_mapping_refresh())
        ctk.CTkEntry(left, textvariable=self.mapping_search_var, height=FIELD_HEIGHT,
                     placeholder_text=I.t("search_medicines"), border_color=LINE).pack(fill="x", pady=(0, 6))
        self.mapping_unclassified_only = tk.BooleanVar(value=False)
        self.mapping_unclassified_only.set(bool(review_unclassified))
        ctk.CTkCheckBox(left, text=I.t("show_unclassified_only"), variable=self.mapping_unclassified_only,
                         text_color=MUTED, fg_color=ACCENT, hover_color=ACCENT_HOVER,
                         command=self.refresh_mapping_list).pack(anchor="w", pady=(0, 6))
        self.mapping_list = tk.Listbox(left, height=5, font=LIST_FONT,
                                       bg="#f8fcfb", fg="#1a302e", relief="flat", borderwidth=0,
                                       highlightthickness=1, highlightbackground=LINE,
                                       selectbackground=ACCENT, selectforeground="white", activestyle="none")
        self.mapping_list.configure(selectmode=tk.EXTENDED, exportselection=False)
        self.mapping_list.pack(fill="both", expand=True)
        self.mapping_list.bind("<<ListboxSelect>>", self.select_mapping_drug)
        review_actions = ctk.CTkFrame(left, fg_color="transparent")
        review_actions.pack(fill="x", pady=(6, 0))
        self.mapping_result_label = ctk.CTkLabel(
            review_actions, text="", text_color=MUTED,
            font=ctk.CTkFont(size=10), anchor="w")
        self.mapping_result_label.pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            review_actions, text=I.t("previous"), width=92, height=32,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=lambda: self.mapping_select_relative(-1)).pack(
                side="left", padx=(0, 4))
        ctk.CTkButton(
            review_actions, text=I.t("skip"), width=78, height=32,
            fg_color=CARD, text_color=ACCENT, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=lambda: self.mapping_select_relative(1)).pack(
                side="left", padx=4)
        ctk.CTkButton(
            review_actions, text=I.t("save_and_next"), width=130, height=32,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self.mapping_save_and_next).pack(side="right")
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
        self.mapping_keep_existing_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            right, text=I.t("keep_existing_classes"),
            variable=self.mapping_keep_existing_var, fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=11)).pack(fill="x", padx=12, pady=(0, 8))
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
        self.mapping_selected_drugs = []
        self.pending_class_mapping = None
        self.mapping_visible_drugs = []
        self.refresh_mapping_list()
        if review_unclassified and self.mapping_visible_drugs:
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
        if self.mapping_unclassified_only.get():
            candidates = self._ensure_classification_cache()["unclassified"]
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
            drugs, key=lambda drug: drug.generic_name.casefold())
        if hasattr(self, "mapping_result_label"):
            suffix = "+" if total_visible > 500 else ""
            self.mapping_result_label.configure(text=f"{len(self.mapping_visible_drugs)}{suffix}")
        self.mapping_list.delete(0, tk.END)
        for drug in self.mapping_visible_drugs:
            found = classes.groups_for(drug)
            suffix = ", ".join(mapping.detail for mapping in found) if found else I.t("unclassified")
            status = found[0].confidence if found else "unclassified"
            self.mapping_list.insert(
                tk.END,
                f"{drug.generic_name}  —  {suffix}  [{I.t('classification_' + status)}]")
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
        if not self.mapping_selected_drugs:
            return
        code = self.mapping_group_labels.get(self.mapping_group_var.get())
        detail = self.mapping_detail_var.get()
        if not code or detail == I.t("choose_detailed_class"):
            self.mapping_status.configure(text=I.t("choose_group_and_class"), text_color=DANGER)
            return
        medicines = tuple(drug.generic_name for drug in self.mapping_selected_drugs)
        self.pending_class_mapping = (medicines, code, detail)
        self.mapping_status.configure(
            text=I.t("mapping_batch_ready", n=len(medicines), detail=detail),
            text_color=GOOD)
        self.save_mapping_button.configure(state="normal")

    def save_class_mapping(self, advance=False):
        if not self.pending_class_mapping:
            return False
        medicines, code, detail = self.pending_class_mapping
        current_indices = self.mapping_list.curselection()
        next_index = current_indices[0] if current_indices else 0
        updated = self.db.update_classifications(
            medicines, code, detail, append=self.mapping_keep_existing_var.get())
        if updated:
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
        ctk.CTkButton(
            toolbar, text="← " + I.t("back"), height=ACTION_HEIGHT,
            width=86, fg_color=CARD, text_color=ACCENT,
            border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
            command=self.back_to_major_groups).grid(row=0, column=0, sticky="w")
        self.all_class_search_var = tk.StringVar()
        self.all_class_search_var.trace_add("write", lambda *_: self.render_all_detailed_classes())
        ctk.CTkEntry(
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
                ctk.CTkButton(
                    grid, text=detail, height=33, corner_radius=8, anchor="w",
                    fg_color="#f8fcfb", text_color="#1a302e", border_width=1,
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
        ctk.CTkButton(bar, text="← " + I.t("back"), height=ACTION_HEIGHT,
                      width=86, fg_color=CARD, text_color=ACCENT,
                      border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
                      command=self.back_from_subclass_page).pack(side="left")
        trail = ctk.CTkFrame(bar, fg_color="transparent")
        trail.pack(side="left", padx=(8, 0))
        ctk.CTkButton(
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
            button = ctk.CTkButton(
                list_frame, text=I.t("detailed_class_with_count", detail=detail,
                                     n=subclass_counts[detail]), height=36, corner_radius=9, anchor="w",
                fg_color="#f8fcfb", text_color="#1a302e", border_width=1,
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
                fg_color=ACCENT if selected else "#f8fcfb",
                text_color="white" if selected else "#1a302e",
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
        ctk.CTkButton(
            bar, text="← " + I.t("back"),
            height=ACTION_HEIGHT, width=86, fg_color=CARD, text_color=ACCENT,
            border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
            command=self.back_to_major_groups).pack(side="left")
        trail = ctk.CTkFrame(bar, fg_color="transparent")
        trail.pack(side="left", padx=(8, 0))
        ctk.CTkButton(
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
        ctk.CTkButton(
            footer, text=I.t("load_more"), width=110, height=32,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
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
            names, text=brand, text_color="#173b38", anchor="w",
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
        ctk.CTkButton(
            row, text="★" if self._class_drug_is_starred(drug) else "☆",
            width=36, height=32, fg_color="transparent", text_color=ACCENT,
            hover_color=ACCENT_SOFT,
            command=lambda item=drug, group=group_code, picked=detail:
                self.toggle_detail_drug_star(item, group, picked)).pack(
                    side="right", padx=(2, 7), pady=5)
        ctk.CTkButton(
            row, text=I.t("favorite_use_rx"), width=88, height=32,
            fg_color=CARD, text_color=ACCENT, border_width=1,
            border_color=LINE, hover_color=ACCENT_SOFT,
            command=lambda item=drug: self.add_drug_database_item(item)).pack(
                side="right", padx=2, pady=5)
        self._bind_class_medicine_context(row, drug)

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

    def select_drug_subclass(self, group_code, detail):
        """Compatibility entry point: open the chosen detail-class page."""
        self.show_detail_medicines_page(group_code, detail)

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
        target = next((row for row in self.rows
                       if not row.name_var.get().strip()
                       and not row.trade_var.get().strip()), None)
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
        """Backward-compatible entry point for selecting the browser group."""
        self.select_class_browser_group(code)

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

    def _set_patient_form(self, record):
        for key, variable in self.patient_vars.items():
            variable.set(str(record.get(key, "")))
        sex_code = str(record.get("sex", ""))
        sex_label = next(
            (label for label, code in self._patient_sex_codes.items() if code == sex_code),
            I.t("sex_m"))
        self.patient_sex_menu.set(sex_label)
        self._loaded_patient_id = str(record.get("id", ""))
        self.patient_delete_button.configure(state="normal")

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
        matches = self.patient_history.find_similar(name, age)
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

    def select_patient_history(self, event=None):
        selected = self.patient_history_list.curselection()
        if not selected:
            return
        record = self._patient_history_records[selected[0]]
        self._current_history_record = record
        self.patient_delete_button.configure(state="normal")
        self.show_patient_prescriptions(record)

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
            card = ctk.CTkFrame(self.patient_prescriptions_body, fg_color="#f8faff",
                                border_color=LINE, border_width=1, corner_radius=9)
            card.pack(fill="x", pady=3)
            prescription_id = str(prescription.get("id", ""))
            expanded = prescription_id in self._expanded_prescription_ids
            arrow = "▾" if expanded else "▸"
            ctk.CTkButton(
                card,
                text=f"{arrow}  {saved_at or '—'}  ·  {I.t('medicine_count', n=len(drugs))}",
                height=36, anchor="w", fg_color="transparent", text_color=ACCENT,
                hover_color=ACCENT_SOFT, font=ctk.CTkFont(size=12, weight="bold"),
                command=lambda item_id=prescription_id, patient=record:
                    self.toggle_patient_prescription(patient, item_id)).pack(
                        fill="x", padx=6, pady=4)
            if not expanded:
                continue
            details = ctk.CTkFrame(card, fg_color="transparent")
            details.pack(fill="x", padx=12, pady=(0, 6))
            for number, drug in enumerate(drugs, 1):
                name = str(drug.get("brand_name", "") or drug.get("generic_name", ""))
                scientific = str(drug.get("generic_name", ""))
                if drug.get("brand_name") and scientific:
                    name = f"{name} ({scientific})"
                regimen = "  ·  ".join(
                    str(drug.get(key, "")).strip()
                    for key in ("dosage", "frequency", "duration", "notes")
                    if str(drug.get(key, "")).strip())
                line = f"{number}. {name}" + (f" — {regimen}" if regimen else "")
                ctk.CTkLabel(
                    details, text=directional_display_text(line), text_color="#243b39",
                    font=ctk.CTkFont(size=11), anchor="w", justify="left",
                    wraplength=720).pack(fill="x", pady=2)
            actions = ctk.CTkFrame(details, fg_color="transparent")
            actions.pack(fill="x", pady=(5, 0))
            ctk.CTkButton(
                actions, text=I.t("load_rx"), width=100, height=32,
                fg_color=CARD, text_color=ACCENT, border_width=1,
                border_color=LINE, hover_color=ACCENT_SOFT,
                command=lambda item=prescription, patient=record:
                    self.load_saved_prescription(patient, item)).pack(side="left")

    def toggle_patient_prescription(self, record, prescription_id):
        if prescription_id in self._expanded_prescription_ids:
            self._expanded_prescription_ids.remove(prescription_id)
        else:
            self._expanded_prescription_ids.add(prescription_id)
        self.show_patient_prescriptions(record)

    def save_patient_history(self):
        if not self._warn_similar_patient():
            return
        try:
            record = self.patient_history.save_patient(
                {key: variable.get() for key, variable in self.patient_vars.items()},
                self._loaded_patient_id)
        except ValueError as exc:
            messagebox.showinfo(APP_TITLE, str(exc))
            return
        self._loaded_patient_id = str(record.get("id", ""))
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

    def delete_selected_patient(self):
        record = self._selected_or_loaded_patient()
        if not record:
            return
        if not messagebox.askyesno(I.t("delete_patient"), I.t("delete_patient_confirm", name=record.get("name", ""))):
            return
        if self.patient_history.delete(record.get("id", "")):
            if record.get("id") == self._loaded_patient_id:
                self.clear_patient_details()
            self.refresh_patient_history()
            self._show_empty_prescriptions()

    def new_patient(self):
        self.clear_patient_details()
        self.patient_search_var.set("")
        self.patient_name_entry.focus_set()

    def clear_patient_details(self):
        for variable in self.patient_vars.values():
            variable.set("")
        self.patient_sex_menu.set(I.t("sex_m"))
        self._loaded_patient_id = ""
        self._current_history_record = None
        self.patient_history_list.selection_clear(0, tk.END)
        self.patient_delete_button.configure(state="disabled")
        self._show_empty_prescriptions()

    def clear_patient_search(self, event=None):
        self.patient_search_var.set("")
        return "break" if event else None

    def _save_patient_shortcut(self, event=None):
        if getattr(self, "active_page", "") == "patient":
            self.save_patient_history()
            return "break"
        return None

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
            self.add_row(data=qu.DrugItem(**{key: saved_drug.get(key, "") for key in
                                              ("generic_name", "brand_name", "dosage", "frequency", "duration", "notes")}))
        if not self.rows:
            self.add_row()
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

    def _render_word_preview_now(self):
        self._word_preview_job = None
        if self.word_preview_visible:
            self.render_word_preview()

    def toggle_word_preview(self):
        self.word_preview_visible = not self.word_preview_visible
        if self.word_preview_visible:
            self.word_preview_body.pack(fill="x", padx=PAD, pady=(0, PAD))
            self.word_preview_toggle.configure(text=I.t("hide_word_preview"))
            self.render_word_preview()
        else:
            self.word_preview_body.pack_forget()
            self.word_preview_toggle.configure(text=I.t("show_word_preview"))

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
                line_row, text=line, text_color="#1a1a1a",
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
        recognized = [drug for key, drug in identities.items() if after.get(key)]
        unclassified = [drug for key, drug in identities.items() if not after.get(key)]
        changed_keys = {
            key for key in set(before).intersection(after)
            if before.get(key) != after.get(key)
        }
        missing_keys = {key for key in set(before) - set(after) if before.get(key)}
        return {
            "new": [identities[key] for key in new_keys if key in identities],
            "recognized": recognized,
            "unclassified": unclassified,
            "changed_missing": list(changed_keys | missing_keys),
        }

    def show_import_classification_assistant(self, report):
        window = ctk.CTkToplevel(self)
        window.title(I.t("import_classification_assistant"))
        window.geometry("720x430")
        window.minsize(640, 390)
        window.transient(self)
        window.grab_set()
        ctk.CTkLabel(
            window, text=I.t("import_classification_assistant"),
            text_color=ACCENT, anchor="w",
            font=ctk.CTkFont(size=24, weight="bold")).pack(
                fill="x", padx=22, pady=(20, 4))
        ctk.CTkLabel(
            window, text=I.t("import_classification_assistant_hint"),
            text_color=MUTED, anchor="w", justify="left",
            font=ctk.CTkFont(size=12)).pack(fill="x", padx=22, pady=(0, 15))
        grid = ctk.CTkFrame(window, fg_color="transparent")
        grid.pack(fill="both", expand=True, padx=22)
        grid.grid_columnconfigure((0, 1), weight=1, uniform="import-report")
        items = (
            ("new", "import_new_medicines", ACCENT_SOFT, ACCENT),
            ("recognized", "import_recognized_mappings", ACCENT_SOFT, ACCENT),
            ("unclassified", "import_unclassified_medicines", WARNING_SOFT, WARNING),
            ("changed_missing", "import_changed_missing", "#f1f3f3", MUTED),
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
                card, text=I.t(label_key), text_color="#263f3d",
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
            ctk.CTkButton(
                actions, text=I.t("review_unclassified"), height=38,
                fg_color=ACCENT, hover_color=ACCENT_HOVER,
                command=lambda: (window.destroy(),
                                 self.show_page("drug_classes"),
                                 self.show_class_mapping_editor(True))).pack(side="left")
        ctk.CTkButton(
            actions, text=I.t("manage_class_mappings"), height=38,
            fg_color=CARD, text_color=ACCENT, border_width=1,
            border_color=LINE, hover_color=ACCENT_SOFT,
            command=lambda: (window.destroy(), self.show_page("drug_classes"),
                             self.show_class_mapping_editor())).pack(side="left", padx=8)
        ctk.CTkButton(
            actions, text=I.t("close"), height=38, width=90,
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
        document = cfg.config.get("document_defaults", {}).copy()
        # _generate_full_document applies the configured margin_mm_value.
        document["_paper_size"] = self.paper_var.get()
        self._generate_full_document(rx, qr, document, path_pdf, path_docx)
        return rx

    def _generate_full_document(self, rx, qr, document, path_pdf=None, path_docx=None):
        """Generate a prepared document; safe to execute on a worker thread."""
        with self._document_lock:
            document_language = document.get("language", "interface")
            previous_language = I.get_lang()
            if document_language in ("en", "ar"):
                I.set_lang(document_language)
            try:
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
                        rx, path_docx, qr_pil_image=qr, **options)
            finally:
                I.set_lang(previous_language)

    def _prepare_full_document(self, action):
        rx = self.collect()
        if not self._validate(rx, action):
            return None
        qr = self._qr(rx)
        if qr is None:
            return None
        document = cfg.config.get("document_defaults", {}).copy()
        document["_paper_size"] = self.paper_var.get()
        return rx, qr, document

    def preview(self):
        path = os.path.join(tempfile.gettempdir(), f"rx_preview_{uuid.uuid4().hex}.pdf")
        prepared = self._prepare_full_document(I.t("preview"))
        if not prepared:
            return
        rx, qr, document = prepared

        def finished(_result):
            try:
                os.startfile(path)
            except Exception:
                webbrowser.open(path)

        self.submit_background(
            lambda: self._generate_full_document(rx, qr, document, path_pdf=path),
            finished, label=I.t("preview") + "…")

    def print_pdf(self):
        path = os.path.join(tempfile.gettempdir(), f"rx_print_{uuid.uuid4().hex}.pdf")
        prepared = self._prepare_full_document(I.t("print"))
        if not prepared:
            return
        rx, qr, document = prepared

        def finished(_result):
            try:
                os.startfile(path, "print")
            except Exception:
                os.startfile(path)

        self.submit_background(
            lambda: self._generate_full_document(rx, qr, document, path_pdf=path),
            finished, label=I.t("print") + "…")

    def export_word(self):
        initial_dir = cfg.config.get("document_defaults", {}).get("export_folder", "")
        path = filedialog.asksaveasfilename(
            title=I.t("export_word"), defaultextension=".docx",
            initialdir=initial_dir if Path(initial_dir).is_dir() else None,
            filetypes=[("Word documents", "*.docx")])
        if not path:
            return
        prepared = self._prepare_full_document(I.t("export_word"))
        if not prepared:
            return
        rx, qr, document = prepared

        def finished(_result):
            messagebox.showinfo(I.t("export_word"), I.t("msg_exported_word", path=path))
            try:
                os.startfile(path)
            except Exception:
                webbrowser.open(path)

        self.submit_background(
            lambda: self._generate_full_document(rx, qr, document, path_docx=path),
            finished, label=I.t("export_word") + "…")

    def export_label(self):
        rx = self.collect()
        if not self._validate(rx, I.t("export_compact")):
            return
        initial_dir = cfg.config.get("document_defaults", {}).get("export_folder", "")
        path = filedialog.asksaveasfilename(
            title=I.t("export_compact"), defaultextension=".docx",
            initialdir=initial_dir if Path(initial_dir).is_dir() else None,
            filetypes=[("Word documents", "*.docx")])
        if not path:
            return
        qr = self._qr(rx)
        if qr is None:
            return
        def finished(_result):
            messagebox.showinfo(
                I.t("export_compact"), I.t("msg_exported_compact", path=path))
            try:
                os.startfile(path)
            except Exception:
                webbrowser.open(path)

        self.submit_background(
            lambda: pdfgen.generate_medication_label_docx(
                rx, path, qr_pil_image=qr),
            finished, label=I.t("export_compact") + "…")

    def open_settings(self):
        SettingsWindow(self)


class SettingsWindow(ctk.CTkToplevel):
    """Organized application settings with category navigation and one save flow."""

    SECTIONS = (
        ("general", "⚒", "General"),
        ("clinic", "✚", "Clinic identity"),
        ("documents", "▤", "Documents"),
        ("qr", "▦", "QR verification"),
        ("database", "⌬", "Database"),
        ("gemini", "✦", "Gemini reference"),
        ("security", "▣", "Backup & security"),
        ("about", "ⓘ", "About"),
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
            self.document_language_var, self.document_header_var,
            self.export_folder_var, self.auto_backup_var,
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
        self.settings_search_entry = ctk.CTkEntry(
            self.sidebar, textvariable=self.settings_search_var, height=38,
            placeholder_text=I.t("settings_search"), border_color=LINE,
            corner_radius=9, font=ctk.CTkFont(size=13))
        self.settings_search_entry.pack(fill="x", padx=12, pady=(16, 5))
        self.settings_search_var.trace_add(
            "write", lambda *_: self._filter_settings_sections())
        ctk.CTkLabel(
            self.sidebar, text=I.t("settings_categories"), text_color=MUTED,
            font=ctk.CTkFont(size=12, weight="bold"), anchor="w").pack(
                fill="x", padx=22, pady=(8, 6))
        for key, symbol, _label in self.SECTIONS:
            icon = _glyph_icon(symbol, ICON_BLUE, 23)
            self._settings_nav_icons[key] = icon
            button = ctk.CTkButton(
                self.sidebar,
                text=I.t('settings_' + key) if key != 'security' else I.t('settings_backup_security'),
                image=icon, compound="left",
                anchor="w", height=43, corner_radius=10, fg_color="transparent",
                hover_color=ACCENT_SOFT, text_color=ACCENT,
                font=ctk.CTkFont(size=16),
                command=lambda section=key: self._show_section(section))
            button.pack(fill="x", padx=12, pady=2)
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
            footer, text=I.t("settings_reset_all"), width=110,
            height=40, fg_color="transparent", hover_color="#fae9e9",
            text_color=DANGER, border_width=1, border_color=LINE,
            command=self.reset_all_settings).pack(side="left", padx=4)
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
            self.content_host, fg_color=CARD, corner_radius=10)
        page.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        page.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            page, text=title, text_color="#153a38", anchor="w",
            font=ctk.CTkFont(size=24, weight="bold")).grid(
                row=0, column=0, sticky="ew", padx=20, pady=(4, 6))
        self._pages[key] = page
        return page

    def _card(self, parent, row, title=None):
        card = ctk.CTkFrame(
            parent, fg_color="#fbfdfd", corner_radius=11, border_width=1,
            border_color=LINE)
        card.grid(row=row, column=0, sticky="ew", padx=20, pady=(0, 9))
        card.grid_columnconfigure(0, weight=1)
        if title:
            ctk.CTkLabel(
                card, text=title, text_color="#244e4b", anchor="w",
                font=ctk.CTkFont(size=14, weight="bold")).grid(
                    row=0, column=0, sticky="ew", padx=14, pady=(10, 6))
        return card

    @staticmethod
    def _entry(parent, variable, row, label, show=None):
        ctk.CTkLabel(
            parent, text=label, text_color=MUTED, anchor="w",
            font=ctk.CTkFont(size=12, weight="bold")).grid(
                row=row, column=0, sticky="ew", padx=14, pady=(5, 3))
        entry = ctk.CTkEntry(
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
            left_pad = 14 if column == 0 else 5
            right_pad = 14 if column == 2 else 5
            ctk.CTkLabel(
                card, text=label, text_color=MUTED, anchor="w",
                font=ctk.CTkFont(size=12, weight="bold")).grid(
                    row=1, column=column, sticky="ew",
                    padx=(left_pad, right_pad), pady=(2, 3))
            ctk.CTkEntry(
                card, textvariable=variable, height=36, corner_radius=8,
                border_color=LINE, font=ctk.CTkFont(size=13)).grid(
                    row=2, column=column, sticky="ew",
                    padx=(left_pad, right_pad), pady=(0, 10))
        logo_card = self._card(page, 3, I.t("settings_logo"))
        logo_card.grid_columnconfigure(0, weight=1)
        self.logo_entry = ctk.CTkEntry(
            logo_card, textvariable=self.logo_var, height=36, corner_radius=8,
            border_color=LINE, font=ctk.CTkFont(size=13))
        self.logo_entry.grid(row=1, column=0, sticky="ew", padx=(14, 6), pady=(0, 10))
        ctk.CTkButton(
            logo_card, text=I.t("settings_browse"), width=88, height=36,
            fg_color=ACCENT_SOFT, hover_color=LINE, text_color=ACCENT,
            command=self.choose_logo).grid(row=1, column=1, padx=(0, 6), pady=(0, 10))
        ctk.CTkButton(
            logo_card, text=I.t("settings_remove"), width=88, height=36,
            fg_color=CARD, hover_color="#fae9e9", text_color=DANGER,
            border_width=1, border_color=LINE,
            command=self.remove_logo).grid(
                row=1, column=2, padx=(0, 14), pady=(0, 10))
        preview = self._card(page, 4, I.t("settings_header_preview"))
        self.clinic_preview_card = preview
        preview.grid_columnconfigure(1, weight=1)
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
            preview, text="", text_color="#244e4b", anchor="w",
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
        ctk.CTkOptionMenu(
            card, values=list(cfg.PAPER_SIZES.keys()), variable=self.paper_var,
            height=44, corner_radius=9, fg_color=ACCENT_SOFT,
            button_color=ACCENT, button_hover_color=ACCENT_HOVER,
            text_color="#244e4b", font=ctk.CTkFont(size=15)).grid(
                row=2, column=0, sticky="ew", padx=20, pady=(0, 10))
        ctk.CTkOptionMenu(
            card, values=[I.t("settings_same_as_interface"), "English", "العربية"],
            variable=self.document_language_var, height=44, corner_radius=9,
            fg_color=ACCENT_SOFT, button_color=ACCENT,
            button_hover_color=ACCENT_HOVER, text_color="#244e4b",
            font=ctk.CTkFont(size=15)).grid(
                row=2, column=1, sticky="ew", padx=20, pady=(0, 10))
        ctk.CTkCheckBox(
            card, text=I.t("settings_show_header"),
            variable=self.document_header_var, fg_color=ACCENT,
            hover_color=ACCENT_HOVER).grid(
                row=3, column=0, columnspan=2, sticky="w",
                padx=20, pady=(4, 14))
        folder = self._card(page, 3, I.t("settings_export_folder"))
        folder.grid_columnconfigure(0, weight=1)
        ctk.CTkEntry(
            folder, textvariable=self.export_folder_var, height=42,
            border_color=LINE).grid(
                row=1, column=0, sticky="ew", padx=(20, 8), pady=(0, 16))
        ctk.CTkButton(
            folder, text=I.t("settings_browse"), width=100, height=42,
            fg_color=ACCENT_SOFT, text_color=ACCENT,
            command=self.choose_export_folder).grid(
                row=1, column=1, padx=(0, 20), pady=(0, 16))

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
        ctk.CTkLabel(
            card, text=I.t("settings_viewer_configured"),
            fg_color=ACCENT_SOFT, text_color=ACCENT, corner_radius=12,
            height=28, font=ctk.CTkFont(size=12, weight="bold")).grid(
                row=5, column=0, sticky="w", padx=20, pady=(0, 16))

    def _build_database_page(self):
        page = self._new_page(
            "database", I.t("settings_database"), I.t("settings_database_tip"))
        card = self._card(page, 2, I.t("settings_current_database"))
        status = ctk.CTkFrame(card, fg_color="transparent")
        status.grid(row=1, column=0, sticky="ew", padx=14, pady=(1, 3))
        status.grid_columnconfigure(0, weight=1)
        self.database_path_label = ctk.CTkLabel(
            status, text=str(self.master.db.path), text_color="#244e4b",
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
            treatment_status, text="", text_color="#244e4b", anchor="w",
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
        foreground = ACCENT if primary else CARD
        text_color = "white" if primary else DANGER if danger else ACCENT
        hover = ACCENT_HOVER if primary else "#fae9e9" if danger else ACCENT_SOFT
        button = ctk.CTkButton(
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
        card = self._card(page, 2, I.t("gemini_settings"))
        ctk.CTkCheckBox(
            card, text=I.t("gemini_enable"), variable=self.gemini_enabled_var,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
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
        self.gemini_test_button = ctk.CTkButton(
            actions, text=I.t("gemini_test"), height=34,
            fg_color=ACCENT_SOFT, hover_color=LINE, text_color=ACCENT,
            command=self.test_gemini_connection)
        self.gemini_test_button.pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            actions, text=I.t("gemini_remove_key"), height=34,
            fg_color=CARD, hover_color="#fae9e9", text_color=DANGER,
            border_width=1, border_color=LINE,
            command=self.remove_gemini_key).pack(side="left")
        privacy = self._card(page, 3, I.t("settings_privacy"))
        self.gemini_privacy_card = privacy
        ctk.CTkLabel(
            privacy, text=I.t("settings_gemini_privacy"), text_color=MUTED,
            justify="left", anchor="w", wraplength=700,
            font=ctk.CTkFont(size=11)).grid(
                row=1, column=0, sticky="ew", padx=14, pady=(1, 9))

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
        ctk.CTkButton(
            actions, text=I.t("settings_create_backup"), height=34,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self.create_backup).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            actions, text=I.t("settings_restore_backup"), height=34,
            fg_color=ACCENT_SOFT, hover_color=LINE, text_color=ACCENT,
            command=self.master.restore_backup).pack(side="left")
        ctk.CTkButton(
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
            variable=self.auto_backup_var, fg_color=ACCENT,
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
            "about": "about version diagnostic log data folder privacy",
        }
        matches = []
        for key, button in self._nav_buttons.items():
            haystack = f"{button.cget('text')} {keywords.get(key, '')}".casefold()
            button.pack_forget()
            if not query or query in haystack:
                button.pack(fill="x", padx=12, pady=2)
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
                cfg.config.viewer_base_url = viewer_url
                cfg.config.paper_size = self.paper_var.get()
                cfg.config.language = language
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
            self.document_language_var.set(I.t("settings_same_as_interface"))
            self.document_header_var.set(True)
            self.export_folder_var.set("")
        elif self._active_section == "qr":
            self.viewer_var.set(cfg.DEFAULT_VIEWER_BASE)
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
        self.clinic_address_var.set("")
        self.clinic_phone_var.set("")
        self.logo_var.set("")
        self.paper_var.set(cfg.DEFAULT_PAPER)
        self.document_language_var.set(I.t("settings_same_as_interface"))
        self.document_header_var.set(True)
        self.export_folder_var.set("")
        self.viewer_var.set(cfg.DEFAULT_VIEWER_BASE)
        self.gemini_enabled_var.set(False)
        self.gemini_key_var.set("")
        self.auto_backup_var.set(False)

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
        if not self.document_header_var.get():
            self.clinic_preview_text.configure(text=I.t("settings_header_hidden"))
            self.clinic_preview_logo.configure(
                text="—", image=self._clinic_preview_placeholder)
            self._clinic_preview_image = None
            return
        name = self.clinic_name_var.get().strip() or I.t("settings_clinic_name")
        details = "  |  ".join(
            item for item in (self.clinic_address_var.get().strip(),
                              self.clinic_phone_var.get().strip()) if item)
        self.clinic_preview_text.configure(text=name + ("\n" + details if details else ""))
        logo_path = self.logo_var.get().strip()
        try:
            if logo_path and Path(logo_path).is_file():
                pixels = 54
                image = Image.open(logo_path).convert("RGBA")
                image.thumbnail((pixels, pixels), Image.Resampling.LANCZOS)
                self._clinic_preview_image = ctk.CTkImage(
                    light_image=image, dark_image=image, size=image.size)
                self.clinic_preview_logo.configure(
                    text="", image=self._clinic_preview_image)
                return
        except Exception:
            logging.exception("Clinic logo preview could not be rendered")
        self._clinic_preview_image = None
        self.clinic_preview_logo.configure(
            text="Rx", image=self._clinic_preview_placeholder)

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

    def open_log(self):
        cfg.LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        cfg.LOG_PATH.touch(exist_ok=True)
        os.startfile(str(cfg.LOG_PATH))

    def copy_diagnostics(self):
        document = cfg.config.get("document_defaults", {})
        text = "\n".join((
            f"App: Rx Prescription Printer {cfg.APP_VERSION}",
            f"Data: {cfg.APP_DIR}", f"Log: {cfg.LOG_PATH}",
            f"Database: {self.master.db.path}",
            f"Medicines: {len(self.master.db.drugs)}",
            f"Paper: {self.paper_var.get()}",
            f"Document defaults: {document}",
        ))
        self.clipboard_clear()
        self.clipboard_append(text)
        messagebox.showinfo(
            I.t("settings_diagnostics"), I.t("settings_diagnostics_copied"), parent=self)


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
