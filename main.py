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
import re
import functools
import logging
import math
import os
import tempfile
import threading
import time
import json
import zipfile
import tkinter as tk
import tkinter.font as tkfont
import unicodedata
import uuid
import weakref
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog
from typing import List
from collections import OrderedDict
from contextlib import contextmanager

if os.environ.get("RX_PERF_DEBUG") == "1":
    import perf_probe

import customtkinter as ctk
from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageOps

import config as cfg
import drug_db as dbmod
class _LazyPdfModule:
    """Keep ReportLab/font setup out of startup; load once before first use."""
    def __init__(self):
        self._module = None
        self._lock = threading.Lock()

    def __getattr__(self, name):
        with self._lock:
            if self._module is None:
                # Static nested import remains visible to PyInstaller's analysis.
                import pdf_generator
                self._module = pdf_generator
            return getattr(self._module, name)


pdfgen = _LazyPdfModule()
import qr_utils as qu
import cloud_rx
import clinic_location
import i18n as I
import openfda
import drug_classes as classes
import print_layout
from workflow_features import (
    calculate_medicine_quantity,
    compact_medicine_summary,
    compare_prescriptions,
    prescription_draft,
    valid_prescription_draft,
    validate_prescription_workflow,
)
from patient_history import PatientHistory, PatientHistoryError

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
CARD_RADIUS = 12
CARD_NAME_SIZE = 15
CARD_DETAIL_SIZE = 12
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


def glass_panel_image(size, surface, region=(0, 0, 1, 1), tint=None, clip=None):
    """Composite a panel against its portion of the shared window backdrop."""
    width, height, radius, edge = size
    left, top, right, bottom = clip or (0, 0, width, height)
    output_size = (right - left, bottom - top)
    x0, y0, x1, y1 = region
    region = (x0 + (x1-x0)*left/width, y0 + (y1-y0)*top/height,
              x0 + (x1-x0)*right/width, y0 + (y1-y0)*bottom/height)
    extent = tuple(value * 96 for value in region)
    image = glass_texture("workspace").transform(
        output_size, Image.Transform.EXTENT, extent, Image.Resampling.BILINEAR)
    veil = {"workspace": 0, "sidebar": .74, "panel": .88, "sheet": .78}.get(surface, .88)
    target = SURFACE if surface == "sheet" else SIDEBAR_SURFACE if surface == "sidebar" else CARD
    image = Image.blend(image, Image.new("RGB", image.size, target), veil)
    if surface != "workspace":
        reflection = glass_reflection(surface).transform(
            image.size, Image.Transform.EXTENT,
            (left*96/width, top*96/height, right*96/width, bottom*96/height),
            Image.Resampling.BILINEAR)
        image = Image.composite(Image.new("RGB", image.size, "white"), image, reflection)
    if tint:
        image = Image.blend(image, Image.new("RGB", image.size, tint), .32)
    image = image.convert("RGBA")
    mask = Image.new("L", image.size)
    if width > 2 * edge and height > 2 * edge:
        bounds = (edge-left, edge-top, width-edge-1-left, height-edge-1-top)
        ImageDraw.Draw(mask).rounded_rectangle(bounds, radius=max(0, radius - edge), fill=255)
        if surface != "workspace" and width > 12 and height > 12:
            # One native outer border; only a quiet top-edge white reflection.
            # No inset outline/shadow, which previously looked double-framed.
            ImageDraw.Draw(image).line(
                (edge+radius-left, edge+1-top, width-edge-radius-1-left, edge+1-top),
                fill=(255, 255, 255, 110), width=1)
    image.putalpha(mask)
    return image


class GlassFrame(ctk.CTkFrame):
    """Portable frosted surface, position-aware and rendered only when visible."""

    def __init__(self, master, *, glass_surface="panel", **kwargs):
        if glass_surface == "panel" and kwargs.get("border_width", 0) in (1, 2):
            kwargs["corner_radius"] = CARD_RADIUS
        self._glass_surface = glass_surface
        self._glass_job = None
        self._glass_photo = None
        self._glass_size = None
        self._glass_viewport = None
        self._glass_generation = 0
        super().__init__(master, **kwargs)
        self.bind("<Map>", lambda _event: self._queue_glass(), add="+")
        self.bind("<Configure>", lambda _event: self._queue_glass(), add="+")
        self.bind("<Unmap>", lambda _event: self._release_glass(), add="+")
        window = self.winfo_toplevel()
        if not hasattr(window, "_glass_panels"):
            window._glass_panels = weakref.WeakSet()
            def window_resized(event, owner=window):
                if event.widget is owner:
                    size = (event.width, event.height)
                    if size == getattr(owner, "_glass_window_size", None):
                        return
                    owner._glass_window_size = size
                    owner._glass_resize_until = time.monotonic() + .20
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
                    def scroll_activity(_event=None, owner=ancestor):
                        for panel in list(owner._glass_frames):
                            panel._queue_glass()

                    # Keep CustomTkinter's native canvas/scrollbar command
                    # untouched. Its scrollbar redraw calls update_idletasks,
                    # so wrapping yscrollcommand can recursively re-enter the
                    # geometry engine on dense card pages. Passive bindings
                    # provide the glass refresh without changing scrolling.
                    ancestor._parent_canvas.bind(
                        "<Configure>", scroll_activity, add="+")
                    ancestor._parent_canvas.bind(
                        "<MouseWheel>", scroll_activity, add="+")
                    ancestor._parent_canvas.bind(
                        "<Button-4>", scroll_activity, add="+")
                    ancestor._parent_canvas.bind(
                        "<Button-5>", scroll_activity, add="+")
                ancestor._glass_frames.add(self)
                break
            ancestor = getattr(ancestor, "master", None)

    def _queue_glass(self):
        if not self.winfo_exists():
            return
        if not self._glass_visible():
            if self._glass_photo is not None:
                self._release_glass()
            return
        if self._glass_job is not None:
            self.after_cancel(self._glass_job)
        remaining = getattr(self.winfo_toplevel(), "_glass_resize_until", 0) - time.monotonic()
        self._glass_job = self.after(max(80, int(remaining * 1000) + 1), self._paint_glass)

    def _release_glass(self):
        self._glass_generation += 1
        self._glass_pending_key = None
        pending = getattr(self, "_glass_future", None)
        if pending is not None:
            pending.cancel()
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
        remaining = getattr(self.winfo_toplevel(), "_glass_resize_until", 0) - time.monotonic()
        if remaining > 0 and not force:
            self._glass_job = self.after(int(remaining * 1000) + 1, self._paint_glass)
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
        clip = (0, 0, width, height)
        if self._glass_viewport is not None:
            viewport = self._glass_viewport
            # Rasterize the exact visible slice, leaving corners in full-panel coordinates.
            top = max(0, viewport.winfo_rooty()-self.winfo_rooty())
            bottom = min(height, viewport.winfo_rooty()+viewport.winfo_height()-self.winfo_rooty())
            if bottom <= top:
                return
            clip = (0, top, width, bottom)
        key += (clip,)
        if key == self._glass_size and self._glass_photo is not None:
            return
        if key == getattr(self, "_glass_pending_key", None):
            return
        root = self._root()
        if not hasattr(root, "_glass_image_cache"):
            root._glass_image_cache = OrderedDict()
            root._glass_cache_bytes = 0
        cached = root._glass_image_cache.get(key)
        if cached is not None:
            root._glass_image_cache.move_to_end(key)
            self._apply_glass_image(key, cached[0], clip)
            return
        self._glass_generation += 1
        generation = self._glass_generation
        if hasattr(root, "_glass_executor"):
            self._glass_pending_key = key
            previous = getattr(self, "_glass_future", None)
            if previous is not None:
                previous.cancel()
            surface = self._glass_surface
            owner = weakref.ref(self)
            future = root._glass_executor.submit(glass_panel_image, size, surface, region, tint, clip)
            self._glass_future = future
            def completed(done):
                try:
                    image = done.result()
                except (Exception, concurrent.futures.CancelledError):
                    return
                def deliver():
                    panel = owner()
                    if (panel is None or root._closing or not panel.winfo_exists()
                            or panel._glass_generation != generation):
                        return
                    panel._glass_pending_key = None
                    panel._cache_glass_image(root, key, image, clip)
                if not root._closing:
                    root._background_callbacks.put(deliver)
            future.add_done_callback(completed)
            return
        # Standalone widget users retain synchronous compatibility.
        self._cache_glass_image(root, key,
            glass_panel_image(size, self._glass_surface, region, tint, clip), clip)

    def _cache_glass_image(self, root, key, image, clip):
        photo = ImageTk.PhotoImage(image, master=root)
        cost = image.width * image.height * 4
        existing = root._glass_image_cache.pop(key, None)
        if existing is not None:
            root._glass_cache_bytes -= existing[1]
        while root._glass_image_cache and (len(root._glass_image_cache) >= 32
                or root._glass_cache_bytes + cost > 32 * 1024 * 1024):
            _, (_, old_cost) = root._glass_image_cache.popitem(last=False)
            root._glass_cache_bytes -= old_cost
        if cost <= 32 * 1024 * 1024:
            root._glass_image_cache[key] = (photo, cost)
            root._glass_cache_bytes += cost
        self._apply_glass_image(key, photo, clip)

    def _apply_glass_image(self, key, photo, clip):
        self._glass_photo = photo
        self._glass_size = key
        self._canvas.delete("glass_surface")
        self._canvas.create_image(clip[0], clip[1], anchor="nw", image=self._glass_photo,
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

    def __init__(self, owner, model_var, *, default_justify="left",
                 dynamic_justify=False):
        self.owner = owner
        self.model_var = model_var
        self.default_justify = "right" if default_justify == "right" else "left"
        self.dynamic_justify = bool(dynamic_justify)
        self.display_var = tk.StringVar(
            master=model_var._root, value=directional_display_text(model_var.get()))
        self.widget = None
        self._syncing = False
        self._model_trace = model_var.trace_add("write", self._from_model)
        self._display_trace = self.display_var.trace_add("write", self._from_display)

    def attach(self, widget):
        self.widget = widget
        self._apply_justification(self.model_var.get())
        return widget

    def _entry_widget(self):
        return getattr(self.widget, "_entry", self.widget)

    def _apply_justification(self, logical=""):
        entry = self._entry_widget()
        if entry is not None:
            justify = self.default_justify
            if self.dynamic_justify and logical:
                justify = "right" if first_strong_direction(logical) == "rtl" else "left"
            try:
                entry.configure(justify=justify)
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
            self._apply_justification(logical)
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
        self._apply_justification(logical)
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


class MappingMedicineList(tk.Frame):
    """Compact, virtualized two-line medicine list with Listbox-like methods."""

    ROW_HEIGHT = 72

    def __init__(self, parent, *, bg, fg, line, select_bg, select_fg):
        super().__init__(parent, bg=bg, highlightthickness=1,
                         highlightbackground=line, bd=0)
        self._bg = bg
        self._fg = fg
        self._line = line
        self._select_bg = select_bg
        self._select_fg = select_fg
        self._rows = []
        self._trade_names = []
        self._selected = set()
        self._active = None
        self._anchor = None
        self._selection_callback = None
        self._selection_guard = None
        self._draw_job = None
        self._name_font = tkfont.Font(family="Segoe UI", size=-45, weight="bold")
        self._trade_font = tkfont.Font(family="Segoe UI", size=-22)
        self._meta_font = tkfont.Font(family="Segoe UI", size=-13)
        self.canvas = tk.Canvas(
            self, bg=bg, highlightthickness=0, bd=0,
            yscrollincrement=self.ROW_HEIGHT, takefocus=True)
        self.scrollbar = tk.Scrollbar(
            self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self._on_yview_changed)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Configure>", lambda _event: self._draw())
        self.canvas.bind("<Up>", lambda event: self._move_active(-1, event))
        self.canvas.bind("<Down>", lambda event: self._move_active(1, event))
        self.canvas.bind("<Prior>", lambda event: self._move_page(-1, event))
        self.canvas.bind("<Next>", lambda event: self._move_page(1, event))
        self.canvas.bind("<Home>", lambda event: self._move_to_edge(0, event))
        self.canvas.bind("<End>", lambda event: self._move_to_edge(-1, event))
        self.canvas.bind("<Control-a>", self._select_all)

    def configure(self, cnf=None, **kwargs):
        # Preserve the small Listbox surface used by the editor.
        kwargs.pop("selectmode", None)
        kwargs.pop("exportselection", None)
        if cnf or kwargs:
            return super().configure(cnf or {}, **kwargs)
        return super().configure()

    config = configure

    def bind(self, sequence=None, func=None, add=None):
        if sequence == "<<ListboxSelect>>":
            self._selection_callback = func
            return None
        return super().bind(sequence, func, add)

    def set_selection_guard(self, callback):
        self._selection_guard = callback

    def insert(self, _index, name, metadata="", *, status="", status_kind="", trade_name=""):
        self._rows.append((str(name), str(metadata), str(status), str(status_kind)))
        self._trade_names.append(str(trade_name))

    def refresh(self):
        """Commit a batch of inserted rows in one redraw."""
        self._update_scrollregion()
        self._draw()

    def delete(self, _first, _last=None):
        self._rows.clear()
        self._trade_names.clear()
        self._selected.clear()
        self._active = None
        self._anchor = None
        self._update_scrollregion()
        self._draw()

    def curselection(self):
        return tuple(sorted(self._selected))

    def selection_set(self, first, last=None):
        if not self._rows:
            return
        start = max(0, int(first))
        finish = start if last in (None, tk.END, "end") else min(int(last), len(self._rows) - 1)
        self._selected.update(range(start, finish + 1))
        self._anchor = start
        self._draw()

    def request_selection(self, proposed, active=None):
        proposed = {int(index) for index in proposed
                    if 0 <= int(index) < len(self._rows)}
        if proposed == self._selected and (active is None or active == self._active):
            return True
        if self._selection_guard:
            guarded = self._selection_guard(set(proposed))
            if guarded is False or guarded is None:
                return False
            if guarded is not True:
                proposed = {int(index) for index in guarded
                            if 0 <= int(index) < len(self._rows)}
        self._selected = proposed
        if active is None or active not in proposed:
            active = min(proposed) if proposed else None
        self._active = active
        self._anchor = active
        self._draw()
        if active is not None:
            self.see(active)
        if self._selection_callback:
            self._selection_callback(None)
        return True

    def selection_clear(self, _first, _last=None):
        self._selected.clear()
        self._draw()

    def activate(self, index):
        if self._rows:
            self._active = max(0, min(int(index), len(self._rows) - 1))
            self.canvas.focus_set()

    def see(self, index):
        if not self._rows:
            return
        index = max(0, min(int(index), len(self._rows) - 1))
        top = self.canvas.canvasy(0)
        bottom = top + max(1, self.canvas.winfo_height())
        row_top = index * self.ROW_HEIGHT
        row_bottom = row_top + self.ROW_HEIGHT
        total = len(self._rows) * self.ROW_HEIGHT
        if row_top < top:
            self.canvas.yview_moveto(row_top / max(total, 1))
        elif row_bottom > bottom:
            self.canvas.yview_moveto(max(0, row_bottom - self.canvas.winfo_height()) / max(total, 1))

    def yview(self, *args):
        return self.canvas.yview(*args)

    def yview_moveto(self, fraction):
        self.canvas.yview_moveto(fraction)

    def _update_scrollregion(self):
        self.canvas.configure(scrollregion=(0, 0, 1, len(self._rows) * self.ROW_HEIGHT))

    def _on_yview_changed(self, first, last):
        self.scrollbar.set(first, last)
        if self._draw_job is None:
            self._draw_job = self.after_idle(self._draw)

    def destroy(self):
        if self._draw_job is not None:
            self.after_cancel(self._draw_job)
            self._draw_job = None
        super().destroy()

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")
        return "break"

    def _on_click(self, event):
        if not self._rows:
            return "break"
        self.canvas.focus_set()
        index = int(self.canvas.canvasy(event.y) // self.ROW_HEIGHT)
        if not 0 <= index < len(self._rows):
            return "break"
        proposed = set(self._selected)
        if event.state & 0x0004:  # Ctrl extends/toggles selection.
            if index in self._selected:
                proposed.remove(index)
            else:
                proposed.add(index)
        elif event.state & 0x0001 and self._anchor is not None:  # Shift range.
            lo, hi = sorted((self._anchor, index))
            proposed = set(range(lo, hi + 1))
        else:
            proposed = {index}
        self.request_selection(proposed, index)
        return "break"

    def _move_active(self, delta, event):
        if not self._rows:
            return "break"
        current = self._active
        if current is None:
            current = min(self._selected) if self._selected else (0 if delta > 0 else len(self._rows) - 1)
        target = max(0, min(len(self._rows) - 1, current + delta))
        if event.state & 0x0001 and self._anchor is not None:
            lo, hi = sorted((self._anchor, target))
            proposed = set(range(lo, hi + 1))
        else:
            proposed = {target}
        self.request_selection(proposed, target)
        return "break"

    def _move_page(self, direction, event):
        page = max(1, self.canvas.winfo_height() // self.ROW_HEIGHT)
        return self._move_active(direction * page, event)

    def _move_to_edge(self, edge, event):
        if not self._rows:
            return "break"
        target = 0 if edge == 0 else len(self._rows) - 1
        if event.state & 0x0001 and self._anchor is not None:
            lo, hi = sorted((self._anchor, target))
            proposed = set(range(lo, hi + 1))
        else:
            proposed = {target}
        self.request_selection(proposed, target)
        return "break"

    def _select_all(self, _event):
        self.request_selection(set(range(len(self._rows))), 0 if self._rows else None)
        return "break"

    def _draw(self):
        if self._draw_job is not None:
            self.after_cancel(self._draw_job)
            self._draw_job = None
        self.canvas.delete("all")
        width = max(1, self.canvas.winfo_width())
        badge_palette = {
            "confirmed": ("#e2f2ec", SUCCESS),
            "suggested": (WARNING_SOFT, WARNING),
            "conflicting": (DANGER_SOFT, DANGER),
            "unclassified": (CARD, MUTED),
            "unrecognized": (DANGER_SOFT, DANGER),
        }
        # Fixed row heights permit viewport rendering with one-row overscan.
        top_visible = max(0, self.canvas.canvasy(0))
        first = max(0, int(top_visible // self.ROW_HEIGHT) - 1)
        last = min(len(self._rows),
                   int((top_visible + self.canvas.winfo_height()) // self.ROW_HEIGHT) + 2)
        for index in range(first, last):
            name, metadata, status, status_kind = self._rows[index]
            top = index * self.ROW_HEIGHT
            selected = index in self._selected
            fill = self._select_bg if selected else self._bg
            text_fill = self._select_fg if selected else self._fg
            meta_fill = self._select_fg if selected else MUTED
            self.canvas.create_rectangle(
                0, top, width, top + self.ROW_HEIGHT,
                fill=fill, outline=self._line, width=1)
            trade = self._trade_names[index]
            if trade and trade.casefold() == name.casefold():
                trade = ""
            name_room = max(40, int((width-28) * .58)) if trade else max(40, width-28)
            shown_name = self._fit_inline_name(name, self._name_font, name_room)
            self.canvas.create_text(
                12, top + 4, anchor="nw", text=shown_name, fill=text_fill,
                font=self._name_font)
            if trade:
                x = 24 + self._name_font.measure(shown_name)
                self.canvas.create_text(x, top + 18, anchor="nw",
                    text=self._fit_inline_name(trade, self._trade_font, max(20, width-x-12)),
                    fill=meta_fill, font=self._trade_font)
            badge_width = max(62, self._meta_font.measure(status) + 20) if status else 0
            self.canvas.create_text(
                14, top + 51, anchor="nw", text=metadata, fill=meta_fill,
                font=self._meta_font, width=max(40, width - badge_width - 42))
            if status:
                badge_fill, badge_text = badge_palette.get(status_kind, (CARD, MUTED))
                if selected:
                    badge_fill, badge_text = "#ffffff", self._select_bg
                left = max(14, width - badge_width - 12)
                self.canvas.create_rectangle(
                    left, top + 48, width - 12, top + 68,
                    fill=badge_fill, outline=badge_fill)
                self.canvas.create_text(
                    (left + width - 12) / 2, top + 58, text=status,
                    fill=badge_text, font=self._meta_font)

    @staticmethod
    def _fit_inline_name(text, font, available):
        if font.measure(text) <= available:
            return text
        low, high = 0, len(text)
        while low < high:
            middle = (low + high + 1) // 2
            if font.measure(text[:middle] + "…") <= available:
                low = middle
            else:
                high = middle - 1
        return text[:low] + "…"

# Common prescription frequencies.  The editable combo box keeps the list
# convenient while still allowing a clinician to enter a non-standard schedule.
FREQUENCY_OPTIONS = (
    "مرة واحدة يومياً",
    "مرتان يومياً",
    "ثلاث مرات يومياً",
    "أربع مرات يومياً",
    "كل 4 ساعات",
    "كل 6 ساعات",
    "كل 8 ساعات",
    "كل 12 ساعة",
    "عند الحاجة",
    "عند الحاجة كل 4 إلى 6 ساعات",
    "عند الحاجة كل 8 ساعات",
    "فوراً (جرعة واحدة)",
    "يوم بعد يوم",
    "مرة واحدة أسبوعياً",
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


def _font_field_height(font):
    """Use actual ascent/descent so Arabic glyphs survive large display sizes."""
    return max(FIELD_HEIGHT, font.cget("size") + 12, font.metrics("linespace") + 12)


def dropdown_font(baseline=None, master=None, role="medications"):
    """Role-specific display type, never applied to Settings or neutral menus."""
    baseline = baseline or _ui_font(13)
    ancestor = master
    while ancestor is not None:
        if getattr(ancestor, "_fixed_dropdown_fonts", False):
            return baseline
        ancestor = getattr(ancestor, "master", None)
    key = {"medications": "medication_font_size", "instructions": "instruction_font_size",
           "names": "name_font_size"}.get(role)
    size = cfg.config.ui_font_size(key) if key else 0
    if not size:
        return baseline
    if isinstance(baseline, ctk.CTkFont):
        return _ui_font(size, baseline.cget("weight"), baseline.cget("family"))
    return _ui_font(size)


class PopupListbox(tk.Listbox):
    def __init__(self, master, **kwargs):
        self._font_role = kwargs.pop("font_role", "medications")
        self._dropdown_font_baseline = kwargs.get("font", ("Segoe UI", 16))
        kwargs["font"] = dropdown_font(self._dropdown_font_baseline, master, self._font_role)
        super().__init__(master, **kwargs)

    def apply_preferences(self):
        self.configure(font=dropdown_font(self._dropdown_font_baseline, self.master, self._font_role))


class VisualMenu(tk.Menu):
    def __init__(self, master, **kwargs):
        self._dropdown_font_baseline = kwargs.get("font", ("Segoe UI", 16))
        kwargs["font"] = dropdown_font(self._dropdown_font_baseline, master, None)
        super().__init__(master, **kwargs)

    def apply_preferences(self):
        self.configure(font=dropdown_font(self._dropdown_font_baseline, self.master, None))


class VisualOptionMenu(ctk.CTkOptionMenu):
    def __init__(self, master, **kwargs):
        self._font_role = kwargs.pop("font_role", None)
        self._dropdown_font_baseline = kwargs.get("dropdown_font") or _ui_font(13)
        kwargs["dropdown_font"] = dropdown_font(self._dropdown_font_baseline, master, self._font_role)
        super().__init__(master, **kwargs)

    def apply_preferences(self):
        self.configure(dropdown_font=dropdown_font(self._dropdown_font_baseline, self.master, self._font_role))


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
    image = ctk.CTkImage(light_image=artwork, dark_image=artwork,
                         size=(display_size, display_size))
    image._edit_action = True
    return image


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
    elif symbol == "save":
        line([(5, 4), (19, 4), (20, 5), (20, 20), (4, 20), (4, 4), (5, 4)])
        line([(8, 4), (8, 10), (16, 10), (16, 4)])
        line([(8, 20), (8, 14), (16, 14), (16, 20)])
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
                          width=ICON_BUTTON_SIZE, height=ICON_BUTTON_SIZE,
                          border_width=0)
        if getattr(kwargs.get("image"), "_edit_action", False):
            kwargs["image"] = _edit_icon(18)
        if self._action_symbol is not None or getattr(kwargs.get("image"), "_edit_action", False):
            kwargs.update(width=ICON_BUTTON_SIZE, height=ICON_BUTTON_SIZE,
                          border_width=0, fg_color="transparent")
        super().__init__(master, **kwargs)

    def configure(self, **kwargs):
        if self._action_symbol is not None and "text_color" in kwargs and "text" not in kwargs:
            color = self._apply_appearance_mode(kwargs["text_color"])
            kwargs["image"] = action_icon(self._action_symbol, color)
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
        resolved = self._apply_appearance_mode(color)
        if self._focus_border is None and resolved not in {DANGER, WARNING}:
            self._focus_border = color
            self.configure(border_color=ACCENT)

    def _field_blurred(self, _event=None):
        if self._focus_border is not None and self.cget("border_color") == ACCENT:
            self.configure(border_color=self._focus_border)
        self._focus_border = None


class VisualEntry(FieldFocus, ctk.CTkEntry):
    def __init__(self, master, **kwargs):
        self._font_role = kwargs.pop("font_role", None)
        kwargs.setdefault("border_width", 1)
        super().__init__(master, **kwargs)
        self._bind_field_focus()
        self.apply_preferences()

    def apply_preferences(self):
        if self._font_role:
            font = dropdown_font(self.cget("font"), self.master, self._font_role)
            self.configure(font=font, height=_font_field_height(font))


class VisualComboBox(FieldFocus, ctk.CTkComboBox):
    def __init__(self, master, **kwargs):
        self._font_role = kwargs.pop("font_role", None)
        kwargs.setdefault("border_width", 1)
        self._dropdown_font_baseline = kwargs.get("dropdown_font") or _ui_font(13)
        kwargs["dropdown_font"] = dropdown_font(self._dropdown_font_baseline, master, self._font_role)
        super().__init__(master, **kwargs)
        self._bind_field_focus()
        self.apply_preferences()

    def apply_preferences(self):
        font = dropdown_font(self._dropdown_font_baseline, self.master, self._font_role)
        self.configure(dropdown_font=font)
        if self._font_role:
            self.configure(font=font, height=_font_field_height(font))


def polish_toolbar(frame):
    """One compact rhythm for search and text actions, not icon hit areas."""
    for widget in frame.winfo_children():
        if isinstance(widget, (ctk.CTkEntry, ctk.CTkComboBox, ctk.CTkOptionMenu)):
            widget.configure(height=FIELD_HEIGHT, corner_radius=9)
        elif isinstance(widget, VisualButton) and widget._action_symbol is None and widget.cget("image") is None:
            widget.configure(height=FIELD_HEIGHT, corner_radius=9)


def wrap_card_labels(frame):
    """Constrain name labels to their allocated columns, not their text width."""
    labels = [child for child in frame.winfo_children() if isinstance(child, ctk.CTkLabel)]
    if not labels:
        return
    actions = [child for child in frame.winfo_children()
               if child not in labels and child.winfo_manager() == "pack"]
    for child in labels + actions:
        child.pack_forget()
    for column, child in enumerate(actions, len(labels)):
        child.pack_forget()
        child.grid(row=0, column=column, sticky="ne", padx=2)
    for column, label in enumerate(labels):
        label.pack_forget()
        label.configure(width=1, wraplength=100)
        frame.grid_columnconfigure(column, weight=1, uniform="card_text")
        label.grid(row=0, column=column, sticky="nw", padx=(0, CARD_GAP))
    def resized(event):
        reserved = sum(child.winfo_reqwidth() + 4 for child in actions)
        available = max(20, (event.width-reserved) // len(labels) - CARD_GAP)
        if available == getattr(frame, "_text_wrap_width", None):
            return
        frame._text_wrap_width = available
        for label in labels:
            label.configure(wraplength=available)
    frame.bind("<Configure>", resized, add="+")


class DrugRow(GlassFrame):
    """One medication row with linked generic/trade and scientific-name inputs."""

    def __init__(self, master, db, on_change, on_remove,
                 on_move_up, on_move_down, on_drag, **kwargs):
        super().__init__(master, **kwargs)
        self.db = db
        self.on_change = on_change
        self.on_remove = on_remove
        self.on_move_up = on_move_up
        self.on_move_down = on_move_down
        self.on_drag = on_drag
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
        self.expanded = True

        self.name_var = tk.StringVar()
        self.trade_var = tk.StringVar()
        self.summary_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.summary_number = ctk.CTkLabel(
            self.summary_frame, text="1.", width=28, text_color=MUTED,
            font=_ui_font(12, "bold"))
        self.summary_number.pack(side="left", padx=(0, 5))
        self.summary_label = ctk.CTkLabel(
            self.summary_frame, text="", text_color=TEXT, anchor="w",
            justify="left", font=_ui_font(12, "bold"))
        self.summary_label.pack(side="left", fill="x", expand=True)
        VisualButton(
            self.summary_frame, text="", image=_edit_icon(18), width=30, height=30,
            fg_color="transparent", hover_color=ACCENT_SOFT,
            command=lambda: self.set_expanded(True)).pack(side="right", padx=2)
        VisualButton(
            self.summary_frame, text="🗑", width=30, height=30,
            fg_color="transparent", text_color=DANGER, hover_color=DANGER_SOFT,
            command=self.on_remove).pack(side="right", padx=2)
        VisualButton(
            self.summary_frame, text="↓", width=30, height=30,
            fg_color="transparent", text_color=ACCENT, hover_color=ACCENT_SOFT,
            command=lambda: self.on_move_down(self)).pack(side="right", padx=1)
        VisualButton(
            self.summary_frame, text="↑", width=30, height=30,
            fg_color="transparent", text_color=ACCENT, hover_color=ACCENT_SOFT,
            command=lambda: self.on_move_up(self)).pack(side="right", padx=1)
        for widget in (self.summary_frame, self.summary_number, self.summary_label):
            widget.bind("<Button-1>", lambda _event: self.set_expanded(True))

        self.names = ctk.CTkFrame(self, fg_color="transparent")
        self.names.pack(fill="x", padx=PAD, pady=(6, CARD_GAP))
        self.names.grid_columnconfigure((1, 2), weight=1, uniform="medication_names")
        self.number_badge = ctk.CTkLabel(
            self.names, text="1.", width=28, height=36,
            fg_color="transparent", text_color=MUTED,
            font=ctk.CTkFont(weight="bold", size=12))
        self.number_badge.grid(row=0, column=0, sticky="n", padx=(0, 4), pady=(30, 0))
        header_actions = ctk.CTkFrame(self.names, fg_color="transparent")
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
        trade_col = ctk.CTkFrame(self.names, fg_color="transparent")
        trade_col.grid(row=0, column=1, sticky="ew", padx=(0, 4))
        science_col = ctk.CTkFrame(self.names, fg_color="transparent")
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
        ctk.CTkLabel(science_col, text=I.t("scientific_name"), anchor="w",
                     text_color=MUTED,
                     font=LABEL_FONT).pack(fill="x", pady=(0, 2))
        self.name_entry = VisualEntry(
            science_col, textvariable=self.name_var,
            fg_color=SURFACE, text_color=TEXT,
            placeholder_text=I.t("scientific_name"), height=36,
            corner_radius=8, border_color=LINE,
            font=SCIENTIFIC_FONT)
        self.name_entry.pack(fill="x")
        # Each name box searches only its matching imported database column.
        self.name_entry.bind("<KeyRelease>", self._on_scientific_type)
        for entry in (self.trade_entry, self.name_entry):
            entry.bind("<Down>", self._ac_down)
            entry.bind("<Up>", self._ac_up)
            entry.bind("<Return>", self._ac_choose_current)
            entry.bind("<Escape>", lambda _event: self._hide_ac())
        self.trade_picker = VisualOptionMenu(
            trade_col, font_role="medications", values=[I.t("choose_linked_trade_name")], height=28,
            fg_color=ACCENT_SOFT, text_color=ACCENT, button_color=PRIMARY,
            button_hover_color=ACCENT_HOVER, command=self._choose_linked_trade)

        self.class_badge = ctk.CTkLabel(
            self, text="", height=22, corner_radius=11, fg_color=ACCENT_SOFT,
            text_color=ACCENT, font=ctk.CTkFont(size=10, weight="bold"), anchor="w")

        details = ctk.CTkFrame(self, fg_color="transparent")
        self.details = details
        details.pack(fill="x", padx=PAD, pady=(0, 6))
        # Retain the original five column weights; notes span the last two.
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
                                   I.t("ph_duration"), self.dur_var,
                                   default_justify="right")
        self.notes_entry = self._notes_box(details, 3, I.t("notes"),
                                           self.notes_var)
        # Notes occupy their original column plus the former Quantity column.
        self.notes_entry.master.grid_configure(columnspan=2)
        # Preserve the original requested widths (140px controls + 8px gutters)
        # so absorbing Quantity does not redistribute space to other fields.
        details.grid_columnconfigure(3, minsize=148)
        details.grid_columnconfigure(4, minsize=148)
        self.notes_entry.configure(width=288)
        for variable in (self.dosage_var, self.freq_var, self.dur_var):
            variable.trace_add("write", lambda *_: self._auto_quantity())

        self.name_var.trace_add("write", self._scientific_name_changed)
        for variable in (self.name_var, self.trade_var, self.dosage_var,
                         self.freq_var, self.dur_var, self.notes_var,
                         self.quantity_var):
            variable.trace_add("write", self._row_value_changed)
        for widget in (self.trade_entry, self.name_entry, self.dosage_entry,
                       self.freq_entry, self.dur_entry, self.notes_entry):
            widget.bind("<FocusIn>", lambda _event: self.set_expanded(True), add="+")

    def _row_value_changed(self, *_args):
        if getattr(self, "_loading_values", False):
            return
        self._refresh_compact_summary()
        self.on_change()

    def _refresh_compact_summary(self):
        self.summary_label.configure(text=compact_medicine_summary(self.get_data()))

    def set_expanded(self, expanded):
        if expanded:
            app = self.winfo_toplevel()
            for other in getattr(app, "rows", ()):
                if other is not self and other.expanded:
                    other.set_expanded(False)
        if expanded == self.expanded:
            return
        self.expanded = expanded
        if expanded:
            self.summary_frame.pack_forget()
            self.names.pack(fill="x", padx=PAD, pady=(6, CARD_GAP))
            self.details.pack(fill="x", padx=PAD, pady=(0, 6))
            self.after_idle(lambda: self.trade_entry.focus_set()
                            if self.winfo_exists() and self.expanded else None)
        else:
            self._hide_ac()
            self.names.pack_forget()
            self.details.pack_forget()
            self._refresh_compact_summary()
            self.summary_frame.pack(fill="x", padx=PAD, pady=5)

    def collapse_if_complete(self):
        if self.name_var.get().strip() or self.trade_var.get().strip():
            self.set_expanded(False)

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

    def _box(self, parent, column, label, ph, var, *, default_justify="left"):
        col = ctk.CTkFrame(parent, fg_color="transparent")
        col.grid(row=0, column=column, sticky="ew", padx=4)
        ctk.CTkLabel(col, text=label, anchor="w", text_color=MUTED,
                     font=LABEL_FONT).pack(fill="x", pady=(0, 2))
        binding = DirectionalTextBinding(
            self, var, default_justify=default_justify)
        self._bidi_bindings.append(binding)
        e = VisualEntry(col, textvariable=binding.display_var,
                         font_role="instructions",
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
            col, font_role="instructions", values=[directional_display_text(value) for value in FREQUENCY_OPTIONS],
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
        binding = DirectionalTextBinding(
            self, var, default_justify="right")
        self._bidi_bindings.append(binding)
        picker = VisualComboBox(
            col, font_role="instructions", values=[directional_display_text(value) for value in NOTE_OPTIONS],
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
        self._refresh_quantity_status()
        self.on_change()

    def _auto_quantity(self):
        if getattr(self, "_loading_values", False):
            return
        if self._quantity_manual:
            self._refresh_quantity_status()
            return
        self.quantity_var.set(calculate_medicine_quantity(
            self.dosage_var.get(), self.freq_var.get(), self.dur_var.get(), self._form))
        self._refresh_quantity_status()

    def _refresh_quantity_status(self):
        if not hasattr(self, "quantity_status"):
            return
        self.quantity_status.configure(
            text=I.t("workflow_quantity_manual") if self._quantity_manual
            else I.t("workflow_quantity_auto"),
            text_color=WARNING if self._quantity_manual else MUTED)

    def recalculate_quantity(self):
        self._quantity_manual = False
        self._auto_quantity()
        self.on_change()

    def set_position(self, number):
        """Keep the plain numbered name row accurate after reordering."""
        self.number_badge.configure(text=f"{number}.")
        self.summary_number.configure(text=f"{number}.")

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
        app = self.winfo_toplevel()
        if hasattr(app, "cancel_latest_search"):
            app.cancel_latest_search((id(self), "scientific"))
            app.cancel_latest_search((id(self), "trade"))
        if self._autocomplete_job is not None:
            try:
                self.after_cancel(self._autocomplete_job)
            except (tk.TclError, ValueError):
                pass
        self._autocomplete_token += 1
        token = self._autocomplete_token
        self._autocomplete_job = self.after(
            35, lambda: self._run_autocomplete(query, mode, token))

    def _run_autocomplete(self, query, mode, token):
        self._autocomplete_job = None
        app = self.winfo_toplevel()
        if getattr(app, "_database_loading", False):
            self._autocomplete_job = self.after(35, lambda: self._run_autocomplete(query, mode, token))
            return
        search = (self.db.search_scientific if mode == "scientific"
                  else self.db.search_prescribable)

        def finished(matches):
            if not self.winfo_exists():
                return
            current = (self.name_var.get().strip() if mode == "scientific"
                       else self.trade_var.get().strip())
            if token != self._autocomplete_token or current != query:
                return
            if matches:
                self._show_ac(matches, mode=mode)
            else:
                self._hide_ac()

        if hasattr(app, "submit_background"):
            if hasattr(app, "submit_latest_search"):
                app.submit_latest_search((id(self), mode), lambda: search(query, limit=20), finished)
            else:
                app.submit_background(lambda: search(query, limit=20), finished, silent=True)
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
        self.db = dbmod.DrugDatabase(cfg.config.drug_db_path, defer_load=True)
        self._database_loading = True
        self.patient_history = PatientHistory()
        self.rows: List[DrugRow] = []
        self._medication_patient_id = ""
        self._executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=3, thread_name_prefix="rx-worker")
        self._glass_executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=1, thread_name_prefix="rx-glass")
        self._persistence_executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=1, thread_name_prefix="rx-save")
        self._row_batch_depth = 0
        self._history_save_future = None
        self._background_tasks = 0
        self._background_callbacks = queue.SimpleQueue()
        self._closing = False
        self._reference_request_token = 0
        self._export_busy = False
        self._document_lock = threading.Lock()
        self._loaded_pages = set()
        self._classification_cache = None
        self._classification_generation = 0
        self._classification_future = None
        self._classification_waiters = {}
        self._latest_searches = {}
        self._response_times = {}
        self._latest_query_tokens = {}
        self._favorite_ac_job = None
        self._favorite_ac_token = 0
        self._favorite_render_limit = 60
        self._favorite_filter_signature = None
        self._word_preview_job = None
        self._comparison_job = None
        self._favorite_refresh_job = None
        self._class_search_job = None
        self._page_scroll_positions = {}
        self.word_preview_visible = False
        self._suspend_draft = True
        self._build_ui()

        self.after(40, self._drain_background_callbacks)
        self.after(500, self._maintain_card_retention)
        self.after_idle(self._start_automatic_backup)
        for button in self.page_buttons.values():
            button.configure(state="disabled")
        self.db_label.configure(text=I.t("database_loading"))
        def database_ready(_count):
            self._database_loading = False
            for button in self.page_buttons.values():
                button.configure(state="normal")
            self.db_label.configure(text=I.t("db_count", n=_count))
            self._prepare_classification_cache()
        def database_failed(error):
            self._database_loading = False
            for button in self.page_buttons.values():
                button.configure(state="normal")
            self.db_label.configure(text=I.t("database_load_failed"))
            messagebox.showerror(APP_TITLE, str(error), parent=self)
        self.after_idle(lambda: self.submit_background(
            self.db.load, database_ready, on_error=database_failed, silent=True))
        if cfg.config.recovered_unreadable_settings:
            self.after(250, self._show_settings_recovery_notice)

    def _show_settings_recovery_notice(self):
        """Explain the one-time reset caused by a foreign/corrupt DPAPI file."""
        if self._closing:
            return
        messagebox.showwarning(
            I.t("settings_recovered_title"),
            I.t("settings_recovered_message"),
            parent=self,
        )

    def destroy(self):
        self._closing = True
        if hasattr(self, "_executor"):
            self._executor.shutdown(wait=False, cancel_futures=True)
        if hasattr(self, "_glass_executor"):
            self._glass_executor.shutdown(wait=False, cancel_futures=True)
        if hasattr(self, "_persistence_executor"):
            # Explicit saves must finish even if the window is destroyed.
            self._persistence_executor.shutdown(wait=False, cancel_futures=False)
        if hasattr(self, "_glass_image_cache"):
            self._glass_image_cache.clear()
        super().destroy()

    def _drain_background_callbacks(self):
        if self._closing:
            return
        deadline = time.perf_counter() + .008
        for _ in range(64):
            if time.perf_counter() >= deadline:
                break
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
            self.after(16, self._drain_background_callbacks)

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
        self._dashboard_group_labels = []
        self._add_dashboard_group(I.t("nav_create_prescription"))
        self._add_page_button("patient", I.t("patient_details"))
        self._add_page_button("medications", I.t("medication_entry"))
        self._add_dashboard_group(I.t("nav_reusable_content"))
        self._add_page_button("favorites", I.t("favorite_drugs"))
        self._add_page_button("treatment_templates", I.t("treatment_templates"))
        self._add_dashboard_group(I.t("nav_clinical_reference"))
        self._add_page_button("drug_classes", I.t("drug_classes"))
        self._add_page_button("interaction_review", I.t("interaction_review"))
        self._add_page_button("reference", I.t("online_drug_reference"))
        self._add_dashboard_group(I.t("nav_administration"))
        self._add_page_button("prescriber", I.t("prescriber_details"))
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
            database_card, text=I.t("database_loading"),
            text_color=ACCENT, font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w")
        self.db_label.pack(fill="x", padx=10, pady=(0, 7))
        self._apply_dashboard_density()

        self.content = ctk.CTkFrame(self.workspace, fg_color="transparent")
        self.content.pack(side="left", fill="both", expand=True, padx=(0, 4), pady=4)
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
        self._suspend_draft = False
        self._offer_resume_draft()
        self.apply_ui_font_preferences()

    def confirm_close(self):
        """Require an explicit confirmation before closing the desktop app."""
        if getattr(self, "_export_busy", False):
            messagebox.showinfo(APP_TITLE, I.t("export_close_pending"), parent=self)
            return
        if getattr(self, "_history_save_future", None) is not None:
            messagebox.showinfo(APP_TITLE, I.t("patient_save_pending"), parent=self)
            return
        if messagebox.askyesno(
                I.t("confirm_close_title"), I.t("confirm_close_message"), parent=self):
            if (getattr(self, "active_page", "") == "drug_classes" and
                    getattr(self, "_class_subpage_kind", "") == "mapping" and
                    not self._confirm_mapping_unsaved_changes()):
                return
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

    def _record_response_time(self, operation, started):
        """Bounded local diagnostics: fixed operation names and durations only."""
        elapsed = round((time.perf_counter() - started) * 1000, 2)
        samples = self._response_times.setdefault(operation, [])
        samples.append(elapsed)
        del samples[:-30]
        if elapsed >= 100:
            logging.info("UI timing %s: %.2f ms", operation, elapsed)

    def submit_latest_search(self, key, worker, on_success):
        """Cancel queued superseded searches; never deliver an obsolete result."""
        previous = self._latest_searches.get(key)
        if previous:
            previous.cancel()
        started = time.perf_counter()
        future = None

        def finished(result):
            if self._latest_searches.get(key) is future:
                self._latest_searches.pop(key, None)
                self._record_response_time("autocomplete", started)
                on_success(result)

        def failed(_error):
            if self._latest_searches.get(key) is future:
                self._latest_searches.pop(key, None)

        future = self.submit_background(worker, finished, on_error=failed, silent=True)
        self._latest_searches[key] = future
        return future

    def cancel_latest_search(self, key):
        previous = self._latest_searches.pop(key, None)
        if previous:
            previous.cancel()

    def _start_automatic_backup(self):
        """Compress a captured configuration off-thread without overwriting live settings."""
        if self._closing or not cfg.config.get("auto_backup_enabled"):
            return
        today = datetime.datetime.now(datetime.timezone.utc).date().isoformat()
        if str(cfg.config.get("last_backup_at", "")).startswith(today):
            return
        data = copy.deepcopy(cfg.config.data)
        database = Path(cfg.config.drug_db_path)
        started = time.perf_counter()

        def worker():
            directory = cfg.APP_DIR / "backups"
            directory.mkdir(parents=True, exist_ok=True)
            target = directory / f"rx-backup-{today}.rxbackup"
            temporary = target.with_suffix(".rxbackup.tmp")
            protected = {"format": "dpapi-v1", "data": cfg.protect(
                json.dumps(data, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))}
            db_bytes = database.read_bytes() if database.is_file() else None
            history_data = PatientHistory().backup_bytes()
            try:
                with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                    archive.writestr("manifest.json", json.dumps({
                        "format": cfg.BACKUP_FORMAT, "version": 1,
                        "has_drug_database": db_bytes is not None}))
                    archive.writestr("config.json", json.dumps(protected))
                    if db_bytes is not None:
                        archive.writestr("drug_database.csv", db_bytes)
                    if history_data is not None:
                        archive.writestr("patient_history.json", history_data)
                temporary.replace(target)
            finally:
                temporary.unlink(missing_ok=True)
            backups = sorted(directory.glob("rx-backup-*.rxbackup"),
                             key=lambda item: item.stat().st_mtime, reverse=True)
            for old in backups[10:]:
                old.unlink(missing_ok=True)
            return str(target)

        def completed(path):
            cfg.config.data["last_backup_at"] = datetime.datetime.now(
                datetime.timezone.utc).isoformat(timespec="seconds")
            cfg.config.data["last_backup_path"] = path
            cfg.config.save()
            self._record_response_time("startup-backup", started)

        self.submit_background(worker, completed,
                               on_error=lambda _error: logging.error("Automatic backup failed"),
                               silent=True)

    def _prepare_classification_cache(self):
        if self._closing:
            return
        previous = self._classification_future
        if previous:
            previous.cancel()
        generation = self._classification_generation
        database = self.db
        started = time.perf_counter()

        def completed(index):
            if generation != self._classification_generation:
                return
            self._classification_cache = index
            self._classification_future = None
            self._record_response_time("classification-index", started)
            waiters = list(self._classification_waiters.values())
            self._classification_waiters.clear()
            for callback in waiters:
                self.after_idle(callback)
            if getattr(self, "active_page", "") == "drug_classes":
                self.refresh_class_overview()

        def failed(_error):
            if generation == self._classification_generation:
                self._classification_future = None
                logging.error("Classification preparation failed")
                self._classification_waiters.clear()

        self._classification_future = self.submit_background(
            lambda: self._build_classification_index(tuple(database.drugs)), completed,
            on_error=failed, silent=True)

    def _defer_classification_render(self, callback):
        if getattr(self, "_classification_future", None) is None:
            return False
        self._classification_waiters[callback.__name__] = callback
        return True

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

    def _add_dashboard_group(self, text):
        """Add a compact navigation heading that disappears with the icon rail."""
        label = ctk.CTkLabel(
            self.dashboard, text=text.upper(), text_color=MUTED, anchor="w",
            font=_ui_font(9, "bold"))
        label.pack(fill="x", padx=15, pady=(8, 2))
        self._dashboard_group_labels.append((label, text.upper()))

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
            for label, _text in getattr(self, "_dashboard_group_labels", []):
                label.configure(text="", height=1)
                label.pack_configure(pady=0)
        else:
            self.dashboard_title.pack(side="left", fill="x", expand=True,
                                      before=self.dashboard_toggle)
            for label, text in getattr(self, "_dashboard_group_labels", []):
                label.configure(text=text, height=0)
                label.pack_configure(pady=(8, 2))
            self.dashboard_footer.pack(side="bottom", fill="x", padx=10, pady=(4, 6))
            self.dashboard_database_card.pack(side="bottom", fill="x", padx=10, pady=(4, 6))

    def show_page(self, key):
        started = time.perf_counter()
        self._ensure_page_built(key)
        if (key != "drug_classes" and
                getattr(self, "active_page", "") == "drug_classes" and
                getattr(self, "_class_subpage_kind", "") == "mapping"):
            if not self._confirm_mapping_unsaved_changes():
                return False
            self._capture_mapping_editor_state()
        if (key != "treatment_templates" and
                getattr(self, "active_page", "") == "treatment_templates" and
                not self._confirm_treatment_leave()):
            return False
        previous = getattr(self, "active_page", "")
        if previous == key:
            if hasattr(self, "medication_subpage"):
                self.close_medication_subpage()
            if key == "medications":
                self._set_workflow_step(2)
            self._resume_page_render(key)
            self._record_response_time("navigate:" + key, started)
            return True
        if previous and hasattr(self, "scroll"):
            try:
                self._page_scroll_positions[previous] = self.scroll._parent_canvas.yview()[0]
            except (tk.TclError, IndexError):
                pass
        if hasattr(self, "medication_subpage"):
            self.close_medication_subpage()
        for name, page in self.pages.items():
            if name == previous:
                page.pack_forget()
            if name not in {previous, key}:
                continue
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
        self._resume_page_render(key)
        # Hide only the visual rail; the canvas and wheel scrolling remain active.
        if key in SCROLLBAR_HIDDEN_PAGES:
            self.scroll._scrollbar.grid_remove()
        else:
            self.scroll._scrollbar.grid()
        self.active_page = key
        if hasattr(self, "action"):
            if key in {"patient", "medications"}:
                self.action.pack_forget()
            elif not self.action.winfo_manager():
                self.action.pack(side="bottom", fill="x", padx=4, pady=(0, 4))
        if key == "patient":
            self._set_workflow_step(1)
            self.after_idle(self.refresh_prescription_comparison)
            pending = getattr(self, "_pending_history_render", None)
            if pending is not None:
                self._pending_history_render = None
                self.after(1, lambda: self._render_prescription_batch(*pending))
        elif key == "medications" and not self.medication_subpage.winfo_manager():
            self._set_workflow_step(2)
        target_position = self._page_scroll_positions.get(key, 0.0)
        self.after_idle(lambda value=target_position:
                        self.scroll._parent_canvas.yview_moveto(value))
        if key not in self._loaded_pages:
            self._loaded_pages.add(key)
            self.after_idle(lambda page=key: self._load_page_data(page))

        self._record_response_time("navigate:" + key, started)

    def _resume_page_render(self, key):
        attribute = {"favorites": "_pending_favorite_render",
                     "treatment_templates": "_pending_template_render"}.get(key)
        if attribute:
            pending = getattr(self, attribute, None)
            if pending is not None:
                setattr(self, attribute, None)
                self.after(1, pending)

    def _load_page_data(self, key):
        """Populate expensive pages only when the clinician first opens them."""
        if self._closing:
            return
        if key == "drug_classes" and self._classification_future is not None:
            self.after(40, lambda: self._load_page_data(key))
            return
        if key == "patient":
            self.refresh_patient_history()
        elif key == "favorites":
            self.refresh_favorites_page()
        elif key == "drug_classes":
            self.refresh_class_overview()
        elif key == "treatment_templates":
            if (getattr(self, "_treatment_view", "saved") == "saved"
                    and getattr(self, "_treatment_saved_return_state", None)):
                state = self._treatment_saved_return_state
                self._treatment_saved_return_state = None
                self._restore_saved_template_state(state)
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
            row,
            # Only the name, not licence/specialty, participates in the name group.
            font_role="names" if var is getattr(self, "doctor_vars", {}).get("name") else None,
            textvariable=display_var, width=width, height=FIELD_HEIGHT,
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

    def _workflow_patient_data(self):
        return {key: variable.get().strip() for key, variable in self.patient_vars.items()}

    def _workflow_medicine_data(self):
        return [row.get_data().__dict__ for row in self.rows
                if row.get_data().generic_name or row.get_data().brand_name]

    def _build_workflow_bar(self, parent, active_step):
        """Create the compact Patient → Medicines → Export progress row."""
        bar = ctk.CTkFrame(parent, fg_color="transparent")
        bar.pack(fill="x", padx=4, pady=(0, 5))
        labels = ("workflow_patient", "workflow_medicines", "workflow_export")
        buttons = []
        for index, key in enumerate(labels, 1):
            if index > 1:
                ctk.CTkFrame(bar, height=1, fg_color=LINE).pack(
                    side="left", fill="x", expand=True, padx=4)
            button = VisualButton(
                bar, text=f"{index}  {I.t(key)}", height=32,
                width=max(105, _ui_font(11).measure(I.t(key)) + 45),
                corner_radius=16, font=_ui_font(11, "bold"),
                command=lambda step=index: self.go_workflow_step(step))
            button.pack(side="left")
            buttons.append(button)
        self._workflow_bars.append((buttons, active_step))
        self._style_workflow_bar(
            buttons, active_step, self._workflow_export_complete)
        return bar

    @staticmethod
    def _style_workflow_bar(buttons, active_step, export_complete=False):
        for index, button in enumerate(buttons, 1):
            if index < active_step or (index == 3 and export_complete):
                button.configure(
                    text=f"✓  {button.cget('text').split('  ', 1)[-1]}",
                    fg_color="#dff4f0", text_color="#137a71",
                    hover_color="#dff4f0", border_width=0)
            elif index == active_step:
                button.configure(fg_color=PRIMARY, text_color="white",
                                 hover_color=ACCENT_HOVER, border_width=0)
            else:
                button.configure(fg_color=ACCENT_SOFT, text_color=MUTED,
                                 hover_color=ACCENT_SOFT, border_width=0)

    def _set_workflow_step(self, step):
        self._workflow_active_step = step
        if step < 3:
            self._workflow_export_complete = False
        for buttons, _built_step in getattr(self, "_workflow_bars", []):
            # Restore stable labels before adding completed-state marks.
            keys = ("workflow_patient", "workflow_medicines", "workflow_export")
            for index, button in enumerate(buttons, 1):
                button.configure(text=f"{index}  {I.t(keys[index - 1])}")
            self._style_workflow_bar(
                buttons, step, self._workflow_export_complete)
        self._refresh_workflow_summary()

    def go_workflow_step(self, step):
        if step == 1:
            self.show_page("patient")
            self._set_workflow_step(1)
            return
        if step == 2:
            self.show_page("medications")
            self.close_medication_subpage()
            self._set_workflow_step(2)
            return
        self.open_export_subpage()

    def save_patient_and_continue(self):
        if self.save_patient_history():
            self._medication_patient_id = self._loaded_patient_id
            self.show_page("medications")
            self._set_workflow_step(2)

    def _workflow_issues(self):
        return validate_prescription_workflow(
            self._workflow_patient_data(), self._workflow_medicine_data())

    @staticmethod
    def _workflow_issue_message(issue):
        if issue.field == "patient.name":
            return I.t("workflow_patient_required")
        if issue.field == "medicines":
            return I.t("workflow_medicine_required")
        parts = issue.field.split(".")
        if len(parts) == 3 and parts[0] == "medicines":
            key = f"workflow_medicine_missing_{parts[2]}"
            return I.t(key, n=parts[1])
        return issue.message

    def _approve_workflow_issues(self):
        issues = self._workflow_issues()
        errors = [self._workflow_issue_message(issue)
                  for issue in issues if issue.severity == "error"]
        if errors:
            messagebox.showinfo(
                I.t("workflow_validation_title"),
                I.t("workflow_errors") + "\n\n" + "\n".join(f"• {item}" for item in errors),
                parent=self)
            return False
        return True

    def review_prescription(self):
        """Compatibility entry point for the optional document preview."""
        if not self._approve_workflow_issues():
            return
        self.show_page("medications")
        self.open_medication_subpage("preview")
        self.render_word_preview()

    def open_export_subpage(self):
        if not self._approve_workflow_issues():
            return
        self._workflow_export_complete = False
        self.show_page("medications")
        self.open_medication_subpage("export")
        self._set_workflow_step(3)
        self.render_word_preview()

    def _refresh_workflow_summary(self):
        if not hasattr(self, "workflow_patient_label"):
            return
        patient = self.patient_vars["name"].get().strip() or I.t("workflow_no_patient")
        age = self.patient_vars["age"].get().strip()
        sex_code = self.patient_vars["sex"].get().strip()
        sex = I.t("sex_m") if sex_code == "M" else I.t("sex_f") if sex_code == "F" else ""
        count = len(self._workflow_medicine_data())
        details = [patient]
        if age:
            details.append(f"{I.t('age')} {age}")
        if sex:
            details.append(sex)
        details.append(I.t("medicine_count", n=count))
        self.workflow_patient_label.configure(text="   ·   ".join(details))
        saved = (self._workflow_saved_signature
                 and self._workflow_saved_signature == self._draft_signature())
        self.workflow_save_state.configure(
            text=("● " + I.t("workflow_saved_local") if saved
                  else "● " + I.t("workflow_unsaved")),
            text_color=GOOD if saved else WARNING)

    def _draft_signature(self):
        patient = tuple(self.patient_vars[key].get().strip()
                        for key in ("name", "age", "sex"))
        medicines = tuple(tuple(str(drug.get(key, "") or "").strip()
                                for key in ("generic_name", "brand_name", "dosage",
                                            "frequency", "duration", "notes", "quantity"))
                          for drug in self._workflow_medicine_data())
        return patient, medicines

    def build_forms(self):
        for names in getattr(self, "_lazy_page_attributes", {}).values():
            for name in names:
                self.__dict__.pop(name, None)
        self._lazy_page_attributes = {}
        self._loaded_pages = set()
        for key in list(self._latest_searches):
            self.cancel_latest_search(key)
        self._bidi_bindings = []
        self.doctor_vars = {k: tk.StringVar() for k in ["name", "license_no", "specialty"]}
        self.patient_vars = {k: tk.StringVar() for k in ["name", "age", "sex"]}
        self._patient_status_suspend = False
        self._patient_saved_snapshot = ("", "", "")
        self._workflow_bars = []
        self._workflow_active_step = 1
        self._workflow_export_complete = False
        self._workflow_saved_signature = None
        self._draft_save_job = None
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
        self._build_workflow_bar(self.pages["patient"], 1)
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
        patient_name_binding = DirectionalTextBinding(
            self, self.patient_vars["name"], default_justify="right",
            dynamic_justify=True)
        self._bidi_bindings.append(patient_name_binding)
        self.patient_name_entry = VisualEntry(
            name_col,
            font_role="names",
            textvariable=patient_name_binding.display_var,
            width=PATIENT_NAME_WIDTH, height=FIELD_HEIGHT,
            corner_radius=9, border_color=LINE,
            font=_ui_font(cfg.config.ui_font_size("patient_name_font_size")), justify="right")
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
            patient_actions, text="", image=action_icon("save"),
            width=ICON_BUTTON_SIZE, height=ICON_BUTTON_SIZE,
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
        VisualButton(
            patient_actions, text=I.t("workflow_continue_medicines"),
            width=_ui_font(12, "bold").measure(I.t("workflow_continue_medicines")) + 28,
            height=32, fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            font=_ui_font(12, "bold"), command=self.save_patient_and_continue).pack(
                side="right", padx=(8, 0))
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
                                                  font_role="names",
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
        self._build_workflow_bar(self.pages["medications"], 2)
        workflow_summary = ctk.CTkFrame(
            self.pages["medications"], fg_color=SURFACE, border_color=LINE,
            border_width=1, corner_radius=9)
        workflow_summary.pack(fill="x", padx=4, pady=(0, 5))
        self.workflow_patient_label = ctk.CTkLabel(
            workflow_summary, text="", text_color=TEXT, anchor="w",
            font=_ui_font(12, "bold"))
        self.workflow_patient_label.pack(side="left", fill="x", expand=True, padx=12, pady=7)
        self.workflow_save_state = ctk.CTkLabel(
            workflow_summary, text="", text_color=WARNING,
            font=_ui_font(10, "bold"))
        self.workflow_save_state.pack(side="right", padx=8)
        VisualButton(
            workflow_summary, text=I.t("workflow_change_patient"),
            width=_ui_font(10).measure(I.t("workflow_change_patient")) + 20,
            height=28, fg_color="transparent", text_color=ACCENT,
            hover_color=ACCENT_SOFT, font=_ui_font(10),
            command=lambda: self.go_workflow_step(1)).pack(side="right", padx=(4, 8), pady=3)
        self.word_preview_visible = False
        self.medication_main = ctk.CTkFrame(self.pages["medications"], fg_color="transparent")
        self.medication_main.pack(fill="both", expand=True)
        self.medication_subpage = ctk.CTkFrame(self.pages["medications"], fg_color="transparent")
        self.medication_subpage_bar = ctk.CTkFrame(
            self.medication_subpage, fg_color="transparent")
        self.medication_subpage_bar.pack(fill="x", pady=(4, 8))
        VisualButton(
            self.medication_subpage_bar, text=I.t("back"), width=70, height=32,
            fg_color=CARD, text_color=ACCENT, border_color=LINE, border_width=1,
            hover_color=ACCENT_SOFT, command=self.close_medication_subpage).pack(side="left")
        self.medication_subpage_title = ctk.CTkLabel(
            self.medication_subpage_bar, text="", font=_ui_font(18, "bold"), text_color=TEXT)
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
        self.medication_save_button = VisualButton(
            medication_toolbar, text=I.t("save_patient_prescription"),
            width=_ui_font(12).measure(I.t("save_patient_prescription")) + 22,
            height=32, font=_ui_font(12), corner_radius=8,
            fg_color=SURFACE, text_color=ACCENT, border_color=LINE, border_width=1,
            hover_color=ACCENT_SOFT,
            command=self.save_prescription_for_patient)
        self.medication_save_button.pack(side="left", padx=(5, 0))
        self.medication_save_status = ctk.CTkLabel(medication_toolbar, text="", width=72,
            font=_ui_font(11), text_color=MUTED)
        self.medication_save_status.pack(side="left", padx=4)
        self.workflow_export_button = VisualButton(
            medication_toolbar, text=I.t("workflow_export_title") + "  →",
            width=_ui_font(12, "bold").measure(I.t("workflow_export_title")) + 42,
            height=32, font=_ui_font(12, "bold"), corner_radius=8,
            fg_color=PRIMARY, text_color="white", hover_color=ACCENT_HOVER,
            command=self.open_export_subpage)
        self.workflow_export_button.pack(side="right")
        # Keep the historical widget attribute available for integrations that
        # customized it before Review ceased to be a dedicated workflow step.
        self.workflow_review_button = self.workflow_export_button

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
        self.word_preview_body = ctk.CTkFrame(preview, fg_color="transparent")

        self.workflow_export_panel = ctk.CTkFrame(
            self.medication_subpage, fg_color=SURFACE, border_color=LINE,
            border_width=1, corner_radius=10)
        ctk.CTkLabel(
            self.workflow_export_panel, text=I.t("workflow_export_title"),
            text_color=TEXT, anchor="w", font=_ui_font(16, "bold")).pack(
                fill="x", padx=12, pady=(10, 2))
        export_actions = ctk.CTkFrame(self.workflow_export_panel, fg_color="transparent")
        export_actions.pack(fill="x", padx=10, pady=(2, 10))
        for key, command in (("export_word", self.export_word),
                             ("export_compact", self.export_label)):
            VisualButton(
                export_actions, text=I.t(key), height=34,
                width=_ui_font(11).measure(I.t(key)) + 24,
                fg_color=PRIMARY if key == "export_word" else CARD,
                text_color="white" if key == "export_word" else ACCENT,
                border_width=0 if key == "export_word" else 1,
                border_color=LINE, hover_color=ACCENT_HOVER if key == "export_word" else ACCENT_SOFT,
                font=_ui_font(11, "bold"), command=command).pack(side="left", padx=3)

        self._built_pages = {"prescriber", "patient", "medications"}
        self._page_builders = {
            "interaction_review": self._build_interaction_review_page,
            "reference": self._build_reference_page,
            "favorites": self._build_favorites_page,
            "drug_classes": self._build_drug_classes_page,
            "treatment_templates": self._build_treatment_templates_page,
        }

    def _ensure_page_built(self, key):
        if key in self._built_pages:
            return
        started = time.perf_counter()
        previous_attributes = set(self.__dict__)
        self._built_pages.add(key)
        try:
            self._page_builders[key]()
            self._lazy_page_attributes[key] = set(self.__dict__) - previous_attributes
            self.apply_ui_font_preferences()
        except Exception:
            self._built_pages.discard(key)
            for child in self.pages[key].winfo_children():
                child.destroy()
            for name in set(self.__dict__) - previous_attributes:
                self.__dict__.pop(name, None)
            raise
        self._record_response_time("build:" + key, started)

    def _build_interaction_review_page(self):
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


    def _build_reference_page(self):
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


    def _build_favorites_page(self):
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
        self.favorite_saved_status = ctk.CTkLabel(favorite_toolbar, text="", width=72, font=_ui_font(11))
        self.favorite_saved_status.pack(side="left", padx=4)

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
                    self.favorite_editor,
                    font_role="instructions",
                    values=list(values),
                    variable=self.favorite_vars[key], height=FIELD_HEIGHT,
                    border_color=LINE, fg_color=CARD,
                    button_color=PRIMARY, button_hover_color=ACCENT_HOVER,
                    dropdown_fg_color=CARD, dropdown_hover_color=ACCENT_SOFT,
                    dropdown_text_color=TEXT, dropdown_font=_ui_font(16),
                    font=ctk.CTkFont(size=12))
            else:
                widget = VisualEntry(
                    self.favorite_editor, textvariable=self.favorite_vars[key],
                    font_role="instructions" if key in {"dosage", "duration"} else None,
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
        self.favorite_save_status = ctk.CTkLabel(preview_row, text="", width=72, font=_ui_font(11))
        self.favorite_save_status.grid(row=0, column=2, padx=4)
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


    def _build_drug_classes_page(self):
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
            count_badge = ctk.CTkLabel(
                tile, text="0", width=28, height=28, corner_radius=14,
                fg_color=ACCENT_SOFT, text_color=ACCENT,
                font=ctk.CTkFont(size=11, weight="bold"))
            count_badge.grid(row=0, column=1, padx=(5, 0))
            VisualButton(
                tile, text="›", width=28, height=28, corner_radius=14,
                fg_color="transparent", text_color=ACCENT,
                hover_color=ACCENT_SOFT, font=_ui_font(18, "bold"),
                command=lambda selected=code: self.open_therapeutic_group(selected)).grid(
                    row=0, column=2, padx=(2, 0))
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
            toolbar, text="", image=action_icon("save"),
            width=ICON_BUTTON_SIZE, height=ICON_BUTTON_SIZE,
            fg_color="transparent", text_color=ACCENT, border_width=0,
            hover_color=ACCENT_SOFT,
            font=ctk.CTkFont(size=17), command=self.save_treatment_template)
        save_button.pack(side="left", padx=3)
        self.template_save_status = ctk.CTkLabel(toolbar, text="", width=72, font=_ui_font(11))
        self.template_save_status.pack(side="left", padx=4)
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

        category_row = ctk.CTkFrame(editor, fg_color="transparent")
        category_row.pack(fill="x", padx=PAD, pady=(0, 5))
        ctk.CTkLabel(
            category_row, text=I.t("treatment_category"), width=125,
            text_color=MUTED, anchor="w",
            font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")
        self.treatment_category_var = tk.StringVar(
            value=I.t("treatment_category_uncategorized"))
        self.treatment_category_menu = VisualComboBox(
            category_row, values=self._treatment_category_values(),
            variable=self.treatment_category_var, width=260,
            height=FIELD_HEIGHT, fg_color=CARD, text_color=TEXT,
            border_color=LINE, dropdown_fg_color=CARD,
            dropdown_hover_color=ACCENT_SOFT,
            button_color=PRIMARY, button_hover_color=ACCENT_HOVER,
            font=_ui_font(12), dropdown_font=_ui_font(12), justify="left")
        self.treatment_category_menu.pack(side="left")

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
        saved_controls = ctk.CTkFrame(
            self.treatment_saved_view, fg_color="transparent")
        saved_controls.pack(fill="x", padx=PAD, pady=(10, 6))
        self.template_saved_status = ctk.CTkLabel(saved_controls, text="", width=72, font=_ui_font(11))
        self.template_saved_status.pack(side="right", padx=4)
        self.treatment_saved_search_var = tk.StringVar()
        self._treatment_saved_search_job = None
        self._saved_template_render_generation = 0
        self._treatment_saved_refresh_suspended = False
        self.treatment_saved_search_var.trace_add(
            "write", lambda *_: self._schedule_saved_treatment_refresh())
        VisualEntry(
            saved_controls, textvariable=self.treatment_saved_search_var,
            placeholder_text=I.t("treatment_saved_search"), height=FIELD_HEIGHT,
            border_color=LINE).pack(side="left", fill="x", expand=True, padx=(0, 6))
        self._treatment_saved_sort_labels = {
            I.t("treatment_sort_used"): "most_used",
            I.t("treatment_sort_recent"): "recently_used",
            I.t("treatment_sort_modified"): "recently_modified",
            I.t("treatment_sort_name"): "name",
            I.t("treatment_sort_name_reverse"): "name_reverse",
        }
        self._treatment_saved_sort_mode = "most_used"
        self.treatment_saved_sort_var = tk.StringVar(
            value=I.t("treatment_sort_used"))
        self.treatment_saved_sort_menu = VisualOptionMenu(
            saved_controls, values=list(self._treatment_saved_sort_labels),
            variable=self.treatment_saved_sort_var, width=165,
            height=FIELD_HEIGHT, fg_color=CARD, text_color=ACCENT,
            button_color=PRIMARY, button_hover_color=ACCENT_HOVER,
            dropdown_fg_color=CARD, dropdown_hover_color=ACCENT_SOFT,
            command=self._saved_template_sort_changed)
        self.treatment_saved_sort_menu.pack(side="right")
        self.treatment_saved_category_var = tk.StringVar(
            value=I.t("treatment_all_categories"))
        self.treatment_saved_category_menu = VisualOptionMenu(
            saved_controls, values=[I.t("treatment_all_categories")],
            variable=self.treatment_saved_category_var, width=155,
            height=FIELD_HEIGHT, fg_color=CARD, text_color=ACCENT,
            button_color=PRIMARY, button_hover_color=ACCENT_HOVER,
            dropdown_fg_color=CARD, dropdown_hover_color=ACCENT_SOFT,
            command=self._saved_template_category_changed)
        self.treatment_saved_category_menu.pack(side="right", padx=(0, 6))
        self.treatment_saved_cards = ctk.CTkFrame(self.treatment_saved_view, fg_color="transparent")
        self.treatment_saved_cards.pack(fill="x", padx=PAD, pady=(0, PAD))
        self.treatment_saved_cards.bind(
            "<Configure>", self._schedule_saved_template_reflow)
        self._treatment_saved_card_widgets = []
        self._treatment_saved_footer = None
        self._treatment_saved_columns = 0
        self._treatment_saved_reflow_job = None
        self._treatment_saved_expanded_id = ""
        self._treatment_saved_highlight_id = ""
        self._treatment_saved_scroll_position = 0.0
        self._treatment_saved_return_state = None
        self._treatment_saved_limit = 24

    def _treatment_default_categories(self):
        return [I.t(key) for key in (
            "treatment_category_uncategorized",
            "treatment_category_respiratory",
            "treatment_category_cardiovascular",
            "treatment_category_gastrointestinal",
            "treatment_category_neurology",
            "treatment_category_endocrine",
            "treatment_category_antiinfective",
            "treatment_category_emergency",
            "treatment_category_other",
        )]

    def _treatment_category_values(self):
        values = self._treatment_default_categories()
        values.extend(item.get("category", "").strip()
                      for item in cfg.config.treatment_templates())
        return list(dict.fromkeys(value for value in values if value))

    def _current_treatment_category(self):
        variable = getattr(self, "treatment_category_var", None)
        value = variable.get().strip() if variable is not None else ""
        return "" if value == I.t("treatment_category_uncategorized") else value

    @staticmethod
    def _treatment_draft_signature(disease, medicines, category=""):
        return (strip_bidi_display_controls(disease).strip(),
                strip_bidi_display_controls(category).strip(), tuple(
            tuple(sorted((key, str(value)) for key, value in medicine.items()
                         if not key.startswith("_"))) for medicine in medicines))

    def _capture_treatment_baseline(self):
        self._treatment_baseline_disease = self.treatment_disease_var.get()
        self._treatment_baseline_selector = self.treatment_template_selector_var.get()
        self._treatment_baseline_category = self._current_treatment_category()
        self._treatment_baseline_drugs = [dict(item) for item in self._treatment_template_drugs]
        self._treatment_baseline_signature = self._treatment_draft_signature(
            self._treatment_baseline_disease, self._treatment_baseline_drugs,
            self._treatment_baseline_category)

    def _confirm_treatment_leave(self):
        if getattr(self, "_treatment_view", "saved") != "editor":
            return True
        current = self._treatment_draft_signature(
            self._current_treatment_disease(), self._treatment_template_drugs,
            self._current_treatment_category())
        if current == self._treatment_baseline_signature:
            return True
        decision = messagebox.askyesnocancel(
            I.t("treatment_unsaved_title"), I.t("treatment_unsaved_message"),
            parent=self)
        if decision is None:
            return False
        if decision:
            return bool(self.save_treatment_template())
        self.treatment_disease_var.set(self._treatment_baseline_disease)
        self.treatment_template_selector_var.set(self._treatment_baseline_selector)
        self.treatment_category_var.set(
            self._treatment_baseline_category
            or I.t("treatment_category_uncategorized"))
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
        if view == "editor":
            self.scroll._parent_canvas.yview_moveto(0)
        else:
            self.after_idle(self._restore_saved_template_scroll)

    def _saved_template_scroll(self):
        canvas = getattr(getattr(self, "scroll", None), "_parent_canvas", None)
        try:
            return canvas.yview()[0] if canvas is not None else 0.0
        except (tk.TclError, IndexError):
            return 0.0

    def _capture_saved_template_state(self, anchor_id=""):
        """Keep the browser stable while a template is edited or applied."""
        return {
            "query": self.treatment_saved_search_var.get(),
            "sort": self._treatment_saved_sort_mode,
            "category": self.treatment_saved_category_var.get(),
            "expanded": self._treatment_saved_expanded_id,
            "highlight": anchor_id or self._treatment_saved_highlight_id,
            "limit": self._treatment_saved_limit,
            "scroll": self._saved_template_scroll(),
        }

    def _restore_saved_template_state(self, state):
        if not state:
            return
        self._treatment_saved_refresh_suspended = True
        self.treatment_saved_search_var.set(state.get("query", ""))
        mode = state.get("sort", "most_used")
        label = next((text for text, value in self._treatment_saved_sort_labels.items()
                      if value == mode), I.t("treatment_sort_used"))
        self._treatment_saved_sort_mode = mode
        self.treatment_saved_sort_var.set(label)
        self.treatment_saved_category_var.set(
            state.get("category", I.t("treatment_all_categories")))
        self._treatment_saved_expanded_id = state.get("expanded", "")
        self._treatment_saved_highlight_id = state.get("highlight", "")
        self._treatment_saved_limit = max(24, int(state.get("limit", 24)))
        self._treatment_saved_scroll_position = float(state.get("scroll", 0.0))
        self._treatment_saved_refresh_suspended = False

    def _schedule_saved_treatment_refresh(self, reset_scroll=True, delay=180):
        if self._treatment_saved_refresh_suspended:
            return
        if self._treatment_saved_search_job is not None:
            try:
                self.after_cancel(self._treatment_saved_search_job)
            except (tk.TclError, ValueError):
                pass
        if reset_scroll:
            self._treatment_saved_scroll_position = 0.0
            self._treatment_saved_limit = 24
        self._treatment_saved_search_job = self.after(
            delay, self.refresh_saved_treatment_templates)

    def _saved_template_sort_changed(self, selected):
        self._treatment_saved_sort_mode = self._treatment_saved_sort_labels.get(
            selected, "most_used")
        self._schedule_saved_treatment_refresh(delay=0)

    def _saved_template_category_changed(self, _selected=None):
        self._schedule_saved_treatment_refresh(delay=0)

    @staticmethod
    def _template_time_rank(value):
        try:
            return datetime.datetime.fromisoformat(str(value).replace(
                "Z", "+00:00")).timestamp()
        except (TypeError, ValueError):
            return 0.0

    def _sort_saved_treatment_templates(self, templates):
        name = lambda item: strip_bidi_display_controls(
            item.get("disease", "")).casefold()
        mode = self._treatment_saved_sort_mode
        if mode == "recently_used":
            return sorted(templates, key=lambda item: (
                -self._template_time_rank(item.get("last_used", "")), name(item)))
        if mode == "recently_modified":
            return sorted(templates, key=lambda item: (
                -self._template_time_rank(item.get("updated_at", "")), name(item)))
        if mode == "name_reverse":
            return sorted(templates, key=name, reverse=True)
        if mode == "name":
            return sorted(templates, key=name)
        return sorted(templates, key=lambda item: (
            -int(item.get("use_count", 0) or 0),
            -self._template_time_rank(item.get("last_used", "")), name(item)))

    def show_treatment_saved_templates(self, selected_id="", restore_state=None):
        if not self._confirm_treatment_leave():
            return False
        if restore_state:
            self._restore_saved_template_state(restore_state)
        self._switch_treatment_view("saved")
        if selected_id:
            self._treatment_saved_highlight_id = selected_id
            if not restore_state:
                self._treatment_saved_limit = 24
                self._treatment_saved_refresh_suspended = True
                self.treatment_saved_search_var.set("")
                self._treatment_saved_refresh_suspended = False
        self.refresh_saved_treatment_templates()
        if selected_id:
            card = getattr(self, "_treatment_highlight_card", None)
            if card is not None:
                self.after_idle(lambda target=card: self._scroll_medication_row_into_view(target))
        return True

    def refresh_saved_treatment_templates(self):
        if not hasattr(self, "treatment_saved_cards"):
            return
        self._treatment_saved_search_job = None
        self._saved_template_render_generation = getattr(self, "_saved_template_render_generation", 0) + 1
        generation = self._saved_template_render_generation
        render_container = self.treatment_saved_cards
        self._treatment_highlight_card = None
        pool = getattr(self, "_saved_template_card_pool", {})
        self._saved_template_card_pool = pool
        for child in self.treatment_saved_cards.winfo_children():
            if not hasattr(child, "_template_record"):
                child.destroy()
        self._treatment_saved_card_widgets = []
        self._treatment_saved_footer = None
        query = self.treatment_saved_search_var.get().strip().casefold()
        templates = cfg.config.treatment_templates()
        category_values = [I.t("treatment_all_categories")]
        category_values.extend(self._treatment_category_values())
        category_values = list(dict.fromkeys(category_values))
        selected_category = self.treatment_saved_category_var.get()
        if selected_category not in category_values:
            selected_category = I.t("treatment_all_categories")
            self.treatment_saved_category_var.set(selected_category)
        self.treatment_saved_category_menu.configure(values=category_values)
        uncategorized = I.t("treatment_category_uncategorized")

        def category_label(item):
            return item.get("category", "").strip() or uncategorized

        matches = [item for item in templates
                   if (selected_category == I.t("treatment_all_categories")
                       or category_label(item) == selected_category)
                   and query in (strip_bidi_display_controls(
                       item.get("disease", "") + " " + category_label(item)).casefold())]
        matches = self._sort_saved_treatment_templates(matches)
        matching_ids = {item["id"] for item in matches[:max(24, self._treatment_saved_limit)]}
        for identifier, card in pool.items():
            if identifier not in matching_ids:
                card.grid_remove()
        if not matches:
            self.treatment_saved_cards.grid_columnconfigure(0, weight=1)
            self.treatment_saved_cards.grid_columnconfigure(1, weight=1)
            empty = GlassFrame(
                self.treatment_saved_cards, fg_color=CARD, border_color=LINE,
                border_width=1, corner_radius=10)
            empty.grid(row=0, column=0, columnspan=2, sticky="ew", pady=3)
            message = (I.t("treatment_saved_empty_search") if templates
                       else I.t("treatment_saved_empty_none"))
            ctk.CTkLabel(empty, text=message, text_color=MUTED,
                         anchor="w").pack(side="left", fill="x", expand=True,
                                           padx=12, pady=10)
            if not templates:
                VisualButton(
                    empty, text="＋ " + I.t("treatment_new_view"), width=135,
                    height=32, fg_color=CARD, text_color=ACCENT,
                    border_width=1, border_color=LINE,
                    hover_color=ACCENT_SOFT,
                    command=self.new_treatment_template).pack(
                        side="right", padx=8, pady=6)
            return
        limit = max(24, self._treatment_saved_limit)
        highlighted = next((index for index, item in enumerate(matches)
                            if item["id"] == self._treatment_saved_highlight_id), -1)
        if highlighted >= limit:
            limit = highlighted + 1
        visible_ids = {item["id"] for item in matches[:limit]}
        valid_ids = {item["id"] for item in templates}
        for identifier, card in list(pool.items()):
            if identifier not in valid_ids or (len(pool) > max(96, len(visible_ids)) and identifier not in visible_ids):
                pool.pop(identifier).destroy()
        if len(matches) > limit:
            self._treatment_saved_footer = VisualButton(
                self.treatment_saved_cards, text=I.t("load_more"), width=110,
                height=32, command=self._load_more_saved_treatments)
        self._reflow_saved_template_cards()
        def render_batch(start=0):
            if (self._closing or not render_container.winfo_exists()
                    or generation != self._saved_template_render_generation):
                return
            if getattr(self, "active_page", "") != "treatment_templates":
                self._pending_template_render = lambda: render_batch(start)
                return
            if start and not self._card_near_viewport(self._treatment_saved_card_widgets[-1]):
                self.after(150, lambda: render_batch(start))
                return
            deadline = time.perf_counter() + .008
            while start < min(limit, len(matches)):
                template = matches[start]
                signature = json.dumps({key: value for key, value in template.items()
                    if key not in {"created_at", "updated_at", "last_used", "use_count"}},
                    sort_keys=True, ensure_ascii=False)
                card = pool.get(template["id"])
                if card is not None and (getattr(card, "_card_placeholder", False)
                                         or card._template_signature != signature):
                    card.destroy()
                    card = None
                if card is None:
                    card = self._build_saved_treatment_card(template)
                    card._template_signature = signature
                    pool[template["id"]] = card
                else:
                    card._template_record = template
                    selected = template["id"] == self._treatment_saved_highlight_id
                    if card.cget("fg_color") != (ACCENT_SOFT if selected else CARD):
                        card.configure(fg_color=ACCENT_SOFT if selected else CARD,
                                       border_color=ACCENT if selected else LINE)
                    self._set_saved_template_card_expanded(
                        card, template["id"] == self._treatment_saved_expanded_id)
                    if selected:
                        self._treatment_highlight_card = card
                self._treatment_saved_card_widgets.append(card)
                start += 1
                if time.perf_counter() >= deadline or not self._card_near_viewport(card):
                    break
            self._reflow_saved_template_cards()
            if start < min(limit, len(matches)):
                self.after(1, lambda: render_batch(start))
            else:
                self.after_idle(self._restore_saved_template_scroll)
                if self._treatment_saved_footer is not None:
                    self._queue_card_loading("treatment_templates", generation,
                        self._treatment_saved_footer, self._load_more_saved_treatments)
        render_batch()

    def _build_saved_treatment_card(self, template):
        template_id = template["id"]
        selected = template_id == self._treatment_saved_highlight_id
        expanded = self._treatment_saved_expanded_id == template_id
        card = GlassFrame(
            self.treatment_saved_cards,
            fg_color=ACCENT_SOFT if selected else CARD,
            border_color=ACCENT if selected else LINE,
            border_width=1, corner_radius=10)
        if selected:
            self._treatment_highlight_card = card
        toggle = lambda _event=None, picked=template_id: (
            self.toggle_saved_treatment_template(picked))
        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=8, pady=(5, 3))
        actions = ctk.CTkFrame(header, fg_color="transparent")
        actions.pack(side="right")
        for symbol, image, command in (
                ("+", None,
                 lambda target=card: self.use_saved_treatment_template(target._template_record)),
                ("", _edit_icon(18),
                 lambda target=card: self.edit_saved_treatment_template(target._template_record)),
                ("🗑", None,
                 lambda target=card: self.delete_saved_treatment_template(target._template_record))):
            VisualButton(
                actions, text=symbol, image=image, width=30, height=30,
                fg_color="transparent", border_width=0,
                text_color=DANGER if symbol == "🗑" else ACCENT,
                hover_color=DANGER_SOFT if symbol == "🗑" else ACCENT_SOFT,
                font=_ui_font(18), command=command).pack(side="left", padx=1)
        card._template_chevron = VisualButton(
            header, text="⌄" if expanded else "›", width=28, height=28,
            fg_color="transparent", text_color=ACCENT,
            hover_color=ACCENT_SOFT, font=_ui_font(18, "bold"),
            command=toggle)
        card._template_chevron.pack(side="right", padx=(2, 3))
        count = ctk.CTkLabel(
            header, text=str(len(template["medications"])), width=27,
            height=24, corner_radius=12, fg_color=SURFACE,
            text_color=ACCENT, font=_ui_font(10, "bold"))
        count.pack(side="right", padx=(3, 2))
        category = ctk.CTkLabel(
            header, text=directional_display_text(
                template.get("category", "").strip()
                or I.t("treatment_category_uncategorized")),
            height=24, corner_radius=12, fg_color=ACCENT_SOFT,
            text_color=ACCENT, font=_ui_font(10, "bold"))
        category.pack(side="right", padx=(3, 2), ipadx=6)
        title = VisualButton(
            header, text=directional_display_text(template["disease"]),
            height=30, anchor="w", fg_color="transparent", text_color=TEXT,
            hover_color=ACCENT_SOFT, font=_ui_font(14, "bold"),
            command=toggle)
        title.pack(side="left", fill="x", expand=True)
        card._template_record = template
        card._template_body = None
        self._set_saved_template_card_expanded(card, expanded)
        return card

    def _set_saved_template_card_expanded(self, card, expanded):
        if getattr(card, "_card_placeholder", False):
            return
        if getattr(card, "_template_is_expanded", None) == expanded:
            return
        card._template_is_expanded = expanded
        card._template_chevron.configure(text="⌄" if expanded else "›")
        if not expanded:
            if card._template_body is not None:
                card._template_body.pack_forget()
            return
        if card._template_body is not None:
            card._template_body.pack(fill="x")
            return
        body = ctk.CTkFrame(card, fg_color="transparent")
        card._template_body = body
        body.pack(fill="x")
        medicines = card._template_record["medications"]
        for index, medicine in enumerate(medicines, 1):
            primary, secondary = self._treatment_medicine_title(medicine)
            regimen = self._treatment_regimen_summary(medicine)
            line = ctk.CTkFrame(body, fg_color="transparent")
            line.pack(fill="x", padx=PAD, pady=(0, CARD_GAP))
            prefix = (I.t("treatment_or") + "  " if medicine.get("alternative_to_previous") else "")
            parts = [(prefix + f"{index}. {primary}", TEXT,
                      "bold" if medicine.get("brand_name", "").strip() else "normal", CARD_NAME_SIZE),
                     (secondary, TEXT_SECONDARY, "normal", CARD_NAME_SIZE),
                     (regimen, TEXT_SECONDARY, "normal", CARD_DETAIL_SIZE)]
            labels = []
            for column, (text, color, weight, size) in enumerate(parts):
                if not text:
                    continue
                line.grid_columnconfigure(column, weight=1, uniform="medicine")
                label = ctk.CTkLabel(line, text=directional_display_text(text), anchor="w",
                    justify="left", wraplength=140, text_color=color, font=_ui_font(size, weight))
                label.grid(row=0, column=column, sticky="nw", padx=(0, CARD_GAP))
                labels.append(label)
            line.bind("<Configure>", lambda event, items=labels:
                [item.configure(wraplength=max(60, event.width // len(items)-CARD_GAP)) for item in items], add="+")

    def _schedule_saved_template_reflow(self, event=None):
        width = event.width if event is not None else max(
            1, self.treatment_saved_cards.winfo_width())
        desired = 2 if width >= 900 else 1
        cards_are_placed = all(
            card.winfo_manager() == "grid"
            for card in self._treatment_saved_card_widgets)
        footer_is_placed = (self._treatment_saved_footer is None
                            or self._treatment_saved_footer.winfo_manager() == "grid")
        if (desired == self._treatment_saved_columns
                and cards_are_placed and footer_is_placed):
            return
        if self._treatment_saved_reflow_job is not None:
            try:
                self.after_cancel(self._treatment_saved_reflow_job)
            except (tk.TclError, ValueError):
                pass
        self._treatment_saved_reflow_job = self.after(
            80, self._reflow_saved_template_cards)

    def _reflow_saved_template_cards(self):
        self._treatment_saved_reflow_job = None
        width = max(1, self.treatment_saved_cards.winfo_width())
        columns = 2 if width >= 900 else 1
        cards_are_placed = all(
            card.winfo_manager() == "grid"
            for card in self._treatment_saved_card_widgets)
        footer_is_placed = (self._treatment_saved_footer is None
                            or self._treatment_saved_footer.winfo_manager() == "grid")
        if (columns == self._treatment_saved_columns
                and cards_are_placed and footer_is_placed):
            return
        self._treatment_saved_columns = columns
        self.treatment_saved_cards.grid_columnconfigure(0, weight=1)
        self.treatment_saved_cards.grid_columnconfigure(
            1, weight=1 if columns == 2 else 0)
        for index, card in enumerate(self._treatment_saved_card_widgets):
            column = index % columns
            position = (index // columns, column, columns)
            if getattr(card, "_template_grid_position", None) != position or not card.winfo_manager():
                card.grid(
                    row=index // columns, column=column, sticky="new",
                    padx=(0, 3) if column == 0 and columns == 2 else (
                        (3, 0) if columns == 2 else 0),
                    pady=3)
                card._template_grid_position = position
        if self._treatment_saved_footer is not None:
            self._treatment_saved_footer.grid(
                row=(len(self._treatment_saved_card_widgets) + columns - 1) // columns,
                column=0, columnspan=columns, sticky="w", pady=5)

    def _restore_saved_template_scroll(self):
        canvas = getattr(getattr(self, "scroll", None), "_parent_canvas", None)
        if canvas is not None:
            try:
                canvas.yview_moveto(self._treatment_saved_scroll_position)
            except tk.TclError:
                pass

    def _load_more_saved_treatments(self):
        self._treatment_saved_scroll_position = self._saved_template_scroll()
        self._treatment_saved_limit += 24
        self.refresh_saved_treatment_templates()

    def toggle_saved_treatment_template(self, template_id):
        self._treatment_saved_scroll_position = self._saved_template_scroll()
        previous_id = self._treatment_saved_expanded_id
        self._treatment_saved_expanded_id = "" if self._treatment_saved_expanded_id == template_id else template_id
        for card in self._treatment_saved_card_widgets:
            card_id = card._template_record["id"]
            if card_id in {previous_id, template_id}:
                self._set_saved_template_card_expanded(
                    card, card_id == self._treatment_saved_expanded_id)
        self.after_idle(self._restore_saved_template_scroll)

    def edit_saved_treatment_template(self, template):
        self._treatment_saved_return_state = self._capture_saved_template_state(
            template["id"])
        self._load_treatment_template_record(template)

    def use_saved_treatment_template(self, template):
        self._treatment_saved_return_state = self._capture_saved_template_state(
            template["id"])
        self._show_treatment_apply_preview(
            medicines=template["medications"], template_id=template["id"])

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
        category_values = self._treatment_category_values()
        if hasattr(self, "treatment_category_menu"):
            self.treatment_category_menu.configure(values=category_values)
        if hasattr(self, "treatment_saved_category_menu"):
            self.treatment_saved_category_menu.configure(values=list(dict.fromkeys(
                [I.t("treatment_all_categories")] + category_values)))
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
        self.treatment_category_var.set(I.t("treatment_category_uncategorized"))
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
        self.treatment_category_var.set(
            template.get("category", "").strip()
            or I.t("treatment_category_uncategorized"))
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
        self.submit_latest_search("treatment_drug",
            lambda: self.db.search_prescribable(query, 8),
            lambda matches: self._render_treatment_drug_results(
                query, token, matches))

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
        self._set_save_feedback("template", "")
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
        for key in ("dosage", "frequency", "duration", "notes"):
            value = str(medicine.get(key, "")).strip()
            if value:
                parts.append(value)
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
            wrap_card_labels(header)

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
                        font_role="instructions",
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
                        font_role="instructions",
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
            return False
        if not self._treatment_template_drugs:
            self._set_treatment_status(I.t("treatment_medicine_required"), DANGER)
            self.treatment_drug_search_entry.focus_set()
            return False
        editing_existing = bool(self._treatment_template_id)
        template_id = self._treatment_template_id or uuid.uuid4().hex
        saved_id = self._save_with_feedback("template", lambda: cfg.config.save_treatment_template({
            "id": template_id, "disease": disease,
            "category": self._current_treatment_category(),
            "variant": "",
            "medications": self._treatment_template_drugs,
        }))
        if not saved_id:
            self._set_treatment_status(I.t("treatment_save_failed"), DANGER)
            return False
        self._treatment_template_id = saved_id
        self.refresh_treatment_template_menu(saved_id)
        self._capture_treatment_baseline()
        self._set_treatment_status(I.t("treatment_saved"), ACCENT)
        restore_state = (self._treatment_saved_return_state
                         if editing_existing else None)
        self._treatment_saved_return_state = None
        self.show_treatment_saved_templates(
            selected_id=saved_id, restore_state=restore_state)
        return True

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
                    template.get("disease", ""), template.get("category", ""), step,
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
            "treatment_category": "category",
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
            "category": "category", "specialty": "category",
            "group": "category", "templategroup": "category",
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
        current_category = ""
        for row in rows:
            entered_disease = cell(row, "disease")
            if entered_disease:
                current_disease = entered_disease
                current_category = cell(row, "category")
            if not current_disease:
                continue
            brand_name = cell(row, "brand_name")
            generic_name = cell(row, "generic_name")
            if not brand_name and not generic_name:
                continue
            disease_key = current_disease.casefold()
            if disease_key not in templates:
                templates[disease_key] = {
                    "disease": current_disease,
                    "category": current_category,
                    "medications": []}
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
            I.t("treatment_disease"), I.t("treatment_category"),
            I.t("treatment_step"),
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
        self._show_treatment_apply_preview(
            template_id=self._treatment_template_id)

    def _show_treatment_apply_preview(self, medicines=None, template_id=""):
        """Preview treatment steps, optionally excluding any before application."""
        groups = self._treatment_choice_groups(
            self._treatment_template_drugs if medicines is None else medicines)
        dialog = ctk.CTkToplevel(self)
        dialog.title(I.t("treatment_apply_preview"))
        dialog.geometry("700x520")
        dialog.minsize(580, 420)
        dialog.transient(self)
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        ctk.CTkLabel(
            dialog, text=I.t("treatment_apply_preview"), text_color=ACCENT,
            font=ctk.CTkFont(size=18, weight="bold"), anchor="w").pack(
                fill="x", padx=14, pady=(12, 2))
        ctk.CTkLabel(
            dialog, text=I.t("treatment_apply_preview_help"), text_color=MUTED,
            font=ctk.CTkFont(size=12), anchor="w", justify="left",
            wraplength=650).pack(fill="x", padx=14, pady=(0, 7))
        body = ctk.CTkScrollableFrame(dialog, fg_color=BG, corner_radius=10)
        body.pack(fill="both", expand=True, padx=14, pady=(0, 7))
        selections = []
        for step, group in enumerate(groups, 1):
            card = GlassFrame(
                body, fg_color=CARD, border_color=LINE,
                border_width=1, corner_radius=9)
            card.pack(fill="x", pady=3)
            heading = I.t("treatment_step_number", n=step)
            if len(group) > 1:
                heading += "  ·  " + I.t("treatment_choose_one")
            selected = tk.IntVar(value=0)
            included = tk.BooleanVar(value=True)
            selections.append((group, selected, included))
            if len(group) > 1:
                ctk.CTkCheckBox(
                    card, text=heading + "  ·  " + I.t("treatment_include_step"),
                    variable=included, onvalue=True, offvalue=False,
                    fg_color=PRIMARY, hover_color=ACCENT_HOVER,
                    text_color=ACCENT, font=ctk.CTkFont(size=12, weight="bold"),
                    checkbox_width=18, checkbox_height=18).pack(
                        fill="x", padx=10, pady=(6, 2), anchor="w")
            for option, medicine in enumerate(group):
                row = ctk.CTkFrame(card, fg_color="transparent")
                row.pack(fill="x", padx=10, pady=(1, 3))
                primary, secondary = self._treatment_medicine_title(medicine)
                title = primary + (f"  ·  {secondary}" if secondary else "")
                summary = self._treatment_regimen_summary(medicine)
                text = title if not summary else f"{title}   ·   {summary}"
                if len(group) > 1:
                    ctk.CTkRadioButton(
                        row, text=text, variable=selected, value=option,
                        fg_color=PRIMARY, hover_color=ACCENT_HOVER,
                        text_color=TEXT, font=ctk.CTkFont(size=12),
                        command=lambda: None).pack(
                            fill="x", padx=2, pady=2, anchor="w")
                else:
                    ctk.CTkCheckBox(
                        row, text=text, variable=included,
                        onvalue=True, offvalue=False, fg_color=PRIMARY,
                        hover_color=ACCENT_HOVER, text_color=TEXT,
                        font=ctk.CTkFont(size=12), checkbox_width=18,
                        checkbox_height=18).pack(
                            fill="x", padx=2, pady=2, anchor="w")
        actions = ctk.CTkFrame(dialog, fg_color="transparent")
        actions.pack(fill="x", padx=14, pady=(0, 10))
        VisualButton(
            actions, text=I.t("settings_cancel"), width=90, height=ACTION_HEIGHT,
            fg_color=CARD, text_color=MUTED, border_width=1, border_color=LINE,
            hover_color=ACCENT_SOFT, command=dialog.destroy).pack(side="left")
        VisualButton(
            actions, text=I.t("treatment_add_current"), width=150,
            height=ACTION_HEIGHT, fg_color=PRIMARY, hover_color=ACCENT_HOVER,
            command=lambda: self._apply_treatment_selection(
                selections, False, dialog, template_id)).pack(
                    side="right", padx=(6, 0))
        VisualButton(
            actions, text=I.t("treatment_replace_current"), width=165,
            height=ACTION_HEIGHT, fg_color=CARD, text_color=ACCENT,
            border_width=1, border_color=LINE, hover_color=ACCENT_SOFT,
            command=lambda: self._apply_treatment_selection(
                selections, True, dialog, template_id)).pack(side="right")
        dialog.grab_set()
        dialog.focus_set()

    def _apply_treatment_selection(self, selections, replace, dialog,
                                   template_id=""):
        if getattr(self, "_applying_template", False):
            return
        if not self._confirm_treatment_leave():
            return
        scroll_position = self._current_scroll_position()
        selected = [group[choice.get()] for group, choice, included in selections
                    if included.get()]
        if not selected:
            messagebox.showwarning(
                I.t("treatment_apply_preview"),
                I.t("treatment_apply_none_selected"), parent=dialog)
            return
        if replace:
            self._template_row_reuse = list(self.rows)
            for row in self._template_row_reuse:
                row.pack_forget()
                row._autocomplete_token += 1
                if row._autocomplete_job is not None:
                    row.after_cancel(row._autocomplete_job)
                    row._autocomplete_job = None
                row._hide_ac()
            self.rows.clear()
        elif len(self.rows) == 1:
            row = self.rows[0]
            if not any((row.name_var.get().strip(), row.trade_var.get().strip(),
                        row.dosage_var.get().strip(), row.freq_var.get().strip(),
                        row.dur_var.get().strip(), row.notes_var.get().strip())):
                row.destroy()
                self.rows.clear()
        captured = [qu.DrugItem(**{key: str(medicine.get(key, "")) for key in (
            "generic_name", "brand_name", "dosage", "frequency", "duration", "notes")})
                    for medicine in selected]
        self._applying_template = True
        self._row_batch_depth = getattr(self, "_row_batch_depth", 0) + 1
        if hasattr(dialog, "protocol"):
            dialog.protocol("WM_DELETE_WINDOW", lambda: None)
            def disable_actions(widget):
                if isinstance(widget, ctk.CTkButton):
                    widget.configure(state="disabled")
                for child in widget.winfo_children():
                    disable_actions(child)
            disable_actions(dialog)
        def finish(success=True):
            try:
                for unused in getattr(self, "_template_row_reuse", []):
                    unused.destroy()
                self._template_row_reuse = []
                self._row_batch_depth -= 1
                self._number_drug_rows()
                if success and template_id:
                    cfg.config.record_treatment_template_use(template_id)
            finally:
                self._applying_template = False
                dialog.grab_release()
                dialog.destroy()
                self.on_any_change()
                self._restore_scroll_position(scroll_position)
        def append_one(index=0):
            if self._closing:
                return
            try:
                self.add_row(data=captured[index])
                if index:
                    self.rows[-2].collapse_if_complete()
            except Exception:
                finish(False)
                raise
            if index+1 < len(captured):
                self.after(1, lambda: append_one(index+1))
            else:
                finish()
        self.after(1, append_one)

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

    @contextmanager
    def _batch_medication_rows(self):
        self._row_batch_depth = getattr(self, "_row_batch_depth", 0) + 1
        try:
            yield
        finally:
            self._row_batch_depth -= 1
            if not self._row_batch_depth:
                for row in self.rows[:-1]:
                    row.collapse_if_complete()
                self._number_drug_rows()
                self.on_any_change()

    def add_row(self, data=None):
        # Keep the active medicine spacious while completed medicines become
        # compact summaries. This materially reduces scrolling on long Rx lists.
        if not getattr(self, "_row_batch_depth", 0):
            for existing in self.rows:
                if existing.expanded:
                    existing.collapse_if_complete()
        reuse = getattr(self, "_template_row_reuse", [])
        if reuse:
            row = reuse.pop(0)
            row._category = ""
            row.set_expanded(True)
        else:
            row = DrugRow(self.drugs_frame, self.db, self.on_any_change,
                      lambda: self.remove_row(row),
                      lambda r: self.move_row(r, -1), lambda r: self.move_row(r, 1),
                      self.drag_row,
                      fg_color=CARD, border_color=LINE, border_width=1, corner_radius=12)
        if data:
            row._loading_values = True
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
            row._loading_values = False
            if not saved_quantity.strip():
                row.recalculate_quantity()
            row._refresh_compact_summary()
            row._refresh_quantity_status()
        row.pack(fill="x", padx=2, pady=4)
        self.rows.append(row)
        row.set_expanded(True)
        if not getattr(self, "_row_batch_depth", 0):
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
            visible_ids = {favorite["id"] for _, favorite in visible_indexed}
            for favorite_id, widgets in self._favorite_card_widgets.items():
                if favorite_id not in visible_ids:
                    widgets["card"].grid_remove()
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
            generation = getattr(self, "_favorite_render_generation", 0) + 1
            self._favorite_render_generation = generation
            render_container = self.favorite_cards
            def render_batch(start=0):
                if (self._closing or not render_container.winfo_exists()
                        or generation != self._favorite_render_generation):
                    return
                if getattr(self, "active_page", "") != "favorites":
                    self._pending_favorite_render = lambda: render_batch(start)
                    return
                if start:
                    previous = visible_indexed[start-1][1]["id"]
                    if not self._card_near_viewport(self._favorite_card_widgets[previous]["card"]):
                        self.after(150, lambda: render_batch(start))
                        return
                deadline = time.perf_counter() + .008
                while start < len(visible_indexed):
                    source_index, favorite = visible_indexed[start]
                    self._render_favorite_card(start, source_index, favorite)
                    start += 1
                    card = self._favorite_card_widgets[favorite["id"]]["card"]
                    if time.perf_counter() >= deadline or not self._card_near_viewport(card):
                        break
                if start < len(visible_indexed):
                    self.after(1, lambda: render_batch(start))
                else:
                    self._restore_scroll_position(scroll_position)
                    footer = getattr(self, "_favorite_more_frame", None)
                    if footer is not None:
                        self._queue_card_loading("favorites", generation,
                            footer, self._load_more_favorites)
            render_batch()
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

    def _queue_card_loading(self, page, generation, footer, load_more):
        """Load the next slice automatically when its footer enters the buffer."""
        def check():
            current = getattr(self, "_favorite_render_generation" if page == "favorites"
                              else "_saved_template_render_generation", None)
            if self._closing or generation != current or not footer.winfo_exists():
                return
            if self.active_page != page:
                setattr(self, "_pending_favorite_render" if page == "favorites"
                        else "_pending_template_render", check)
                return
            if self._card_near_viewport(footer):
                load_more()
            else:
                self.after(150, check)
        self.after(150, check)

    def _maintain_card_retention(self):
        """Retain geometry/state but bound heavy off-screen card contents."""
        if self._closing:
            return
        canvas = self.scroll._parent_canvas
        top = canvas.winfo_rooty() - 160
        bottom = top + canvas.winfo_height() + 320
        def nearby(card):
            return (card.winfo_viewable() and card.winfo_rooty() < bottom
                    and card.winfo_rooty() + card.winfo_height() > top)
        pools = [list(getattr(self, "_favorite_card_widgets", {}).items()),
                 list(getattr(self, "_saved_template_card_pool", {}).items())]
        deadline = time.perf_counter() + .008
        for kind, entries in enumerate(pools):
            live = sum(not (item.get("placeholder", False) if kind == 0
                           else getattr(item, "_card_placeholder", False)) for _, item in entries)
            for identifier, item in entries:
                card = item["card"] if kind == 0 else item
                placeholder = item.get("placeholder", False) if kind == 0 else getattr(card, "_card_placeholder", False)
                if placeholder and nearby(card):
                    if kind == 0:
                        self._render_favorite_card(*item["render_args"])
                    else:
                        info = card.grid_info()
                        replacement = self._build_saved_treatment_card(card._template_record)
                        replacement._template_signature = card._template_signature
                        replacement.grid(**info)
                        self._saved_template_card_pool[identifier] = replacement
                        self._treatment_saved_card_widgets = [replacement if old is card else old
                                                               for old in self._treatment_saved_card_widgets]
                        card.destroy()
                    live += 1
                elif not placeholder and live > 96 and not nearby(card):
                    # Keep a lightweight frame as a spacer so scroll positions
                    # and category grid ordering do not jump when content leaves.
                    card.configure(height=max(1, card._reverse_widget_scaling(card.winfo_height())))
                    card.grid_propagate(False)
                    card.pack_propagate(False)
                    for child in card.winfo_children():
                        if child is not getattr(card, "_canvas", None):
                            child.destroy()
                    card._release_glass()
                    if kind == 0:
                        item["placeholder"] = True
                        item.pop("star", None)
                        item.pop("selected", None)
                    else:
                        card._card_placeholder = True
                        card._template_body = None
                        card._template_chevron = None
                    live -= 1
                if time.perf_counter() >= deadline:
                    break
        self.after(250, self._maintain_card_retention)

    def _card_near_viewport(self, card):
        """Build only the visible card rows plus a small scrolling buffer."""
        canvas = getattr(getattr(self, "scroll", None), "_parent_canvas", None)
        if canvas is None or not card.winfo_ismapped():
            return True
        return card.winfo_rooty() <= canvas.winfo_rooty() + canvas.winfo_height() + 160

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
        scroll_position = self._current_scroll_position()
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
        for _index, favorite in selected:
            drug_fields = {key: favorite.get(key, "") for key in
                           ("generic_name", "brand_name", "dosage", "frequency",
                            "duration", "notes")}
            self.add_row(data=qu.DrugItem(**drug_fields))
        self._restore_scroll_position(scroll_position)

    @staticmethod
    def _favorite_autocomplete_navigation_key(event):
        return bool(event and event.keysym in {
            "Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R",
            "Control_L", "Control_R", "Alt_L", "Alt_R",
        })

    def _on_favorite_brand_type(self, event=None):
        if self._favorite_autocomplete_navigation_key(event):
            return
        query = self.favorite_vars["brand_name"].get().strip()
        self._schedule_favorite_autocomplete(query, "brand")

    def _schedule_favorite_autocomplete(self, query, mode):
        self.cancel_latest_search("favorite_autocomplete")
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

            self.submit_latest_search("favorite_autocomplete",
                                      lambda: search(query, 20), finished)

        self._favorite_ac_job = self.after(35, launch)

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
        if widgets is not None and (widgets.get("placeholder") or widgets["signature"] != signature):
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
            name_line.grid(row=0, column=0, sticky="ew")
            brand_name = favorite.get("brand_name", "").strip()
            scientific_name = favorite.get("generic_name", "").strip()
            if brand_name:
                ctk.CTkLabel(
                    name_line, text=brand_name, text_color=TEXT,
                    anchor="w", justify="left",
                    font=_ui_font(CARD_NAME_SIZE, "bold")).pack(side="left")
            if scientific_name and scientific_name.casefold() != brand_name.casefold():
                separator = "  ·  " if brand_name else ""
                ctk.CTkLabel(
                    name_line, text=f"{separator}{scientific_name}", text_color=MUTED,
                    anchor="w", justify="left",
                    font=_ui_font(CARD_NAME_SIZE)).pack(side="left")
            wrap_card_labels(name_line)
            actions = ctk.CTkFrame(card, fg_color="transparent")
            actions.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 5))
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
            if favorite.get("category", "").strip():
                ctk.CTkLabel(
                    actions, text=favorite["category"].strip(), height=20,
                    wraplength=100, corner_radius=10,
                    fg_color=ACCENT_SOFT, text_color=ACCENT,
                    font=ctk.CTkFont(size=9, weight="bold")).pack(
                        side="right", padx=(7, 0))
            widgets = {
                "card": card, "star": star_button, "selected": selected_var,
                "signature": signature,
            }
            self._favorite_card_widgets[favorite_id] = widgets
        else:
            card = widgets["card"]
            star_text = "★" if favorite.get("pinned") else "☆"
            if widgets["star"].cget("text") != star_text:
                widgets["star"].configure(text=star_text)
            if widgets["selected"].get() != selected:
                widgets["selected"].set(selected)
        widgets["render_args"] = (display_index, source_index, favorite)

        highlighted = favorite_id == self._favorite_highlight_id
        style = (highlighted, selected)
        if widgets.get("style") != style:
            card.configure(
                fg_color=ACCENT_SOFT if highlighted else CARD,
                border_color=ACCENT if selected or highlighted else LINE,
                border_width=2 if selected or highlighted else 1)
            widgets["style"] = style
        columns = getattr(self, "_favorite_card_columns", 3)
        position = (display_index // columns, display_index % columns)
        if widgets.get("position") != position or not card.winfo_manager():
            card.grid(row=position[0], column=position[1],
                      sticky="new", padx=CARD_GAP, pady=CARD_GAP)
            widgets["position"] = position
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
        self._ensure_page_built("favorites")
        self.favorite_edit_index = None
        for variable in self.favorite_vars.values():
            variable.set("")
        self.favorite_editor_title.configure(text=I.t("new_favorite"))
        self._show_favorite_editor()

    def edit_favorite(self, identifier):
        self._ensure_page_built("favorites")
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
            saved = self._save_with_feedback("favorite", lambda: cfg.config.update_medication_favorite(duplicate_index, item))
        elif self.favorite_edit_index is None:
            saved = self._save_with_feedback("favorite", lambda: cfg.config.add_medication_favorite(item))
        else:
            saved = self._save_with_feedback("favorite", lambda: cfg.config.update_medication_favorite(self.favorite_edit_index, item))
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
        scroll_position = self._current_scroll_position()
        favorites = cfg.config.medication_favorites()
        index = self._favorite_index(favorites, identifier)
        if index is None:
            return
        favorite = favorites[index]
        cfg.config.record_medication_favorite_use(index)
        drug_fields = {key: favorite.get(key, "") for key in
                       ("generic_name", "brand_name", "dosage", "frequency", "duration", "notes", "quantity")}
        self.add_row(data=qu.DrugItem(**drug_fields))
        self._restore_scroll_position(scroll_position)

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
        self._classification_generation += 1
        self._prepare_classification_cache()

    def _ensure_classification_cache(self):
        """Classify each medicine once and reuse counts/lists across the browser."""
        if self._classification_cache is not None:
            return self._classification_cache
        self._classification_cache = self._build_classification_index(tuple(self.db.drugs))
        return self._classification_cache

    @staticmethod
    def _build_classification_index(drugs):
        by_group = {code: [] for code in classes.GROUPS}
        by_pair = {}
        unclassified = []
        suggested = []
        medicines_by_name = {}
        mappings_by_name = {}
        medicines_by_identity = {}
        mappings_by_identity = {}
        for drug in drugs:
            mappings = classes.groups_for(drug)
            normalized_name = drug.generic_name.strip().casefold()
            identity = dbmod.DrugDatabase.drug_identity(drug)
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
        return {
            "by_group": by_group, "by_pair": by_pair,
            "unclassified": unclassified, "suggested": suggested,
            "variations": variations, "conflicting": conflicting,
        }

    def unclassified_medicine_count(self):
        return len(self._ensure_classification_cache()["unclassified"])

    def therapeutic_group_medicine_count(self, code):
        return len(self._ensure_classification_cache()["by_group"].get(code, ()))

    def refresh_class_overview(self):
        if not hasattr(self, "class_tiles"):
            return
        if self._defer_classification_render(self.refresh_class_overview):
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
        if self._defer_classification_render(self.refresh_class_browser):
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
            detail_row = ctk.CTkFrame(
                self.class_detail_list, fg_color="transparent")
            detail_row.pack(fill="x", pady=2)
            button = VisualButton(
                detail_row,
                text=I.t("detailed_class_with_count", detail=detail, n=count),
                height=34, corner_radius=8, anchor="w",
                fg_color=PRIMARY if selected else CARD,
                text_color="white" if selected else TEXT,
                border_width=1, border_color=ACCENT if selected else LINE,
                hover_color=ACCENT_SOFT,
                command=lambda picked=detail:
                    self.schedule_class_detail_selection(picked))
            button.pack(side="left", fill="x", expand=True)
            VisualButton(
                detail_row, text="›", width=32, height=32, corner_radius=16,
                fg_color="transparent", text_color=ACCENT,
                hover_color=ACCENT_SOFT, font=_ui_font(18, "bold"),
                command=lambda group=code, picked=detail:
                    self.show_detail_medicines_page(group, picked)).pack(
                        side="right", padx=(4, 0))
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
        self._ensure_page_built("drug_classes")
        """Edit local class metadata without requiring CSV editing."""
        restored_state = (getattr(self, "_mapping_editor_state", {})
                          if not review_unclassified and target_drug is None else {})
        restored_mode = restored_state.get("filter", "all")
        if restored_mode not in {"all", "unclassified", "suggested", "conflicting"}:
            restored_mode = "all"
        initial_mode = "unclassified" if review_unclassified else restored_mode
        initial_query = "" if (review_unclassified or target_drug) else restored_state.get("query", "")
        self._capture_class_overview_state()
        self._class_subpage_kind = "mapping"
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
        right = GlassFrame(body, width=500, fg_color=SURFACE,
                             border_color=LINE, border_width=1, corner_radius=12)
        right.pack(side="left", fill="both")
        self.mapping_search_var = tk.StringVar(value=initial_query)
        self._mapping_target_identity = (
            self.db.drug_identity(target_drug) if target_drug else None)
        self._mapping_search_job = None
        self._mapping_search_syncing = False
        self._mapping_search_previous = initial_query
        self.mapping_search_var.trace_add("write", lambda *_: self._schedule_mapping_refresh())
        mapping_control_font = ctk.CTkFont(size=13)
        VisualEntry(
            left, textvariable=self.mapping_search_var, height=FIELD_HEIGHT,
            placeholder_text=I.t("search_medicines"), border_width=2,
            border_color=ACCENT, corner_radius=9,
            font=mapping_control_font).pack(fill="x", pady=(0, 6))
        self.mapping_unclassified_only = tk.BooleanVar(value=initial_mode == "unclassified")
        self.mapping_suggested_only = tk.BooleanVar(value=initial_mode == "suggested")
        self.mapping_conflicting_only = tk.BooleanVar(value=initial_mode == "conflicting")
        self.mapping_filter_var = tk.StringVar(value=initial_mode)
        mapping_filters = ctk.CTkFrame(left, fg_color="transparent")
        mapping_filters.pack(fill="x", pady=(0, 6))
        self.mapping_filter_buttons = {}
        for mode, label in (
                ("all", I.t("mapping_filter_all")),
                ("unclassified", I.t("unclassified")),
                ("suggested", I.t("classification_suggested")),
                ("conflicting", I.t("classification_conflicting"))):
            button = VisualButton(
                mapping_filters, text=label, width=88, height=30,
                corner_radius=15, border_width=1, border_color=LINE,
                font=_ui_font(10, "bold"),
                command=lambda selected=mode: self._set_mapping_filter(selected))
            button.pack(side="left", padx=(0, 5))
            self.mapping_filter_buttons[mode] = button
        self.mapping_filter_count_label = ctk.CTkLabel(
            mapping_filters, text="", text_color=MUTED,
            font=_ui_font(10, "bold"), anchor="e")
        self.mapping_filter_count_label.pack(side="right")
        self._style_mapping_filter_buttons()
        # Keep the medicine name prominent while class/status stays compact on
        # a separate secondary line.  Canvas drawing avoids hundreds of widgets.
        self.mapping_list = MappingMedicineList(
            left, bg=SURFACE, fg=TEXT, line=LINE,
            select_bg=PRIMARY, select_fg="white")
        self.mapping_list.configure(selectmode=tk.EXTENDED, exportselection=False)
        self.mapping_list.pack(fill="both", expand=True)
        self.mapping_list.bind("<<ListboxSelect>>", self.select_mapping_drug)
        self.mapping_list.set_selection_guard(self._guard_mapping_selection)
        review_actions = ctk.CTkFrame(left, fg_color="transparent")
        review_actions.pack(fill="x", pady=(6, 0))
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
                                                        weight="bold"), wraplength=430)
        self.mapping_selected_label.pack(fill="x", padx=12, pady=(10, 8))
        self._mapping_name_syncing = False
        self.mapping_brand_var = tk.StringVar()
        self.mapping_generic_var = tk.StringVar()
        self.mapping_brand_var.trace_add("write", self._mapping_name_changed)
        self.mapping_generic_var.trace_add("write", self._mapping_name_changed)
        self.mapping_name_entries = []
        mapping_names = ctk.CTkFrame(right, fg_color="transparent")
        mapping_names.pack(fill="x", padx=12, pady=(0, 4))
        mapping_names.grid_columnconfigure((0, 1), weight=1)
        for label_key, variable in (("brand_name", self.mapping_brand_var),
                                    ("scientific_name", self.mapping_generic_var)):
            column = len(self.mapping_name_entries)
            field = ctk.CTkFrame(mapping_names, fg_color="transparent")
            field.grid(row=0, column=column, sticky="ew",
                       padx=(0, 4) if column == 0 else (4, 0))
            ctk.CTkLabel(
                field, text=I.t(label_key), text_color=MUTED,
                font=ctk.CTkFont(size=10, weight="bold"), anchor="w").pack(fill="x")
            entry = VisualEntry(
                field, textvariable=variable, height=FIELD_HEIGHT,
                corner_radius=9, border_color=LINE, font=mapping_control_font,
                state="disabled")
            entry.pack(fill="x", pady=(3, 3))
            self.mapping_name_entries.append(entry)
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
        self.mapping_unsaved_label = ctk.CTkLabel(
            right, text="", text_color=WARNING,
            font=ctk.CTkFont(size=10, weight="bold"), anchor="w")
        self.mapping_unsaved_label.pack(fill="x", padx=12, pady=(0, 4))
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
            text_color="white", text_color_disabled="#ffffff",
            font=_ui_font(11, "bold"),
            state="disabled", command=self.save_class_mapping)
        self.save_mapping_button.pack(
            side="left", fill="x", expand=True, padx=(3, 0))
        self.mapping_save_status = ctk.CTkLabel(mapping_actions, text="", width=72, font=_ui_font(11))
        self.mapping_save_status.pack(side="left", padx=4)
        self.mapping_status = ctk.CTkLabel(right, text="", text_color=MUTED,
                                           font=ctk.CTkFont(size=10), wraplength=430, justify="left")
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
        elif restored_state:
            self._restore_mapping_editor_state(restored_state)

    def _mapping_has_unsaved_changes(self):
        return bool(getattr(self, "pending_class_mapping", None) or
                    self._mapping_names_changed())

    def _discard_mapping_changes(self):
        self.pending_class_mapping = None
        if hasattr(self, "mapping_list") and self.mapping_list.curselection():
            self.select_mapping_drug()
        else:
            self._refresh_mapping_save_state()
        if hasattr(self, "mapping_status"):
            self.mapping_status.configure(text="", text_color=MUTED)

    def _confirm_mapping_unsaved_changes(self):
        """Save, discard, or keep editing before mapping context changes."""
        if getattr(self, "_mapping_save_busy", False):
            return False
        if not self._mapping_has_unsaved_changes():
            return True
        choice = messagebox.askyesnocancel(
            I.t("mapping_unsaved_title"), I.t("mapping_unsaved_message"), parent=self)
        if choice is None:
            return False
        if choice:
            return bool(self.save_class_mapping())
        self._discard_mapping_changes()
        return True

    def _guard_mapping_selection(self, proposed_indices):
        if getattr(self, "_mapping_save_busy", False):
            return False
        if not self._mapping_has_unsaved_changes():
            return proposed_indices
        target_identities = {
            self.db.drug_identity(self.mapping_visible_drugs[index])
            for index in proposed_indices if index < len(self.mapping_visible_drugs)
        }
        if not self._confirm_mapping_unsaved_changes():
            return False
        return {
            index for index, drug in enumerate(self.mapping_visible_drugs)
            if self.db.drug_identity(drug) in target_identities
        }

    def _capture_mapping_editor_state(self):
        if (getattr(self, "_class_subpage_kind", "") != "mapping" or
                not hasattr(self, "mapping_list")):
            return
        selected = {
            self.db.drug_identity(self.mapping_visible_drugs[index])
            for index in self.mapping_list.curselection()
            if index < len(self.mapping_visible_drugs)
        }
        try:
            scroll = self.mapping_list.yview()[0]
        except (tk.TclError, IndexError):
            scroll = 0.0
        self._mapping_editor_state = {
            "query": self.mapping_search_var.get(),
            "filter": self.mapping_filter_var.get(),
            "selected": selected,
            "scroll": scroll,
            "keep_existing": self.mapping_keep_existing_var.get(),
        }

    def _restore_mapping_editor_state(self, state):
        selected = state.get("selected", set())
        indices = [
            index for index, drug in enumerate(self.mapping_visible_drugs)
            if self.db.drug_identity(drug) in selected
        ]
        for index in indices:
            self.mapping_list.selection_set(index)
        if indices:
            self.mapping_list.activate(indices[0])
            self.select_mapping_drug()
        self.mapping_keep_existing_var.set(bool(state.get("keep_existing", True)))
        self.mapping_list.yview_moveto(float(state.get("scroll", 0.0)))

    def _schedule_mapping_refresh(self, delay=180):
        if getattr(self, "_mapping_search_syncing", False):
            return
        if self._mapping_has_unsaved_changes():
            requested = self.mapping_search_var.get()
            if not self._confirm_mapping_unsaved_changes():
                self._mapping_search_syncing = True
                try:
                    self.mapping_search_var.set(self._mapping_search_previous)
                finally:
                    self._mapping_search_syncing = False
                return
            self._mapping_search_previous = requested
        if self._mapping_search_job is not None:
            try:
                self.after_cancel(self._mapping_search_job)
            except (tk.TclError, ValueError):
                pass
        self._mapping_search_job = self.after(delay, self.refresh_mapping_list)

    def _set_mapping_filter(self, mode):
        """Use one explicit mapping review mode instead of overlapping filters."""
        if mode not in {"all", "unclassified", "suggested", "conflicting"}:
            mode = "all"
        if mode != self.mapping_filter_var.get() and not self._confirm_mapping_unsaved_changes():
            return
        self.mapping_filter_var.set(mode)
        self.mapping_unclassified_only.set(mode == "unclassified")
        self.mapping_suggested_only.set(mode == "suggested")
        self.mapping_conflicting_only.set(mode == "conflicting")
        self._style_mapping_filter_buttons()
        self.refresh_mapping_list()

    def _style_mapping_filter_buttons(self):
        current = self.mapping_filter_var.get()
        for mode, button in getattr(self, "mapping_filter_buttons", {}).items():
            selected = mode == current
            button.configure(
                fg_color=PRIMARY if selected else CARD,
                text_color="white" if selected else ACCENT,
                hover_color=ACCENT_HOVER if selected else ACCENT_SOFT,
                border_color=PRIMARY if selected else LINE)

    def refresh_mapping_list(self):
        if getattr(self, "_mapping_save_busy", False):
            return
        if not hasattr(self, "mapping_list"):
            return
        if self._defer_classification_render(self.refresh_mapping_list):
            return
        self._mapping_search_job = None
        self._mapping_search_previous = self.mapping_search_var.get()
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
        if hasattr(self, "mapping_filter_count_label"):
            suffix = "+" if total_visible > 500 else ""
            self.mapping_filter_count_label.configure(
                text=I.t("medicine_count", n=f"{len(self.mapping_visible_drugs)}{suffix}"))
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
                medicine_name,
                suffix,
                status=I.t("classification_" + status), status_kind=status,
                trade_name=drug.brand_name.strip() if drug.generic_name.strip() else "")
        self.mapping_list.refresh()
        self.mapping_selected_drugs = []
        self.mapping_selected_drug = None
        self.pending_class_mapping = None
        if hasattr(self, "mapping_selected_label"):
            self.mapping_selected_label.configure(text=I.t("no_medicine_selected"))
        if hasattr(self, "mapping_name_entries"):
            self._mapping_name_syncing = True
            try:
                self.mapping_brand_var.set("")
                self.mapping_generic_var.set("")
            finally:
                self._mapping_name_syncing = False
            for entry in self.mapping_name_entries:
                entry.configure(state="disabled")
        if hasattr(self, "add_mapping_drug_button"):
            self.add_mapping_drug_button.configure(state="disabled")
        if hasattr(self, "save_mapping_button"):
            self.save_mapping_button.configure(state="disabled")
        if hasattr(self, "mapping_unsaved_label"):
            self.mapping_unsaved_label.configure(text="")

    def select_mapping_drug(self, event=None):
        selected_indices = self.mapping_list.curselection()
        if not selected_indices:
            return
        selected = [self.mapping_visible_drugs[index] for index in selected_indices]
        self.mapping_selected_drugs = selected
        drug = selected[0]
        self.mapping_selected_drug = drug
        if len(selected) == 1:
            selected_name = drug.generic_name.strip() or drug.brand_name.strip()
            if drug.brand_name.strip() and drug.brand_name.strip().casefold() != selected_name.casefold():
                selected_name += "  ·  " + drug.brand_name.strip()
            self.mapping_selected_label.configure(text=selected_name)
            self._mapping_name_syncing = True
            try:
                self.mapping_brand_var.set(drug.brand_name)
                self.mapping_generic_var.set(drug.generic_name)
            finally:
                self._mapping_name_syncing = False
            for entry in self.mapping_name_entries:
                entry.configure(state="normal")
        else:
            self.mapping_selected_label.configure(
                text=I.t("mapping_selected_count", n=len(selected)))
            self._mapping_name_syncing = True
            try:
                self.mapping_brand_var.set("")
                self.mapping_generic_var.set("")
            finally:
                self._mapping_name_syncing = False
            for entry in self.mapping_name_entries:
                entry.configure(state="disabled")
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
        self._refresh_mapping_save_state()
        self.mapping_status.configure(text="", text_color=MUTED)

    def _mapping_names_changed(self):
        if len(getattr(self, "mapping_selected_drugs", [])) != 1:
            return False
        drug = self.mapping_selected_drugs[0]
        return (self.mapping_brand_var.get().strip() != drug.brand_name.strip()
                or self.mapping_generic_var.get().strip() != drug.generic_name.strip())

    def _refresh_mapping_save_state(self):
        if not hasattr(self, "save_mapping_button"):
            return
        enabled = (not getattr(self, "_mapping_save_busy", False)
                   and bool(self.pending_class_mapping or self._mapping_names_changed()))
        self.save_mapping_button.configure(state="normal" if enabled else "disabled")
        if hasattr(self, "mapping_unsaved_label"):
            self.mapping_unsaved_label.configure(
                text=I.t("mapping_unsaved_changes") if enabled else "")

    def _mapping_name_changed(self, *_args):
        if getattr(self, "_mapping_name_syncing", False):
            return
        self._refresh_mapping_save_state()

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
        self._refresh_mapping_save_state()
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
        self._refresh_mapping_save_state()

    def _set_mapping_save_busy(self, busy):
        self._mapping_save_busy = busy
        controls = list(getattr(self, "mapping_name_entries", []))
        controls.extend(getattr(self, "mapping_filter_buttons", {}).values())
        controls.extend(getattr(self, name, None) for name in (
            "mapping_group_menu", "mapping_detail_menu",
            "mapping_keep_existing_checkbox", "add_mapping_drug_button",
            "save_mapping_button"))
        for control in controls:
            if control is not None:
                control.configure(state="disabled" if busy else "normal")
        if busy:
            self._set_save_feedback("mapping", "saving")
            self.mapping_status.configure(text=I.t("mapping_saving"), text_color=MUTED)

    def save_class_mapping(self, advance=False):
        if getattr(self, "_mapping_save_busy", False):
            return False
        mapping = self.pending_class_mapping
        medicines = tuple(mapping[0]) if mapping else tuple(self.mapping_selected_drugs)
        rename_requested = bool(len(medicines) == 1 and self._mapping_names_changed())
        if not mapping and not rename_requested:
            return False
        brand_name = self.mapping_brand_var.get().strip() if rename_requested else ""
        generic_name = self.mapping_generic_var.get().strip() if rename_requested else ""
        if rename_requested and not (brand_name or generic_name):
            self.mapping_status.configure(
                text=I.t("mapping_name_required"), text_color=DANGER)
            return False
        code = mapping[1] if mapping else ""
        detail = mapping[2] if mapping else ""
        current_indices = self.mapping_list.curselection()
        next_index = current_indices[0] if current_indices else 0
        selected_identities = {self.db.drug_identity(drug) for drug in medicines}
        try:
            previous_scroll = self.mapping_list.yview()[0]
        except (tk.TclError, IndexError):
            previous_scroll = 0.0
        filter_mode = (self.mapping_filter_var.get()
                       if hasattr(self, "mapping_filter_var") else "all")
        database = self.db
        append = self.mapping_keep_existing_var.get()
        rename = (generic_name, brand_name) if rename_requested else None
        self._set_mapping_save_busy(True)

        def worker():
            return database.update_drug_mapping_and_names(
                medicines, code, detail, append=append, rename=rename)

        def completed(result):
            nonlocal selected_identities
            self._set_mapping_save_busy(False)
            mapped, renamed, prior_states = result
            getattr(self, "_set_save_feedback", lambda *_: None)("mapping", "saved" if mapped or renamed else "failed")
            if renamed:
                for state in prior_states:
                    state.update({
                        "current_generic_name": generic_name,
                        "current_brand_name": brand_name,
                        "current_strength": medicines[0].strength,
                        "current_form": medicines[0].form,
                    })
                renamed_drug = dbmod.Drug(
                    generic_name=generic_name, brand_name=brand_name,
                    strength=medicines[0].strength, form=medicines[0].form)
                selected_identities = {self.db.drug_identity(renamed_drug)}
                if getattr(self, "_mapping_target_identity", None):
                    self._mapping_target_identity = next(iter(selected_identities))
            if mapped or renamed:
                if prior_states:
                    cfg.config.add_recovery_item(
                        "mapping", I.t("mapping_recovery_label", n=len(medicines)), prior_states)
                self._invalidate_database_caches()
                success_message = (I.t(
                    "class_mapping_batch_saved", n=mapped,
                    group=I.t("class_" + code), detail=detail)
                    if mapping else I.t("mapping_name_updated"))
                self.refresh_mapping_list()
                self.refresh_class_overview()
                self.pending_class_mapping = None
                self.mapping_selected_drugs = []
                self.mapping_selected_drug = None
                self.save_mapping_button.configure(state="disabled")
                if self.mapping_visible_drugs:
                    if advance:
                        # In All mode the saved row remains visible, so move one
                        # further. Filtered queues usually remove it automatically.
                        target_index = next_index + (1 if filter_mode == "all" else 0)
                    else:
                        target_index = next((index for index, drug in enumerate(
                            self.mapping_visible_drugs)
                            if self.db.drug_identity(drug) in selected_identities), next_index)
                    target_index = min(target_index, len(self.mapping_visible_drugs) - 1)
                    self.mapping_list.selection_set(target_index)
                    self.mapping_list.activate(target_index)
                    if advance:
                        self.mapping_list.see(target_index)
                    else:
                        self.mapping_list.yview_moveto(previous_scroll)
                    self.select_mapping_drug()
                self.mapping_status.configure(text=success_message, text_color=GOOD)
                return True

            self.mapping_status.configure(text=I.t("mapping_save_failed"), text_color=DANGER)

        def failed(_error):
            getattr(self, "_set_save_feedback", lambda *_: None)("mapping", "failed")
            self._set_mapping_save_busy(False)
            self._refresh_mapping_save_state()
            self.mapping_status.configure(text=I.t("mapping_save_failed"), text_color=DANGER)

        try:
            self.submit_background(worker, completed, on_error=failed,
                                   label=I.t("mapping_saving"))
        except Exception:
            failed(None)
        # A navigation guard must stay on the editor until completion.
        return False

    def mapping_select_relative(self, step):
        if not self.mapping_visible_drugs:
            return
        selected = self.mapping_list.curselection()
        current = selected[0] if selected else (0 if step > 0 else len(self.mapping_visible_drugs) - 1)
        target = max(0, min(len(self.mapping_visible_drugs) - 1, current + step))
        self.mapping_list.request_selection({target}, target)

    def mapping_save_and_next(self):
        if not self.mapping_selected_drugs:
            self.mapping_select_relative(1)
            return
        if self._mapping_names_changed() and not self.pending_class_mapping:
            self.save_class_mapping(advance=True)
            return
        if not self.pending_class_mapping:
            self.add_drug_to_class()
        if self.pending_class_mapping:
            self.save_class_mapping(advance=True)

    def open_therapeutic_group(self, code):
        """Open a focused subpage containing this group's detailed classes."""
        self._capture_class_overview_state()
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
        self._ensure_page_built("drug_classes")
        """Show every configured detailed class, grouped on one scrollable page."""
        self._capture_class_overview_state()
        self._class_subpage_kind = "all"
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

    def show_subclass_page(self, group_code, selected_detail=None, return_to_all=False,
                           open_selected=True):
        self._ensure_page_built("drug_classes")
        self._class_subpage_kind = "subclass"
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
        list_frame.grid_columnconfigure((0, 1), weight=1, uniform="subclass_cards")
        self.subclass_list_frame = list_frame
        self.subclass_buttons = {}
        subclass_counts = {}
        for detail in classes.subclasses_for(group_code):
            subclass_counts[detail] = len(self._drugs_in_class(group_code, detail))
        for index, detail in enumerate(classes.subclasses_for(group_code)):
            tile = ctk.CTkFrame(list_frame, fg_color="transparent")
            tile.grid(row=index // 2, column=index % 2, sticky="ew", padx=3, pady=3)
            tile.grid_columnconfigure(0, weight=1)
            button = VisualButton(
                tile, text=I.t("detailed_class_with_count", detail=detail,
                               n=subclass_counts[detail]), height=36, corner_radius=9, anchor="w",
                fg_color=SURFACE, text_color=TEXT, border_width=1,
                border_color=LINE, hover_color=ACCENT_SOFT,
                command=lambda picked=detail:
                    self.schedule_subclass_selection(picked))
            button.grid(row=0, column=0, sticky="ew")
            VisualButton(
                tile, text="›", width=32, height=32, corner_radius=16,
                fg_color="transparent", text_color=ACCENT,
                hover_color=ACCENT_SOFT, font=_ui_font(18, "bold"),
                command=lambda group=group_code, picked=detail:
                    self.show_detail_medicines_page(group, picked)).grid(
                        row=0, column=1, padx=(4, 0))
            self.subclass_buttons[detail] = button
        if selected_detail and open_selected:
            if return_to_all:
                self._class_subpage_kind = "all"
            self.show_detail_medicines_page(group_code, selected_detail)
        elif selected_detail:
            self._finish_subclass_selection(selected_detail)
            position = getattr(self, "_subclass_scroll_position", 0.0)
            self.after_idle(lambda: list_frame._parent_canvas.yview_moveto(position))

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
        self._ensure_page_built("drug_classes")
        if self.class_overview.winfo_manager():
            self._capture_class_overview_state()
            self._detail_return_target = "overview"
        else:
            self._detail_return_target = getattr(self, "_class_subpage_kind", "subclass")
        self._class_subpage_kind = "detail"
        subclass_list = getattr(self, "subclass_list_frame", None)
        if subclass_list is not None and subclass_list.winfo_exists():
            try:
                self._subclass_scroll_position = subclass_list._parent_canvas.yview()[0]
            except (tk.TclError, IndexError):
                self._subclass_scroll_position = 0.0
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
            command=lambda: self.back_to_subclass_page(group_code, detail)).pack(side="left")
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

        medicines = self._drugs_in_class(group_code, detail)
        ctk.CTkLabel(
            bar, text=I.t("medicine_count", n=len(medicines)),
            text_color=MUTED, font=_ui_font(11, "bold")).pack(
                side="right", padx=(8, 10))

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

        if not medicines:
            empty = ctk.CTkFrame(
                card, fg_color=SURFACE, border_color=LINE,
                border_width=1, corner_radius=10)
            empty.pack(fill="x", padx=PAD, pady=(0, PAD))
            ctk.CTkLabel(
                empty, text=I.t("no_mapped_medicines"),
                text_color=MUTED, anchor="w").pack(
                    side="left", fill="x", expand=True, padx=12, pady=10)
            VisualButton(
                empty, text="+", width=32, height=32,
                fg_color="transparent", text_color=ACCENT,
                hover_color=ACCENT_SOFT,
                command=self.toggle_detail_drug_creator).pack(
                    side="right", padx=8, pady=6)
            return
        body = ctk.CTkScrollableFrame(card, height=450, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=PAD, pady=(0, PAD))
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
                           border_width=1, corner_radius=CARD_RADIUS)
        row.pack(fill="x", pady=CARD_GAP)
        names = ctk.CTkFrame(row, fg_color="transparent")
        row.grid_columnconfigure(0, weight=1)
        names.grid(row=0, column=0, sticky="ew", padx=10, pady=7)
        actions = ctk.CTkFrame(row, fg_color="transparent")
        actions.grid(row=0, column=1, sticky="ne", padx=4, pady=5)
        brand = drug.brand_name.strip() or drug.generic_name.strip()
        scientific = drug.generic_name.strip() if drug.brand_name.strip() else ""
        ctk.CTkLabel(
            names, text=brand, text_color=TEXT, anchor="w",
            font=_ui_font(CARD_NAME_SIZE, "bold" if drug.brand_name.strip() else "normal")).pack(side="left")
        if scientific and scientific.casefold() != brand.casefold():
            ctk.CTkLabel(
                names, text="  " + scientific, text_color=MUTED, anchor="w",
                font=_ui_font(CARD_NAME_SIZE)).pack(side="left")
        status = self._classification_status_for(drug, group_code, detail)
        status_background, status_foreground = self._classification_status_colors(status)
        ctk.CTkLabel(
            names, text=I.t("classification_" + status), height=22,
            corner_radius=11, fg_color=status_background,
            text_color=status_foreground,
            font=_ui_font(9, "bold")).pack(
                side="left", padx=(8, 0))
        delete_button = VisualButton(
            actions, text="🗑", width=32, height=32,
            fg_color="transparent", text_color=DANGER, border_width=0,
            hover_color=DANGER_SOFT, font=_ui_font(14),
            command=lambda item=drug, group=group_code, picked=detail:
                self.delete_detail_class_drug(item, group, picked))
        delete_button.pack(side="right", padx=(2, 7), pady=5)
        VisualButton(
            actions, text="★" if self._class_drug_is_starred(drug) else "☆",
            width=36, height=32, fg_color="transparent", text_color=ACCENT,
            hover_color=ACCENT_SOFT,
            command=lambda item=drug, group=group_code, picked=detail:
                self.toggle_detail_drug_star(item, group, picked)).pack(
                    side="right", padx=(2, 7), pady=5)
        VisualButton(
            actions, text="", image=_edit_icon(20), width=32, height=32, corner_radius=8,
            fg_color="transparent", text_color=ACCENT,
            border_width=0, hover_color=ACCENT_SOFT,
            font=_ui_font(14, "bold"),
            command=lambda item=drug: self.edit_class_drug_mapping(item)).pack(
                side="right", padx=2, pady=5)
        add_button = VisualButton(
            actions, text="+", width=32, height=32,
            fg_color="transparent", text_color=ACCENT, border_width=0,
            hover_color=ACCENT_SOFT, font=_ui_font(22),
            command=lambda item=drug: self.add_drug_database_item(item))
        add_button.pack(side="right", padx=2, pady=5)
        wrap_card_labels(names)
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
        if getattr(self, "_class_subpage_kind", "") == "mapping":
            if not self._confirm_mapping_unsaved_changes():
                return
            self._capture_mapping_editor_state()
        self.class_subpage.pack_forget()
        if not self.class_page_header.winfo_manager():
            self.class_page_header.pack(fill="x", padx=4, pady=(0, 3))
        self.class_overview.pack(fill="both", expand=True)
        code = self._active_subclass_group
        if code not in classes.GROUPS:
            code = classes.GROUPS[0]
        self.select_therapeutic_group(code)
        self.after_idle(self._restore_class_overview_state)

    def back_to_subclass_page(self, group_code, detail):
        """Return one level while retaining the selected class and scroll position."""
        target = getattr(self, "_detail_return_target", "subclass")
        if target == "all":
            self.show_all_detailed_classes()
            return
        if target == "overview":
            self.back_to_major_groups()
            return
        self.show_subclass_page(
            group_code, selected_detail=detail, open_selected=False)

    def _capture_class_overview_state(self):
        """Remember browser position before opening a Drug Classes subpage."""
        if not hasattr(self, "class_group_list"):
            return
        state = {
            "group": self._selected_therapeutic_group,
            "detail": self._selected_detailed_class,
            "query": self.class_search_var.get() if hasattr(self, "class_search_var") else "",
        }
        for key, widget_name in (("group_scroll", "class_group_list"),
                                 ("detail_scroll", "class_detail_list")):
            widget = getattr(self, widget_name, None)
            try:
                state[key] = widget._parent_canvas.yview()[0]
            except (AttributeError, tk.TclError, IndexError):
                state[key] = 0.0
        self._class_navigation_state = state

    def _restore_class_overview_state(self):
        state = getattr(self, "_class_navigation_state", {})
        if not state:
            return
        self._selected_therapeutic_group = state.get(
            "group", self._selected_therapeutic_group)
        self._selected_detailed_class = state.get(
            "detail", self._selected_detailed_class)
        if hasattr(self, "class_search_var") and self.class_search_var.get() != state.get("query", ""):
            self.class_search_var.set(state.get("query", ""))
        self._class_browser_render_signature = None
        self.refresh_class_browser()
        for key, widget_name in (("group_scroll", "class_group_list"),
                                 ("detail_scroll", "class_detail_list")):
            widget = getattr(self, widget_name, None)
            try:
                widget._parent_canvas.yview_moveto(state.get(key, 0.0))
            except (AttributeError, tk.TclError):
                pass

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
        self._clear_prescription_draft()
        self._workflow_saved_signature = None
        self._set_workflow_step(1)
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
        self.submit_latest_search("patient_history",
            lambda: self.patient_history.search(query),
            lambda records: self._render_patient_history(query, token, records))

    def _render_patient_history(self, query, token, records):
        if (token != self._latest_query_tokens.get("patient_history")
                or self.patient_search_var.get() != query):
            return
        self._patient_history_records = records
        self.patient_history_list.delete(0, tk.END)
        generation = getattr(self, "_patient_list_render_generation", 0) + 1
        self._patient_list_render_generation = generation
        def append_batch(start=0):
            if self._closing or generation != self._patient_list_render_generation:
                return
            labels = []
            for record in records[start:start+100]:
                prescriptions = record.get("prescriptions", [])
                last_saved = max((str(item.get("saved_at", ""))[:10] for item in prescriptions), default="")
                detail = I.t("last_prescription", date=last_saved or I.t("none_short"))
                display = f"{record.get('name', '')} — {detail}"
                labels.append(directional_display_text(display))
            if labels:
                self.patient_history_list.insert(tk.END, *labels)
            if start+100 < len(records):
                self.after(1, lambda: append_batch(start+100))
            else:
                self._select_patient_record(self._loaded_patient_id)
        append_batch()

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
        self._workflow_saved_signature = None
        self._refresh_workflow_summary()
        if not self._suspend_draft:
            self._schedule_draft_save()

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
        try:
            matches = self.patient_history.find_similar(name, age, sex=sex)
        except PatientHistoryError as exc:
            messagebox.showerror(APP_TITLE, str(exc), parent=self)
            return False
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
        self._prescription_render_generation = getattr(self, "_prescription_render_generation", 0) + 1
        self._pending_history_render = None
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
        if getattr(self, "active_page", "") != "patient":
            return
        loaded_id = str(getattr(self, "_loaded_patient_id", ""))
        cached = getattr(self, "_current_history_record", None)
        visible_id = str((cached or {}).get("id", ""))
        # Never compare the current medicine workspace with another patient's
        # history. Browsing a different patient hides this panel until that
        # patient is explicitly loaded and owns the current medicine set.
        if (not loaded_id or (visible_id and visible_id != loaded_id)
                or self._medication_patient_id != loaded_id):
            self._comparison_render_signature = None
            self.patient_comparison_body.pack_forget()
            return
        record = self.patient_history.get(loaded_id)
        current = [row.get_data().__dict__ for row in self.rows
                   if row.get_data().generic_name or row.get_data().brand_name]
        prescriptions = list((record or {}).get("prescriptions", []))
        if not record or not prescriptions or not current:
            self._comparison_render_signature = None
            self.patient_comparison_body.pack_forget()
            return
        if not self.patient_comparison_body.winfo_manager():
            self.patient_comparison_body.pack(
                fill="x", pady=(0, 3), before=self.patient_prescriptions_body)
        previous = max(prescriptions, key=lambda item: str(item.get("saved_at", "")))
        result = compare_prescriptions(current, previous.get("drugs", []))
        signature = (id(self.patient_comparison_body), loaded_id, cfg.config.language, result)
        if signature == getattr(self, "_comparison_render_signature", None):
            return
        self._comparison_render_signature = signature
        for child in self.patient_comparison_body.winfo_children():
            child.destroy()
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
        self._prescription_render_generation = getattr(self, "_prescription_render_generation", 0) + 1
        generation = self._prescription_render_generation
        prescriptions = list(record.get("prescriptions", []))
        if not prescriptions:
            self._show_empty_prescriptions()
            return
        for child in self.patient_prescriptions_body.winfo_children():
            child.destroy()
        prescriptions.sort(key=lambda item: str(item.get("saved_at", "")), reverse=True)
        self._render_prescription_batch(record, prescriptions, generation)

    def _render_prescription_batch(self, record, prescriptions, generation, start=0):
        if self._closing or generation != self._prescription_render_generation:
            return
        if getattr(self, "active_page", "") != "patient":
            self._pending_history_render = (record, prescriptions, generation, start)
            return
        for prescription in prescriptions[start:start+1]:
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

        next_start = start + 1
        if next_start < len(prescriptions):
            self.after(1, lambda: self._render_prescription_batch(
                record, prescriptions, generation, next_start))

    def toggle_patient_prescription(self, record, prescription_id):
        if prescription_id in self._expanded_prescription_ids:
            self._expanded_prescription_ids.remove(prescription_id)
        else:
            self._expanded_prescription_ids.add(prescription_id)
        self.show_patient_prescriptions(record)

    def save_patient_history(self):
        if not self._warn_similar_patient():
            return False
        if (self._loaded_patient_id
                and self._patient_form_snapshot() != self._patient_saved_snapshot
                and not messagebox.askyesno(
                    I.t("patient_overwrite_title"),
                    I.t("patient_overwrite_message"), parent=self)):
            return False
        try:
            record = self._save_with_feedback("patient", lambda: self.patient_history.save_patient(
                {key: variable.get() for key, variable in self.patient_vars.items()},
                self._loaded_patient_id))
        except (ValueError, PatientHistoryError) as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return False
        self._loaded_patient_id = str(record.get("id", ""))
        self._current_history_record = record
        self._patient_saved_snapshot = self._patient_form_snapshot()
        self._set_patient_status("saved")
        self.patient_delete_button.configure(state="normal")
        self.patient_search_var.set("")
        self.refresh_patient_history()
        self._refresh_workflow_summary()
        return True

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

    def _set_save_feedback(self, scope, state):
        labels = [getattr(self, scope + suffix, None) for suffix in ("_save_status", "_saved_status")]
        if scope == "patient":
            labels.append(getattr(self, "patient_status_label", None))
        for label in labels:
            if label is None:
                continue
            label.configure(text=I.t("save_feedback_" + state) if state else "",
                text_color={"saving": MUTED, "saved": GOOD, "failed": DANGER}.get(state, MUTED))

    def _save_with_feedback(self, scope, operation):
        self._set_save_feedback(scope, "saving")
        try:
            result = operation()
        except Exception:
            self._set_save_feedback(scope, "failed")
            raise
        self._set_save_feedback(scope, "saved" if result else "failed")
        return result

    def save_prescription_for_patient(self):
        if getattr(self, "_applying_template", False):
            return
        if getattr(self, "_history_save_future", None) is not None:
            return
        patient = {key: variable.get() for key, variable in self.patient_vars.items()}
        drugs = [row.get_data() for row in self.rows
                 if row.get_data().generic_name or row.get_data().brand_name]
        if not patient.get("name", "").strip() or not drugs:
            messagebox.showinfo(I.t("save_patient_prescription"), I.t("patient_save_required"))
            return
        record_id = self._loaded_patient_id
        signature = self._draft_signature()
        snapshot = copy.deepcopy([drug.__dict__ for drug in drugs])

        def publish(worker, callback):
            getattr(self, "_set_save_feedback", lambda *_: None)("medication", "saving")
            button = getattr(self, "medication_save_button", None)
            if button is not None:
                button.configure(state="disabled", text=I.t("patient_saving"))
            future = self._persistence_executor.submit(worker)
            self._history_save_future = future
            def completed(done):
                try:
                    value, error = done.result(), None
                except Exception as exc:
                    value, error = None, exc
                def deliver():
                    self._history_save_future = None
                    if button is not None:
                        button.configure(state="normal", text=I.t("save_patient_prescription"))
                    if error is not None:
                        getattr(self, "_set_save_feedback", lambda *_: None)("medication", "failed")
                        messagebox.showerror(I.t("save_patient_prescription"), str(error), parent=self)
                    else:
                        callback(value)
                if not self._closing:
                    self._background_callbacks.put(deliver)
            future.add_done_callback(completed)

        def saved(record):
            getattr(self, "_set_save_feedback", lambda *_: None)("medication",
                "saved" if self._draft_signature() == signature else "")
            self.refresh_patient_history()
            # Never mark newer edits or a different patient as the saved snapshot.
            if self._draft_signature() == signature and self._loaded_patient_id == record_id:
                self._loaded_patient_id = str(record.get("id", ""))
                self._medication_patient_id = self._loaded_patient_id
                self._current_history_record = record
                self._patient_saved_snapshot = self._patient_form_snapshot()
                self._set_patient_status("saved")
                self.patient_search_var.set("")
                self.show_patient_prescriptions(record)
                self._workflow_saved_signature = signature
                self._clear_prescription_draft()
                self._refresh_workflow_summary()
            messagebox.showinfo(I.t("save_patient_prescription"), I.t("patient_prescription_saved"))

        def save_now():
            publish(lambda: self.patient_history.save_prescription(patient, snapshot, record_id), saved)

        def checked(matches):
            if matches:
                match = matches[0]
                if messagebox.askyesno(I.t("similar_patient_title"),
                        I.t("similar_patient_message", name=match.get("name", ""),
                            age=match.get("age", "") or I.t("none_short")), parent=self):
                    if self._draft_signature() == signature:
                        self._load_patient_record(match)
                    getattr(self, "_set_save_feedback", lambda *_: None)("medication", "")
                    return
            save_now()

        if record_id:
            save_now()
        else:
            publish(lambda: self.patient_history.find_similar(
                patient["name"], patient.get("age", ""), sex=patient.get("sex", "")), checked)

    # -- data ----------------------------------------------------------------
    def collect(self):
        doctor = qu.Doctor(**{k: v.get().strip() for k, v in self.doctor_vars.items()})
        patient = qu.Patient(**{k: v.get().strip() for k, v in self.patient_vars.items()})
        drugs = [r.get_data() for r in self.rows
                 if r.get_data().generic_name or r.get_data().brand_name]
        rx_id = "RX-" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        clinic = qu.Clinic(**{k: v for k, v in cfg.config.get_clinic().items()
                              if k in {"name", "address", "phone", "website", "logo_path",
                                       "latitude", "longitude", "include_location"}})
        return qu.Prescription(clinic=clinic, doctor=doctor, patient=patient, drugs=drugs,
                               date=datetime.datetime.now().strftime("%Y-%m-%d"), rx_id=rx_id)

    def on_any_change(self, *a):
        if getattr(self, "_row_batch_depth", 0):
            return
        if getattr(self, "_history_save_future", None) is None:
            getattr(self, "_set_save_feedback", lambda *_: None)("medication", "")
        # Text-entry widgets can fire several events for one visible edit.
        # Coalesce them so the Word preview is rebuilt once after typing pauses.
        if self._word_preview_job is not None:
            try:
                self.after_cancel(self._word_preview_job)
            except (tk.TclError, ValueError):
                pass
        self._word_preview_job = self.after(180, self._render_word_preview_now)
        self._workflow_saved_signature = None
        self._refresh_workflow_summary()
        if not self._suspend_draft:
            self._schedule_draft_save()
        if self._loaded_patient_id:
            self._medication_patient_id = self._loaded_patient_id
            if self._comparison_job is not None:
                try:
                    self.after_cancel(self._comparison_job)
                except (tk.TclError, ValueError):
                    pass
            self._comparison_job = None
            if getattr(self, "active_page", "") == "patient":
                self._comparison_job = self.after(250, self._render_prescription_comparison_now)

    def _schedule_draft_save(self):
        """Compatibility hook: unfinished prescriptions are never persisted."""
        return

    def _save_draft_now(self):
        self._clear_prescription_draft()

    def _clear_prescription_draft(self):
        if self._draft_save_job is not None:
            try:
                self.after_cancel(self._draft_save_job)
            except (tk.TclError, ValueError):
                pass
            self._draft_save_job = None
        if cfg.config.get("prescription_draft"):
            cfg.config.set("prescription_draft", {})

    def _offer_resume_draft(self):
        self._clear_prescription_draft()
        self.show_page("patient")
        self._set_workflow_step(1)

    def _resume_prescription_draft(self, draft):
        patient = draft.get("patient", {})
        self._patient_status_suspend = True
        try:
            for key in ("name", "age", "sex"):
                self.patient_vars[key].set(str(patient.get(key, "") or ""))
            sex_code = self.patient_vars["sex"].get()
            sex_label = next((label for label, code in self._patient_sex_codes.items()
                              if code == sex_code), I.t("sex_m"))
            self.patient_sex_menu.set(sex_label)
            self._loaded_patient_id = str(patient.get("id", "") or "")
        finally:
            self._patient_status_suspend = False
        for row in list(self.rows):
            row.destroy()
        self.rows = []
        for item in draft.get("medicines", []):
            self.add_row(qu.DrugItem(**{
                key: str(item.get(key, "") or "") for key in
                ("generic_name", "brand_name", "dosage", "frequency",
                 "duration", "notes", "quantity")
            }))
        if not self.rows:
            self.add_row()
        target = draft.get("active_page", "patient")
        self.show_page(target if target in {"patient", "medications"} else "patient")
        self._set_workflow_step(2 if target == "medications" else 1)
        self._workflow_saved_signature = None
        self._refresh_workflow_summary()

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
        self.review_prescription()

    def open_medication_subpage(self, kind):
        """Show secondary medication tools without expanding the entry form."""
        for row in self.rows:
            row._hide_ac()
        self.medication_main.pack_forget()
        self.medication_favorite_panel.pack_forget()
        self.word_preview_section.pack_forget()
        export_panel = getattr(self, "workflow_export_panel", None)
        if export_panel is not None:
            export_panel.pack_forget()
        self.word_preview_visible = kind in {"preview", "review", "export"}
        self.medication_subpage_title.configure(
            text=(I.t("workflow_export_title") if kind == "export" else
                  I.t("workflow_review") if self.word_preview_visible else
                  I.t("starred_drugs")))
        self.medication_subpage.pack(fill="both", expand=True)
        subpage_bar = getattr(self, "medication_subpage_bar", None)
        if subpage_bar is not None:
            if kind == "export":
                # The export card already names the action. Remove the repeated
                # Back/Export prescription row from workflow step four.
                subpage_bar.pack_forget()
            else:
                subpage_bar.pack(fill="x", pady=(4, 8))
        panel = self.word_preview_section if self.word_preview_visible else self.medication_favorite_panel
        panel.pack(fill="x", padx=2, pady=CARD_GAP)
        if self.word_preview_visible:
            self.word_preview_body.pack(fill="x", padx=PAD, pady=PAD)
            if kind == "export" and export_panel is not None:
                export_panel.pack(fill="x", padx=2, pady=(0, CARD_GAP))
        self.scroll._parent_canvas.yview_moveto(0)

    def close_medication_subpage(self):
        self.word_preview_visible = False
        self.medication_subpage.pack_forget()
        self.medication_favorite_panel.pack_forget()
        self.word_preview_section.pack_forget()
        export_panel = getattr(self, "workflow_export_panel", None)
        if export_panel is not None:
            export_panel.pack_forget()
        self.medication_main.pack(fill="both", expand=True)
        if hasattr(self, "_set_workflow_step"):
            self._set_workflow_step(2)
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
        self._ensure_page_built("reference")
        medicines = [row.get_data().generic_name.strip() for row in self.rows
                     if row.get_data().generic_name.strip()]
        if not medicines:
            messagebox.showinfo(I.t("online_drug_reference"), I.t("openfda_no_drugs"))
            return
        self._start_openfda_lookup(medicines)

    def lookup_openfda_search(self, _event=None):
        self._ensure_page_built("reference")
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
        self._reference_request_token = getattr(self, "_reference_request_token", 0) + 1
        token = self._reference_request_token
        self._reference_last_medicines = medicines
        self._set_reference_lookup_state(True)
        self.reference_status.configure(text=I.t("openfda_searching"))
        self._clear_reference_cards()
        threading.Thread(target=self._lookup_openfda_worker,
                         args=(medicines, force, token), daemon=True).start()

    def refresh_openfda_labels(self):
        self._ensure_page_built("reference")
        if self._reference_last_medicines:
            self._start_openfda_lookup(self._reference_last_medicines, force=True)
        else:
            self.lookup_openfda_search()

    def clear_openfda_cache(self):
        self._ensure_page_built("reference")
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

    def _lookup_openfda_worker(self, medicines, force=False, token=None):
        def publish(callback):
            def deliver():
                if (not self._closing and (token is None or token == getattr(self, "_reference_request_token", 0))):
                    callback()
            if not self._closing:
                self._background_callbacks.put(deliver)
        try:
            references, cached_names, checked = self._lookup_openfda_with_cache(medicines, force)
            publish(lambda: self._show_openfda_results(references, medicines, cached_names, checked))
        except openfda.OpenFDALookupError as exc:
            publish(lambda detail=str(exc): self._show_openfda_error(detail))
        except Exception as exc:
            logging.exception("openFDA lookup failed")
            publish(lambda detail=str(exc): self._show_openfda_error(detail))

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
        body = ctk.CTkFrame(line, fg_color="transparent")
        def set_expanded(expanded):
            line._expanded = expanded
            heading.configure(text=("▾  " if expanded else "▸  ") + title)
            if expanded:
                body.pack(fill="x", padx=10, pady=(0, 8))
            else:
                body.pack_forget()
        def toggle():
            expanded = not line._expanded
            previous = getattr(self, "_expanded_reference_panel", None)
            if previous is not None and previous is not line:
                try:
                    previous._set_expanded(False)
                except tk.TclError:
                    pass
            set_expanded(expanded)
            self._expanded_reference_panel = line if expanded else None
        heading = VisualButton(
            line, text="▸  " + title, text_color=foreground,
            fg_color="transparent", hover_color=background,
            font=_ui_font(13, "bold"), anchor="w", height=32, command=toggle)
        heading.pack(fill="x", padx=6, pady=4)
        line._expanded = False
        line._set_expanded = set_expanded
        if status:
            ctk.CTkLabel(body, text=status, text_color=foreground,
                         font=_ui_font(10, "bold"), anchor="w").pack(fill="x")
        value = "\n\n".join(paragraphs) or I.t("openfda_not_stated")
        content = ctk.CTkLabel(body, text=value, text_color=TEXT,
                              font=_ui_font(12), justify="left", wraplength=300,
                              anchor="nw")
        content.pack(fill="x")
        if source_sections:
            ctk.CTkLabel(body, text=I.t("reference_section_source", sections=", ".join(source_sections)),
                         text_color=foreground, font=_ui_font(10),
                         wraplength=300).pack(fill="x", pady=(4, 0))
        line.bind("<Configure>", lambda event: content.configure(
            wraplength=max(100, event.width - 24)))
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
        # Session-only capture. Nothing is persisted or offered on next launch.
        snapshot = {
            "patient": {key: variable.get() for key, variable in self.patient_vars.items()},
            "doctor": {key: variable.get() for key, variable in self.doctor_vars.items()},
            "medicines": [(row.get_data(), row.expanded) for row in self.rows],
            "patient_id": getattr(self, "_loaded_patient_id", ""),
            "saved_patient": self._patient_saved_snapshot,
            "saved_workflow": self._workflow_saved_signature,
            "page": getattr(self, "active_page", "patient"),
        }
        self._reference_request_token += 1
        self._suspend_draft = True
        for name, job in list(self.__dict__.items()):
            if name.endswith("_job") and isinstance(job, str) and job.startswith("after#"):
                self.after_cancel(job)
                setattr(self, name, None)
        for child in list(self.children.values()):
            child.destroy()
        self.rows = []
        self._build_ui()
        self._favorite_card_widgets = {}
        self._saved_template_card_pool = {}
        self._pending_favorite_render = None
        self._pending_template_render = None
        self._patient_status_suspend = True
        try:
            for key, value in snapshot["patient"].items():
                self.patient_vars[key].set(value)
            for key, value in snapshot["doctor"].items():
                self.doctor_vars[key].set(value)
            self.patient_sex_menu.set(next(
                (label for label, code in self._patient_sex_codes.items()
                 if code == snapshot["patient"].get("sex")), I.t("sex_m")))
            self._loaded_patient_id = snapshot["patient_id"]
            self._medication_patient_id = snapshot["patient_id"]
            self._patient_saved_snapshot = snapshot["saved_patient"]
            for row in list(self.rows):
                row.destroy()
            self.rows = []
            for medicine, expanded in snapshot["medicines"]:
                self.add_row(medicine)
            for row in self.rows:
                row.set_expanded(False)
            for row, (_, expanded) in zip(self.rows, snapshot["medicines"]):
                row.set_expanded(expanded)
            self._workflow_saved_signature = snapshot["saved_workflow"]
        finally:
            self._patient_status_suspend = False
            self._suspend_draft = False
        self._on_patient_form_change()
        self._workflow_saved_signature = snapshot["saved_workflow"]
        # Reusable/reference pages load lazily when revisited.
        self.show_page(snapshot["page"] if snapshot["page"] in {"patient", "medications", "prescriber"} else "patient")
        self._refresh_workflow_summary()

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
        if getattr(self, "_database_import_busy", False):
            return
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
            self._database_import_busy = False
            self._database_loading = False
            for button in self.page_buttons.values():
                button.configure(state="normal")
            total, report = result
            self._invalidate_database_caches()
            self.db_label.configure(text=I.t("db_count", n=len(self.db.drugs)))
            self.refresh_class_overview()
            self.refresh_favorites_page()
            self.show_import_classification_assistant(report)
            if on_done:
                on_done()

        def failed(exc):
            self._database_import_busy = False
            messagebox.showerror(I.t("import_db"), str(exc), parent=self)

        self._database_import_busy = True
        self.submit_background(
            worker, finished,
            on_error=failed,
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
    def _validate(self, rx, action):
        errors, warnings = qu.validate_prescription(rx)
        if errors:
            messagebox.showwarning(action, "Please correct the following:\n\n" + "\n".join(f"• {x}" for x in errors))
            return False
        if warnings and not messagebox.askyesno(action, "Warnings:\n\n" + "\n".join(
                f"• {x}" for x in warnings) + "\n\nContinue?"):
            return False
        return True

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
                        paper_size=document.get("_paper_size", "A4"),
                        font_size=int(document.get("font_size", 10)), **options)

    def _prepare_full_document(self, action):
        rx = copy.deepcopy(self.collect())
        if not self._validate(rx, action):
            return None
        document = copy.deepcopy(cfg.config.get("document_defaults", {}))
        document["_paper_size"] = self.paper_var.get()
        document["_layout_profile"] = print_layout.active_profile(document)
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
        if self._export_busy or self._closing or getattr(self, "_applying_template", False):
            return None
        prepared = self._prepare_full_document(action)
        if prepared is None:
            return None
        rx, document = prepared
        try:
            payload = copy.deepcopy(rx.to_cloud_payload())
        except clinic_location.LocationError:
            messagebox.showerror(I.t("location_title"), I.t("location_invalid"), parent=self)
            return None
        operation = {"rx": rx, "document": document, "payload": payload,
                     "api_key": cfg.config.cloud_rx_api_key.strip(), "action": action,
                     "path_pdf": path_pdf, "path_docx": path_docx,
                     "compact": compact, "on_success": on_success}
        self._set_export_busy(True)
        self._upload_export_link(operation)
        return operation

    def _upload_export_link(self, operation):
        def ready(link):
            try:
                cfg.config.set(
                    "cloud_last_upload_at",
                    datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"))
            except Exception:
                logging.exception("Could not save the last successful cloud-upload time")
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
                            paper_size=operation["document"]["_paper_size"],
                            font_size=int(operation["document"].get("font_size", 10)),
                            layout_profile=operation["document"].get("_layout_profile"))
            return self._generate_full_document(
                operation["rx"], qr, operation["document"],
                operation["path_pdf"], operation["path_docx"])

        def finished(result):
            self._set_export_busy(False)
            if getattr(self, "_workflow_active_step", 0) == 3:
                self._workflow_export_complete = True
                self._set_workflow_step(3)
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
            # The user may have corrected and saved the key after the failed
            # request. Refresh it instead of reusing the stale export snapshot.
            operation["api_key"] = cfg.config.cloud_rx_api_key.strip()
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
            initialfile=self._word_export_filename(),
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
            initialfile=self._word_export_filename(),
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

    def _word_export_filename(self):
        """Use patient + date when supplied while keeping a valid Windows filename."""
        patient_name = self.patient_vars["name"].get().strip()
        export_date = datetime.date.today().isoformat()
        base = f"{patient_name or 'Prescription'}_{export_date}"
        invalid = '<>:"/\\|?*'
        safe_name = "".join("_" if character in invalid else character
                            for character in base).strip().rstrip(".")
        return (safe_name[:100] or "Prescription") + ".docx"

    def open_settings(self):
        if getattr(self, "_mapping_save_busy", False):
            return
        SettingsWindow(self)

    def apply_ui_font_preferences(self):
        """Update existing controls without rebuilding pages or losing edits."""
        pending = [self]
        while pending:
            widget = pending.pop()
            if isinstance(widget, (VisualEntry, VisualComboBox, VisualOptionMenu, PopupListbox, VisualMenu)):
                widget.apply_preferences()
            pending.extend(widget.winfo_children())
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
        self.clinic_website_var = tk.StringVar(value=clinic.get("website", ""))
        self.logo_var = tk.StringVar(value=clinic.get("logo_path", ""))
        self.location_input_var = tk.StringVar()
        self.latitude_var = tk.StringVar(value=clinic.get("latitude", ""))
        self.longitude_var = tk.StringVar(value=clinic.get("longitude", ""))
        self.include_location_var = tk.BooleanVar(value=clinic.get("include_location") is True)
        self._confirmed_location = (self.latitude_var.get(), self.longitude_var.get())
        self._location_cancel_event = threading.Event()
        self._location_request_token = 0
        self._location_busy = False
        self._cloud_test_running = False
        self.paper_var = tk.StringVar(value=cfg.config.paper_size)
        self.language_var = tk.StringVar(
            value="العربية" if cfg.config.language == "ar" else "English")
        dropdown_size = cfg.config.ui_font_size("dropdown_font_size")
        self.dropdown_font_var = tk.StringVar(value=str(dropdown_size) if dropdown_size else I.t("settings_font_default"))
        self.instruction_font_var = tk.StringVar(value=str(cfg.config.ui_font_size("instruction_font_size")))
        self.patient_font_var = tk.StringVar(value=str(cfg.config.ui_font_size("patient_name_font_size")))
        document = cfg.config.get("document_defaults", {})
        language_labels = {"interface": I.t("settings_same_as_interface"),
                           "en": "English", "ar": "العربية"}
        self.document_language_var = tk.StringVar(
            value=language_labels.get(document.get("language", "interface"),
                                      I.t("settings_same_as_interface")))
        self.document_header_var = tk.BooleanVar(
            value=bool(document.get("show_header", True)))
        self.document_font_var = tk.StringVar(value=str(document.get("font_size", 10)))
        self.export_folder_var = tk.StringVar(value=document.get("export_folder", ""))
        self._printer_profiles = print_layout.normalize_profiles(
            document.get("printer_profiles"))
        self._calibration_presets = print_layout.normalize_presets(
            document.get("calibration_presets"))
        self.calibration_preset_var = tk.StringVar(value=I.t("settings_no_preset"))
        self._preset_revision_var = tk.StringVar(value="0")
        self._layout_preview_window = None
        self._layout_preview_label = None
        self._layout_preview_image = None
        self._layout_preview_after = None
        self._print_layout_dialog = None
        selected_printer = str(document.get(
            "selected_printer", print_layout.DEFAULT_PROFILE_NAME))
        if selected_printer not in self._printer_profiles:
            self._printer_profiles[selected_printer] = print_layout.default_profile()
        self.printer_profile_var = tk.StringVar(value=selected_printer)
        self._loaded_printer_name = selected_printer
        initial_layout = print_layout.normalize_profile(
            self._printer_profiles[selected_printer])
        self.layout_horizontal_var = tk.StringVar(
            value=self._format_layout_number(initial_layout["horizontal_offset_mm"]))
        self.layout_vertical_var = tk.StringVar(
            value=self._format_layout_number(initial_layout["vertical_offset_mm"]))
        self.layout_medication_top_var = tk.StringVar(
            value=self._format_layout_number(initial_layout["medication_top_mm"]))
        self.layout_line_spacing_var = tk.StringVar(
            value=self._format_layout_number(initial_layout["line_spacing"]))
        self.layout_field_gap_var = tk.StringVar(
            value=self._format_layout_number(initial_layout["field_gap_mm"]))
        self.layout_qr_position_var = tk.StringVar(value=initial_layout["qr_position"])
        self.layout_qr_size_var = tk.StringVar(
            value=self._format_layout_number(initial_layout["qr_size_mm"]))
        self.layout_qr_side_var = tk.StringVar(
            value=self._format_layout_number(initial_layout["qr_side_margin_mm"]))
        self.layout_qr_bottom_var = tk.StringVar(
            value=self._format_layout_number(initial_layout["qr_bottom_margin_mm"]))
        self.layout_background_var = tk.StringVar(
            value=initial_layout["preview_background_path"])
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
            self.clinic_phone_var, self.clinic_website_var, self.logo_var, self.paper_var,
            self.language_var,
            self.document_language_var, self.document_header_var,
            self.document_font_var, self.export_folder_var, self.auto_backup_var,
            self.dropdown_font_var, self.patient_font_var,
            self.instruction_font_var,
            self.location_input_var, self.latitude_var, self.longitude_var,
            self.include_location_var,
            self.printer_profile_var, self.layout_horizontal_var,
            self.layout_vertical_var, self.layout_medication_top_var,
            self.layout_line_spacing_var, self.layout_field_gap_var,
            self.layout_qr_position_var, self.layout_qr_size_var,
            self.layout_qr_side_var, self.layout_qr_bottom_var,
            self.layout_background_var,
            self._preset_revision_var,
        )
        self._saved_snapshot = self._snapshot()
        for variable in self._tracked_variables:
            variable.trace_add("write", self._mark_dirty)
        for variable in (
                self.layout_horizontal_var, self.layout_vertical_var,
                self.layout_medication_top_var, self.layout_line_spacing_var,
                self.layout_field_gap_var, self.layout_qr_position_var,
                self.layout_qr_size_var, self.layout_qr_side_var,
                self.layout_qr_bottom_var, self.layout_background_var):
            variable.trace_add("write", self._schedule_live_layout_preview)
        for variable in (self.clinic_name_var, self.clinic_address_var,
                         self.clinic_phone_var, self.clinic_website_var, self.logo_var,
                         self.document_header_var):
            variable.trace_add("write", lambda *_: self.update_clinic_preview())
        self.cloud_key_var.trace_add("write", lambda *_: self._refresh_cloud_key_status())
        self.update_clinic_preview()
        self._refresh_cloud_key_status(reset_tests=False)
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
        display.grid_columnconfigure((0, 1, 2), weight=1, uniform="font_sections")
        display.winfo_children()[-1].grid_configure(columnspan=3)
        sizes = [str(size) for size in range(18, 57)]
        for column, (label, variable, values) in enumerate((
                (I.t("settings_medication_font"), self.dropdown_font_var, sizes),
                (I.t("settings_instruction_font"), self.instruction_font_var, sizes),
                (I.t("settings_name_font"), self.patient_font_var, sizes))):
            ctk.CTkLabel(display, text=label, text_color=MUTED, anchor="w", font=_ui_font(12)).grid(
                row=1, column=column, sticky="ew", padx=14, pady=(2, 4))
            VisualOptionMenu(display, variable=variable, values=values, height=FIELD_HEIGHT,
                             font=_ui_font(13), fg_color=SURFACE, text_color=TEXT,
                             button_color=ACCENT_SOFT, button_hover_color=LINE).grid(
                row=2, column=column, sticky="ew", padx=14, pady=(0, 10))
            preview = ctk.CTkLabel(display, text=("Drug", "1×2", "أحمد Ali")[column],
                                   text_color=TEXT, anchor="w", font=_ui_font(int(variable.get())))
            preview.grid(row=3, column=column, sticky="ew", padx=14, pady=(0, 10))
            variable.trace_add("write", lambda *_args, var=variable, label=preview:
                               label.configure(font=_ui_font(int(var.get()))))

    def _build_clinic_page(self):
        page = self._new_page(
            "clinic", I.t("settings_clinic"), I.t("settings_clinic_tip"))
        page.grid_rowconfigure(1, weight=1)
        body = ctk.CTkScrollableFrame(page, fg_color="transparent", corner_radius=0)
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        card = self._card(body, 0, I.t("settings_clinic_contact"))
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
        ctk.CTkLabel(
            card, text=I.t("settings_clinic_website"), text_color=MUTED,
            anchor="w", font=ctk.CTkFont(size=12, weight="bold")).grid(
                row=3, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 3))
        self.clinic_website_entry = VisualEntry(
            card, textvariable=self.clinic_website_var, height=36, corner_radius=8,
            border_color=LINE, font=ctk.CTkFont(size=13),
            placeholder_text=I.t("settings_clinic_website_hint"))
        self.clinic_website_entry.grid(
            row=4, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 10))
        ctk.CTkLabel(
            card, text=I.t("settings_clinic_website_privacy"), text_color=MUTED,
            anchor="w", justify="left", wraplength=720,
            font=ctk.CTkFont(size=11)).grid(
                row=5, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 8))
        self._build_location_card(body)
        logo_card = self._card(body, 2, I.t("settings_logo"))
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
        preview = self._card(body, 3, I.t("settings_header_preview"))
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
        VisualButton(
            preview, text=I.t("settings_preview_cloud"), height=34,
            fg_color=ACCENT_SOFT, hover_color=LINE, text_color=ACCENT,
            command=self._open_cloud_viewer_preview).grid(
                row=2, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 10))

    def _open_cloud_viewer_preview(self):
        """Render a private, fictitious mobile viewer preview without network or files."""
        title = self.clinic_name_var.get().strip() or "Electronic Prescription"
        phone = self.clinic_phone_var.get().strip()
        dial_digits = re.sub(r"[^0-9+]", "", phone)
        phone_valid = bool(re.fullmatch(r"\+?[0-9]{7,15}", dial_digits))
        try:
            website_valid = bool(qu.normalize_clinic_website(self.clinic_website_var.get()))
        except ValueError:
            website_valid = False
        location_valid = False
        if self.include_location_var.get():
            try:
                clinic_location.validate_coordinates(
                    self.latitude_var.get(), self.longitude_var.get())
                location_valid = True
            except clinic_location.LocationError:
                pass

        dialog = ctk.CTkToplevel(self)
        dialog.title(I.t("cloud_preview_title"))
        dialog.geometry("420x700")
        dialog.minsize(380, 620)
        dialog.configure(fg_color="#94a3b8")
        dialog.transient(self)

        canvas = ctk.CTkScrollableFrame(
            dialog, width=390, fg_color="#f1f5f9", corner_radius=0)
        canvas.pack(fill="both", expand=True, padx=14, pady=14)
        canvas.grid_columnconfigure(0, weight=1)
        header = ctk.CTkFrame(canvas, fg_color="#1d4ed8", corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            header, text="℞", text_color="#93c5fd", width=30,
            font=ctk.CTkFont(size=23, weight="bold")).grid(
                row=0, column=0, padx=(12, 5), pady=12)
        ctk.CTkLabel(
            header, text=title, text_color="white", anchor="w",
            font=ctk.CTkFont(size=15, weight="bold")).grid(
                row=0, column=1, sticky="ew", pady=12)
        actions = ctk.CTkFrame(header, fg_color="transparent")
        actions.grid(row=0, column=2, padx=(5, 10), pady=8)
        for visible, glyph in ((phone_valid, "☎"), (location_valid, "⌖"),
                               (website_valid, "◎")):
            if visible:
                ctk.CTkLabel(
                    actions, text=glyph, width=32, height=32, corner_radius=16,
                    fg_color="#4775df", text_color="white",
                    font=ctk.CTkFont(size=15, weight="bold")).pack(
                        side="left", padx=3)

        def preview_card(row, heading, lines, *, medication=False):
            card = ctk.CTkFrame(
                canvas, fg_color="white", corner_radius=12,
                border_width=1, border_color="#cbd5e1")
            card.grid(row=row, column=0, sticky="ew", padx=12, pady=(12, 0))
            card.grid_columnconfigure(0, weight=1)
            if heading:
                ctk.CTkLabel(
                    card, text=heading, fg_color="#e2e8f0", corner_radius=8,
                    text_color="#0f172a", anchor="w", padx=12, height=34,
                    font=ctk.CTkFont(size=13, weight="bold")).grid(
                        row=0, column=0, sticky="ew", padx=1, pady=1)
            start_row = 1 if heading else 0
            for offset, line in enumerate(lines):
                index = start_row + offset
                ctk.CTkLabel(
                    card, text=line, text_color="#0f172a", anchor="w",
                    justify="left", wraplength=330,
                    font=ctk.CTkFont(size=13, weight="bold" if medication and offset == 0 else "normal")).grid(
                        row=index, column=0, sticky="ew", padx=14,
                        pady=(10 if offset == 0 else 2,
                              10 if offset == len(lines) - 1 else 2))

        preview_card(1, I.t("cloud_preview_prescriber"),
                     [I.t("cloud_preview_doctor")])
        preview_card(2, I.t("cloud_preview_patient"),
                     [I.t("cloud_preview_patient_name")])
        preview_card(3, "", [I.t("cloud_preview_medication"),
                              I.t("cloud_preview_instruction")], medication=True)
        dialog.after(100, dialog.grab_set)

    def _build_location_card(self, body):
        card = self._card(body, 1, I.t("location_title"))
        row = ctk.CTkFrame(card, fg_color="transparent")
        row.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 5))
        row.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(row, text=I.t("location_link_hint"), font=_ui_font(11),
                     text_color=MUTED, anchor="w").grid(row=0, column=0, columnspan=2, sticky="ew")
        VisualEntry(row, textvariable=self.location_input_var, height=32,
                    placeholder_text=I.t("location_link_hint"), font=_ui_font(12),
                    border_color=LINE).grid(row=1, column=0, sticky="ew", padx=(0, 6))
        self._location_buttons = []
        def button(parent, key, command):
            widget = VisualButton(parent, text=I.t(key), command=command, height=30,
                width=_ui_font(12).measure(I.t(key)) + 18, font=_ui_font(12),
                fg_color=SURFACE, text_color=ACCENT, hover_color=ACCENT_SOFT,
                border_width=1, border_color=LINE)
            self._location_buttons.append(widget)
            return widget
        button(row, "location_check", self.check_clinic_location).grid(row=1, column=1)
        coords = ctk.CTkFrame(card, fg_color="transparent")
        coords.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 5))
        coords.grid_columnconfigure((0, 1), weight=1)
        for col, (variable, label) in enumerate(((self.latitude_var, "location_latitude"),
                                               (self.longitude_var, "location_longitude"))):
            ctk.CTkLabel(coords, text=I.t(label), font=_ui_font(11), anchor="w",
                         text_color=MUTED).grid(row=0, column=col, sticky="ew", padx=(0, 6))
            VisualEntry(coords, textvariable=variable, height=30, font=_ui_font(12),
                        border_color=LINE).grid(row=1, column=col, sticky="ew", padx=(0, 6))
        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="w", padx=14, pady=(0, 5))
        for col, (key, command) in enumerate((("location_open", self.open_clinic_maps),
                ("location_current", self.use_current_clinic_location),
                ("location_remove", self.remove_clinic_location))):
            button(actions, key, command).grid(row=0, column=col, padx=(0, 6))
        ctk.CTkCheckBox(card, text=I.t("location_include"), variable=self.include_location_var,
            font=_ui_font(12), checkbox_width=18, checkbox_height=18).grid(
                row=4, column=0, sticky="w", padx=14, pady=(0, 4))
        ctk.CTkLabel(card, text=I.t("location_privacy"), font=_ui_font(11),
            text_color=MUTED, anchor="w", justify="left", wraplength=560).grid(
                row=5, column=0, sticky="ew", padx=14, pady=(0, 6))
        self.location_status = ctk.CTkLabel(card, text="", font=_ui_font(11),
                                          anchor="w", text_color=ACCENT)
        self.location_status.grid(row=6, column=0, sticky="ew", padx=14)
        self.location_status.grid_remove()

    def _location_staging(self):
        return (self.location_input_var.get(), self.latitude_var.get(), self.longitude_var.get())

    def _set_location_busy(self, busy):
        self._location_busy = busy
        for button in self._location_buttons:
            button.configure(state="disabled" if busy else "normal")
        if busy:
            self.location_status.configure(text=I.t("location_waiting"), text_color=ACCENT)
            self.location_status.grid()
        else:
            self.location_status.grid_remove()

    def _request_location(self, worker):
        if self._location_busy:
            return
        self._location_cancel_event = threading.Event()
        self._location_request_token += 1
        token, snapshot = self._location_request_token, self._location_staging()
        event = self._location_cancel_event
        self._set_location_busy(True)
        def current():
            return self.winfo_exists() and token == self._location_request_token
        def success(pin):
            if not current():
                return
            self._set_location_busy(False)
            if snapshot != self._location_staging():
                return  # Never overwrite fields edited while the request was pending.
            accuracy = "" if pin.accuracy is None else I.t("location_accuracy", metres=round(pin.accuracy))
            if messagebox.askyesno(I.t("location_title"), I.t("location_confirm",
                    latitude=f"{pin.latitude:.7f}", longitude=f"{pin.longitude:.7f}",
                    accuracy=accuracy), parent=self):
                self.latitude_var.set(f"{pin.latitude:.7f}")
                self.longitude_var.set(f"{pin.longitude:.7f}")
                self._confirmed_location = (self.latitude_var.get(), self.longitude_var.get())
                self.location_input_var.set("")
        def failure(exc):
            if not current():
                return
            self._set_location_busy(False)
            code = exc.code if isinstance(exc, clinic_location.LocationError) else "unavailable"
            if code != "cancelled":
                key = "location_" + code
                messagebox.showerror(I.t("location_title"), I.t(key), parent=self)
        self.master.submit_background(lambda: worker(event), success, on_error=failure, silent=True)

    def check_clinic_location(self):
        value = self.location_input_var.get().strip()
        if value:
            self._request_location(lambda _event: clinic_location.resolve_location(value))
            return
        try:
            pin = clinic_location.validate_coordinates(self.latitude_var.get(), self.longitude_var.get())
            webbrowser.open(clinic_location.maps_link(pin))
        except clinic_location.LocationError:
            messagebox.showerror(I.t("location_title"), I.t("location_invalid"), parent=self)

    def open_clinic_maps(self):
        webbrowser.open("https://www.google.com/maps")

    def use_current_clinic_location(self):
        if not messagebox.askyesno(I.t("location_title"), I.t("location_detect_confirm"), parent=self):
            return
        labels = {key: I.t("location_browser_" + key)
                  for key in ("title", "detail", "button", "waiting", "done", "failed")}
        labels["language"] = I.get_lang()
        self._request_location(lambda event: clinic_location.detect_current_location(labels, cancel_event=event))

    def remove_clinic_location(self):
        self._location_cancel_event.set()
        self._location_request_token += 1
        self._set_location_busy(False)
        self.latitude_var.set("")
        self.longitude_var.set("")
        self.location_input_var.set("")
        self.include_location_var.set(False)

    def destroy(self):
        if hasattr(self, "_location_cancel_event"):
            self._location_cancel_event.set()
            self._location_request_token += 1
        self._close_layout_preview()
        self._close_print_layout_dialog()
        super().destroy()

    @staticmethod
    def _format_layout_number(value):
        number = float(value)
        return str(int(number)) if number.is_integer() else f"{number:g}"

    def _current_layout_profile(self):
        return print_layout.normalize_profile({
            "horizontal_offset_mm": self.layout_horizontal_var.get(),
            "vertical_offset_mm": self.layout_vertical_var.get(),
            "medication_top_mm": self.layout_medication_top_var.get(),
            "line_spacing": self.layout_line_spacing_var.get(),
            "field_gap_mm": self.layout_field_gap_var.get(),
            "qr_position": self.layout_qr_position_var.get(),
            "qr_size_mm": self.layout_qr_size_var.get(),
            "qr_side_margin_mm": self.layout_qr_side_var.get(),
            "qr_bottom_margin_mm": self.layout_qr_bottom_var.get(),
            "preview_background_path": self.layout_background_var.get().strip(),
        })

    def _store_current_printer_profile(self):
        name = str(self._loaded_printer_name or print_layout.DEFAULT_PROFILE_NAME)
        self._printer_profiles[name] = self._current_layout_profile()

    def _load_printer_profile(self, name):
        self._store_current_printer_profile()
        profile = print_layout.normalize_profile(
            self._printer_profiles.get(name, print_layout.default_profile()))
        self._printer_profiles[name] = profile
        self._loaded_printer_name = name
        self.printer_profile_var.set(name)
        self._apply_layout_profile_variables(profile)

    def _apply_layout_profile_variables(self, profile):
        variables = (
            (self.layout_horizontal_var, profile["horizontal_offset_mm"]),
            (self.layout_vertical_var, profile["vertical_offset_mm"]),
            (self.layout_medication_top_var, profile["medication_top_mm"]),
            (self.layout_line_spacing_var, profile["line_spacing"]),
            (self.layout_field_gap_var, profile["field_gap_mm"]),
            (self.layout_qr_size_var, profile["qr_size_mm"]),
            (self.layout_qr_side_var, profile["qr_side_margin_mm"]),
            (self.layout_qr_bottom_var, profile["qr_bottom_margin_mm"]),
        )
        for variable, value in variables:
            variable.set(self._format_layout_number(value))
        self.layout_qr_position_var.set(profile["qr_position"])
        self.layout_background_var.set(profile["preview_background_path"])

    def _preset_names(self):
        return [I.t("settings_no_preset")] + sorted(
            self._calibration_presets, key=str.casefold)

    def _refresh_preset_menu(self):
        menu = getattr(self, "calibration_preset_menu", None)
        if menu is not None and menu.winfo_exists():
            menu.configure(values=self._preset_names())

    def _touch_presets(self):
        self._preset_revision_var.set(str(int(self._preset_revision_var.get()) + 1))
        self._refresh_preset_menu()

    def _save_calibration_preset(self):
        parent = self._layout_dialog_parent()
        name = simpledialog.askstring(
            I.t("settings_save_preset"), I.t("settings_preset_name"), parent=parent)
        name = str(name or "").strip()
        if not name:
            return
        if name in self._calibration_presets and not messagebox.askyesno(
                I.t("settings_save_preset"),
                I.t("settings_preset_exists", name=name), parent=parent):
            return
        self._calibration_presets[name] = self._current_layout_profile().copy()
        self.calibration_preset_var.set(name)
        self._touch_presets()
        messagebox.showinfo(
            I.t("settings_save_preset"),
            I.t("settings_preset_saved", name=name), parent=parent)

    def _apply_calibration_preset(self, name):
        if name == I.t("settings_no_preset"):
            return
        profile = self._calibration_presets.get(name)
        if profile is None:
            return
        self._apply_layout_profile_variables(print_layout.normalize_profile(profile))
        self._store_current_printer_profile()

    def _delete_calibration_preset(self):
        name = self.calibration_preset_var.get()
        if name not in self._calibration_presets:
            return
        if not messagebox.askyesno(
                I.t("settings_delete_preset"),
                I.t("settings_delete_preset_confirm", name=name),
                parent=self._layout_dialog_parent()):
            return
        del self._calibration_presets[name]
        self.calibration_preset_var.set(I.t("settings_no_preset"))
        self._touch_presets()

    @staticmethod
    def _layout_option(parent, label, variable, values, row, column):
        ctk.CTkLabel(
            parent, text=label, text_color=MUTED, anchor="w",
            font=ctk.CTkFont(size=12, weight="bold")).grid(
                row=row, column=column, sticky="ew", padx=8, pady=(4, 3))
        VisualOptionMenu(
            parent, values=values, variable=variable, height=36,
            corner_radius=8, fg_color=SURFACE, button_color=PRIMARY,
            button_hover_color=ACCENT_HOVER, text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(size=13)).grid(
                row=row + 1, column=column, sticky="ew", padx=8, pady=(0, 5))

    def _open_print_layout_settings(self):
        if self._print_layout_dialog is not None:
            try:
                if self._print_layout_dialog.winfo_exists():
                    self._print_layout_dialog.lift()
                    self._print_layout_dialog.focus_force()
                    return
            except tk.TclError:
                pass
        dialog = ctk.CTkToplevel(self)
        self._print_layout_dialog = dialog
        dialog.title(I.t("settings_print_layout"))
        dialog.geometry("780x650")
        dialog.minsize(720, 570)
        dialog.transient(self)
        dialog.configure(fg_color=BG)
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(0, weight=1)
        body = ctk.CTkScrollableFrame(
            dialog, fg_color=SURFACE, corner_radius=10,
            scrollbar_button_color=LINE, scrollbar_button_hover_color=MUTED)
        body.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 6))
        body.grid_columnconfigure((0, 1, 2), weight=1)
        ctk.CTkLabel(
            body, text=I.t("settings_print_layout"), text_color=TEXT,
            anchor="w", font=ctk.CTkFont(size=22, weight="bold")).grid(
                row=0, column=0, columnspan=3, sticky="ew", padx=8, pady=(4, 10))

        printers = [print_layout.DEFAULT_PROFILE_NAME]
        printers.extend(print_layout.installed_printers())
        printers.extend(self._printer_profiles)
        printers = list(dict.fromkeys(printers))
        ctk.CTkLabel(
            body, text=I.t("settings_printer_profile"), text_color=MUTED,
            anchor="w", font=ctk.CTkFont(size=12, weight="bold")).grid(
                row=1, column=0, columnspan=3, sticky="ew", padx=8, pady=(0, 3))
        VisualOptionMenu(
            body, values=printers, variable=self.printer_profile_var,
            command=self._load_printer_profile, height=38, corner_radius=8,
            fg_color=SURFACE, button_color=PRIMARY,
            button_hover_color=ACCENT_HOVER, text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(size=13)).grid(
                row=2, column=0, columnspan=3, sticky="ew", padx=8, pady=(0, 8))

        ctk.CTkLabel(
            body, text=I.t("settings_calibration_preset"), text_color=MUTED,
            anchor="w", font=ctk.CTkFont(size=12, weight="bold")).grid(
                row=3, column=0, columnspan=3, sticky="ew", padx=8, pady=(0, 3))
        preset_row = ctk.CTkFrame(body, fg_color="transparent")
        preset_row.grid(row=4, column=0, columnspan=3, sticky="ew", padx=8, pady=(0, 8))
        preset_row.grid_columnconfigure(0, weight=1)
        self.calibration_preset_menu = VisualOptionMenu(
            preset_row, values=self._preset_names(),
            variable=self.calibration_preset_var,
            command=self._apply_calibration_preset, height=36, corner_radius=8,
            fg_color=SURFACE, button_color=PRIMARY,
            button_hover_color=ACCENT_HOVER, text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(size=13))
        self.calibration_preset_menu.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        VisualButton(
            preset_row, text=I.t("settings_save_preset"), height=36,
            fg_color=ACCENT_SOFT, text_color=ACCENT, border_width=1,
            border_color=LINE, command=self._save_calibration_preset).grid(
                row=0, column=1, padx=(0, 6))
        VisualButton(
            preset_row, text=I.t("settings_delete_preset"), height=36,
            fg_color=SURFACE, text_color=DANGER, border_width=1,
            border_color=LINE, command=self._delete_calibration_preset).grid(
                row=0, column=2)

        offset_values = [str(value) for value in range(-25, 26)]
        self._layout_option(body, I.t("settings_horizontal_offset"),
                            self.layout_horizontal_var, offset_values, 5, 0)
        self._layout_option(body, I.t("settings_vertical_offset"),
                            self.layout_vertical_var, offset_values, 5, 1)
        self._layout_option(body, I.t("settings_medication_top"),
                            self.layout_medication_top_var,
                            [str(value) for value in range(25, 161)], 5, 2)
        self._layout_option(body, I.t("settings_line_spacing"),
                            self.layout_line_spacing_var,
                            [f"{value / 10:.1f}" for value in range(8, 21)], 7, 0)
        self._layout_option(body, I.t("settings_field_gap"),
                            self.layout_field_gap_var,
                            [str(value) for value in range(2, 31)], 7, 1)
        self._layout_option(body, I.t("settings_qr_position"),
                            self.layout_qr_position_var,
                            ["Left", "Center", "Right"], 7, 2)
        self._layout_option(body, I.t("settings_qr_size"), self.layout_qr_size_var,
                            [str(value) for value in range(20, 56)], 9, 0)
        self._layout_option(body, I.t("settings_qr_side_margin"), self.layout_qr_side_var,
                            [str(value) for value in range(3, 41)], 9, 1)
        self._layout_option(body, I.t("settings_qr_bottom_margin"), self.layout_qr_bottom_var,
                            [str(value) for value in range(5, 61)], 9, 2)

        ctk.CTkLabel(
            body, text=I.t("settings_preprinted_image"), text_color=MUTED,
            anchor="w", font=ctk.CTkFont(size=12, weight="bold")).grid(
                row=11, column=0, columnspan=3, sticky="ew", padx=8, pady=(7, 3))
        background_row = ctk.CTkFrame(body, fg_color="transparent")
        background_row.grid(row=12, column=0, columnspan=3, sticky="ew", padx=8, pady=(0, 4))
        background_row.grid_columnconfigure(0, weight=1)
        VisualEntry(
            background_row, textvariable=self.layout_background_var,
            height=36, border_color=LINE, font=ctk.CTkFont(size=12)).grid(
                row=0, column=0, sticky="ew", padx=(0, 6))
        VisualButton(
            background_row, text=I.t("settings_browse"), height=36,
            width=90, fg_color=ACCENT_SOFT, text_color=ACCENT,
            command=self._choose_preprint_background).grid(row=0, column=1)
        ctk.CTkLabel(
            body, text=I.t("settings_preview_overlay_note"), text_color=MUTED,
            anchor="w", font=ctk.CTkFont(size=11)).grid(
                row=13, column=0, columnspan=3, sticky="ew", padx=8, pady=(0, 6))

        actions = ctk.CTkFrame(body, fg_color="transparent")
        actions.grid(row=14, column=0, columnspan=3, sticky="ew", padx=8, pady=(4, 12))
        for label, command in (
                (I.t("settings_preview_layout"), self._preview_print_layout),
                (I.t("settings_test_print"), self._create_calibration_test),
                (I.t("settings_reset_profile"), self._reset_current_print_profile)):
            VisualButton(
                actions, text=label, height=36, fg_color=ACCENT_SOFT,
                text_color=ACCENT, border_width=1, border_color=LINE,
                command=command).pack(side="left", padx=(0, 6))
        VisualButton(
            actions, text=I.t("close"), height=36, fg_color=PRIMARY,
            text_color="white", command=self._close_print_layout_dialog).pack(side="right")
        dialog.protocol("WM_DELETE_WINDOW", self._close_print_layout_dialog)
        dialog.after(100, dialog.grab_set)

    def _layout_dialog_parent(self):
        dialog = getattr(self, "_print_layout_dialog", None)
        try:
            if dialog is not None and dialog.winfo_exists():
                return dialog
        except tk.TclError:
            pass
        return self

    def _close_print_layout_dialog(self):
        dialog = getattr(self, "_print_layout_dialog", None)
        self._print_layout_dialog = None
        if dialog is not None:
            try:
                if dialog.winfo_exists():
                    dialog.grab_release()
                    dialog.destroy()
            except tk.TclError:
                pass

    def _choose_preprint_background(self):
        path = filedialog.askopenfilename(
            parent=self._layout_dialog_parent(), title=I.t("settings_preprinted_image"),
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.webp")])
        if path:
            self.layout_background_var.set(path)
            self._preview_print_layout()

    def _render_layout_preview_image(self, profile, width=420, height=596):
        background = profile.get("preview_background_path", "")
        try:
            if background and Path(background).is_file():
                image = ImageOps.fit(Image.open(background).convert("RGB"), (width, height))
            else:
                image = Image.new("RGB", (width, height), "white")
        except OSError:
            image = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(image, "RGBA")
        x_scale, y_scale = width / 148.0, height / 210.0
        x_offset = float(profile["horizontal_offset_mm"])
        y_offset = float(profile["vertical_offset_mm"])
        med_y = int((float(profile["medication_top_mm"]) + y_offset) * y_scale)
        text_x = int(max(4, (9 + x_offset) * x_scale))
        draw.rectangle((text_x - 4, med_y - 5, width - 18, med_y + 86),
                       fill=(255, 255, 255, 205), outline=(9, 100, 220, 210), width=2)
        for index, label in enumerate((
                "1. Medicine name    dosage    frequency",
                "2. Medicine name    dosage    frequency",
                "3. Medicine name    dosage    frequency")):
            draw.text((text_x, med_y + index * 24), label, fill=(25, 35, 50, 255))
        qr_size = float(profile["qr_size_mm"])
        side = float(profile["qr_side_margin_mm"])
        if profile["qr_position"] == "Center":
            qr_left = (148 - qr_size) / 2
        elif profile["qr_position"] == "Right":
            qr_left = 148 - side - qr_size
        else:
            qr_left = side
        qr_left = max(0, min(148 - qr_size, qr_left + x_offset))
        qr_top = max(0, min(210 - qr_size,
            210 - float(profile["qr_bottom_margin_mm"]) - qr_size + y_offset))
        qr_box = (int(qr_left * x_scale), int(qr_top * y_scale),
                  int((qr_left + qr_size) * x_scale),
                  int((qr_top + qr_size) * y_scale))
        draw.rectangle(qr_box, fill=(255, 255, 255, 235), outline=(9, 100, 220, 255), width=3)
        draw.line((qr_box[0], qr_box[1], qr_box[2], qr_box[3]), fill=(9, 100, 220, 255), width=2)
        draw.line((qr_box[2], qr_box[1], qr_box[0], qr_box[3]), fill=(9, 100, 220, 255), width=2)
        return image

    def _preview_print_layout(self):
        if self._layout_preview_window is not None:
            try:
                if self._layout_preview_window.winfo_exists():
                    self._layout_preview_window.lift()
                    self._layout_preview_window.focus_force()
                    self._refresh_live_layout_preview()
                    return
            except tk.TclError:
                pass
        dialog = ctk.CTkToplevel(self)
        dialog.title(I.t("settings_live_preview"))
        dialog.geometry("420x590")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.configure(fg_color=BG)
        self._layout_preview_window = dialog
        self._layout_preview_label = ctk.CTkLabel(
            dialog, text="", fg_color="white", corner_radius=8)
        self._layout_preview_label.pack(padx=20, pady=(20, 8))
        ctk.CTkLabel(
            dialog, text=I.t("settings_live_preview_hint"),
            text_color=MUTED, font=ctk.CTkFont(size=12)).pack(
                padx=20, pady=(0, 12))
        dialog.protocol("WM_DELETE_WINDOW", self._close_layout_preview)
        self._refresh_live_layout_preview()

    def _refresh_live_layout_preview(self):
        window = self._layout_preview_window
        label = self._layout_preview_label
        if window is None or label is None:
            return
        try:
            if not window.winfo_exists():
                self._close_layout_preview()
                return
            image = self._render_layout_preview_image(self._current_layout_profile())
            display = ctk.CTkImage(
                light_image=image, dark_image=image, size=(360, 511))
            self._layout_preview_image = display
            label.configure(image=display)
            label.image = display
            self._layout_preview_after = None
        except tk.TclError:
            self._close_layout_preview()

    def _schedule_live_layout_preview(self, *_args):
        window = getattr(self, "_layout_preview_window", None)
        if window is None:
            return
        try:
            if not window.winfo_exists():
                self._close_layout_preview()
                return
            if self._layout_preview_after is not None:
                self.after_cancel(self._layout_preview_after)
            self._layout_preview_after = self.after(
                80, self._refresh_live_layout_preview)
        except tk.TclError:
            self._close_layout_preview()

    def _close_layout_preview(self):
        after_id = getattr(self, "_layout_preview_after", None)
        if after_id is not None:
            try:
                self.after_cancel(after_id)
            except tk.TclError:
                pass
        self._layout_preview_after = None
        window = getattr(self, "_layout_preview_window", None)
        self._layout_preview_window = None
        self._layout_preview_label = None
        self._layout_preview_image = None
        if window is not None:
            try:
                if window.winfo_exists():
                    window.destroy()
            except tk.TclError:
                pass

    def _create_calibration_test(self):
        initial_dir = self.export_folder_var.get().strip()
        path = filedialog.asksaveasfilename(
            parent=self._layout_dialog_parent(), title=I.t("settings_test_print"),
            defaultextension=".docx", initialfile="A5-printer-calibration.docx",
            initialdir=initial_dir if Path(initial_dir).is_dir() else None,
            filetypes=[("Word documents", "*.docx")])
        if not path:
            return
        try:
            pdfgen.generate_calibration_docx(
                path, paper_size=self.paper_var.get(),
                layout_profile=self._current_layout_profile())
            os.startfile(path)
        except Exception as exc:
            logging.exception("Could not create printer calibration document")
            messagebox.showerror(
                I.t("settings_test_print"), str(exc),
                parent=self._layout_dialog_parent())

    def _reset_current_print_profile(self):
        profile = print_layout.default_profile()
        name = self.printer_profile_var.get() or print_layout.DEFAULT_PROFILE_NAME
        self._printer_profiles[name] = profile
        self._loaded_printer_name = name
        self._apply_layout_profile_variables(profile)

    def _build_documents_page(self):
        page = self._new_page(
            "documents", I.t("settings_documents"), I.t("settings_documents_tip"))
        card = self._card(page, 2, I.t("settings_page_format"))
        card.grid_columnconfigure((0, 1, 2), weight=1)
        fields = ((I.t("paper"), 0), (I.t("settings_document_language"), 1),
                  (I.t("settings_word_font_size"), 2))
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
        VisualOptionMenu(
            card, values=[str(size) for size in range(8, 25)],
            variable=self.document_font_var, height=44, corner_radius=9,
            fg_color=ACCENT_SOFT, button_color=PRIMARY,
            button_hover_color=ACCENT_HOVER, text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(size=15)).grid(
                row=2, column=2, sticky="ew", padx=20, pady=(0, 10))
        ctk.CTkCheckBox(
            card, text=I.t("settings_show_header"),
            variable=self.document_header_var, fg_color=PRIMARY,
            hover_color=ACCENT_HOVER).grid(
                row=3, column=0, columnspan=3, sticky="w",
                padx=20, pady=(4, 14))
        layout = self._card(page, 3, I.t("settings_print_layout"))
        layout.grid_columnconfigure(0, weight=1)
        self.print_layout_summary = ctk.CTkLabel(
            layout, text=I.t("settings_print_layout_tip"), text_color=MUTED,
            anchor="w", justify="left", font=ctk.CTkFont(size=12))
        self.print_layout_summary.grid(
            row=1, column=0, sticky="ew", padx=14, pady=(0, 6))
        VisualButton(
            layout, text=I.t("settings_configure_layout"), height=34,
            width=150, fg_color=ACCENT_SOFT, text_color=ACCENT,
            border_width=1, border_color=LINE,
            command=self._open_print_layout_settings).grid(
                row=2, column=0, sticky="w", padx=14, pady=(0, 10))
        folder = self._card(page, 4, I.t("settings_export_folder"))
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

        readiness = self._card(page, 3, I.t("cloud_readiness_title"))
        readiness.grid_columnconfigure(0, weight=1)
        self.cloud_readiness_labels = {}
        rows = (
            ("key", I.t("cloud_status_api_key")),
            ("api", I.t("cloud_status_api")),
            ("viewer", I.t("cloud_status_viewer")),
            ("upload", I.t("cloud_status_last_upload")),
        )
        for row, (key, label) in enumerate(rows, start=1):
            line = ctk.CTkFrame(readiness, fg_color="transparent")
            line.grid(row=row, column=0, sticky="ew", padx=14, pady=2)
            line.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(
                line, text=label, text_color=MUTED, anchor="w",
                font=_ui_font(12)).grid(row=0, column=0, sticky="w", padx=(0, 12))
            value = ctk.CTkLabel(
                line, text="", text_color=MUTED, anchor="e", justify="right",
                wraplength=430, font=_ui_font(12, "bold"))
            value.grid(row=0, column=1, sticky="ew")
            self.cloud_readiness_labels[key] = value
        self.cloud_test_button = VisualButton(
            readiness, text=I.t("cloud_test_connection"), height=34,
            fg_color=ACCENT_SOFT, hover_color=LINE, text_color=ACCENT,
            command=self._test_cloud_readiness)
        self.cloud_test_button.grid(row=5, column=0, sticky="w", padx=14, pady=(7, 10))

    @staticmethod
    def _format_cloud_timestamp(value):
        if not value:
            return I.t("cloud_status_never")
        try:
            timestamp = datetime.datetime.fromisoformat(str(value))
            if timestamp.tzinfo is None:
                timestamp = timestamp.replace(tzinfo=datetime.timezone.utc)
            return timestamp.astimezone().strftime("%Y-%m-%d %H:%M")
        except (TypeError, ValueError):
            return I.t("cloud_status_never")

    def _set_cloud_status(self, key, text, color=MUTED):
        label = getattr(self, "cloud_readiness_labels", {}).get(key)
        if label is not None:
            label.configure(text=text, text_color=color)

    def _refresh_cloud_key_status(self, *_args, reset_tests=True):
        if not hasattr(self, "cloud_readiness_labels"):
            return
        configured = bool(self.cloud_key_var.get().strip())
        self._set_cloud_status(
            "key", I.t("cloud_status_configured") if configured
            else I.t("cloud_status_not_configured"), ACCENT if configured else DANGER)
        self._set_cloud_status(
            "upload", self._format_cloud_timestamp(
                cfg.config.get("cloud_last_upload_at", "")), TEXT_SECONDARY)
        if reset_tests and not self._cloud_test_running:
            self._set_cloud_status("api", I.t("cloud_status_not_tested"), MUTED)
            self._set_cloud_status("viewer", I.t("cloud_status_not_tested"), MUTED)

    def _test_cloud_readiness(self):
        if self._cloud_test_running:
            return
        key = self.cloud_key_var.get().strip()
        if not key:
            self._refresh_cloud_key_status()
            return
        self._cloud_test_running = True
        self.cloud_test_button.configure(
            state="disabled", text=I.t("cloud_status_testing"))
        self._set_cloud_status("api", I.t("cloud_status_testing"), WARNING)
        self._set_cloud_status("viewer", I.t("cloud_status_testing"), WARNING)

        def finish_button():
            self._cloud_test_running = False
            self.cloud_test_button.configure(
                state="normal", text=I.t("cloud_test_connection"))

        def ready(result):
            if not self.winfo_exists():
                return
            finish_button()
            self._set_cloud_status("api", I.t("cloud_status_ready"), GOOD)
            self._set_cloud_status("viewer", I.t("cloud_status_reachable"), GOOD)

        def failed(error):
            if not self.winfo_exists():
                return
            finish_button()
            code = error.code if isinstance(error, cloud_rx.CloudRxError) else "unexpected"
            text = I.t("cloud_error_" + code)
            if code == "viewer":
                self._set_cloud_status("api", I.t("cloud_status_ready"), GOOD)
                self._set_cloud_status("viewer", text, DANGER)
            else:
                self._set_cloud_status("api", text, DANGER)
                self._set_cloud_status("viewer", I.t("cloud_status_not_tested"), MUTED)

        self.master.submit_background(
            lambda: cloud_rx.check_readiness(key), ready, on_error=failed,
            label=I.t("cloud_status_testing"), silent=True)

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
            "clinic": "clinic identity name address phone logo header preview location maps latitude longitude",
            "documents": "document pdf word paper margin export folder header logo language",
            "qr": "qr verification viewer key link",
            "database": "drug database medicine import replace remove validate classification",
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
        if (getattr(self.master, "_export_busy", False)
                or getattr(self.master, "_history_save_future", None) is not None):
            messagebox.showinfo(APP_TITLE, I.t("export_close_pending"), parent=self)
            return
        if self._location_busy or self.location_input_var.get().strip():
            self._show_section("clinic")
            messagebox.showerror(I.t("location_title"), I.t("location_apply_first"), parent=self)
            return
        latitude, longitude = self.latitude_var.get().strip(), self.longitude_var.get().strip()
        if latitude or longitude or self.include_location_var.get():
            try:
                pin = clinic_location.validate_coordinates(latitude, longitude)
                latitude, longitude = f"{pin.latitude:.7f}", f"{pin.longitude:.7f}"
                if (self.latitude_var.get(), self.longitude_var.get()) != self._confirmed_location:
                    if not messagebox.askyesno(I.t("location_title"), I.t("location_confirm",
                            latitude=latitude, longitude=longitude, accuracy=""), parent=self):
                        return
                    self._confirmed_location = (self.latitude_var.get(), self.longitude_var.get())
            except clinic_location.LocationError:
                self._show_section("clinic")
                messagebox.showerror(I.t("location_title"), I.t("location_invalid"), parent=self)
                return
        try:
            website = qu.normalize_clinic_website(self.clinic_website_var.get())
        except ValueError:
            self._show_section("clinic")
            messagebox.showerror(
                I.t("set_title"), I.t("settings_clinic_website_invalid"), parent=self)
            return
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
        language = "ar" if self.language_var.get() == "العربية" else "en"
        language_changed = language != cfg.config.language
        try:
            self._store_current_printer_profile()
            with cfg.config.batch_save():
                cfg.config.set_clinic(name=self.clinic_name_var.get().strip(),
                                      address=self.clinic_address_var.get().strip(),
                                      phone=self.clinic_phone_var.get().strip(),
                                      website=website, logo_path=logo_path,
                                      latitude=latitude, longitude=longitude,
                                      include_location=bool(self.include_location_var.get()))
                cfg.config.cloud_rx_api_key = self.cloud_key_var.get().strip()
                cfg.config.paper_size = self.paper_var.get()
                cfg.config.language = language
                cfg.config.set_ui_font_sizes(
                    int(self.dropdown_font_var.get()), int(self.patient_font_var.get()),
                    int(self.instruction_font_var.get()))
                document_language = {
                    "English": "en", "العربية": "ar",
                    I.t("settings_same_as_interface"): "interface",
                }.get(self.document_language_var.get(), "interface")
                cfg.config.data["document_defaults"] = {
                    "language": document_language,
                    "show_header": bool(self.document_header_var.get()),
                    "logo_size": "medium",
                    "margin_mm": 16,
                    "font_size": int(self.document_font_var.get()),
                    "export_folder": export_folder,
                    "selected_printer": self.printer_profile_var.get(),
                    "printer_profiles": print_layout.normalize_profiles(
                        self._printer_profiles),
                    "calibration_presets": print_layout.normalize_presets(
                        self._calibration_presets),
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
            self.dropdown_font_var.set("20")
            self.patient_font_var.set("18")
            self.instruction_font_var.set("18")
        elif self._active_section == "clinic":
            self.remove_clinic_location()
            self.clinic_name_var.set("")
            self.clinic_address_var.set("")
            self.clinic_phone_var.set("")
            self.clinic_website_var.set("")
            self.logo_var.set("")
        elif self._active_section == "documents":
            self.paper_var.set(cfg.DEFAULT_PAPER)
            self.document_language_var.set(I.t("settings_same_as_interface"))
            self.document_header_var.set(True)
            self.document_font_var.set("10")
            self.export_folder_var.set("")
            self._printer_profiles = {
                print_layout.DEFAULT_PROFILE_NAME: print_layout.default_profile()}
            self._calibration_presets = {}
            self.calibration_preset_var.set(I.t("settings_no_preset"))
            self._touch_presets()
            self._loaded_printer_name = print_layout.DEFAULT_PROFILE_NAME
            self.printer_profile_var.set(print_layout.DEFAULT_PROFILE_NAME)
            profile = print_layout.default_profile()
            for variable, key in (
                    (self.layout_horizontal_var, "horizontal_offset_mm"),
                    (self.layout_vertical_var, "vertical_offset_mm"),
                    (self.layout_medication_top_var, "medication_top_mm"),
                    (self.layout_line_spacing_var, "line_spacing"),
                    (self.layout_field_gap_var, "field_gap_mm"),
                    (self.layout_qr_size_var, "qr_size_mm"),
                    (self.layout_qr_side_var, "qr_side_margin_mm"),
                    (self.layout_qr_bottom_var, "qr_bottom_margin_mm")):
                variable.set(self._format_layout_number(profile[key]))
            self.layout_qr_position_var.set(profile["qr_position"])
            self.layout_background_var.set("")
        elif self._active_section == "qr":
            self.cloud_key_var.set("")
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
        self.dropdown_font_var.set("20")
        self.patient_font_var.set("18")
        self.instruction_font_var.set("18")
        self.clinic_address_var.set("")
        self.clinic_phone_var.set("")
        self.clinic_website_var.set("")
        self.remove_clinic_location()
        self.logo_var.set("")
        self.paper_var.set(cfg.DEFAULT_PAPER)
        self.document_language_var.set(I.t("settings_same_as_interface"))
        self.document_header_var.set(True)
        self.document_font_var.set("10")
        self.export_folder_var.set("")
        self._printer_profiles = {
            print_layout.DEFAULT_PROFILE_NAME: print_layout.default_profile()}
        self._calibration_presets = {}
        self.calibration_preset_var.set(I.t("settings_no_preset"))
        self._touch_presets()
        self._loaded_printer_name = print_layout.DEFAULT_PROFILE_NAME
        self.printer_profile_var.set(print_layout.DEFAULT_PROFILE_NAME)
        self._apply_layout_profile_variables(print_layout.default_profile())
        self.cloud_key_var.set("")
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
                              self.clinic_phone_var.get().strip(),
                              self.clinic_website_var.get().strip()) if item)
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

if os.environ.get("RX_PERF_DEBUG") == "1":
    perf_probe.install(globals())

if __name__ == "__main__":
    app = App()
    app.mainloop()
