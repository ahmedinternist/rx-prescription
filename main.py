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
import tkinter as tk
import uuid
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog
from typing import List

import customtkinter as ctk

import config as cfg
import drug_db as dbmod
import pdf_generator as pdfgen
import qr_utils as qu
import i18n as I

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

APP_TITLE = "Rx Prescription Printer"

logging.basicConfig(filename=cfg.LOG_PATH, level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s")


def _report_exception(exc_type, exc_value, exc_traceback):
    logging.exception("Unexpected application error", exc_info=(exc_type, exc_value, exc_traceback))
    messagebox.showerror(APP_TITLE, f"An unexpected error occurred. Details were saved to:\n{cfg.LOG_PATH}")


tk.Tk.report_callback_exception = staticmethod(_report_exception)

# --- Softer, cohesive modern palette (Windows 11 card style) ----------------
BG = "#eef1f7"          # neutral app background
CARD = "#ffffff"        # card surface
ACCENT = "#3a66d8"      # primary accent (softer blue)
ACCENT_HOVER = "#2f55bd"
ACCENT_SOFT = "#e7edfb" # tinted fill for chips/accents
LINE = "#e3e8f1"        # subtle border
MUTED = "#69748c"       # secondary text
DANGER = "#d24b4b"      # delete
DANGER_HOVER = "#b53c3c"
GOOD = "#2f9e6b"        # save / print
GOOD_HOVER = "#268a5c"

PAD = 14                # inner card padding


class DrugRow(ctk.CTkFrame):
    """One medication row: NAME on top, 4 detail boxes beneath it."""

    def __init__(self, master, db, on_change, on_remove, on_duplicate, **kwargs):
        super().__init__(master, **kwargs)
        self.db = db
        self.on_change = on_change
        self.on_remove = on_remove
        self.on_duplicate = on_duplicate
        self._ac_top = None
        self._ac_listbox = None
        self._brand = ""

        self.name_var = tk.StringVar()
        name_row = ctk.CTkFrame(self, fg_color="transparent")
        name_row.pack(fill="x", padx=PAD, pady=(PAD - 4, 2))
        ctk.CTkLabel(name_row, text=I.t("drug"), width=48, anchor="w",
                     text_color=ACCENT, font=ctk.CTkFont(weight="bold", size=11)
                     ).pack(side="left")
        self.name_entry = ctk.CTkEntry(self, textvariable=self.name_var,
                                       placeholder_text=I.t("drug"), height=32,
                                       corner_radius=8, border_color=LINE)
        self.name_entry.pack(fill="x", padx=PAD, pady=(0, 8))
        self.name_entry.bind("<KeyRelease>", self._on_type)
        # Sticky dropdown: stays open until selection or Escape.
        self.name_entry.bind("<Down>", self._ac_down)
        self.name_entry.bind("<Up>", self._ac_up)
        self.name_entry.bind("<Return>", self._ac_choose_current)
        self.name_entry.bind("<Escape>", lambda e: self._hide_ac())

        details = ctk.CTkFrame(self, fg_color="transparent")
        details.pack(fill="x", padx=PAD, pady=(0, 4))

        self.dosage_var = tk.StringVar()
        self.freq_var = tk.StringVar()
        self.dur_var = tk.StringVar()
        self.notes_var = tk.StringVar()
        self.dosage_entry = self._box(details, I.t("dosage"),
                                      I.t("ph_dosage"), self.dosage_var)
        self.freq_entry = self._box(details, I.t("frequency"),
                                    I.t("ph_frequency"), self.freq_var)
        self.dur_entry = self._box(details, I.t("duration"),
                                   I.t("ph_duration"), self.dur_var)
        self.notes_entry = self._box(details, I.t("notes"),
                                     I.t("ph_notes"), self.notes_var)

        # visible action buttons
        act = ctk.CTkFrame(self, fg_color="transparent")
        act.pack(fill="x", padx=PAD, pady=(2, PAD - 4))
        ctk.CTkLabel(act, text="", width=48).pack(side="left")  # spacer align
        ctk.CTkButton(act, text=I.t("duplicate"), width=96, height=30,
                      corner_radius=8, fg_color=ACCENT_SOFT,
                      text_color=ACCENT, hover_color="#d8e1f7",
                      border_width=1, border_color=LINE,
                      command=lambda: self.on_duplicate(self)).pack(side="right", padx=4)
        ctk.CTkButton(act, text=I.t("delete"), width=90, height=30,
                      corner_radius=8, fg_color=DANGER, hover_color=DANGER_HOVER,
                      command=self.on_remove).pack(side="right", padx=4)

    def _box(self, parent, label, ph, var):
        col = ctk.CTkFrame(parent, fg_color="transparent")
        col.pack(side="left", fill="both", expand=True, padx=4)
        ctk.CTkLabel(col, text=label, anchor="w", text_color=MUTED,
                     font=ctk.CTkFont(size=10, weight="bold")).pack(fill="x", pady=(0, 2))
        e = ctk.CTkEntry(col, textvariable=var, height=30, corner_radius=8,
                         border_color=LINE, placeholder_text=ph)
        e.pack(fill="x")
        e.bind("<KeyRelease>", lambda ev: self.on_change())
        return e

    # autocomplete (sticky, large)
    def _on_type(self, event=None):
        self.on_change()
        q = self.name_var.get().strip()
        if len(q) < 1:
            self._hide_ac()
            return
        matches = self.db.search(q, limit=20)
        if not matches:
            self._hide_ac()
            return
        self._show_ac(matches)

    def _show_ac(self, matches):
        self._hide_ac()
        top = tk.Toplevel(self)
        top.wm_overrideredirect(True)
        x = self.name_entry.winfo_rootx()
        y = self.name_entry.winfo_rooty() + self.name_entry.winfo_height()
        top.geometry(f"+{x}+{y}")
        self._ac_top = top
        self._ac_matches = matches
        self._ac_index = -1
        lb = tk.Listbox(top,
                        height=min(36, len(matches) * 2),
                        font=("Segoe UI", 30),
                        bg="white", fg="#111", relief="solid", borderwidth=3,
                        activestyle="dotbox",
                        highlightthickness=3, highlightbackground=ACCENT,
                        selectbackground=ACCENT, selectforeground="white")
        max_chars = max((len(d.generic_name) + len(d.brand_name or "")
                         + len(d.strength or "") + 8) for d in matches)
        width = max(110, min(190, max_chars * 2))
        lb.configure(width=width)
        for d in matches:
            label = d.generic_name
            if d.brand_name:
                label += f"  ({d.brand_name})"
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
        self.name_var.set(d.generic_name)
        self._brand = d.brand_name if d.brand_name else ""
        self._real_hide()
        self.on_change()

    def get_data(self) -> qu.DrugItem:
        return qu.DrugItem(
            generic_name=self.name_var.get().strip(),
            brand_name=self._brand,
            dosage=self.dosage_var.get().strip(),
            frequency=self.freq_var.get().strip(),
            duration=self.dur_var.get().strip(),
            notes=self.notes_var.get().strip(),
        )


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1120x820")
        self.configure(fg_color=BG)
        I.set_lang(cfg.config.language)
        cfg.ensure_seed_db()
        self.db = dbmod.DrugDatabase(cfg.config.drug_db_path)
        self.rows: List[DrugRow] = []
        self._build_ui()

    def _build_ui(self):
        # toolbar
        self.toolbar = ctk.CTkFrame(self, height=52, fg_color="white",
                                   border_color=LINE, border_width=1,
                                   corner_radius=12)
        self.toolbar.pack(side="top", fill="x", padx=12, pady=12)
        ctk.CTkButton(self.toolbar, text=I.t("settings"), width=96, height=34,
                      corner_radius=8, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.open_settings).pack(side="left", padx=4)
        ctk.CTkButton(self.toolbar, text=I.t("import_db"), width=104, height=34,
                      corner_radius=8, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.import_db).pack(side="left", padx=2)
        ctk.CTkButton(self.toolbar, text=I.t("export_db"), width=104, height=34,
                      corner_radius=8, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.export_db).pack(side="left", padx=2)
        ctk.CTkLabel(self.toolbar, text=I.t("paper"), text_color=MUTED).pack(
            side="left", padx=(16, 4))
        self.paper_var = tk.StringVar(value=cfg.config.paper_size)
        self.paper_menu = ctk.CTkOptionMenu(self.toolbar,
                                            values=list(cfg.PAPER_SIZES.keys()),
                                            variable=self.paper_var, width=92,
                                            height=34, corner_radius=8,
                                            fg_color=ACCENT_SOFT, text_color=ACCENT,
                                            button_color=ACCENT, button_hover_color=ACCENT_HOVER,
                                            dropdown_hover_color=ACCENT_SOFT,
                                            command=self.on_paper_change)
        self.paper_menu.pack(side="left", padx=2)
        ctk.CTkLabel(self.toolbar, text=I.t("language"), text_color=MUTED).pack(
            side="left", padx=(16, 4))
        self.lang_var = tk.StringVar(value=cfg.config.language)
        self.lang_menu = ctk.CTkOptionMenu(self.toolbar, values=["en", "ar"],
                                           variable=self.lang_var, width=72,
                                           height=34, corner_radius=8,
                                           fg_color=ACCENT_SOFT, text_color=ACCENT,
                                           button_color=ACCENT, button_hover_color=ACCENT_HOVER,
                                           dropdown_hover_color=ACCENT_SOFT,
                                           command=self.on_lang_change)
        self.lang_menu.pack(side="left", padx=2)
        ctk.CTkLabel(self.toolbar, text="Profile:", text_color=MUTED).pack(side="left", padx=(16, 4))
        self.profile_var = tk.StringVar(value=cfg.config.get("active_profile", "Default"))
        self.profile_menu = ctk.CTkOptionMenu(
            self.toolbar, values=cfg.config.profile_names() + ["+ New profile…"],
            variable=self.profile_var, width=150, height=34, corner_radius=8,
            fg_color=ACCENT_SOFT, text_color=ACCENT, button_color=ACCENT,
            button_hover_color=ACCENT_HOVER, command=self.on_profile_change)
        self.profile_menu.pack(side="left", padx=2)
        self.db_label = ctk.CTkLabel(self.toolbar, text=I.t("db_count", n=len(self.db.drugs)),
                                     text_color=MUTED)
        self.db_label.pack(side="right", padx=10)

        # scroll area
        self.scroll = ctk.CTkScrollableFrame(self, fg_color=BG, corner_radius=12)
        self.scroll.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        self.build_forms()

        # action bar (fixed footer)
        self.action = ctk.CTkFrame(self, height=62, fg_color="white",
                                   border_color=LINE, border_width=1,
                                   corner_radius=12)
        self.action.pack(side="bottom", fill="x", padx=12, pady=(0, 12))
        # left cluster
        ctk.CTkButton(self.action, text=I.t("add_drug"), width=120, height=40,
                      corner_radius=8, fg_color=ACCENT_SOFT, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color="#d8e1f7",
                      command=self.add_row).pack(side="left", padx=6)
        ctk.CTkButton(self.action, text=I.t("preview"), width=120, height=40,
                      corner_radius=8, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.preview).pack(side="left", padx=6)
        ctk.CTkButton(self.action, text=I.t("clear"), width=100, height=40,
                      corner_radius=8, fg_color=CARD, text_color=MUTED,
                      border_color=LINE, border_width=1, hover_color="#eef0f4",
                      command=self.clear_all).pack(side="left", padx=6)
        # right cluster (primary actions)
        ctk.CTkButton(self.action, text=I.t("export_compact"), width=160, height=40,
                      corner_radius=8, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.export_label).pack(side="right", padx=6)
        ctk.CTkButton(self.action, text=I.t("export_word"), width=120, height=40,
                      corner_radius=8, fg_color=CARD, text_color=ACCENT,
                      border_color=LINE, border_width=1, hover_color=ACCENT_SOFT,
                      command=self.export_word).pack(side="right", padx=6)
        ctk.CTkButton(self.action, text=I.t("save_profile"), width=130, height=40,
                      corner_radius=8, fg_color=GOOD, hover_color=GOOD_HOVER,
                      command=self.save_profile).pack(side="right", padx=6)
        ctk.CTkButton(self.action, text=I.t("print"), width=110, height=40,
                      corner_radius=8, fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      font=ctk.CTkFont(weight="bold", size=13),
                      command=self.print_pdf).pack(side="right", padx=6)

        self.add_row()
        self.load_profile_into_ui()

    # -- forms ---------------------------------------------------------------
    def section(self, parent, title, accent=True):
        f = ctk.CTkFrame(parent, fg_color=CARD, border_color=LINE,
                         border_width=1, corner_radius=14)
        f.pack(fill="x", padx=4, pady=8)
        if accent:
            bar = ctk.CTkFrame(f, height=4, fg_color=ACCENT, corner_radius=0)
            bar.pack(fill="x", side="top")
        ctk.CTkLabel(f, text=title, font=ctk.CTkFont(weight="bold", size=14),
                     text_color=ACCENT, anchor="w").pack(anchor="w", padx=PAD, pady=(10, 4))
        return f

    def field(self, parent, label, var, width=260, placeholder=None):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=PAD, pady=7)
        ctk.CTkLabel(row, text=label, width=130, anchor="w",
                     text_color=MUTED, font=ctk.CTkFont(size=11),
                     height=32).pack(side="left")
        ctk.CTkEntry(row, textvariable=var, width=width, height=32,
                     corner_radius=8, border_color=LINE,
                     placeholder_text=placeholder).pack(
            side="left", fill="x", expand=True)
        return row

    def sex_field(self, parent, label, var):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=PAD, pady=7)
        ctk.CTkLabel(row, text=label, width=130, anchor="w",
                     text_color=MUTED, font=ctk.CTkFont(size=11),
                     height=32).pack(side="left")
        opts = [I.t("sex_m"), I.t("sex_f")]
        code_of = {I.t("sex_m"): "M", I.t("sex_f"): "F"}
        cur = var.get()
        sel = next((lab for lab, c in code_of.items() if c == cur), opts[0])
        om = ctk.CTkOptionMenu(row, values=opts, width=120, height=32,
                               corner_radius=8, fg_color=ACCENT_SOFT,
                               text_color=ACCENT, button_color=ACCENT,
                               button_hover_color=ACCENT_HOVER,
                               dropdown_hover_color=ACCENT_SOFT,
                               command=lambda v, code_of=code_of, var=var: var.set(code_of.get(v, "")))
        om.set(sel)
        om.pack(side="left")
        return row

    def build_forms(self):
        # Prescriber + Patient SIDE BY SIDE (compact)
        self.doctor_vars = {k: tk.StringVar() for k in ["name", "license_no", "specialty"]}
        self.patient_vars = {k: tk.StringVar() for k in ["name", "age", "sex", "id_number", "allergies"]}
        self.rx_vars = {k: tk.StringVar() for k in ["diagnosis", "refills"]}

        side = ctk.CTkFrame(self.scroll, fg_color="transparent")
        side.pack(fill="x", padx=4, pady=8)

        d = self.section(side, I.t("prescriber"))
        d.pack(side="left", fill="both", expand=True, padx=4)
        self.field(d, I.t("f_name"), self.doctor_vars["name"], 220)
        self.field(d, I.t("license"), self.doctor_vars["license_no"], 160)
        self.field(d, I.t("specialty"), self.doctor_vars["specialty"], 180)

        p = self.section(side, I.t("patient"))
        p.pack(side="left", fill="both", expand=True, padx=4)
        self.field(p, I.t("f_name"), self.patient_vars["name"], 220)
        self.field(p, I.t("age"), self.patient_vars["age"], 90)
        self.sex_field(p, I.t("sex"), self.patient_vars["sex"])
        self.field(p, I.t("patient_id"), self.patient_vars["id_number"], 160)
        self.field(p, I.t("allergies"), self.patient_vars["allergies"], 220,
                   "None known / enter allergies")

        details = self.section(self.scroll, I.t("clinical_details"))
        self.field(details, I.t("diagnosis"), self.rx_vars["diagnosis"], 280, "optional")
        self.field(details, I.t("refills"), self.rx_vars["refills"], 100, "optional")

        # Medications
        dr = self.section(self.scroll, I.t("medications"))
        hint = ctk.CTkLabel(dr, text=I.t("drug_hint"),
                            text_color=MUTED, font=ctk.CTkFont(size=10), anchor="w")
        hint.pack(fill="x", padx=PAD, pady=(0, 8))
        self.drugs_frame = ctk.CTkFrame(dr, fg_color="transparent")
        self.drugs_frame.pack(fill="x", padx=6, pady=2)

    # -- drug rows ----------------------------------------------------------
    def add_row(self, data=None):
        row = DrugRow(self.drugs_frame, self.db, self.on_any_change,
                      lambda: self.remove_row(row), lambda r: self.duplicate_row(r),
                      fg_color=CARD, border_color=LINE, border_width=1, corner_radius=12)
        if data:
            row.name_var.set(data.generic_name)
            row._brand = data.brand_name
            row.dosage_var.set(data.dosage)
            row.freq_var.set(data.frequency)
            row.dur_var.set(data.duration)
            row.notes_var.set(data.notes)
        row.pack(fill="x", padx=2, pady=6)
        self.rows.append(row)

    def duplicate_row(self, row):
        self.add_row(data=row.get_data())

    def remove_row(self, row):
        if len(self.rows) <= 1:
            messagebox.showinfo("Info", I.t("msg_at_least_one"))
            return
        row.destroy()
        self.rows.remove(row)
        self.on_any_change()

    def clear_all(self):
        for v in self.doctor_vars.values():
            v.set("")
        for v in self.patient_vars.values():
            v.set("")
        for v in self.rx_vars.values():
            v.set("")
        for r in list(self.rows):
            r.destroy()
        self.rows = []
        self.add_row()
        self.on_any_change()

    # -- data ----------------------------------------------------------------
    def collect(self):
        doctor = qu.Doctor(**{k: v.get().strip() for k, v in self.doctor_vars.items()})
        patient = qu.Patient(**{k: v.get().strip() for k, v in self.patient_vars.items()})
        drugs = [r.get_data() for r in self.rows if r.get_data().generic_name]
        rx_id = "RX-" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        clinic = qu.Clinic(**{k: v for k, v in cfg.config.get_clinic().items()
                              if k in {"name", "address", "phone", "logo_path"}})
        return qu.Prescription(clinic=clinic, doctor=doctor, patient=patient, drugs=drugs,
                               date=datetime.datetime.now().strftime("%Y-%m-%d"), rx_id=rx_id,
                               diagnosis=self.rx_vars["diagnosis"].get().strip(),
                               refills=self.rx_vars["refills"].get().strip())

    def on_any_change(self, *a):
        pass

    def on_paper_change(self, value):
        cfg.config.paper_size = value

    def on_lang_change(self, value):
        cfg.config.language = value
        I.set_lang(value)
        self.reload_texts()

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

    # -- DB ops --------------------------------------------------------------
    def import_db(self):
        path = filedialog.askopenfilename(
            title=I.t("import_db"),
            filetypes=[("Drug database", "*.csv;*.xlsx"),
                       ("CSV files", "*.csv"),
                       ("Excel files", "*.xlsx"),
                       ("All files", "*.*")])
        if not path:
            return
        choice = messagebox.askyesno(I.t("import_mode_title"), I.t("import_mode_body"))
        replace = bool(choice)
        try:
            added = self.db.import_file(path, replace=replace)
            cfg.config.drug_db_path = str(self.db.path)
            self.db_label.configure(text=I.t("db_count", n=len(self.db.drugs)))
            msg = I.t("msg_imported", n=len(self.db.drugs)) if replace \
                else I.t("msg_imported_merge", n=len(self.db.drugs), a=added)
            messagebox.showinfo(I.t("import_db"), msg)
        except Exception as e:
            messagebox.showerror(I.t("import_db"), str(e))

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
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.title(I.t("set_title"))
        self.geometry("620x650")
        self.transient(master)
        self.grab_set()

        ctk.CTkLabel(self, text=I.t("set_viewer"), anchor="w").pack(
            fill="x", padx=16, pady=(16, 2))
        self.viewer_var = tk.StringVar(value=cfg.config.viewer_base_url)
        ctk.CTkEntry(self, textvariable=self.viewer_var, width=480,
                     height=32, corner_radius=8, border_color=LINE).pack(padx=16, fill="x")
        ctk.CTkLabel(self, text=I.t("set_tip"), text_color="#777", anchor="w").pack(
            fill="x", padx=16, pady=6)
        clinic = cfg.config.get_clinic()
        ctk.CTkLabel(self, text="Clinic name:", anchor="w").pack(fill="x", padx=16, pady=(8, 2))
        self.clinic_name_var = tk.StringVar(value=clinic.get("name", ""))
        ctk.CTkEntry(self, textvariable=self.clinic_name_var, height=32, corner_radius=8,
                     border_color=LINE).pack(padx=16, fill="x")
        ctk.CTkLabel(self, text="Clinic address:", anchor="w").pack(fill="x", padx=16, pady=(8, 2))
        self.clinic_address_var = tk.StringVar(value=clinic.get("address", ""))
        ctk.CTkEntry(self, textvariable=self.clinic_address_var, height=32, corner_radius=8,
                     border_color=LINE).pack(padx=16, fill="x")
        ctk.CTkLabel(self, text="Clinic phone:", anchor="w").pack(fill="x", padx=16, pady=(8, 2))
        self.clinic_phone_var = tk.StringVar(value=clinic.get("phone", ""))
        ctk.CTkEntry(self, textvariable=self.clinic_phone_var, height=32, corner_radius=8,
                     border_color=LINE).pack(padx=16, fill="x")
        ctk.CTkLabel(self, text="Clinic logo (optional PNG/JPG):", anchor="w").pack(
            fill="x", padx=16, pady=(8, 2))
        logo_row = ctk.CTkFrame(self, fg_color="transparent")
        logo_row.pack(fill="x", padx=16)
        self.logo_var = tk.StringVar(value=clinic.get("logo_path", ""))
        ctk.CTkEntry(logo_row, textvariable=self.logo_var, height=32, corner_radius=8,
                     border_color=LINE).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(logo_row, text="Browse…", width=88, command=self.choose_logo).pack(
            side="left", padx=(8, 0))
        ctk.CTkLabel(self, text=I.t("set_db_path"), anchor="w").pack(
            fill="x", padx=16, pady=(8, 2))
        ctk.CTkLabel(self, text=str(self.master.db.path), anchor="w").pack(fill="x", padx=16)
        ctk.CTkButton(self, text=I.t("set_open"), width=200,
                      command=lambda: webbrowser.open(
                          cfg.PROJECT_DIR / "viewer.html")).pack(pady=12)
        ctk.CTkButton(self, text="Copy QR verification key", width=220,
                      command=self.copy_verification_key).pack(pady=(0, 4))
        ctk.CTkButton(self, text=I.t("set_save"), fg_color=GOOD, hover_color=GOOD_HOVER,
                      command=self.save).pack(pady=4)

    def save(self):
        try:
            cfg.config.set_clinic(name=self.clinic_name_var.get().strip(),
                                  address=self.clinic_address_var.get().strip(),
                                  phone=self.clinic_phone_var.get().strip(),
                                  logo_path=self.logo_var.get().strip())
            cfg.config.viewer_base_url = self.viewer_var.get().strip().rstrip("/")
        except Exception as exc:
            logging.exception("Could not save clinic settings")
            messagebox.showerror(I.t("set_title"), f"Settings could not be saved:\n{exc}")
            return
        self.destroy()
        messagebox.showinfo(I.t("set_title"), "Clinic settings saved.")

    def copy_verification_key(self):
        import json
        try:
            key = json.dumps(qu.verification_key(), separators=(",", ":"))
            self.clipboard_clear()
            self.clipboard_append(key)
            messagebox.showinfo(APP_TITLE,
                "Verification key copied. Add it to TRUSTED_SIGNERS in the static viewer before relying on QR verification.")
        except Exception as exc:
            logging.exception("Could not export verification key")
            messagebox.showerror(APP_TITLE, str(exc))

    def choose_logo(self):
        path = filedialog.askopenfilename(title="Clinic logo",
                                          filetypes=[("Images", "*.png;*.jpg;*.jpeg"),
                                                     ("All files", "*.*")])
        if path:
            self.logo_var.set(path)


if __name__ == "__main__":
    app = App()
    app.mainloop()
