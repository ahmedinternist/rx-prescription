"""Internationalisation (English / Arabic) for the prescription app.

Use t(key) to get the string for the currently selected language. Call
set_lang(code) to switch. Strings cover both the desktop GUI labels and the
PDF/Word document text.
"""
from __future__ import annotations

LANG = "en"

STRINGS = {
    "en": {
        "app_title": "Rx Prescription Printer",
        "settings": "Settings",
        "import_db": "Import DB…",
        "export_db": "Export DB…",
        "paper": "Paper:",
        "db_count": "DB: {n} drugs",
        "language": "Language:",
        "add_drug": "+ Add Drug",
        "preview": "Preview PDF",
        "print": "Print",
        "clear": "Clear",
        "duplicate": "Duplicate",
        "delete": "Delete",
        "export_word": "Export Word",
        "export_compact": "Medication Label (Word)",
        "save_profile": "Save Profile",
        "prescriber": "Prescriber (Doctor)",
        "patient": "Patient",
        "medications": "Medications",
        "drug_hint": "Tip: type a drug name to autocomplete from the database. "
                    "Each drug shows its name on top and Dosage / Frequency / "
                    "Duration / Notes directly below.",
        "f_name": "Full name",
        "license": "License No.",
        "specialty": "Specialty",
        "age": "Age",
        "sex": "Sex (M/F)",
        "patient_id": "Patient ID",
        "allergies": "Allergies",
        "clinical_details": "Clinical details",
        "diagnosis": "Diagnosis",
        "refills": "Refills",
        "drug": "Drug",
        "dosage": "Dosage",
        "frequency": "Frequency",
        "duration": "Duration",
        "notes": "Notes",
        # field placeholders (units)
        "ph_dosage": "e.g. 500 mg",
        "ph_frequency": "e.g. 3 times/day",
        "ph_duration": "e.g. 7 days",
        "ph_notes": "optional notes",
        "sex_m": "Male",
        "sex_f": "Female",
        # PDF / Word text
        "pdf_title": "PRESCRIPTION",
        "pdf_subtitle": "روشة طبية",
        "pdf_prescriber": "Prescriber",
        "pdf_patient": "Patient",
        "pdf_medications": "Medications",
        "pdf_signature": "Signature",
        "pdf_date": "Date:",
        "pdf_rx": "Rx No:",
        "pdf_scan": "Scan with a phone camera + internet to view this prescription.",
        "pdf_col_no": "#",
        "pdf_col_drug": "Drug (generic / brand)",
        "pdf_col_dosage": "Dosage",
        "pdf_col_freq": "Frequency",
        "pdf_col_duration": "Duration",
        "pdf_col_notes": "Notes",
        # Settings window
        "set_viewer": "Viewer page URL (the link a pharmacist lands on):",
        "set_tip": "Tip: host the included viewer.html on GitHub Pages / "
                  "Netlify for free. The data stays in the QR code (#hash).",
        "set_db_path": "Current drug database path:",
        "set_open": "Open viewer.html in browser",
        "set_save": "Save settings",
        "set_title": "Settings",
        # messages
        "msg_saved_profile": "Doctor profile saved.",
        "msg_at_least_one": "At least one drug row is required.",
        "msg_no_drugs": "Add at least one drug with a name.",
        "msg_imported": "Imported. Database now has {n} drugs.",
        "msg_imported_merge": "Imported. Database now has {n} drugs. ({a} added)",
        "msg_exported_db": "Database exported to:\n{path}",
        "msg_exported_word": "Editable Word document saved:\n{path}",
        "msg_exported_compact": "Medication label (Word) saved:\n{path}",
        "msg_settings_saved": "Settings saved.",
        "import_mode_title": "Import mode",
        "import_mode_body": "Replace the current database with this file?\n\n"
                           "YES = replace\nNO = merge (keep existing, add new)",
    },
    "ar": {
        "app_title": "طابعة الوصفات الطبية",
        "settings": "الإعدادات",
        "import_db": "استيراد قاعدة البيانات…",
        "export_db": "تصدير قاعدة البيانات…",
        "paper": "الورق:",
        "db_count": "قاعدة البيانات: {n} دواء",
        "language": "اللغة:",
        "add_drug": "+ إضافة دواء",
        "preview": "معاينة PDF",
        "print": "طباعة",
        "clear": "مسح",
        "duplicate": "تكرار",
        "delete": "حذف",
        "export_word": "تصدير Word",
        "export_compact": "ملصق الأدوية (Word)",
        "save_profile": "حفظ الملف",
        "prescriber": "الطبيب المعالج",
        "patient": "المريض",
        "medications": "الأدوية",
        "drug_hint": "تلميح: اكتب اسم الدواء للإكمال التلقائي من قاعدة البيانات. "
                    "يظهر اسم الدواء في الأعلى و الجرعة / التكرار / المدة / الملاحظات تحته.",
        "f_name": "الاسم الكامل",
        "license": "رقم الترخيص",
        "specialty": "التخصص",
        "age": "العمر",
        "sex": "الجنس (ذ/أ)",
        "patient_id": "رقم المريض",
        "allergies": "الحساسية",
        "clinical_details": "التفاصيل السريرية",
        "diagnosis": "التشخيص",
        "refills": "إعادات الصرف",
        "drug": "الدواء",
        "dosage": "الجرعة",
        "frequency": "التكرار",
        "duration": "المدة",
        "notes": "ملاحظات",
        "ph_dosage": "مثال: 500 mg",
        "ph_frequency": "مثال: 3 مرات/يوم",
        "ph_duration": "مثال: 7 أيام",
        "ph_notes": "ملاحظات اختيارية",
        "sex_m": "ذكر",
        "sex_f": "أنثى",
        "pdf_title": "وصفة طبية",
        "pdf_subtitle": "Prescription",
        "pdf_prescriber": "الطبيب المعالج",
        "pdf_patient": "المريض",
        "pdf_medications": "الأدوية",
        "pdf_signature": "التوقيع",
        "pdf_date": "التاريخ:",
        "pdf_rx": "رقم الوصفة:",
        "pdf_scan": "امسح بكاميرا الهاتف والإنترنت لعرض هذه الوصفة.",
        "pdf_col_no": "#",
        "pdf_col_drug": "الدواء (علمي / تجاري)",
        "pdf_col_dosage": "الجرعة",
        "pdf_col_freq": "التكرار",
        "pdf_col_duration": "المدة",
        "pdf_col_notes": "ملاحظات",
        "set_viewer": "رابط صفحة العرض (الرابط الذي يفتحه الصيدلي):",
        "set_tip": "تلميح: استضف ملف viewer.html مجاناً على GitHub Pages / "
                  "Netlify. البيانات تبقى داخل رمز الاستجابة السريعة (#hash).",
        "set_db_path": "مسار قاعدة بيانات الأدوية الحالية:",
        "set_open": "فتح viewer.html في المتصفح",
        "set_save": "حفظ الإعدادات",
        "set_title": "الإعدادات",
        "msg_saved_profile": "تم حفظ ملف الطبيب.",
        "msg_at_least_one": "يجب وجود صف دواء واحد على الأقل.",
        "msg_no_drugs": "أضف دواء واحداً على الأقل باسم.",
        "msg_imported": "تم الاستيراد. قاعدة البيانات تحتوي الآن على {n} دواء.",
        "msg_imported_merge": "تم الاستيراد. قاعدة البيانات تحتوي الآن على {n} دواء. (تمت إضافة {a})",
        "msg_exported_db": "تم تصدير قاعدة البيانات إلى:\n{path}",
        "msg_exported_word": "تم حفظ مستند Word قابل للتحرير:\n{path}",
        "msg_exported_compact": "تم حفظ ملصق الأدوية (Word):\n{path}",
        "msg_settings_saved": "تم حفظ الإعدادات.",
        "import_mode_title": "وضع الاستيراد",
        "import_mode_body": "استبدال قاعدة البيانات الحالية بهذا الملف؟\n\n"
                           "نعم = استبدال\nلا = دمج (الاحتفاظ الحالي، إضافة الجديد)",
    },
}


def set_lang(code: str) -> None:
    global LANG
    if code in STRINGS:
        LANG = code


def get_lang() -> str:
    return LANG


def t(key: str, **kw) -> str:
    s = STRINGS.get(LANG, STRINGS["en"]).get(key, key)
    if kw:
        try:
            s = s.format(**kw)
        except Exception:
            pass
    return s
