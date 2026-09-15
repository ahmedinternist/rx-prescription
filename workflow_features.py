"""Small, testable workflow helpers used by the desktop interface."""
from __future__ import annotations

import re
import unicodedata
from typing import Any


def _norm(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return " ".join(text.split())


def medicine_identity(item: Any) -> str:
    get = item.get if isinstance(item, dict) else lambda key, default="": getattr(item, key, default)
    return _norm(get("generic_name", "") or get("brand_name", ""))


def compare_prescriptions(current: list[Any], previous: list[Any]) -> dict[str, list[dict]]:
    """Classify current medicines against a previous snapshot by medicine identity."""
    current_map = {medicine_identity(item): item for item in current if medicine_identity(item)}
    previous_map = {medicine_identity(item): item for item in previous if medicine_identity(item)}
    fields = ("generic_name", "brand_name", "dosage", "frequency", "duration", "notes", "quantity")

    def as_dict(item):
        return dict(item) if isinstance(item, dict) else {
            key: str(getattr(item, key, "") or "") for key in fields}

    result = {"added": [], "removed": [], "changed": [], "unchanged": []}
    for key, item in current_map.items():
        if key not in previous_map:
            result["added"].append(as_dict(item))
            continue
        before = as_dict(previous_map[key])
        after = as_dict(item)
        changes = [field for field in fields[2:] if _norm(before.get(field)) != _norm(after.get(field))]
        if changes:
            result["changed"].append({"before": before, "after": after, "fields": changes})
        else:
            result["unchanged"].append(after)
    for key, item in previous_map.items():
        if key not in current_map:
            result["removed"].append(as_dict(item))
    return result


def _quantity_unit_from_form(form: str) -> str:
    """Return a countable dispensing unit for a database dosage form."""
    value = _norm(form)
    for pattern, unit in (
        (r"tablet|\btab\b|قرص|اقراص|أقراص", "tablet"),
        (r"capsule|\bcap\b|كبسول", "capsule"),
        (r"syrup|solution|suspension|\bml\b|مل", "mL"),
        (r"inhaler|puff|بخ", "puff"),
        (r"unit|insulin|وحد", "units"),
    ):
        if re.search(pattern, value, re.IGNORECASE):
            return unit
    return "dose"


def calculate_medicine_quantity(dosage: str, frequency: str, duration: str,
                                form: str = "") -> str:
    """Calculate quantity only when dose, frequency and duration are explicit.

    PRN and ambiguous strength-only doses intentionally return an empty string.
    """
    dose = _norm(dosage)
    freq = _norm(frequency)
    dur = _norm(duration)
    if not dose or not freq or not dur or "prn" in freq or "عند الحاجة" in freq:
        return ""

    unit_patterns = (
        (r"([\d.]+)\s*(tablets?|tabs?|قرص|اقراص|أقراص)", "tablet"),
        (r"([\d.]+)\s*(capsules?|caps?|كبسول(?:ة|ات)?)", "capsule"),
        (r"([\d.]+)\s*(ml|مل)", "mL"),
        (r"([\d.]+)\s*(units?|وحد(?:ة|ات)?)", "units"),
        (r"([\d.]+)\s*(puffs?|بخ(?:ة|ات)?)", "puffs"),
    )
    dose_value = None
    unit = ""
    for pattern, canonical in unit_patterns:
        match = re.search(pattern, dose, re.IGNORECASE)
        if match:
            dose_value, unit = float(match.group(1)), canonical
            break
    if dose_value is None:
        # Many prescription workflows enter a simple administration count
        # (for example dosage ``1``) and keep the tablet/capsule form in the
        # selected database medicine.  Continue to reject strength-only input
        # such as ``500 mg`` because that is not a dispensable quantity.
        bare_dose = re.fullmatch(r"[\d.]+", dose)
        if not bare_dose:
            return ""
        dose_value = float(bare_dose.group(0))
        unit = _quantity_unit_from_form(form)

    per_day = None
    for pattern, value in (
        (r"1x4|qid|q6h|كل 6", 4), (r"1x3|tid|q8h|كل 8", 3),
        (r"1x2|bid|q12h|كل 12", 2), (r"1x1|\bod\b|\bqd\b", 1),
        (r"q4h|كل 4", 6), (r"qod|يوم بعد يوم", .5),
    ):
        if re.search(pattern, freq, re.IGNORECASE):
            per_day = value
            break
    weekly = re.search(r"(^|\D)([12])x/week", freq)
    if weekly:
        per_day = float(weekly.group(2)) / 7
    if per_day is None or "stat" in freq or "جرعة واحدة" in freq:
        return ""

    match = re.search(r"([\d.]+)", dur)
    if not match:
        return ""
    duration_value = float(match.group(1))
    if re.search(r"week|أسبوع|اسبوع", dur):
        days = duration_value * 7
    elif re.search(r"month|شهر", dur):
        days = duration_value * 30
    elif re.search(r"day|days|يوم|أيام|ايام", dur):
        days = duration_value
    elif re.fullmatch(r"[\d.]+", dur):
        # A bare duration is treated as days, matching the duration field's
        # most common clinical use.
        days = duration_value
    else:
        return ""

    total = dose_value * per_day * days
    display = str(int(total)) if total.is_integer() else f"{total:.1f}".rstrip("0").rstrip(".")
    if unit in {"tablet", "capsule", "puff", "dose"} and total != 1:
        unit += "s"
    return f"{display} {unit}"
