"""Small, testable workflow helpers used by the desktop interface."""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime, timezone
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


@dataclass(frozen=True)
class WorkflowIssue:
    """A clinician-facing validation result with explicit severity."""

    severity: str
    field: str
    message: str


def medicine_present(item: Any) -> bool:
    """Return True when a row has either a trade or scientific medicine name."""
    get = item.get if isinstance(item, dict) else lambda key, default="": getattr(item, key, default)
    return bool(str(get("brand_name", "") or "").strip()
                or str(get("generic_name", "") or "").strip())


def compact_medicine_summary(item: Any) -> str:
    """Build the one-line summary used by collapsed medication rows."""
    get = item.get if isinstance(item, dict) else lambda key, default="": getattr(item, key, default)
    brand = str(get("brand_name", "") or "").strip()
    scientific = str(get("generic_name", "") or "").strip()
    if brand and scientific and brand.casefold() != scientific.casefold():
        name = f"{brand}  ·  {scientific}"
    else:
        name = brand or scientific or "—"
    details = [str(get(key, "") or "").strip()
               for key in ("dosage", "frequency", "duration", "notes", "quantity")]
    details = [value for value in details if value]
    return name + (("   ·   " + "   ·   ".join(details)) if details else "")


def validate_prescription_workflow(patient: Any, medicines: list[Any]) -> list[WorkflowIssue]:
    """Validate entry completeness without making clinical decisions.

    Blocking errors are structural only.  Missing regimen details remain warnings;
    the clinician can still continue after reviewing them.
    """
    issues: list[WorkflowIssue] = []
    # Patient identity is optional for medicine-only review and export. Patient
    # saving remains a separate workflow and validates its required fields.
    present = [item for item in medicines if medicine_present(item)]
    if not present:
        issues.append(WorkflowIssue("error", "medicines", "Add at least one medicine."))
        return issues
    for index, item in enumerate(present, 1):
        get = item.get if isinstance(item, dict) else lambda key, default="": getattr(item, key, default)
        for key, label in (("dosage", "dosage"), ("frequency", "frequency"),
                           ("duration", "duration")):
            if not str(get(key, "") or "").strip():
                issues.append(WorkflowIssue(
                    "warning", f"medicines.{index}.{key}",
                    f"Medicine {index} has no {label}."))
    return issues


def workflow_step(patient: Any, medicines: list[Any], view: str = "entry") -> int:
    """Return the visible 1-based Patient/Medicines/Export step.

    ``review`` remains accepted for callers from older integrations, but it now
    resolves to the Export step because review is no longer a separate stage.
    """
    if view in {"review", "export"}:
        return 3
    get = patient.get if isinstance(patient, dict) else lambda key, default="": getattr(patient, key, default)
    return 2 if str(get("name", "") or "").strip() else 1


def prescription_draft(patient: Any, medicines: list[Any], *, active_page="patient") -> dict:
    """Create a minimal local resume snapshot without changing export schemas."""
    if isinstance(patient, dict):
        patient_data = {key: str(patient.get(key, "") or "")
                        for key in ("name", "age", "sex", "id")}
    else:
        patient_data = {key: str(getattr(patient, key, "") or "")
                        for key in ("name", "age", "sex", "id")}
    fields = ("generic_name", "brand_name", "dosage", "frequency",
              "duration", "notes", "quantity")
    medicine_data = []
    for item in medicines:
        get = item.get if isinstance(item, dict) else lambda key, default="": getattr(item, key, default)
        values = {key: str(get(key, "") or "") for key in fields}
        if medicine_present(values):
            medicine_data.append(values)
    return {
        "schema": 1,
        "saved_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "active_page": active_page if active_page in {"patient", "medications"} else "patient",
        "patient": patient_data,
        "medicines": medicine_data,
    }


def valid_prescription_draft(value: Any) -> bool:
    """Accept only the small, versioned local draft shape."""
    if not isinstance(value, dict) or value.get("schema") != 1:
        return False
    if not isinstance(value.get("patient"), dict) or not isinstance(value.get("medicines"), list):
        return False
    return bool(str(value["patient"].get("name", "") or "").strip()
                or any(medicine_present(item) for item in value["medicines"]
                       if isinstance(item, dict)))


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
