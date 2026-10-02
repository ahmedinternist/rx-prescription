"""Printer-aware Word layout profiles for pre-printed prescriptions."""
from __future__ import annotations

from copy import deepcopy
from typing import Any


DEFAULT_PROFILE_NAME = "Default printer"


def default_profile() -> dict[str, Any]:
    """Return an independent, conservative A5 pre-printed-paper layout."""
    return {
        "horizontal_offset_mm": 0.0,
        "vertical_offset_mm": 0.0,
        "medication_top_mm": 82.0,
        "line_spacing": 1.0,
        "field_gap_mm": 8.0,
        "qr_position": "Left",
        "qr_size_mm": 34.0,
        "qr_side_margin_mm": 12.0,
        "qr_bottom_margin_mm": 20.0,
        "preview_background_path": "",
    }


_LIMITS = {
    "horizontal_offset_mm": (-25.0, 25.0),
    "vertical_offset_mm": (-25.0, 25.0),
    "medication_top_mm": (25.0, 160.0),
    "line_spacing": (0.8, 2.0),
    "field_gap_mm": (2.0, 30.0),
    "qr_size_mm": (20.0, 55.0),
    "qr_side_margin_mm": (3.0, 40.0),
    "qr_bottom_margin_mm": (5.0, 60.0),
}


def normalize_profile(value: Any) -> dict[str, Any]:
    """Coerce persisted or UI-supplied profile values to safe layout bounds."""
    result = default_profile()
    if isinstance(value, dict):
        result.update({key: value[key] for key in result if key in value})
    for key, (minimum, maximum) in _LIMITS.items():
        try:
            number = float(result[key])
        except (TypeError, ValueError):
            number = float(default_profile()[key])
        result[key] = max(minimum, min(maximum, number))
    if result.get("qr_position") not in {"Left", "Center", "Right"}:
        result["qr_position"] = "Left"
    result["preview_background_path"] = str(
        result.get("preview_background_path", "") or "")
    return result


def normalize_profiles(value: Any) -> dict[str, dict[str, Any]]:
    profiles = {}
    if isinstance(value, dict):
        for name, profile in value.items():
            clean_name = str(name or "").strip()
            if clean_name:
                profiles[clean_name] = normalize_profile(profile)
    profiles.setdefault(DEFAULT_PROFILE_NAME, default_profile())
    return profiles


def normalize_presets(value: Any) -> dict[str, dict[str, Any]]:
    """Normalize reusable named layouts without adding a printer default."""
    presets = {}
    if isinstance(value, dict):
        for name, profile in value.items():
            clean_name = str(name or "").strip()
            if clean_name:
                presets[clean_name] = normalize_profile(profile)
    return presets


def active_profile(document_defaults: Any) -> dict[str, Any]:
    """Resolve the selected printer profile from a document settings snapshot."""
    if not isinstance(document_defaults, dict):
        return default_profile()
    profiles = normalize_profiles(document_defaults.get("printer_profiles"))
    selected = str(document_defaults.get(
        "selected_printer", DEFAULT_PROFILE_NAME) or DEFAULT_PROFILE_NAME)
    return deepcopy(profiles.get(selected, profiles[DEFAULT_PROFILE_NAME]))


def installed_printers() -> list[str]:
    """Return Windows printer names without requiring pywin32.

    Registry lookup is intentionally best-effort so the app remains usable in
    restricted Windows sessions and in tests on non-Windows hosts.
    """
    names: set[str] = set()
    try:
        import winreg
        locations = (
            (winreg.HKEY_CURRENT_USER,
             r"Software\Microsoft\Windows NT\CurrentVersion\Devices"),
            (winreg.HKEY_LOCAL_MACHINE,
             r"SYSTEM\CurrentControlSet\Control\Print\Printers"),
        )
        for hive, path in locations:
            try:
                with winreg.OpenKey(hive, path) as key:
                    index = 0
                    while True:
                        try:
                            if hive == winreg.HKEY_LOCAL_MACHINE:
                                names.add(winreg.EnumKey(key, index))
                            else:
                                names.add(winreg.EnumValue(key, index)[0])
                            index += 1
                        except OSError:
                            break
            except OSError:
                continue
    except (ImportError, OSError):
        pass
    return sorted((name for name in names if name.strip()), key=str.casefold)
