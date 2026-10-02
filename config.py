"""Application configuration, local encryption, and profile persistence."""
from __future__ import annotations

import json
import hashlib
import os
import shutil
import sys
import threading
import tempfile
import uuid
import zipfile
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict

from security import DataProtectionError, protect, unprotect
from print_layout import (
    DEFAULT_PROFILE_NAME,
    default_profile,
    normalize_presets,
    normalize_profiles,
)

APP_VERSION = "8.0"
RECOVERY_RETENTION_DAYS = 30
# ISO portrait sizes in points: A5 is exactly 148 × 210 mm, A4 210 × 297 mm.
PAPER_SIZES: Dict[str, tuple[float, float]] = {
    "A5": (419.528, 595.276), "A4": (595.276, 841.889),
}
DEFAULT_PAPER = "A4"
DEFAULT_VIEWER_BASE = "https://ahmedinternist.github.io/rx-viewer/"
APP_DIR = Path(os.environ.get("RX_APP_DATA_DIR") or os.environ.get(
    "APPDATA", str(Path.home() / ".prescription_app"))) / "prescription_app"
CONFIG_PATH = APP_DIR / "config.json"
DATA_DIR = APP_DIR / "data"
DEFAULT_DB_PATH = DATA_DIR / "drugs.csv"
LOG_PATH = APP_DIR / "rx-prescription.log"
BACKUP_FORMAT = "rx-prescription-backup-v1"
TEMPLATE_BACKUP_FORMAT = "rx-treatment-templates-v1"
PROJECT_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
SEED_DB_PATH = PROJECT_DIR / "data" / "drugs.csv"

APP_DIR.mkdir(parents=True, exist_ok=True)
DATA_DIR.mkdir(parents=True, exist_ok=True)


def ensure_seed_db() -> None:
    if not DEFAULT_DB_PATH.exists() and SEED_DB_PATH.exists():
        shutil.copyfile(SEED_DB_PATH, DEFAULT_DB_PATH)


def _doctor() -> Dict[str, str]:
    return {"name": "", "license_no": "", "specialty": ""}


def _replace_bytes(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=path.parent, prefix=path.name + ".",
                suffix=".tmp", delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def _default_config() -> Dict[str, Any]:
    doctor = _doctor()
    return {
        "schema": 2,
        "paper_size": DEFAULT_PAPER,
        "language": "en",
        "ui_fonts": {"medication_font_size": 20, "instruction_font_size": 18, "name_font_size": 18},
        "viewer_base_url": DEFAULT_VIEWER_BASE,
        "cloud_rx_api_key": "",
        "cloud_last_upload_at": "",
        "drug_db_path": str(DEFAULT_DB_PATH),
        "clinic": {"name": "", "address": "", "phone": "", "website": "", "logo_path": "",
                   "latitude": "", "longitude": "", "include_location": False},
        "doctor": doctor,
        "profiles": {"Default": doctor.copy()},
        "active_profile": "Default",
        "signing_private_key": "",
        "medication_favorites": [],
        "treatment_templates": [],
        "recovery_bin": [],
        "favorite_therapeutic_groups": [],
        "dosage_presets": [],
        "openfda_cache": {},
        "prescription_draft": {},
        "document_defaults": {
            "language": "interface", "show_header": True,
            "logo_size": "medium", "margin_mm": 16,
            "font_size": 10, "export_folder": "",
            "selected_printer": DEFAULT_PROFILE_NAME,
            "printer_profiles": {DEFAULT_PROFILE_NAME: default_profile()},
            "calibration_presets": {},
        },
        "auto_backup_enabled": False,
        "last_backup_at": "",
        "last_backup_path": "",
        "drug_db_imported_at": "",
    }


def _validated_config(data):
    """Repair malformed sections without discarding unrelated valid values."""
    defaults = _default_config()
    clean = dict(data)
    for key, default in defaults.items():
        value = clean.get(key, default)
        if isinstance(default, dict):
            value = dict(value) if isinstance(value, dict) else dict(default)
            for field, fallback in default.items():
                current = value.get(field, fallback)
                if isinstance(fallback, dict) and not isinstance(current, dict):
                    current = fallback
                elif isinstance(fallback, str) and not isinstance(current, str):
                    current = fallback
                elif isinstance(fallback, bool) and not isinstance(current, bool):
                    current = fallback
                elif isinstance(fallback, int) and (not isinstance(current, int) or isinstance(current, bool)):
                    current = fallback
                value[field] = current
        elif isinstance(default, list):
            value = value if isinstance(value, list) else []
            element_type = str if key in {"dosage_presets", "favorite_therapeutic_groups"} else dict
            value = [item for item in value if isinstance(item, element_type)]
        elif isinstance(default, str) and not isinstance(value, str):
            value = default
        elif isinstance(default, bool) and not isinstance(value, bool):
            value = default
        clean[key] = value
    clean["profiles"] = {name: profile for name, profile in clean["profiles"].items()
                         if isinstance(name, str) and isinstance(profile, dict)}
    clean["profiles"].setdefault("Default", clean["doctor"].copy())
    if clean["language"] not in {"en", "ar"}:
        clean["language"] = "en"
    return clean


class Config:
    def __init__(self) -> None:
        self.data: Dict[str, Any] = _default_config()
        self.recovered_unreadable_settings = False
        self.unreadable_settings_backup = ""
        self._save_lock = threading.RLock()
        self._save_batch_depth = 0
        self._save_pending = False
        self._last_saved_digest = None
        self._last_saved_signature = None
        self.load()

    def load(self) -> None:
        if CONFIG_PATH.exists():
            try:
                raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                if not isinstance(raw, dict):
                    raise ValueError("Configuration root must be an object")
                if raw.get("format") == "dpapi-v1":
                    raw = json.loads(unprotect(raw["data"]).decode("utf-8"))
                if not isinstance(raw, dict):
                    raise ValueError("Decrypted configuration must be an object")
                self.data.update(_validated_config(raw))
            except (OSError, ValueError, TypeError, KeyError, DataProtectionError):
                # DPAPI data is intentionally bound to the Windows user that
                # created it. A config copied from another PC/account therefore
                # cannot be decrypted. Preserve the unreadable file for support,
                # then create a fresh local config instead of repeatedly loading
                # an unusable envelope on every launch.
                self.recovered_unreadable_settings = True
                self.unreadable_settings_backup = self._preserve_unreadable_config()
                try:
                    self._write_config()
                except OSError:
                    # A read-only profile must not prevent the application from
                    # opening. Saving a setting later will report the normal I/O
                    # failure while this session continues with safe defaults.
                    pass
        self.data.setdefault("clinic", {})
        self.data["clinic"].setdefault("website", "")
        self.data.setdefault("doctor", _doctor())
        self.data.setdefault("profiles", {"Default": self.data["doctor"].copy()})
        self.data.setdefault("active_profile", "Default")
        self.data.setdefault("signing_private_key", "")
        self.data.setdefault("medication_favorites", [])
        self.data.setdefault("treatment_templates", [])
        self.data.setdefault("recovery_bin", [])
        self.data.setdefault("favorite_therapeutic_groups", [])
        self.data.setdefault("dosage_presets", [])
        removed_obsolete_reference_settings = False
        for key in (
                "gemini_enabled", "gemini_api_key", "gemini_last_test",
                "clinical_ai_enabled", "openai_api_key", "openai_last_test"):
            removed_obsolete_reference_settings |= self.data.pop(key, None) is not None
        self.data.setdefault("openfda_cache", {})
        self.data.setdefault("prescription_draft", {})
        self.data.setdefault("document_defaults", {
            "language": "interface", "show_header": True,
            "logo_size": "medium", "margin_mm": 16,
            "font_size": 10, "export_folder": "",
            "selected_printer": DEFAULT_PROFILE_NAME,
            "printer_profiles": {DEFAULT_PROFILE_NAME: default_profile()},
            "calibration_presets": {},
        })
        self.data["document_defaults"].setdefault("font_size", 10)
        self.data["document_defaults"].setdefault(
            "selected_printer", DEFAULT_PROFILE_NAME)
        self.data["document_defaults"]["printer_profiles"] = normalize_profiles(
            self.data["document_defaults"].get("printer_profiles"))
        self.data["document_defaults"]["calibration_presets"] = normalize_presets(
            self.data["document_defaults"].get("calibration_presets"))
        self.data.setdefault("auto_backup_enabled", False)
        self.data.setdefault("last_backup_at", "")
        self.data.setdefault("last_backup_path", "")
        self.data.setdefault("drug_db_imported_at", "")
        if self.data.get("paper_size") not in PAPER_SIZES:
            self.data["paper_size"] = DEFAULT_PAPER
        if removed_obsolete_reference_settings:
            try:
                self._write_config()
            except OSError:
                pass
        for cache_name in ("clinical-reference-cache.json", "gemini-cache.json"):
            try:
                (APP_DIR / cache_name).unlink(missing_ok=True)
            except OSError:
                pass

    @staticmethod
    def _preserve_unreadable_config() -> str:
        """Copy an unreadable settings envelope aside before regenerating it."""
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        candidate = CONFIG_PATH.with_name(f"config-unreadable-{stamp}.json")
        suffix = 1
        while candidate.exists():
            candidate = CONFIG_PATH.with_name(
                f"config-unreadable-{stamp}-{suffix}.json")
            suffix += 1
        try:
            shutil.copy2(CONFIG_PATH, candidate)
            return str(candidate)
        except OSError:
            return ""

    def save(self) -> None:
        with self._save_lock:
            if self._save_batch_depth:
                self._save_pending = True
                return
            self._write_config()

    def _write_config(self) -> None:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(self.data, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        digest = hashlib.sha256(payload).digest()
        try:
            stat = CONFIG_PATH.stat()
            signature = stat.st_ino, stat.st_mtime_ns, stat.st_size
        except OSError:
            signature = None
        if digest == self._last_saved_digest and signature == self._last_saved_signature:
            return
        protected = {"format": "dpapi-v1", "data": protect(payload)}
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8",
                    dir=CONFIG_PATH.parent, prefix=CONFIG_PATH.name + ".", suffix=".tmp",
                    delete=False) as stream:
                temporary = Path(stream.name)
                json.dump(protected, stream)
                stream.flush()
                os.fsync(stream.fileno())
                stat = os.fstat(stream.fileno())
            os.replace(temporary, CONFIG_PATH)
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)
        self._last_saved_digest = digest
        self._last_saved_signature = stat.st_ino, stat.st_mtime_ns, stat.st_size

    @contextmanager
    def batch_save(self):
        """Combine several related setting changes into one encrypted write."""
        with self._save_lock:
            self._save_batch_depth += 1
            try:
                yield self
            finally:
                self._save_batch_depth -= 1
                if self._save_batch_depth == 0 and self._save_pending:
                    self._save_pending = False
                    self._write_config()

    def get(self, key: str, default=None):
        return self.data.get(key, default)

    @property
    def cloud_rx_api_key(self) -> str:
        return str(self.data.get("cloud_rx_api_key", ""))

    @cloud_rx_api_key.setter
    def cloud_rx_api_key(self, value: str) -> None:
        self.set("cloud_rx_api_key", str(value or "").strip())

    def set(self, key: str, value: Any) -> None:
        self.data[key] = value
        self.save()

    # -- recoverable deletion ---------------------------------------------
    def recovery_items(self) -> list[Dict[str, Any]]:
        """Return non-expired deleted items, newest first."""
        now = datetime.now(timezone.utc)
        kept = []
        for item in self.data.get("recovery_bin", []):
            if not isinstance(item, dict):
                continue
            try:
                expires = datetime.fromisoformat(str(item.get("expires_at", "")))
                if expires.tzinfo is None:
                    expires = expires.replace(tzinfo=timezone.utc)
            except ValueError:
                continue
            if expires > now:
                kept.append(item)
        kept.sort(key=lambda item: str(item.get("deleted_at", "")), reverse=True)
        if kept != self.data.get("recovery_bin", []):
            self.data["recovery_bin"] = kept
            self.save()
        return [dict(item) for item in kept]

    def add_recovery_item(self, kind: str, label: str, payload: Any) -> str:
        now = datetime.now(timezone.utc)
        item = {
            "id": uuid.uuid4().hex,
            "kind": str(kind),
            "label": str(label).strip() or str(kind),
            "deleted_at": now.isoformat(timespec="seconds"),
            "expires_at": (now + timedelta(days=RECOVERY_RETENTION_DAYS)).isoformat(
                timespec="seconds"),
            "payload": payload,
        }
        items = self.recovery_items()
        items.insert(0, item)
        self.data["recovery_bin"] = items[:250]
        self.save()
        return item["id"]

    def discard_recovery_item(self, item_id: str) -> bool:
        items = self.recovery_items()
        kept = [item for item in items if item.get("id") != item_id]
        if len(kept) == len(items):
            return False
        self.data["recovery_bin"] = kept
        self.save()
        return True

    @property
    def paper_size(self) -> str:
        value = self.data.get("paper_size", DEFAULT_PAPER)
        return value if value in PAPER_SIZES else DEFAULT_PAPER

    @paper_size.setter
    def paper_size(self, value: str) -> None:
        if value not in PAPER_SIZES:
            raise ValueError(f"Unknown paper size: {value}")
        self.set("paper_size", value)

    def ui_font_size(self, key: str) -> int:
        aliases = {"dropdown_font_size": "medication_font_size", "patient_name_font_size": "name_font_size"}
        key = aliases.get(key, key)
        defaults = {"medication_font_size": 20, "instruction_font_size": 18, "name_font_size": 18}
        default = defaults.get(key, 0)
        fonts = self.data.get("ui_fonts", {})
        try:
            legacy = {"medication_font_size": "dropdown_font_size", "name_font_size": "patient_name_font_size"}
            value = int(fonts.get(key, fonts.get(legacy.get(key), default)))
        except (TypeError, ValueError, AttributeError):
            return default
        if 10 <= value < 18:
            return 18  # Preserve older preferences within the new supported range.
        return value if 18 <= value <= 56 else default

    def set_ui_font_sizes(self, dropdown: int, patient: int, instructions: int = 18) -> None:
        if any(isinstance(size, bool) or not isinstance(size, int) or not 18 <= size <= 56
               for size in (dropdown, patient, instructions)):
            raise ValueError("Display font sizes must be 18–56")
        self.set("ui_fonts", {"medication_font_size": dropdown, "name_font_size": patient,
                              "instruction_font_size": instructions})

    @property
    def language(self) -> str:
        return self.data.get("language", "en")

    @language.setter
    def language(self, value: str) -> None:
        self.set("language", value if value in ("en", "ar") else "en")

    @property
    def viewer_base_url(self) -> str:
        """Deprecated legacy viewer setting; retained for old backups only."""
        return self.data.get("viewer_base_url", DEFAULT_VIEWER_BASE)

    @viewer_base_url.setter
    def viewer_base_url(self, value: str) -> None:
        self.set("viewer_base_url", value)

    @property
    def drug_db_path(self) -> str:
        return self.data.get("drug_db_path", str(DEFAULT_DB_PATH))

    @drug_db_path.setter
    def drug_db_path(self, value: str) -> None:
        self.set("drug_db_path", value)

    @property
    def signing_private_key(self) -> str:
        return self.data.get("signing_private_key", "")

    @signing_private_key.setter
    def signing_private_key(self, value: str) -> None:
        self.set("signing_private_key", value)

    def get_doctor(self) -> Dict[str, Any]:
        return self.data.get("doctor", {}).copy()

    def set_doctor(self, **kw) -> None:
        self.data.setdefault("doctor", _doctor()).update(kw)
        active = self.data.get("active_profile", "Default")
        self.data.setdefault("profiles", {})[active] = self.data["doctor"].copy()
        self.save()

    def profile_names(self) -> list[str]:
        return sorted(self.data.get("profiles", {}).keys())

    def use_profile(self, name: str) -> Dict[str, Any]:
        profile = self.data.get("profiles", {}).get(name)
        if profile is None:
            raise KeyError(name)
        self.data["active_profile"] = name
        self.data["doctor"] = profile.copy()
        self.save()
        return self.get_doctor()

    def save_profile(self, name: str, doctor: Dict[str, str]) -> None:
        name = name.strip() or "Default"
        self.data.setdefault("profiles", {})[name] = doctor.copy()
        self.data["active_profile"] = name
        self.data["doctor"] = doctor.copy()
        self.save()

    def get_clinic(self) -> Dict[str, Any]:
        return self.data.get("clinic", {}).copy()

    def set_clinic(self, **kw) -> None:
        self.data.setdefault("clinic", {}).update(kw)
        self.save()

    # -- medication helpers -------------------------------------------------
    @staticmethod
    def _clean_medication_favorite(item: Dict[str, Any], existing=None) -> Dict[str, Any]:
        existing = existing or {}
        text_fields = (
            "generic_name", "brand_name", "category", "regimen_name",
            "dosage", "frequency", "duration", "notes",
        )
        cleaned = {
            field: str(item.get(field, existing.get(field, ""))).strip()
            for field in text_fields
        }
        cleaned["pinned"] = bool(item.get("pinned", existing.get("pinned", False)))
        try:
            cleaned["use_count"] = max(
                0, int(item.get("use_count", existing.get("use_count", 0))))
        except (TypeError, ValueError):
            cleaned["use_count"] = 0
        cleaned["last_used"] = str(
            item.get("last_used", existing.get("last_used", ""))).strip()
        cleaned["id"] = str(
            item.get("id", existing.get("id", ""))).strip() or uuid.uuid4().hex
        return cleaned

    def medication_favorites(self) -> list[Dict[str, Any]]:
        raw_favorites = self.data.get("medication_favorites", [])
        cleaned = [self._clean_medication_favorite(item)
                   for item in raw_favorites
                   if isinstance(item, dict)
                   and (item.get("generic_name") or item.get("brand_name"))]
        # One-time, backward-compatible migration.  Stable IDs let the UI reuse
        # cards safely when favorites are filtered, starred, edited, or deleted.
        if (len(cleaned) != len(raw_favorites)
                or any(not str(item.get("id", "")).strip()
                       for item in raw_favorites if isinstance(item, dict))):
            self.data["medication_favorites"] = cleaned
            self.save()
        return cleaned

    def add_medication_favorite(self, item: Dict[str, str]) -> bool:
        cleaned = self._clean_medication_favorite(item)
        if not (cleaned["generic_name"] or cleaned["brand_name"]):
            return False
        favorites = self.medication_favorites()
        comparable_fields = tuple(key for key in cleaned if key != "id")
        if not any(all(existing.get(key) == cleaned.get(key)
                       for key in comparable_fields) for existing in favorites):
            favorites.append(cleaned)
            self.data["medication_favorites"] = favorites
            self.save()
        return True

    def update_medication_favorite(self, index: int, item: Dict[str, str]) -> bool:
        favorites = self.medication_favorites()
        if not 0 <= index < len(favorites):
            return False
        cleaned = self._clean_medication_favorite(item, favorites[index])
        if not (cleaned["generic_name"] or cleaned["brand_name"]):
            return False
        favorites[index] = cleaned
        self.data["medication_favorites"] = favorites
        self.save()
        return True

    def insert_medication_favorite(self, index: int, item: Dict[str, Any]) -> bool:
        cleaned = self._clean_medication_favorite(item)
        if not (cleaned["generic_name"] or cleaned["brand_name"]):
            return False
        favorites = self.medication_favorites()
        favorites.insert(max(0, min(index, len(favorites))), cleaned)
        self.data["medication_favorites"] = favorites
        self.save()
        return True

    def toggle_medication_favorite_pin(self, index: int) -> bool:
        favorites = self.medication_favorites()
        if not 0 <= index < len(favorites):
            return False
        favorites[index]["pinned"] = not favorites[index].get("pinned", False)
        self.data["medication_favorites"] = favorites
        self.save()
        return bool(favorites[index]["pinned"])

    def record_medication_favorite_use(self, index: int) -> bool:
        favorites = self.medication_favorites()
        if not 0 <= index < len(favorites):
            return False
        favorites[index]["use_count"] = int(favorites[index].get("use_count", 0)) + 1
        favorites[index]["last_used"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
        self.data["medication_favorites"] = favorites
        self.save()
        return True

    def remove_medication_favorite(self, index: int) -> bool:
        favorites = self.medication_favorites()
        if not 0 <= index < len(favorites):
            return False
        favorites.pop(index)
        self.data["medication_favorites"] = favorites
        self.save()
        return True

    # -- treatment templates ----------------------------------------------
    @staticmethod
    def _clean_treatment_template(item: Dict[str, Any], existing=None,
                                  touch_updated: bool = False) -> Dict[str, Any]:
        """Normalize one disease template before encrypted local storage."""
        existing = existing or {}
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        medications = []
        raw_medications = item.get("medications", existing.get("medications", []))
        for raw in raw_medications if isinstance(raw_medications, list) else []:
            if not isinstance(raw, dict):
                continue
            medicine = {
                key: str(raw.get(key, "")).strip()
                for key in ("generic_name", "brand_name", "dosage", "frequency",
                            "duration", "notes", "quantity")
            }
            medicine["alternative_to_previous"] = bool(
                raw.get("alternative_to_previous", False)) and bool(medications)
            if medicine["generic_name"] or medicine["brand_name"]:
                medications.append(medicine)
        try:
            use_count = max(0, int(item.get(
                "use_count", existing.get("use_count", 0)) or 0))
        except (TypeError, ValueError):
            use_count = 0
        created_at = str(item.get(
            "created_at", existing.get("created_at", ""))).strip()
        updated_at = str(item.get(
            "updated_at", existing.get("updated_at", ""))).strip()
        return {
            "id": str(item.get("id", existing.get("id", ""))).strip() or uuid.uuid4().hex,
            "disease": str(item.get("disease", existing.get("disease", ""))).strip(),
            "category": str(item.get(
                "category", existing.get("category", ""))).strip(),
            "variant": str(item.get("variant", existing.get("variant", ""))).strip(),
            "medications": medications,
            "created_at": created_at or now,
            "updated_at": now if touch_updated else (updated_at or created_at or now),
            "last_used": str(item.get(
                "last_used", existing.get("last_used", ""))).strip(),
            "use_count": use_count,
        }

    def treatment_templates(self) -> list[Dict[str, Any]]:
        raw_templates = self.data.get("treatment_templates", [])
        return [self._clean_treatment_template(item, item)
                for item in raw_templates if isinstance(item, dict)
                and str(item.get("disease", "")).strip()]

    def save_treatment_template(self, item: Dict[str, Any]) -> str:
        templates = self.treatment_templates()
        template_id = str(item.get("id", "")).strip()
        previous = next((entry for entry in templates
                         if entry["id"] == template_id), None)
        cleaned = self._clean_treatment_template(
            item, previous or item, touch_updated=True)
        if not cleaned["disease"] or not cleaned["medications"]:
            return ""
        for index, existing in enumerate(templates):
            if existing["id"] == cleaned["id"]:
                templates[index] = cleaned
                break
        else:
            templates.append(cleaned)
        self.data["treatment_templates"] = templates
        self.save()
        return cleaned["id"]

    def merge_treatment_templates(self, items, replace: bool = False) -> int:
        """Import editable templates, updating matching diseases without duplicates."""
        imported = [self._clean_treatment_template(item, item, touch_updated=True)
                    for item in items if isinstance(item, dict)]
        imported = [item for item in imported
                    if item["disease"] and item["medications"]]
        if replace:
            merged = imported
        else:
            merged = self.treatment_templates()
            positions = {
                item["disease"].strip().casefold(): index
                for index, item in enumerate(merged)
            }
            for item in imported:
                key = item["disease"].strip().casefold()
                if key in positions:
                    index = positions[key]
                    item["id"] = merged[index]["id"]
                    item["created_at"] = merged[index].get("created_at", "")
                    item["last_used"] = merged[index].get("last_used", "")
                    item["use_count"] = merged[index].get("use_count", 0)
                    merged[index] = self._clean_treatment_template(
                        item, merged[index], touch_updated=True)
                else:
                    positions[key] = len(merged)
                    merged.append(item)
        self.data["treatment_templates"] = merged
        self.save()
        return len(imported)

    def record_treatment_template_use(self, template_id: str) -> bool:
        """Record a successful application without changing clinical content."""
        templates = self.treatment_templates()
        for template in templates:
            if template["id"] != template_id:
                continue
            template["use_count"] = int(template.get("use_count", 0)) + 1
            template["last_used"] = datetime.now(timezone.utc).isoformat(
                timespec="seconds")
            self.data["treatment_templates"] = templates
            self.save()
            return True
        return False

    def remove_treatment_template(self, template_id: str) -> bool:
        templates = self.treatment_templates()
        remaining = [item for item in templates if item["id"] != template_id]
        if len(remaining) == len(templates):
            return False
        self.data["treatment_templates"] = remaining
        self.save()
        return True

    def export_treatment_templates(self, destination: str) -> str:
        """Write a Windows-encrypted backup containing treatment templates only."""
        target = Path(destination)
        target.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "format": TEMPLATE_BACKUP_FORMAT,
            "version": 1,
            "templates": self.treatment_templates(),
        }
        protected = {
            "format": "dpapi-v1",
            "kind": TEMPLATE_BACKUP_FORMAT,
            "data": protect(json.dumps(
                payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")),
        }
        temporary = target.with_suffix(target.suffix + ".tmp")
        temporary.write_text(json.dumps(protected), encoding="utf-8")
        temporary.replace(target)
        return str(target)

    def import_treatment_templates(self, source: str, replace: bool = False) -> int:
        """Restore a template-only backup, either replacing or merging local templates."""
        protected = json.loads(Path(source).read_text(encoding="utf-8"))
        if (protected.get("format") != "dpapi-v1"
                or protected.get("kind") != TEMPLATE_BACKUP_FORMAT):
            raise ValueError("This is not a valid encrypted treatment-template backup.")
        payload = json.loads(unprotect(protected["data"]).decode("utf-8"))
        if payload.get("format") != TEMPLATE_BACKUP_FORMAT:
            raise ValueError("This treatment-template backup is not supported.")
        raw_templates = payload.get("templates", [])
        if not isinstance(raw_templates, list):
            raise ValueError("The treatment-template backup is invalid.")
        imported = [self._clean_treatment_template(item, item)
                    for item in raw_templates if isinstance(item, dict)]
        imported = [item for item in imported
                    if item["disease"] and item["medications"]]
        if replace:
            merged = imported
        else:
            merged = self.treatment_templates()
            positions = {item["id"]: index for index, item in enumerate(merged)}
            for item in imported:
                if item["id"] in positions:
                    merged[positions[item["id"]]] = item
                else:
                    positions[item["id"]] = len(merged)
                    merged.append(item)
        self.data["treatment_templates"] = merged
        self.save()
        return len(imported)

    def favorite_therapeutic_groups(self) -> list[str]:
        return [str(code) for code in self.data.get("favorite_therapeutic_groups", [])
                if isinstance(code, str) and code.strip()]

    def toggle_favorite_therapeutic_group(self, code: str) -> bool:
        """Add or remove a major therapeutic group while preserving pin order."""
        code = code.strip()
        groups = self.favorite_therapeutic_groups()
        if code in groups:
            groups.remove(code)
            selected = False
        else:
            groups.append(code)
            selected = True
        self.data["favorite_therapeutic_groups"] = groups
        self.save()
        return selected

    def dosage_presets(self) -> list[str]:
        defaults = ["1 tablet", "1 capsule", "5 mL", "500 mg", "1 puff", "1 injection"]
        saved = [str(value).strip() for value in self.data.get("dosage_presets", []) if str(value).strip()]
        return list(dict.fromkeys(defaults + saved))

    def add_dosage_preset(self, value: str) -> bool:
        value = value.strip()
        if not value:
            return False
        saved = [str(item).strip() for item in self.data.get("dosage_presets", []) if str(item).strip()]
        if value not in saved:
            saved.append(value)
            self.data["dosage_presets"] = saved
            self.save()
        return True

    # -- backup / restore ---------------------------------------------------
    def create_backup(self, destination: str) -> str:
        """Create a portable archive of encrypted settings and the active drug DB."""
        self.save()
        target = Path(destination)
        target.parent.mkdir(parents=True, exist_ok=True)
        database = Path(self.drug_db_path)
        from patient_history import PatientHistory
        history_data = PatientHistory().backup_bytes()
        manifest = {"format": BACKUP_FORMAT, "version": 1, "has_drug_database": database.is_file()}
        with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("manifest.json", json.dumps(manifest, separators=(",", ":")))
            archive.writestr("config.json", CONFIG_PATH.read_bytes())
            if database.is_file():
                archive.write(database, "drug_database.csv")
            if history_data is not None:
                archive.writestr("patient_history.json", history_data)
        self.data["last_backup_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
        self.data["last_backup_path"] = str(target)
        self.save()
        return str(target)

    def maybe_create_automatic_backup(self) -> str:
        """Create at most one automatic backup per UTC day and retain the newest ten."""
        if not self.data.get("auto_backup_enabled"):
            return ""
        today = datetime.now(timezone.utc).date().isoformat()
        if str(self.data.get("last_backup_at", "")).startswith(today):
            return ""
        backup_dir = APP_DIR / "backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        destination = backup_dir / f"rx-backup-{today}.rxbackup"
        result = self.create_backup(str(destination))
        backups = sorted(
            backup_dir.glob("rx-backup-*.rxbackup"),
            key=lambda item: item.stat().st_mtime, reverse=True)
        for old_backup in backups[10:]:
            old_backup.unlink(missing_ok=True)
        return result

    def restore_backup(self, source: str) -> None:
        """Restore settings and database after validating the backup archive."""
        with zipfile.ZipFile(source, "r") as archive:
            names = set(archive.namelist())
            if not {"manifest.json", "config.json"}.issubset(names):
                raise ValueError("This is not a valid prescription backup.")
            manifest = json.loads(archive.read("manifest.json").decode("utf-8"))
            if manifest.get("format") != BACKUP_FORMAT:
                raise ValueError("This backup was created by an unsupported app version.")
            protected = json.loads(archive.read("config.json").decode("utf-8"))
            if protected.get("format") != "dpapi-v1":
                raise ValueError("The settings backup is not protected correctly.")
            restored = json.loads(unprotect(protected["data"]).decode("utf-8"))
            db_data = archive.read("drug_database.csv") if "drug_database.csv" in names else None
            history_data = archive.read("patient_history.json") if "patient_history.json" in names else None

        if not isinstance(restored, dict):
            raise ValueError("The settings backup is invalid.")
        from patient_history import HISTORY_PATH, PatientHistory, validate_history_bytes
        if history_data is not None:
            validate_history_bytes(history_data)
        replacements = []
        if db_data is not None:
            replacements.append((DEFAULT_DB_PATH, db_data))
            restored["drug_db_path"] = str(DEFAULT_DB_PATH)
        if history_data is not None:
            replacements.append((HISTORY_PATH, history_data))
        new_data = _default_config()
        new_data.update(_validated_config(restored))
        if not isinstance(new_data.get("clinic"), dict) or not isinstance(new_data.get("doctor"), dict):
            raise ValueError("The settings backup is invalid.")
        replacements.append((CONFIG_PATH, json.dumps({"format": "dpapi-v1", "data": protect(
            json.dumps(new_data, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))}).encode("utf-8")))
        from drug_db import database_lock
        with self._save_lock, database_lock(DEFAULT_DB_PATH), PatientHistory()._lock:
            originals = {path: path.read_bytes() if path.exists() else None for path, _ in replacements}
            changed = []
            try:
                for path, contents in replacements:
                    _replace_bytes(path, contents)
                    changed.append(path)
            except Exception:
                for path in reversed(changed):
                    if originals[path] is None:
                        path.unlink(missing_ok=True)
                    else:
                        _replace_bytes(path, originals[path])
                raise
            self.data = new_data
            self._last_saved_digest = None
            self._last_saved_signature = None


config = Config()
