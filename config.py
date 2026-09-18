"""Application configuration, local encryption, and profile persistence."""
from __future__ import annotations

import json
import os
import shutil
import sys
import threading
import uuid
import zipfile
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict

from security import DataProtectionError, protect, unprotect

APP_VERSION = "4.82.0"
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


def _default_config() -> Dict[str, Any]:
    doctor = _doctor()
    return {
        "schema": 2,
        "paper_size": DEFAULT_PAPER,
        "language": "en",
        "ui_fonts": {"dropdown_font_size": 0, "patient_name_font_size": 14},
        "viewer_base_url": DEFAULT_VIEWER_BASE,
        "cloud_rx_api_key": "",
        "drug_db_path": str(DEFAULT_DB_PATH),
        "clinic": {"name": "", "address": "", "phone": "", "logo_path": ""},
        "doctor": doctor,
        "profiles": {"Default": doctor.copy()},
        "active_profile": "Default",
        "signing_private_key": "",
        "medication_favorites": [],
        "treatment_templates": [],
        "recovery_bin": [],
        "favorite_therapeutic_groups": [],
        "dosage_presets": [],
        "gemini_enabled": False,
        "gemini_api_key": "",
        "gemini_last_test": "",
        "openfda_cache": {},
        "document_defaults": {
            "language": "interface", "show_header": True,
            "logo_size": "medium", "margin_mm": 16,
            "export_folder": "",
        },
        "auto_backup_enabled": False,
        "last_backup_at": "",
        "last_backup_path": "",
        "drug_db_imported_at": "",
    }


class Config:
    def __init__(self) -> None:
        self.data: Dict[str, Any] = _default_config()
        self._save_lock = threading.RLock()
        self._save_batch_depth = 0
        self._save_pending = False
        self.load()

    def load(self) -> None:
        if CONFIG_PATH.exists():
            try:
                raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                if raw.get("format") == "dpapi-v1":
                    raw = json.loads(unprotect(raw["data"]).decode("utf-8"))
                self.data.update(raw)
            except (OSError, ValueError, KeyError, DataProtectionError):
                pass
        self.data.setdefault("clinic", {})
        self.data.setdefault("doctor", _doctor())
        self.data.setdefault("profiles", {"Default": self.data["doctor"].copy()})
        self.data.setdefault("active_profile", "Default")
        self.data.setdefault("signing_private_key", "")
        self.data.setdefault("medication_favorites", [])
        self.data.setdefault("treatment_templates", [])
        self.data.setdefault("recovery_bin", [])
        self.data.setdefault("favorite_therapeutic_groups", [])
        self.data.setdefault("dosage_presets", [])
        self.data.setdefault("gemini_enabled", False)
        self.data.setdefault("gemini_api_key", "")
        self.data.setdefault("gemini_last_test", "")
        self.data.setdefault("openfda_cache", {})
        self.data.setdefault("document_defaults", {
            "language": "interface", "show_header": True,
            "logo_size": "medium", "margin_mm": 16,
            "export_folder": "",
        })
        self.data.setdefault("auto_backup_enabled", False)
        self.data.setdefault("last_backup_at", "")
        self.data.setdefault("last_backup_path", "")
        self.data.setdefault("drug_db_imported_at", "")
        if self.data.get("paper_size") not in PAPER_SIZES:
            self.data["paper_size"] = DEFAULT_PAPER

    def save(self) -> None:
        with self._save_lock:
            if self._save_batch_depth:
                self._save_pending = True
                return
            self._write_config()

    def _write_config(self) -> None:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(self.data, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        protected = {"format": "dpapi-v1", "data": protect(payload)}
        temporary = CONFIG_PATH.with_suffix(".tmp")
        temporary.write_text(json.dumps(protected), encoding="utf-8")
        temporary.replace(CONFIG_PATH)

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
        default = 14 if key == "patient_name_font_size" else 0
        fonts = self.data.get("ui_fonts", {})
        try:
            value = int(fonts.get(key, default))
        except (TypeError, ValueError, AttributeError):
            return default
        if key == "dropdown_font_size" and value == 0:
            return 0
        return value if 10 <= value <= 56 else default

    def set_ui_font_sizes(self, dropdown: int, patient: int) -> None:
        if dropdown != 0 and not 10 <= dropdown <= 56:
            raise ValueError("Dropdown font size must be 10–56 or Default")
        if not 10 <= patient <= 56:
            raise ValueError("Patient name font size must be 10–56")
        self.set("ui_fonts", {"dropdown_font_size": dropdown, "patient_name_font_size": patient})

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
    def _clean_treatment_template(item: Dict[str, Any], existing=None) -> Dict[str, Any]:
        """Normalize one disease template before encrypted local storage."""
        existing = existing or {}
        medications = []
        for raw in item.get("medications", existing.get("medications", [])):
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
        return {
            "id": str(item.get("id", existing.get("id", ""))).strip() or uuid.uuid4().hex,
            "disease": str(item.get("disease", existing.get("disease", ""))).strip(),
            "variant": str(item.get("variant", existing.get("variant", ""))).strip(),
            "medications": medications,
            "updated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        }

    def treatment_templates(self) -> list[Dict[str, Any]]:
        raw_templates = self.data.get("treatment_templates", [])
        return [self._clean_treatment_template(item, item)
                for item in raw_templates if isinstance(item, dict)
                and str(item.get("disease", "")).strip()]

    def save_treatment_template(self, item: Dict[str, Any]) -> str:
        cleaned = self._clean_treatment_template(item)
        if not cleaned["disease"] or not cleaned["medications"]:
            return ""
        templates = self.treatment_templates()
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
        imported = [self._clean_treatment_template(item)
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
                    merged[index] = item
                else:
                    positions[key] = len(merged)
                    merged.append(item)
        self.data["treatment_templates"] = merged
        self.save()
        return len(imported)

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

    # -- Gemini drug reference ---------------------------------------------
    @property
    def gemini_enabled(self) -> bool:
        return bool(self.data.get("gemini_enabled", False))

    @gemini_enabled.setter
    def gemini_enabled(self, value: bool) -> None:
        self.set("gemini_enabled", bool(value))

    @property
    def gemini_api_key(self) -> str:
        return str(self.data.get("gemini_api_key", ""))

    def set_gemini(self, api_key: str, enabled: bool) -> None:
        self.data["gemini_api_key"] = str(api_key or "").strip()
        self.data["gemini_enabled"] = bool(enabled and self.data["gemini_api_key"])
        self.save()

    def remove_gemini_key(self) -> None:
        self.data["gemini_api_key"] = ""
        self.data["gemini_enabled"] = False
        self.save()

    # -- backup / restore ---------------------------------------------------
    def create_backup(self, destination: str) -> str:
        """Create a portable archive of encrypted settings and the active drug DB."""
        self.save()
        target = Path(destination)
        target.parent.mkdir(parents=True, exist_ok=True)
        database = Path(self.drug_db_path)
        manifest = {"format": BACKUP_FORMAT, "version": 1, "has_drug_database": database.is_file()}
        with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("manifest.json", json.dumps(manifest, separators=(",", ":")))
            archive.writestr("config.json", CONFIG_PATH.read_bytes())
            if database.is_file():
                archive.write(database, "drug_database.csv")
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

        if not isinstance(restored, dict):
            raise ValueError("The settings backup is invalid.")
        if db_data is not None:
            temporary_db = DEFAULT_DB_PATH.with_suffix(".restore.tmp")
            temporary_db.write_bytes(db_data)
            temporary_db.replace(DEFAULT_DB_PATH)
            restored["drug_db_path"] = str(DEFAULT_DB_PATH)

        self.data = _default_config()
        self.data.update(restored)
        self.data.setdefault("clinic", {})
        self.data.setdefault("doctor", _doctor())
        self.data.setdefault("profiles", {"Default": self.data["doctor"].copy()})
        self.data.setdefault("active_profile", "Default")
        self.data.setdefault("signing_private_key", "")
        self.data.setdefault("medication_favorites", [])
        self.data.setdefault("treatment_templates", [])
        self.data.setdefault("favorite_therapeutic_groups", [])
        self.data.setdefault("dosage_presets", [])
        self.save()


config = Config()
