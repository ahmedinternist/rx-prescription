"""Application configuration, local encryption, and profile persistence."""
from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path
from typing import Any, Dict

from security import DataProtectionError, protect, unprotect

APP_VERSION = "0.2.0"
PAPER_SIZES: Dict[str, tuple[float, float]] = {
    "A5": (420.945, 595.276), "A4": (595.276, 841.889), "Letter": (612.0, 792.0),
}
DEFAULT_PAPER = "A4"
DEFAULT_VIEWER_BASE = "https://ahmedinternist.github.io/rx-viewer/"
APP_DIR = Path(os.environ.get("RX_APP_DATA_DIR") or os.environ.get(
    "APPDATA", str(Path.home() / ".prescription_app"))) / "prescription_app"
CONFIG_PATH = APP_DIR / "config.json"
DATA_DIR = APP_DIR / "data"
DEFAULT_DB_PATH = DATA_DIR / "drugs.csv"
LOG_PATH = APP_DIR / "rx-prescription.log"
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
        "viewer_base_url": DEFAULT_VIEWER_BASE,
        "drug_db_path": str(DEFAULT_DB_PATH),
        "clinic": {"name": "", "address": "", "phone": "", "logo_path": ""},
        "doctor": doctor,
        "profiles": {"Default": doctor.copy()},
        "active_profile": "Default",
        "signing_private_key": "",
    }


class Config:
    def __init__(self) -> None:
        self.data: Dict[str, Any] = _default_config()
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

    def save(self) -> None:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(self.data, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        protected = {"format": "dpapi-v1", "data": protect(payload)}
        temporary = CONFIG_PATH.with_suffix(".tmp")
        temporary.write_text(json.dumps(protected), encoding="utf-8")
        temporary.replace(CONFIG_PATH)

    def get(self, key: str, default=None):
        return self.data.get(key, default)

    def set(self, key: str, value: Any) -> None:
        self.data[key] = value
        self.save()

    @property
    def paper_size(self) -> str:
        return self.data.get("paper_size", DEFAULT_PAPER)

    @paper_size.setter
    def paper_size(self, value: str) -> None:
        if value not in PAPER_SIZES:
            raise ValueError(f"Unknown paper size: {value}")
        self.set("paper_size", value)

    @property
    def language(self) -> str:
        return self.data.get("language", "en")

    @language.setter
    def language(self, value: str) -> None:
        self.set("language", value if value in ("en", "ar") else "en")

    @property
    def viewer_base_url(self) -> str:
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


config = Config()
