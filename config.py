"""Application configuration and persistence.

Stores user settings and the doctor's saved profile in a JSON file under the
user's AppData/Roaming directory so they survive between launches.
"""
from __future__ import annotations

import json
import os
import shutil
from pathlib import Path
from typing import Any, Dict

# ---------------------------------------------------------------------------
# Paper sizes (in PostScript points, 1 pt = 1/72 inch) as (width, height)
# ---------------------------------------------------------------------------
PAPER_SIZES: Dict[str, tuple[float, float]] = {
    "A5": (420.945, 595.276),
    "A4": (595.276, 841.889),
    "Letter": (612.0, 792.0),
}

DEFAULT_PAPER = "A4"

# The base URL of the free static "viewer" page a pharmacist opens when they
# scan the QR code. The encoded prescription is appended after the '#' so the
# data never travels to a server (privacy-friendly).
DEFAULT_VIEWER_BASE = "https://ahmedinternist.github.io/rx-viewer/"

APP_DIR = Path(os.environ.get("APPDATA", str(Path.home() / ".prescription_app"))) / "prescription_app"
CONFIG_PATH = APP_DIR / "config.json"
DATA_DIR = APP_DIR / "data"
DEFAULT_DB_PATH = DATA_DIR / "drugs.csv"

# When the app is run from the project folder, a bundled seed DB lives here.
# On first launch (or when the user has no DB yet) we copy it into AppData.
PROJECT_DIR = Path(__file__).resolve().parent
SEED_DB_PATH = PROJECT_DIR / "data" / "drugs.csv"

APP_DIR.mkdir(parents=True, exist_ok=True)
DATA_DIR.mkdir(parents=True, exist_ok=True)


def ensure_seed_db() -> None:
    """Copy the bundled seed drug DB into AppData if no DB exists yet."""
    if not DEFAULT_DB_PATH.exists() and SEED_DB_PATH.exists():
        DEFAULT_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(SEED_DB_PATH, DEFAULT_DB_PATH)


def _default_config() -> Dict[str, Any]:
    return {
        "paper_size": DEFAULT_PAPER,
        "language": "en",
        "viewer_base_url": DEFAULT_VIEWER_BASE,
        "drug_db_path": str(DEFAULT_DB_PATH),
        "doctor": {
            "name": "",
            "license_no": "",
            "specialty": "",
        },
    }


class Config:
    """Thin wrapper around the JSON config file."""

    def __init__(self) -> None:
        self.data: Dict[str, Any] = _default_config()
        self.load()

    # -- persistence --------------------------------------------------------
    def load(self) -> None:
        if CONFIG_PATH.exists():
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
                    loaded = json.load(fh)
                # merge so new keys are not lost
                self.data.update(loaded)
            except Exception:
                pass
        # ensure nested dicts exist
        self.data.setdefault("clinic", {})
        self.data.setdefault("doctor", {})

    def save(self) -> None:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
            json.dump(self.data, fh, indent=2, ensure_ascii=False)

    # -- accessors ----------------------------------------------------------
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
        self.data["paper_size"] = value
        self.save()

    @property
    def language(self) -> str:
        return self.data.get("language", "en")

    @language.setter
    def language(self, value: str) -> None:
        if value not in ("en", "ar"):
            value = "en"
        self.data["language"] = value
        self.save()

    @property
    def viewer_base_url(self) -> str:
        return self.data.get("viewer_base_url", DEFAULT_VIEWER_BASE)

    @viewer_base_url.setter
    def viewer_base_url(self, value: str) -> None:
        self.data["viewer_base_url"] = value
        self.save()

    @property
    def drug_db_path(self) -> str:
        return self.data.get("drug_db_path", str(DEFAULT_DB_PATH))

    @drug_db_path.setter
    def drug_db_path(self, value: str) -> None:
        self.data["drug_db_path"] = value
        self.save()

    def get_doctor(self) -> Dict[str, Any]:
        return self.data.get("doctor", {})

    def set_doctor(self, **kw) -> None:
        self.data.setdefault("doctor", {}).update(kw)
        self.save()


# module-level singleton
config = Config()
