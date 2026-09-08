"""Encrypted, local-only patient history.

The history intentionally stores only patient demographics. Prescription content
is not retained automatically, which limits the amount of sensitive data kept
on the device. Data is protected with Windows DPAPI for the current account.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from uuid import uuid4

from config import APP_DIR
from security import DataProtectionError, protect, unprotect


HISTORY_PATH = APP_DIR / "patient_history.json"


class PatientHistory:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or HISTORY_PATH

    def _load(self) -> list[dict[str, Any]]:
        if not self.path.exists():
            return []
        try:
            protected = json.loads(self.path.read_text(encoding="utf-8"))
            if protected.get("format") != "dpapi-v1":
                return []
            payload = json.loads(unprotect(protected["data"]).decode("utf-8"))
            records = payload.get("records", [])
            return [record for record in records if isinstance(record, dict)]
        except (OSError, ValueError, KeyError, DataProtectionError):
            return []

    def _save(self, records: list[dict[str, Any]]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps({"schema": 1, "records": records}, ensure_ascii=False,
                             separators=(",", ":")).encode("utf-8")
        protected = {"format": "dpapi-v1", "data": protect(payload)}
        temporary = self.path.with_suffix(".tmp")
        temporary.write_text(json.dumps(protected), encoding="utf-8")
        temporary.replace(self.path)

    @staticmethod
    def _key(name: str) -> str:
        return " ".join(name.casefold().split())

    def search(self, query: str = "") -> list[dict[str, Any]]:
        query = self._key(query)
        records = self._load()
        if query:
            records = [record for record in records
                       if query in self._key(str(record.get("name", "")))]
        return sorted(records, key=lambda record: record.get("updated_at", ""), reverse=True)

    def save_patient(self, patient: dict[str, str]) -> dict[str, Any]:
        name = str(patient.get("name", "")).strip()
        if not name:
            raise ValueError("Patient name is required to save history.")
        records = self._load()
        key = self._key(name)
        record = next((item for item in records if self._key(str(item.get("name", ""))) == key), None)
        now = datetime.now(timezone.utc).isoformat()
        if record is None:
            record = {"id": uuid4().hex, "created_at": now}
            records.append(record)
        record.update({"name": name, "age": str(patient.get("age", "")).strip(),
                       "sex": str(patient.get("sex", "")).strip(), "updated_at": now})
        record.setdefault("prescriptions", [])
        self._save(records)
        return record.copy()

    def save_prescription(self, patient: dict[str, str], drugs: list[dict[str, str]]) -> dict[str, Any]:
        """Append a local-only medication snapshot for a named patient."""
        name = str(patient.get("name", "")).strip()
        if not name:
            raise ValueError("Patient name is required before saving a prescription.")
        clean_drugs = []
        for drug in drugs:
            generic_name = str(drug.get("generic_name", "")).strip()
            brand_name = str(drug.get("brand_name", "")).strip()
            if not (generic_name or brand_name):
                continue
            clean_drugs.append({key: str(drug.get(key, "")).strip() for key in
                                ("generic_name", "brand_name", "dosage", "frequency", "duration", "notes")})
        if not clean_drugs:
            raise ValueError("Add at least one medicine before saving a prescription.")
        records = self._load()
        key = self._key(name)
        record = next((item for item in records if self._key(str(item.get("name", ""))) == key), None)
        now = datetime.now(timezone.utc).isoformat()
        if record is None:
            record = {"id": uuid4().hex, "created_at": now, "prescriptions": []}
            records.append(record)
        record.update({"name": name, "age": str(patient.get("age", "")).strip(),
                       "sex": str(patient.get("sex", "")).strip(), "updated_at": now})
        prescriptions = record.setdefault("prescriptions", [])
        prescriptions.append({"id": uuid4().hex, "saved_at": now, "drugs": clean_drugs})
        self._save(records)
        return record.copy()

    def delete(self, record_id: str) -> bool:
        records = self._load()
        kept = [record for record in records if record.get("id") != record_id]
        if len(kept) == len(records):
            return False
        self._save(kept)
        return True
