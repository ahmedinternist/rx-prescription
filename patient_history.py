"""Encrypted, local-only patient records and explicitly saved prescriptions.

Prescription snapshots are retained only when the clinician chooses to save
them for a patient. Data is protected with Windows DPAPI for the current account.
"""
from __future__ import annotations

import json
import copy
import functools
import os
import tempfile
import threading
import unicodedata
from difflib import SequenceMatcher
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from uuid import uuid4

from config import APP_DIR
from security import DataProtectionError, protect, unprotect


HISTORY_PATH = APP_DIR / "patient_history.json"


class PatientHistoryError(RuntimeError):
    """Unreadable history must never be treated as a new, empty database."""


def validate_history_bytes(data: bytes) -> list[dict[str, Any]]:
    try:
        protected = json.loads(data.decode("utf-8"))
        if not isinstance(protected, dict) or protected.get("format") != "dpapi-v1":
            raise ValueError("Invalid history envelope")
        payload = json.loads(unprotect(protected["data"]).decode("utf-8"))
        if not isinstance(payload, dict) or not isinstance(payload.get("records"), list):
            raise ValueError("Invalid history records")
        records = payload["records"]
        if any(not isinstance(record, dict) or not isinstance(record.get("prescriptions", []), list)
               for record in records):
            raise ValueError("Invalid patient record")
        return records
    except (ValueError, TypeError, KeyError, DataProtectionError) as exc:
        raise PatientHistoryError(
            "Patient history cannot be read. The original file is preserved; "
            "restore a valid backup before saving patients.") from exc

_path_locks = {}
_path_locks_guard = threading.Lock()


def _serialized(method):
    @functools.wraps(method)
    def transaction(self, *args, **kwargs):
        with self._lock:
            return method(self, *args, **kwargs)
    return transaction


class PatientHistory:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or HISTORY_PATH
        with _path_locks_guard:
            self._lock = _path_locks.setdefault(str(self.path.resolve()), threading.RLock())
        self._cached_records = None
        self._cached_signature = None

    def _signature(self):
        try:
            stat = self.path.stat()
            return stat.st_mtime_ns, stat.st_size, stat.st_ino
        except OSError:
            return None

    @_serialized
    def _load(self) -> list[dict[str, Any]]:
        return copy.deepcopy(self._read_records())

    def _read_records(self):
        """Borrow the cache only while the caller holds the path lock."""
        signature = self._signature()
        if self._cached_records is not None and signature == self._cached_signature:
            return self._cached_records
        if not self.path.exists():
            self._cached_records = []
            self._cached_signature = None
            return []
        try:
            records = validate_history_bytes(self.path.read_bytes())
            # Only cache a consistent read; an external replacement invalidates it.
            if signature == self._signature():
                self._cached_records = records
                self._cached_signature = signature
            return records
        except OSError as exc:
            raise PatientHistoryError("Patient history is unavailable. No records were changed.") from exc

    @_serialized
    def backup_bytes(self) -> bytes | None:
        if not self.path.exists():
            return None
        data = self.path.read_bytes()
        validate_history_bytes(data)
        return data

    @_serialized
    def _save(self, records: list[dict[str, Any]]) -> None:
        expected_signature = self._signature()
        if self._cached_records is not None and expected_signature != self._cached_signature:
            raise RuntimeError("Patient history changed externally. Reload it before saving again.")
        # Also protect callers of this compatibility entry point.
        self._read_records()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps({"schema": 1, "records": records}, ensure_ascii=False,
                             separators=(",", ":")).encode("utf-8")
        protected = {"format": "dpapi-v1", "data": protect(payload)}
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8",
                    dir=self.path.parent, prefix=self.path.name + ".", suffix=".tmp",
                    delete=False) as stream:
                temporary = Path(stream.name)
                json.dump(protected, stream)
                stream.flush()
                os.fsync(stream.fileno())
                stat = os.fstat(stream.fileno())
                written_signature = stat.st_mtime_ns, stat.st_size, stat.st_ino
            if expected_signature != self._signature():
                raise RuntimeError("Patient history changed externally. Reload it before saving again.")
            os.replace(temporary, self.path)
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)
        self._cached_records = copy.deepcopy(records)
        self._cached_signature = written_signature

    @staticmethod
    def _key(name: str) -> str:
        normalized = unicodedata.normalize("NFKD", str(name).casefold())
        normalized = "".join(
            char for char in normalized
            if not unicodedata.combining(char) and not ("\u064b" <= char <= "\u065f"))
        normalized = normalized.translate(str.maketrans({
            "أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا", "ى": "ي",
        }))
        return " ".join(normalized.split())

    @_serialized
    def search(self, query: str = "") -> list[dict[str, Any]]:
        query = self._key(query)
        records = self._read_records()
        if query:
            records = [record for record in records
                       if query in self._key(str(record.get("name", "")))]
        return copy.deepcopy(sorted(records, key=lambda record: record.get("updated_at", ""), reverse=True))

    @_serialized
    def find_similar(self, name: str, age: str = "", threshold: float = 0.82,
                     exclude_id: str = "", sex: str = "") -> list[dict[str, Any]]:
        """Return likely duplicates using normalized name, age, and sex matching."""
        key = self._key(name)
        if not key:
            return []
        age = str(age or "").strip()
        sex = str(sex or "").strip().casefold()
        matches = []
        for record in self._read_records():
            if exclude_id and record.get("id") == exclude_id:
                continue
            record_key = self._key(str(record.get("name", "")))
            score = SequenceMatcher(None, key, record_key).ratio()
            same_age = not age or not record.get("age") or str(record.get("age", "")).strip() == age
            record_sex = str(record.get("sex", "")).strip().casefold()
            same_sex = not sex or not record_sex or record_sex == sex
            if score >= threshold and same_age and same_sex:
                item = copy.deepcopy(record)
                item["similarity"] = score
                matches.append(item)
        return sorted(matches, key=lambda item: item["similarity"], reverse=True)

    @_serialized
    def get(self, record_id: str) -> dict[str, Any] | None:
        record = next((item for item in self._read_records() if item.get("id") == record_id), None)
        return copy.deepcopy(record) if record else None

    @_serialized
    def save_patient(self, patient: dict[str, str], record_id: str = "") -> dict[str, Any]:
        name = str(patient.get("name", "")).strip()
        if not name:
            raise ValueError("Patient name is required to save history.")
        records = self._load()
        key = self._key(name)
        record = next((item for item in records if record_id and item.get("id") == record_id), None)
        if record_id and record is None:
            raise ValueError("The selected patient no longer exists. Reload patient history.")
        now = datetime.now(timezone.utc).isoformat()
        if record is None:
            record = {"id": uuid4().hex, "created_at": now}
            records.append(record)
        record.update({"name": name, "age": str(patient.get("age", "")).strip(),
                       "sex": str(patient.get("sex", "")).strip(), "updated_at": now})
        record.setdefault("prescriptions", [])
        self._save(records)
        return record.copy()

    @_serialized
    def save_prescription(self, patient: dict[str, str], drugs: list[dict[str, str]],
                          record_id: str = "") -> dict[str, Any]:
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
                                ("generic_name", "brand_name", "dosage", "frequency",
                                 "duration", "notes", "quantity")})
        if not clean_drugs:
            raise ValueError("Add at least one medicine before saving a prescription.")
        records = self._load()
        key = self._key(name)
        record = next((item for item in records if record_id and item.get("id") == record_id), None)
        if record_id and record is None:
            raise ValueError("The selected patient no longer exists. Reload patient history.")
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

    @_serialized
    def delete(self, record_id: str) -> bool:
        records = self._load()
        kept = [record for record in records if record.get("id") != record_id]
        if len(kept) == len(records):
            return False
        self._save(kept)
        return True

    @_serialized
    def restore_record(self, record: dict[str, Any]) -> bool:
        """Restore one previously deleted patient without duplicating its ID."""
        if not isinstance(record, dict) or not record.get("id") or not record.get("name"):
            return False
        records = self._load()
        records = [item for item in records if item.get("id") != record.get("id")]
        restored = dict(record)
        restored["updated_at"] = datetime.now(timezone.utc).isoformat()
        restored.setdefault("prescriptions", [])
        records.append(restored)
        self._save(records)
        return True

    @_serialized
    def delete_prescription(self, record_id: str, prescription_id: str) -> dict[str, Any] | None:
        records = self._load()
        for record in records:
            if record.get("id") != record_id:
                continue
            prescriptions = record.get("prescriptions", [])
            deleted = next((item for item in prescriptions
                            if item.get("id") == prescription_id), None)
            if deleted is None:
                return None
            record["prescriptions"] = [item for item in prescriptions
                                       if item.get("id") != prescription_id]
            record["updated_at"] = datetime.now(timezone.utc).isoformat()
            self._save(records)
            return dict(deleted)
        return None

    @_serialized
    def restore_prescription(self, record_id: str, prescription: dict[str, Any]) -> bool:
        records = self._load()
        for record in records:
            if record.get("id") != record_id:
                continue
            items = record.setdefault("prescriptions", [])
            if any(item.get("id") == prescription.get("id") for item in items):
                return False
            items.append(dict(prescription))
            record["updated_at"] = datetime.now(timezone.utc).isoformat()
            self._save(records)
            return True
        return False
