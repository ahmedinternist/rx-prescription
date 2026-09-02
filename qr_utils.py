"""QR-code payload encoding for prescriptions.

Design (hybrid, privacy-friendly):
  * The FULL prescription (doctor + patient + drugs) is encoded directly into
    the QR code, so it works offline with any free QR scanner.
  * The payload is shaped as a URL:  <viewer_base>#<encoded-data>
    so that when a pharmacist scans it with a phone camera + internet, the
    phone opens the free static viewer page (viewer.html) which decodes the
    part after '#'. Data never leaves the device / is never sent to a server.

Encoding pipeline (kept deliberately simple for maximum compatibility):
  Python dict -> json (utf-8) -> base64url (no padding)
This keeps the payload readable and decodable in EVERY browser without any
dependency on the Compression Streams API. A typical prescription (a few
drugs) is well within QR capacity (version <= ~15). Compression is available
via compress=True but OFF by default for maximal compatibility.
"""
from __future__ import annotations

import base64
import hashlib
import json
import zlib
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import qrcode
from qrcode.constants import ERROR_CORRECT_H

from config import config


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------
@dataclass
class Doctor:
    name: str = ""
    license_no: str = ""
    specialty: str = ""


@dataclass
class Clinic:
    name: str = ""
    address: str = ""
    phone: str = ""
    logo_path: str = ""


@dataclass
class Patient:
    name: str = ""
    age: str = ""
    sex: str = ""
    id_number: str = ""
    allergies: str = ""


@dataclass
class DrugItem:
    generic_name: str = ""
    brand_name: str = ""
    dosage: str = ""
    frequency: str = ""
    duration: str = ""
    notes: str = ""


@dataclass
class Prescription:
    clinic: Clinic = field(default_factory=Clinic)
    doctor: Doctor = field(default_factory=Doctor)
    patient: Patient = field(default_factory=Patient)
    drugs: List[DrugItem] = field(default_factory=list)
    date: str = ""
    rx_id: str = ""
    diagnosis: str = ""
    refills: str = ""

    def to_payload(self) -> Dict[str, Any]:
        """Serialise to a plain dict (versioned)."""
        return {
            "v": 2,
            "rx_id": self.rx_id,
            "date": self.date,
            "diagnosis": self.diagnosis,
            "refills": self.refills,
            "clinic": self.clinic.__dict__,
            "doctor": self.doctor.__dict__,
            "patient": self.patient.__dict__,
            "drugs": [d.__dict__ for d in self.drugs],
        }

    def to_qr_payload(self) -> Dict[str, Any]:
        """Minimal signed data for phone verification and reliable QR scanning.

        Optional values are omitted instead of encoded as empty strings.  This
        lets the web viewer hide unused medication columns and keeps the QR as
        small as practical.
        """
        doctor: Dict[str, str] = {"name": self.doctor.name}
        if self.doctor.specialty:
            doctor["specialty"] = self.doctor.specialty

        drugs: List[Dict[str, str]] = []
        for drug in self.drugs:
            item = {"generic_name": drug.generic_name}
            for field_name in ("dosage", "frequency", "duration", "notes"):
                value = getattr(drug, field_name)
                if value:
                    item[field_name] = value
            drugs.append(item)

        return {
            "v": 4,
            "date": self.date,
            "doctor": doctor,
            "patient": {"name": self.patient.name},
            "drugs": drugs,
        }


def canonical_payload(payload: Dict[str, Any]) -> bytes:
    """Stable JSON representation shared with the static viewer verifier."""
    return json.dumps(payload, ensure_ascii=False, sort_keys=True,
                      separators=(",", ":")).encode("utf-8")


def _b64(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode("ascii").rstrip("=")


def _sign(payload: Dict[str, Any]) -> Dict[str, Any]:
    """Attach an ES256 signature, creating one clinic-local signing key if needed.

    The private key is saved only in the DPAPI-protected local configuration.
    A viewer must contain the matching public key in its trusted key registry
    before it may call a prescription *verified*.
    """
    try:
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import ec
        from cryptography.hazmat.primitives.asymmetric.utils import decode_dss_signature
    except ImportError as exc:  # pragma: no cover - dependency checked at startup
        raise RuntimeError("QR signing requires the 'cryptography' package") from exc

    pem = config.signing_private_key
    if pem:
        private_key = serialization.load_pem_private_key(pem.encode("ascii"), password=None)
    else:
        private_key = ec.generate_private_key(ec.SECP256R1())
        pem = private_key.private_bytes(serialization.Encoding.PEM,
                                        serialization.PrivateFormat.PKCS8,
                                        serialization.NoEncryption()).decode("ascii")
        config.signing_private_key = pem

    public = private_key.public_key()
    public_der = public.public_bytes(serialization.Encoding.DER,
                                     serialization.PublicFormat.SubjectPublicKeyInfo)
    signature_der = private_key.sign(canonical_payload(payload), ec.ECDSA(hashes.SHA256()))
    r, s = decode_dss_signature(signature_der)
    raw_signature = r.to_bytes(32, "big") + s.to_bytes(32, "big")
    signed = dict(payload)
    signed["sig"] = {
        "alg": "ES256",
        "kid": hashlib.sha256(public_der).hexdigest()[:16],
        "value": _b64(raw_signature),
    }
    return signed


def verification_key() -> Dict[str, str]:
    """Return the current signer's public JWK for the static viewer trust list."""
    try:
        from cryptography.hazmat.primitives import serialization
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("QR signing requires the 'cryptography' package") from exc
    # Ensure a signing key exists; its signed sample is intentionally discarded.
    _sign({"v": 2, "bootstrap": True})
    private = serialization.load_pem_private_key(
        config.signing_private_key.encode("ascii"), password=None)
    numbers = private.public_key().public_numbers()
    public_der = private.public_key().public_bytes(serialization.Encoding.DER,
                                                   serialization.PublicFormat.SubjectPublicKeyInfo)
    return {
        "kid": hashlib.sha256(public_der).hexdigest()[:16],
        "kty": "EC", "crv": "P-256",
        "x": _b64(numbers.x.to_bytes(32, "big")),
        "y": _b64(numbers.y.to_bytes(32, "big")),
    }


# ---------------------------------------------------------------------------
# Encode / decode
# ---------------------------------------------------------------------------
def encode_payload(prescription: Prescription, compress: bool = True) -> str:
    """Return the encoded (base64url) string WITHOUT the viewer base URL."""
    raw = canonical_payload(_sign(prescription.to_qr_payload()))
    if compress:
        raw = zlib.compress(raw, level=9)
    s = base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")
    return s


def build_qr_url(prescription: Prescription, viewer_base: Optional[str] = None,
                 compress: bool = True) -> str:
    """Return the full QR content: viewer_base + '#' + encoded payload.

    If the base URL already ends with '#', we don't double it.
    """
    base = (viewer_base or config.viewer_base_url).rstrip("/")
    base = base.rstrip("#")  # avoid ".../#" + "#" => double-hash
    return f"{base}#{encode_payload(prescription, compress=compress)}"


def decode_payload(encoded: str, compressed: Optional[bool] = None) -> Dict[str, Any]:
    """Decode a base64url payload back to a dict.

    If `compressed` is None we auto-detect: try plain JSON first, then zlib.
    """
    s = encoded.strip()
    s += "=" * (-len(s) % 4)  # re-add padding if missing
    data = base64.urlsafe_b64decode(s.encode("ascii"))
    if compressed is None:
        # auto-detect: valid JSON (uncompressed) vs zlib (compressed)
        try:
            return json.loads(data.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            compressed = True
    if compressed:
        data = zlib.decompress(data)
    return json.loads(data.decode("utf-8"))


def decode_from_url(url: str) -> Dict[str, Any]:
    """Decode from a full URL (takes the part after the last '#')."""
    if "#" in url:
        frag = url.rsplit("#", 1)[1]
    else:
        frag = url
    return decode_payload(frag)


# ---------------------------------------------------------------------------
# QR image generation
# ---------------------------------------------------------------------------
def make_qr_image(url: str, box_size: int = 6, border: int = 2):
    """Generate a PIL Image of the QR code (high error correction)."""
    qr = qrcode.QRCode(
        version=None,  # auto-size to fit
        error_correction=ERROR_CORRECT_H,
        box_size=box_size,
        border=border,
    )
    qr.add_data(url)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white")


def qr_info(url: str) -> Dict[str, Any]:
    """Return the number of chars and the chosen QR version/size (for UX)."""
    qr = qrcode.QRCode(error_correction=ERROR_CORRECT_H, box_size=6, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    matrix = qr.get_matrix()
    return {
        "char_count": len(url),
        "qr_version": qr.version,
        "matrix_size": len(matrix),
    }


def validate_prescription(prescription: Prescription) -> tuple[list[str], list[str]]:
    """Return (errors, warnings) before a prescription is exported or printed."""
    errors: list[str] = []
    warnings: list[str] = []
    if not prescription.doctor.name:
        errors.append("Prescriber name is required.")
    if not prescription.doctor.license_no:
        errors.append("Prescriber license number is required.")
    if prescription.refills:
        try:
            refills = int(prescription.refills)
            if not 0 <= refills <= 99:
                errors.append("Refills must be between 0 and 99.")
        except ValueError:
            errors.append("Refills must be a whole number.")
    if not prescription.drugs:
        errors.append("At least one medication is required.")

    seen: set[str] = set()
    for index, drug in enumerate(prescription.drugs, 1):
        if not drug.generic_name:
            errors.append(f"Medication {index}: drug name is required.")
        normalized = " ".join(drug.generic_name.casefold().split())
        if normalized in seen:
            warnings.append(f"{drug.generic_name}: duplicate medication.")
        seen.add(normalized)
    return errors, warnings
