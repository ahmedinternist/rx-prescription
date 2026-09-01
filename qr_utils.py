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
class Patient:
    name: str = ""
    age: str = ""
    sex: str = ""


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
    doctor: Doctor = field(default_factory=Doctor)
    patient: Patient = field(default_factory=Patient)
    drugs: List[DrugItem] = field(default_factory=list)
    date: str = ""
    rx_id: str = ""

    def to_payload(self) -> Dict[str, Any]:
        """Serialise to a plain dict (versioned)."""
        return {
            "v": 1,
            "rx_id": self.rx_id,
            "date": self.date,
            "doctor": self.doctor.__dict__,
            "patient": self.patient.__dict__,
            "drugs": [d.__dict__ for d in self.drugs],
        }


# ---------------------------------------------------------------------------
# Encode / decode
# ---------------------------------------------------------------------------
def encode_payload(prescription: Prescription, compress: bool = True) -> str:
    """Return the encoded (base64url) string WITHOUT the viewer base URL."""
    raw = json.dumps(prescription.to_payload(), ensure_ascii=False).encode("utf-8")
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
