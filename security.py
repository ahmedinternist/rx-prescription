"""Windows data-protection helpers used for local application secrets.

The app deliberately does not retain patient records. The only persistent
secrets are the doctor's signing key and clinic settings, which are encrypted
with the current Windows user's DPAPI key before they are written to disk.
"""
from __future__ import annotations

import base64
import ctypes
from ctypes import wintypes


class DataProtectionError(RuntimeError):
    """Raised when Windows cannot protect or unprotect local configuration."""


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


_crypt32 = ctypes.windll.crypt32
_kernel32 = ctypes.windll.kernel32


def _blob(data: bytes):
    buffer = ctypes.create_string_buffer(data)
    return _DataBlob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte))), buffer


def _raise(message: str) -> None:
    raise DataProtectionError(f"{message} (Windows error {ctypes.get_last_error()})")


def protect(data: bytes) -> str:
    """Encrypt bytes for the current Windows account and return base64 text."""
    source, source_buffer = _blob(data)
    result = _DataBlob()
    if not _crypt32.CryptProtectData(ctypes.byref(source), "RxPrescription", None, None,
                                     None, 0, ctypes.byref(result)):
        _raise("Could not protect local configuration")
    try:
        return base64.b64encode(ctypes.string_at(result.pbData, result.cbData)).decode("ascii")
    finally:
        _kernel32.LocalFree(result.pbData)


def unprotect(encoded: str) -> bytes:
    """Decrypt base64 data protected by :func:`protect` for this Windows user."""
    source, source_buffer = _blob(base64.b64decode(encoded))
    result = _DataBlob()
    if not _crypt32.CryptUnprotectData(ctypes.byref(source), None, None, None, None, 0,
                                       ctypes.byref(result)):
        _raise("Could not decrypt local configuration")
    try:
        return ctypes.string_at(result.pbData, result.cbData)
    finally:
        _kernel32.LocalFree(result.pbData)
