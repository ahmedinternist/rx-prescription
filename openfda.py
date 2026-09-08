"""Compact openFDA Drug Label reference client.

The client queries only medicine names.  Returned label text is presented as
reference material for clinician review, not as automated prescribing advice.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Callable, Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


API_URL = "https://api.fda.gov/drug/label.json"
USER_AGENT = "RxPrescriptionPrinter/3.0 (openFDA label reference)"


class OpenFDALookupError(RuntimeError):
    pass


@dataclass(frozen=True)
class LabelReference:
    medicine: str
    scientific_name: str
    label_name: str
    contraindications: tuple[str, ...] = ()
    pregnancy: tuple[str, ...] = ()
    interactions: tuple[str, ...] = ()


_STRENGTH_OR_FORM = re.compile(
    r"\s+(?:\d+(?:[.,]\d+)?\s*(?:mcg|μg|ug|mg|g|ml|mL|iu|units?|%)(?:\s*/\s*\w+)?"
    r"|tab(?:let)?s?|caps?(?:ule)?s?|syrup|susp(?:ension)?|solution|drops?|"
    r"cream|ointment|gel|spray|inhaler|amp(?:oule)?s?|vials?|injection|inj)\b.*$",
    re.IGNORECASE,
)
_MANUFACTURER_SUFFIX = re.compile(r"\s*[-–—]\s*[^-–—]+$")


def scientific_name_candidate(medicine: str) -> str:
    name = " ".join(medicine.split())
    name = _MANUFACTURER_SUFFIX.sub("", name)
    name = _STRENGTH_OR_FORM.sub("", name).strip(" ,;:-")
    return name or " ".join(medicine.split())


def _request_json(url: str, opener: Callable = urlopen) -> dict:
    request = Request(url, headers={"Accept": "application/json", "User-Agent": USER_AGENT})
    try:
        with opener(request, timeout=10) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        if exc.code == 404:
            return {}
        raise OpenFDALookupError(f"openFDA returned HTTP {exc.code}.") from exc
    except (URLError, TimeoutError, OSError) as exc:
        raise OpenFDALookupError("Could not connect to openFDA. Check your internet connection and try again.") from exc
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise OpenFDALookupError("openFDA returned an unreadable response.") from exc


def _sections(record: dict, key: str) -> tuple[str, ...]:
    values = record.get(key, [])
    return tuple(" ".join(str(value).split()) for value in values if str(value).strip()) if isinstance(values, list) else ()


def _label_name(record: dict, fallback: str) -> str:
    metadata = record.get("openfda", {}) if isinstance(record.get("openfda"), dict) else {}
    for key in ("generic_name", "brand_name", "substance_name"):
        values = metadata.get(key, [])
        if isinstance(values, list) and values and str(values[0]).strip():
            return str(values[0]).strip()
    return fallback


def concise(paragraphs: Iterable[str], limit: int = 180) -> str:
    """Return one readable, short label statement for the dashboard card."""
    text = " ".join(paragraph for paragraph in paragraphs if paragraph)
    text = text.replace("\ufffd", "")
    # FDA label extracts frequently begin with a numbered section heading.  The
    # card already supplies that heading, so omit it from the short statement.
    text = re.sub(
        r"^\s*\d+(?:\.\d+)?\s+(?:CONTRAINDICATIONS|PREGNANCY|DRUG INTERACTIONS|"
        r"USE IN SPECIFIC POPULATIONS)\s*",
        "", text, flags=re.IGNORECASE,
    )
    if not text:
        return ""
    sentence = re.split(r"(?<=[.!?])\s+", text)[0]
    if len(sentence) > limit:
        sentence = sentence[:limit - 3].rsplit(" ", 1)[0] + "..."
    return sentence


def lookup_label(medicine: str, opener: Callable = urlopen) -> LabelReference | None:
    medicine = " ".join(medicine.split())
    if not medicine:
        return None
    scientific_name = scientific_name_candidate(medicine)
    escaped = scientific_name.replace('"', "")
    for field in ("openfda.generic_name", "openfda.brand_name", "generic_name", "brand_name"):
        url = API_URL + "?" + urlencode({"search": f'{field}:"{escaped}"', "limit": 1})
        payload = _request_json(url, opener)
        results = payload.get("results", []) if isinstance(payload, dict) else []
        if not results:
            continue
        record = results[0]
        return LabelReference(
            medicine=medicine, scientific_name=scientific_name,
            label_name=_label_name(record, scientific_name),
            contraindications=_sections(record, "contraindications"),
            pregnancy=_sections(record, "pregnancy") + _sections(record, "use_in_specific_populations"),
            interactions=_sections(record, "drug_interactions"),
        )
    return None


def lookup_labels(medicines: Iterable[str], opener: Callable = urlopen) -> list[LabelReference]:
    results: list[LabelReference] = []
    seen: set[str] = set()
    for medicine in medicines:
        normalized = " ".join(str(medicine).split())
        if not normalized or normalized.casefold() in seen:
            continue
        seen.add(normalized.casefold())
        result = lookup_label(normalized, opener)
        if result:
            results.append(result)
    return results
