"""Compact openFDA Drug Label reference client.

The client queries only medicine names.  Returned label text is presented as
reference material for clinician review, not as automated prescribing advice.
"""
from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass
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
    indications: tuple[str, ...] = ()
    dosage: tuple[str, ...] = ()
    renal_adjustment: tuple[str, ...] = ()
    adult_dose: tuple[str, ...] = ()
    maximum_dose: tuple[str, ...] = ()
    route: tuple[str, ...] = ()
    hepatic_adjustment: tuple[str, ...] = ()
    renal_status: str = "not_found"
    effective_date: str = ""
    field_sources: tuple[tuple[str, tuple[str, ...]], ...] = ()
    full_sections: tuple[tuple[str, tuple[str, ...]], ...] = ()
    source_url: str = ""


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


def _topic_excerpts(record: dict, keys: Iterable[str], pattern: str) -> tuple[str, ...]:
    """Keep verbatim topic-containing sections, including their dosing context.

    Renal adjustment is not a standalone openFDA field. Split numbered label
    subsections where available; otherwise retain the whole matching paragraph
    rather than inventing or separating a dose from its qualifications.
    """
    excerpts = []
    for key in keys:
        for paragraph in _sections(record, key):
            chunks = re.split(r"(?<![\w.])(?=\d{1,2}\.\d{1,2}\s+[A-Z])", paragraph)
            for chunk in chunks:
                chunk = chunk.strip()
                if chunk and re.search(pattern, chunk, flags=re.IGNORECASE) and chunk not in excerpts:
                    excerpts.append(chunk)
    return tuple(excerpts)


def _topic_excerpts_with_sources(record: dict, keys: Iterable[str], pattern: str):
    excerpts, sources = [], []
    for key in keys:
        found = _topic_excerpts(record, (key,), pattern)
        if found:
            excerpts.extend(value for value in found if value not in excerpts)
            sources.append(key.replace("_", " ").title())
    return tuple(excerpts), tuple(sources)


def _dose_sentences(values: Iterable[str], pattern: str) -> tuple[str, ...]:
    matches = []
    for value in values:
        for sentence in re.split(r"(?<=[.!?])\s+", value):
            if re.search(pattern, sentence, flags=re.IGNORECASE):
                normalized = sentence.strip()
                if normalized and normalized not in matches:
                    matches.append(normalized)
    return tuple(matches)


def _renal_status(excerpts: Iterable[str]) -> str:
    text = " ".join(excerpts)
    if not text:
        return "not_found"
    if re.search(r"\b(?:adjust(?:ment|ed)?|reduce[sd]?|decrease[sd]?|increase[sd]?|"
                 r"no\s+(?:dose|dosage)\s+adjustment|not\s+recommended|avoid(?:ed)?|"
                 r"contraindicated)\b", text, flags=re.IGNORECASE):
        return "dose_stated"
    return "precaution_only"


_FULL_SECTION_ORDER = (
    "indications_and_usage", "dosage_and_administration", "dosage_forms_and_strengths",
    "contraindications", "warnings_and_cautions", "warnings", "precautions",
    "adverse_reactions", "drug_interactions", "use_in_specific_populations",
    "pregnancy", "pediatric_use", "geriatric_use", "clinical_pharmacology",
    "clinical_studies", "how_supplied", "storage_and_handling", "patient_medication_information",
)


def _full_sections(record: dict):
    sections = []
    for key in _FULL_SECTION_ORDER:
        values = _sections(record, key)
        if values:
            sections.append((key.replace("_", " ").title(), values))
    return tuple(sections)


def reference_to_dict(reference: LabelReference) -> dict:
    """Return a JSON-safe public-label cache record."""
    return asdict(reference)


def reference_from_dict(data: dict) -> LabelReference:
    """Restore a cache record while safely accepting older cache shapes."""
    tuple_fields = {
        "contraindications", "pregnancy", "interactions", "indications", "dosage",
        "renal_adjustment", "adult_dose", "maximum_dose", "route", "hepatic_adjustment",
    }
    nested_fields = {"field_sources", "full_sections"}
    values = {}
    for name in LabelReference.__dataclass_fields__:
        if name not in data:
            continue
        value = data[name]
        if name in tuple_fields:
            value = tuple(value or ())
        elif name in nested_fields:
            value = tuple((str(item[0]), tuple(item[1] or ())) for item in (value or ())
                          if isinstance(item, (list, tuple)) and len(item) == 2)
        values[name] = value
    return LabelReference(**values)


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
        dosage = _sections(record, "dosage_and_administration")
        renal, renal_sources = _topic_excerpts_with_sources(
            record, ("dosage_and_administration", "use_in_specific_populations",
                     "warnings_and_cautions", "warnings", "precautions", "contraindications"),
            r"\b(?:renal|kidney|eGFR|CrCl|dialysis|hemodialysis|haemodialysis)\b|"
            r"\bcreatinine\s+clearance\b")
        hepatic, hepatic_sources = _topic_excerpts_with_sources(
            record, ("dosage_and_administration", "use_in_specific_populations",
                     "warnings_and_cautions", "warnings", "precautions"),
            r"\b(?:hepatic|liver|Child-Pugh)\b")
        metadata = record.get("openfda", {}) if isinstance(record.get("openfda"), dict) else {}
        route = tuple(str(value).strip() for value in metadata.get("route", [])
                      if str(value).strip()) or _sections(record, "route")
        pregnancy = _sections(record, "pregnancy") or _topic_excerpts(
            record, ("use_in_specific_populations", "pregnancy_or_breast_feeding"),
            r"\bpregnan\w*\b")
        dose_sources = (["Dosage And Administration"] if dosage else [])
        if route:
            dose_sources.append("openFDA Route Metadata")
        dose_sources.extend(source for source in hepatic_sources if source not in dose_sources)
        sources = (
            ("indication", ("Indications And Usage",) if _sections(record, "indications_and_usage") else ()),
            ("dose", tuple(dose_sources)),
            ("contraindication", ("Contraindications",) if _sections(record, "contraindications") else ()),
            ("pregnancy", (("Pregnancy",) if _sections(record, "pregnancy") else
                           (("Use In Specific Populations",) if pregnancy else ()))),
            ("renal", renal_sources),
            ("hepatic", hepatic_sources),
        )
        return LabelReference(
            medicine=medicine, scientific_name=scientific_name,
            label_name=_label_name(record, scientific_name),
            contraindications=_sections(record, "contraindications"),
            pregnancy=pregnancy,
            interactions=_sections(record, "drug_interactions"),
            indications=_sections(record, "indications_and_usage"),
            dosage=dosage,
            renal_adjustment=renal,
            adult_dose=_dose_sentences(dosage, r"\badults?\b"),
            maximum_dose=_dose_sentences(dosage, r"\b(?:maximum|max\.)\b"),
            route=route,
            hepatic_adjustment=hepatic,
            renal_status=_renal_status(renal),
            effective_date=str(record.get("effective_time", "")).strip(),
            field_sources=sources,
            full_sections=_full_sections(record),
            source_url=url,
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
