"""Grounded Gemini drug-reference lookup with a short encrypted local cache."""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

import config as cfg
from security import DataProtectionError, protect, unprotect

MODEL_CANDIDATES = (
    "gemini-flash-latest",
    "gemini-3.8-flash",
    "gemini-3.7-flash",
    "gemini-3.6-flash",
)
MODEL_NAME = MODEL_CANDIDATES[0]
CACHE_TTL = timedelta(days=7)
CACHE_PATH = cfg.APP_DIR / "gemini-drug-cache.json"
SCHEMA_VERSION = 2
_working_model: str | None = None

SECTIONS = (
    ("indications", "1. Indications"),
    ("minimum_starting_dose", "2. Minimum adult starting dose"),
    ("usual_starting_dose", "3. Usual adult starting dose"),
    ("usual_frequency", "4. Minimum and usual administration frequency"),
    ("maximum_dose_frequency", "5. Maximum adult dose and frequency"),
    ("adverse_effects", "6. Common and serious adverse effects"),
    ("contraindications", "7. Major contraindications"),
    ("pregnancy", "8. Use in pregnancy"),
    ("renal_adjustment", "9. Renal adjustment"),
)

REFERENCE_SCHEMA = {
    "type": "object",
    "properties": {
        key: {
            "type": "array",
            "items": {"type": "string"},
            "description": f"Concise clinician-reference bullets for {heading}.",
        }
        for key, heading in SECTIONS
    },
    "required": [key for key, _heading in SECTIONS],
}


class GeminiDrugError(RuntimeError):
    """A concise user-facing Gemini lookup failure."""


@dataclass(frozen=True)
class DrugReference:
    drug_name: str
    text: str
    sources: tuple[tuple[str, str], ...]
    checked_at: str
    cache_hit: bool = False
    grounded: bool = True


def normalize_drug_name(value: str) -> str:
    """Accept one compact scientific name and reject prompt-like input."""
    raw = str(value or "")
    if any(char in raw for char in "\r\n{}<>`"):
        raise GeminiDrugError("The scientific medicine name is not valid for lookup.")
    name = " ".join(raw.split())
    if len(name) < 2:
        raise GeminiDrugError("Enter a scientific medicine name first.")
    if len(name) > 120:
        raise GeminiDrugError("The scientific medicine name is not valid for lookup.")
    return name


def build_prompt(drug_name: str, grounded: bool = True) -> str:
    name = normalize_drug_name(drug_name)
    source_instruction = (
        f"Search current authoritative sources for the scientific medicine name: {name}"
        if grounded else
        f"Use your internal medical knowledge for the scientific medicine name: {name}. "
        "Web search is unavailable. Do not claim that the answer is current and do not invent citations."
    )
    closing_instruction = (
        "Keep the complete answer concise and include the sources used."
        if grounded else
        "Keep the complete answer concise. Do not include a Sources section."
    )
    return f"""You are preparing a concise drug-reference summary for a licensed clinician.
{source_instruction}

Return a complete structured answer for all nine requested fields:
1. Indications
2. Minimum adult starting dose
3. Usual adult starting dose
4. Minimum and usual administration frequency
5. Maximum adult dose and maximum frequency
6. Common and serious adverse effects
7. Major contraindications
8. Use in pregnancy
9. Renal adjustment

Every field must contain at least one concise item. Never omit a field. If a value
is not established or varies, state that explicitly rather than leaving it blank.
For every dose, state route, formulation, indication, and frequency when relevant.
State clearly when dosing depends on age, pregnancy stage, or kidney-function range.
Do not infer patient-specific advice and do not guess.
Prefer official product labels, regulators, and established medical references.
{closing_instruction}"""


def _clean_text(value: str) -> str:
    lines: list[str] = []
    for raw in str(value or "").splitlines():
        line = raw.strip()
        line = re.sub(r"^#{1,6}\s*", "", line)
        line = line.replace("**", "").replace("__", "")
        if line.startswith("* "):
            line = "• " + line[2:]
        elif line.startswith("- "):
            line = "• " + line[2:]
        lines.append(line)
    return "\n".join(lines).strip()


def _parse_monograph(response: Any) -> dict[str, list[str]]:
    parsed = getattr(response, "parsed", None)
    if hasattr(parsed, "model_dump"):
        parsed = parsed.model_dump()
    if not isinstance(parsed, dict):
        try:
            parsed = json.loads(str(getattr(response, "text", "") or ""))
        except (TypeError, ValueError, json.JSONDecodeError) as exc:
            raise GeminiDrugError("Gemini returned an incomplete drug reference.") from exc
    cleaned: dict[str, list[str]] = {}
    for key, _heading in SECTIONS:
        values = parsed.get(key)
        if not isinstance(values, list):
            raise GeminiDrugError("Gemini returned an incomplete drug reference.")
        items: list[str] = []
        for value in values:
            item = _clean_text(str(value)).strip(" •*-:\t")
            if item:
                items.append(item)
        if not items:
            raise GeminiDrugError("Gemini returned an incomplete drug reference.")
        cleaned[key] = items[:6]
    return cleaned


def _format_monograph(value: dict[str, list[str]]) -> str:
    lines: list[str] = []
    for key, heading in SECTIONS:
        if lines:
            lines.append("")
        lines.append(heading)
        lines.extend(f"• {item}" for item in value[key])
    return "\n".join(lines)


def _extract_sources(response: Any) -> tuple[tuple[str, str], ...]:
    found: list[tuple[str, str]] = []
    for candidate in getattr(response, "candidates", None) or ():
        metadata = getattr(candidate, "grounding_metadata", None)
        for chunk in getattr(metadata, "grounding_chunks", None) or ():
            web = getattr(chunk, "web", None)
            uri = str(getattr(web, "uri", "") or "").strip()
            title = str(getattr(web, "title", "") or uri).strip()
            if uri and (title, uri) not in found:
                found.append((title, uri))
    return tuple(found[:8])


def _model_unavailable(exc: Exception) -> bool:
    message = str(exc).casefold()
    return ("404" in message or "not_found" in message or "not found" in message
            or "no longer available" in message) and "model" in message


def _quota_exhausted(exc: Exception) -> bool:
    message = str(exc).casefold()
    return "429" in message or "resource_exhausted" in message or "quota" in message


def _friendly_error(exc: Exception, action: str) -> GeminiDrugError:
    message = str(exc).casefold()
    if ("api_key_invalid" in message or "api key not valid" in message
            or "invalid api key" in message or "401" in message):
        detail = "The Gemini API key was rejected. Create or verify the key in Google AI Studio."
    elif "403" in message or "permission_denied" in message:
        detail = "This API key does not have permission to use Gemini. Check its Google project restrictions."
    elif "429" in message or "resource_exhausted" in message or "quota" in message:
        detail = "The Gemini request quota has been reached. Check the project's quota or try again later."
    elif "timeout" in message or "timed out" in message:
        detail = "The Gemini request timed out. Check the internet connection and try again."
    elif _model_unavailable(exc):
        detail = "No compatible Gemini Flash model is available for this API key."
    else:
        detail = "Check the internet connection and Gemini API-key settings, then try again."
    return GeminiDrugError(f"Gemini {action} failed. {detail}")


class GeminiDrugClient:
    def __init__(self, api_key: str, cache_path: Path | str = CACHE_PATH,
                 client: Any | None = None) -> None:
        self.api_key = str(api_key or "").strip()
        self.cache_path = Path(cache_path)
        self._client = client

    def _api_client(self):
        if self._client is not None:
            return self._client
        if not self.api_key:
            raise GeminiDrugError("Add a Gemini API key in Settings first.")
        try:
            from google import genai
            from google.genai import types
        except ImportError as exc:
            raise GeminiDrugError("The Gemini component is not installed in this app.") from exc
        self._client = genai.Client(
            api_key=self.api_key,
            http_options=types.HttpOptions(timeout=30000),
        )
        return self._client

    def _generate(self, contents: str, config: Any | None = None):
        """Use the rolling Flash alias, with stable fallbacks for account rollout gaps."""
        global _working_model
        models = list(MODEL_CANDIDATES)
        if _working_model in models:
            models.remove(_working_model)
            models.insert(0, _working_model)
        last_unavailable: Exception | None = None
        for model in models:
            try:
                response = self._api_client().models.generate_content(
                    model=model, contents=contents, config=config)
            except Exception as exc:
                if _model_unavailable(exc):
                    last_unavailable = exc
                    continue
                raise
            _working_model = model
            return response
        if last_unavailable is not None:
            raise last_unavailable
        raise GeminiDrugError("No Gemini Flash model is configured.")

    def _generate_monograph(self, name: str, grounded: bool):
        from google.genai import types

        options: dict[str, Any] = {
            "max_output_tokens": 3200,
            "response_mime_type": "application/json",
            "response_json_schema": REFERENCE_SCHEMA,
            "thinking_config": types.ThinkingConfig(thinking_level="low"),
        }
        if grounded:
            options["tools"] = [types.Tool(google_search=types.GoogleSearch())]
        config = types.GenerateContentConfig(**options)
        last_error: GeminiDrugError | None = None
        for attempt in range(2):
            prompt = build_prompt(name, grounded=grounded)
            if attempt:
                prompt += "\nYour previous response was incomplete. Return every required field now."
            response = self._generate(prompt, config)
            try:
                return response, _parse_monograph(response)
            except GeminiDrugError as exc:
                last_error = exc
        raise last_error or GeminiDrugError("Gemini returned an incomplete drug reference.")

    def _load_cache(self) -> dict[str, Any]:
        if not self.cache_path.is_file():
            return {}
        try:
            envelope = json.loads(self.cache_path.read_text(encoding="utf-8"))
            if envelope.get("format") != "dpapi-v1":
                return {}
            value = json.loads(unprotect(envelope["data"]).decode("utf-8"))
            return value if isinstance(value, dict) else {}
        except (OSError, ValueError, KeyError, DataProtectionError):
            return {}

    def _save_cache(self, value: dict[str, Any]) -> None:
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        envelope = {"format": "dpapi-v1", "data": protect(payload)}
        temporary = self.cache_path.with_suffix(".tmp")
        temporary.write_text(json.dumps(envelope), encoding="utf-8")
        temporary.replace(self.cache_path)

    @staticmethod
    def _from_cache(item: dict[str, Any], cache_hit: bool) -> DrugReference:
        sources = tuple((str(source[0]), str(source[1])) for source in item.get("sources", [])
                        if isinstance(source, list) and len(source) == 2)
        return DrugReference(
            drug_name=str(item.get("drug_name", "")),
            text=str(item.get("text", "")),
            sources=sources,
            checked_at=str(item.get("checked_at", "")),
            cache_hit=cache_hit,
            grounded=bool(item.get("grounded", True)),
        )

    def fetch(self, drug_name: str, force: bool = False) -> DrugReference:
        name = normalize_drug_name(drug_name)
        cache = self._load_cache()
        key = name.casefold()
        cached = cache.get(key)
        if (isinstance(cached, dict) and not force
                and cached.get("schema_version") == SCHEMA_VERSION):
            try:
                checked = datetime.fromisoformat(str(cached["checked_at"]))
                if datetime.now(timezone.utc) - checked <= CACHE_TTL:
                    return self._from_cache(cached, cache_hit=True)
            except (KeyError, TypeError, ValueError):
                pass

        grounded = True
        try:
            response, monograph = self._generate_monograph(name, grounded=True)
        except Exception as exc:
            if not _quota_exhausted(exc):
                if isinstance(exc, GeminiDrugError):
                    raise
                raise _friendly_error(exc, "lookup") from exc
            # Google Search grounding is unavailable on the free tier. Retry
            # once without tools and mark the answer clearly as ungrounded.
            grounded = False
            try:
                response, monograph = self._generate_monograph(name, grounded=False)
            except Exception as fallback_exc:
                if isinstance(fallback_exc, GeminiDrugError):
                    raise
                raise _friendly_error(fallback_exc, "lookup") from fallback_exc

        text = _format_monograph(monograph)
        checked_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
        sources = _extract_sources(response) if grounded else ()
        result = DrugReference(name, text, sources, checked_at, grounded=grounded)
        cache[key] = {
            "drug_name": result.drug_name,
            "text": result.text,
            "sources": [list(source) for source in result.sources],
            "checked_at": result.checked_at,
            "grounded": result.grounded,
            "schema_version": SCHEMA_VERSION,
        }
        # Keep the cache bounded even when many medicines have been queried.
        if len(cache) > 100:
            cache = dict(sorted(cache.items(), key=lambda pair: pair[1].get("checked_at", ""))[-100:])
        self._save_cache(cache)
        return result

    def test_connection(self) -> bool:
        try:
            response = self._generate("Reply with OK only.")
        except Exception as exc:
            if isinstance(exc, GeminiDrugError):
                raise
            raise _friendly_error(exc, "connection") from exc
        return bool(str(getattr(response, "text", "")).strip())


def clear_cache(path: Path | str = CACHE_PATH) -> None:
    target = Path(path)
    if target.is_file():
        target.unlink()
