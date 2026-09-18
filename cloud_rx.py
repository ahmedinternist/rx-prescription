"""Cloud prescription link client. Never logs credentials or prescription data."""
from __future__ import annotations

import json
import socket
from dataclasses import dataclass
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener

API_URL = "https://rx-v2.vercel.app/api/rx"
VIEWER_URL = "https://rx-v2.vercel.app"
TIMEOUT = 15


class CloudRxError(RuntimeError):
    """A safe error code, without server bodies, patient data or credentials."""

    def __init__(self, code: str, status: int | None = None):
        self.code, self.status = code, status
        super().__init__(f"Cloud prescription request failed: {code}")


@dataclass(frozen=True)
class CloudRxLink:
    rx_id: str
    url: str


class _NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def upload_prescription(payload: dict, api_key: str, *, opener=None) -> CloudRxLink:
    """POST minimal JSON once; callers run this blocking function on a worker."""
    key = str(api_key or "").strip()
    if not key:
        raise CloudRxError("missing_key")
    if "\r" in key or "\n" in key:
        raise CloudRxError("auth")
    request = Request(API_URL, method="POST",
                      data=json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8"),
                      headers={"Content-Type": "application/json", "Accept": "application/json",
                               "x-api-key": key})
    try:
        send = opener or build_opener(_NoRedirect()).open
        with send(request, timeout=TIMEOUT) as response:
            status = response.status
            if not 200 <= status < 300:
                raise CloudRxError("http", status)
            raw = response.read(65537)
        if len(raw) > 65536:
            raise CloudRxError("response")
        result = json.loads(raw.decode("utf-8"))
        if not isinstance(result, dict):
            raise CloudRxError("response")
        rx_id, url = result.get("rxId"), result.get("url")
        if not isinstance(rx_id, str) or not rx_id.strip() or not isinstance(url, str):
            raise CloudRxError("response")
        parsed = urlsplit(url)
        if (parsed.scheme != "https" or parsed.netloc != "rx-v2.vercel.app"
                or not parsed.path.startswith("/p/") or not parsed.path[3:].strip()
                or parsed.query or parsed.fragment or any(char.isspace() for char in url)):
            raise CloudRxError("response")
        return CloudRxLink(rx_id.strip(), url)
    except HTTPError as exc:
        code = ("auth" if exc.code in (401, 403) else "quota" if exc.code == 429
                else "server" if exc.code >= 500 else "redirect" if 300 <= exc.code < 400 else "http")
        raise CloudRxError(code, exc.code) from None
    except (TimeoutError, socket.timeout):
        raise CloudRxError("timeout") from None
    except URLError as exc:
        raise CloudRxError("timeout" if isinstance(exc.reason, TimeoutError) else "offline") from None
    except OSError:
        raise CloudRxError("offline") from None
    except (ValueError, UnicodeDecodeError):
        raise CloudRxError("response") from None
