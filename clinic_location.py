"""Clinic pin parsing and a one-shot, permission-based local browser locator.

No API key, external geocoder, automatic tracking or location logging. The browser
helper binds only to loopback, expires, and requires a random token and same origin.
"""
from __future__ import annotations

import html
import json
import math
import re
import secrets
import time
import webbrowser
from dataclasses import dataclass
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, unquote, urljoin, urlsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener


class LocationError(ValueError):
    def __init__(self, code="invalid"):
        self.code = code
        super().__init__("Clinic location operation failed: " + code)


@dataclass(frozen=True)
class Coordinates:
    latitude: float
    longitude: float
    accuracy: float | None = None


def validate_coordinates(latitude, longitude, accuracy=None):
    try:
        if isinstance(latitude, bool) or isinstance(longitude, bool) or isinstance(accuracy, bool):
            raise ValueError()
        lat, lng = float(latitude), float(longitude)
        if not math.isfinite(lat) or not math.isfinite(lng) or abs(lat) > 90 or abs(lng) > 180:
            raise ValueError()
        acc = None if accuracy is None else float(accuracy)
        if acc is not None and (not math.isfinite(acc) or acc < 0):
            raise ValueError()
        return Coordinates(lat, lng, acc)
    except (TypeError, ValueError, OverflowError):
        raise LocationError() from None


_NUMBER = r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)"
_PAIR = re.compile(rf"\s*({_NUMBER})\s*,\s*({_NUMBER})\s*", re.ASCII)
_HOSTS = {"google.com", "www.google.com", "maps.google.com", "maps.app.goo.gl", "goo.gl"}


def _maps_url_parts(value):
    try:
        parts = urlsplit(value)
        if (parts.scheme != "https" or parts.hostname not in _HOSTS or
                parts.username or parts.password or parts.port not in (None, 443)):
            raise ValueError()
        if parts.hostname == "goo.gl" and not parts.path.startswith("/maps/"):
            raise ValueError()
        if parts.hostname in {"google.com", "www.google.com"} and not parts.path.startswith("/maps"):
            raise ValueError()
        return parts
    except ValueError:
        raise LocationError() from None


def parse_location(value):
    """Accept explicit coordinate queries or a Google place pin, never @viewport."""
    value = str(value).strip()
    if len(value) > 8192:
        raise LocationError()
    match = _PAIR.fullmatch(value)
    if match:
        return validate_coordinates(*match.groups())
    parts = _maps_url_parts(value)
    decoded = unquote(value)
    pin = re.search(rf"!3d({_NUMBER})!4d({_NUMBER})", decoded, re.ASCII)
    if pin:
        return validate_coordinates(*pin.groups())
    queries = parse_qs(parts.query)
    for key in ("query", "q"):
        for candidate in queries.get(key, []):
            match = _PAIR.fullmatch(candidate)
            if match:
                return validate_coordinates(*match.groups())
    raise LocationError("pin_missing")


class _NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        return None


def resolve_location(value, *, opener=None):
    """Resolve only Google short links, with TLS, bounded redirects and no secrets."""
    try:
        return parse_location(value)
    except LocationError as exc:
        if exc.code != "pin_missing":
            raise
    parts = _maps_url_parts(value)
    if parts.hostname not in {"maps.app.goo.gl", "goo.gl"}:
        raise LocationError("pin_missing")
    send = opener or build_opener(_NoRedirect()).open
    deadline = time.monotonic() + 15
    url = value.strip()
    for _ in range(5):
        _maps_url_parts(url)  # Reject hostile redirects before sending anything.
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise LocationError("timeout")
        try:
            with send(Request(url, method="HEAD"), timeout=min(5, remaining)) as response:
                status, headers = response.status, response.headers
        except HTTPError as exc:
            status, headers = exc.code, exc.headers
            exc.close()
        except TimeoutError:
            raise LocationError("timeout") from None
        except (URLError, OSError):
            raise LocationError("offline") from None
        if status not in (301, 302, 303, 307, 308) or not headers.get("Location"):
            raise LocationError("pin_missing")
        url = urljoin(url, headers["Location"])
        _maps_url_parts(url)
        try:
            return parse_location(url)
        except LocationError as exc:
            if exc.code != "pin_missing":
                raise
    raise LocationError("pin_missing")


def maps_link(coordinates):
    pin = validate_coordinates(coordinates.latitude, coordinates.longitude)
    return f"https://www.google.com/maps/search/?api=1&query={pin.latitude:.7f},{pin.longitude:.7f}"


class BrowserLocationSession:
    """One private loopback session; no browser permission is requested on load."""
    def __init__(self, labels=None):
        labels = labels or {}
        self.token = secrets.token_urlsafe(32)
        self.result = None
        session = self

        class Handler(BaseHTTPRequestHandler):
            def setup(self):
                super().setup()
                self.connection.settimeout(2)

            def log_message(self, *_args):
                pass

            def reply(self, status, content=b"", kind="text/plain; charset=utf-8"):
                self.send_response(status)
                self.send_header("Content-Type", kind)
                self.send_header("Content-Length", str(len(content)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("Referrer-Policy", "no-referrer")
                self.send_header("X-Content-Type-Options", "nosniff")
                self.send_header("Content-Security-Policy", "default-src 'none'; connect-src 'self'; "
                                 f"script-src 'nonce-{session.token}'; style-src 'nonce-{session.token}'; "
                                 "frame-ancestors 'none'; base-uri 'none'")
                self.end_headers()
                try:
                    self.wfile.write(content)
                except OSError:
                    pass

            def do_GET(self):
                if self.path != "/" + session.token or self.headers.get("Host") != session.host:
                    self.reply(404)
                    return
                self.reply(200, session.page.encode("utf-8"), "text/html; charset=utf-8")

            def do_POST(self):
                if (self.path != "/" + session.token or self.headers.get("Host") != session.host or
                        self.headers.get("Origin") != session.origin or
                        self.headers.get("X-Location-Token") != session.token):
                    self.reply(403)
                    return
                try:
                    size = int(self.headers.get("Content-Length", "0"))
                    if not 0 < size <= 1024 or session.result is not None:
                        raise ValueError()
                    data = json.loads(self.rfile.read(size))
                    if not isinstance(data, dict):
                        raise ValueError()
                    if data.get("error") in {"denied", "unavailable", "timeout"}:
                        result = LocationError(data["error"])
                    else:
                        result = validate_coordinates(data["latitude"], data["longitude"], data.get("accuracy"))
                except (ValueError, KeyError, TypeError):
                    self.reply(400)
                    return
                session.result = result
                self.reply(200, b"OK")

        self.server = HTTPServer(("127.0.0.1", 0), Handler)
        self.server.timeout = 0.25
        self.host = f"127.0.0.1:{self.server.server_port}"
        self.origin = "http://" + self.host
        self.url = self.origin + "/" + self.token
        title = html.escape(labels.get("title", "Clinic Location"))
        detail = html.escape(labels.get("detail", "Only use this while at your clinic. Allow browser location access, then return to the app to confirm the pin. No prescription data is sent here."))
        button = html.escape(labels.get("button", "Get Current Location"))
        waiting = json.dumps(labels.get("waiting", "Requesting location…"), ensure_ascii=False)
        done = json.dumps(labels.get("done", "Return to the app to confirm. You can close this tab."), ensure_ascii=False)
        failed = json.dumps(labels.get("failed", "Location unavailable or permission denied. Return to the app and enter your clinic pin manually."), ensure_ascii=False)
        self.page = f'''<!doctype html><html lang="{labels.get('language', 'en')}"><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>{title}</title>
<style nonce="{self.token}">body{{font:16px system-ui;background:#f6f7f9;color:#25272a;margin:32px}}
main{{max-width:520px;padding:24px;background:white;border:1px solid #cdd2d9;border-radius:12px;margin:auto}}
button{{background:white;border:1px solid #0960c7;color:#0960c7;padding:12px;border-radius:8px;font:inherit}}
button:focus-visible{{outline:3px solid #0960c7;outline-offset:3px}}</style>
<main><h1>{title}</h1><p>{detail}</p><button id="locate">{button}</button><p id="status" role="status" aria-live="polite"></p></main>
<script nonce="{self.token}">const button=document.getElementById('locate'),status=document.getElementById('status');
async function send(data){{try{{const r=await fetch(location.pathname,{{method:'POST',headers:{{'Content-Type':'application/json','X-Location-Token':'{self.token}'}},body:JSON.stringify(data)}});status.textContent=r.ok&&!data.error?{done}:{failed};}}catch{{status.textContent={failed};}}}}
button.addEventListener('click',()=>{{button.disabled=true;status.textContent={waiting};
if(!navigator.geolocation){{send({{error:'unavailable'}});return;}}
navigator.geolocation.getCurrentPosition(p=>send({{latitude:p.coords.latitude,longitude:p.coords.longitude,accuracy:p.coords.accuracy}}),e=>send({{error:e.code===1?'denied':e.code===3?'timeout':'unavailable'}}),{{enableHighAccuracy:true,timeout:20000,maximumAge:0}});}});</script></html>'''

    def close(self):
        self.server.server_close()


def detect_current_location(labels=None, *, cancel_event=None, open_browser=webbrowser.open, timeout=120):
    session = BrowserLocationSession(labels)
    try:
        if not open_browser(session.url):
            raise LocationError("unavailable")
        deadline = time.monotonic() + timeout
        while session.result is None:
            if cancel_event is not None and cancel_event.is_set():
                raise LocationError("cancelled")
            if time.monotonic() >= deadline:
                raise LocationError("timeout")
            session.server.handle_request()
        if isinstance(session.result, LocationError):
            raise session.result
        return session.result
    finally:
        session.close()
