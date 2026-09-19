"""Fictitious pins only; never starts a browser or reads production settings."""
import http.client
import json
import threading
from urllib.error import URLError

import pytest
import clinic_location as loc
from test_core import run_isolated


@pytest.mark.parametrize("value", [
    "12.345, 45.678", "https://www.google.com/maps/search/?api=1&query=12.345%2C45.678",
    "https://maps.google.com/?q=12.345,45.678",
    "https://www.google.com/maps/place/Test/@1,2,5z/data=!3d12.345!4d45.678",
])
def test_explicit_pins(value):
    assert loc.parse_location(value) == loc.Coordinates(12.345, 45.678)
    assert "query=12.3450000,45.6780000" in loc.maps_link(loc.parse_location(value))


@pytest.mark.parametrize("value", ["", "91,0", "0,-181", "nan,4",
    "https://evil.example/maps?q=1,2", "https://www.google.com.evil.example/maps?q=1,2",
    "http://maps.google.com/?q=1,2", "https://user:pass@www.google.com/maps?q=1,2",
    "https://www.google.com:8000/maps?q=1,2", "https://goo.gl/notmaps/abc"])
def test_reject_invalid_or_untrusted(value):
    with pytest.raises(loc.LocationError) as error:
        loc.resolve_location(value, opener=lambda *_a, **_k: pytest.fail("must not connect"))
    assert value not in str(error.value) if value else True


@pytest.mark.parametrize("lat,lng", [(True, 0), (0, False), (float("inf"),0),
    (0,float("nan")), ("",0), (None,0), (90.0001,0), (0,180.0001)])
def test_coordinate_bounds(lat, lng):
    with pytest.raises(loc.LocationError):
        loc.validate_coordinates(lat, lng)


def test_zero_and_boundaries_are_valid():
    assert loc.validate_coordinates(0, 0).latitude == 0
    assert loc.validate_coordinates(-90, 180).longitude == 180


def test_no_viewport_guessing():
    with pytest.raises(loc.LocationError, match="pin_missing"):
        loc.parse_location("https://www.google.com/maps/@12.345,45.678,10z")


class Response:
    status = 302
    def __init__(self, url): self.headers = {"Location": url}
    def __enter__(self): return self
    def __exit__(self, *args): pass


def test_short_link_resolution_and_no_hostile_redirect():
    calls = []
    def send(request, timeout):
        calls.append(request.full_url)
        assert request.method == "HEAD" and timeout <= 5
        assert not request.data and not request.get_header("X-api-key")
        return Response("https://www.google.com/maps?q=12.345,45.678")
    assert loc.resolve_location("https://maps.app.goo.gl/fictitious", opener=send).latitude == 12.345
    assert len(calls) == 1
    for hostile in ("http://www.google.com/maps?q=1,2", "https://127.0.0.1/maps?q=1,2",
                    "https://evil.example/maps?q=1,2"):
        calls.clear()
        def redirect(request, timeout):
            calls.append(request.full_url)
            return Response(hostile)
        with pytest.raises(loc.LocationError):
            loc.resolve_location("https://maps.app.goo.gl/fictitious", opener=redirect)
        assert len(calls) == 1


def test_short_link_offline_and_redirect_limit():
    def offline(*args, **kwargs): raise URLError("secret should not appear")
    with pytest.raises(loc.LocationError, match="offline") as error:
        loc.resolve_location("https://maps.app.goo.gl/fictitious", opener=offline)
    assert "secret" not in str(error.value)
    calls = []
    def loop(*args, **kwargs):
        calls.append(1)
        return Response("https://maps.app.goo.gl/loop")
    with pytest.raises(loc.LocationError, match="pin_missing"):
        loc.resolve_location("https://maps.app.goo.gl/fictitious", opener=loop)
    assert len(calls) == 5


def request_session(session, method, body=None, headers=None, path=None):
    client = http.client.HTTPConnection(session.host, timeout=3)
    try:
        client.request(method, path or "/" + session.token, body=body, headers=headers or {})
        response = client.getresponse()
        return response.status, response.read(), dict(response.getheaders())
    finally:
        client.close()


def test_loopback_helper_security_and_single_result(capsys):
    session = loc.BrowserLocationSession()
    done = threading.Event()
    def serve():
        while not done.is_set(): session.server.handle_request()
    thread = threading.Thread(target=serve)
    thread.start()
    try:
        status, page, headers = request_session(session, "GET")
        assert status == 200 and headers["Cache-Control"] == "no-store"
        assert "default-src 'none'" in headers["Content-Security-Policy"]
        assert b"button.addEventListener('click'" in page
        assert b"getCurrentPosition" in page and b"https://" not in page
        assert request_session(session,"GET",path="/wrong")[0] == 404
        body = json.dumps({"latitude":12.345,"longitude":45.678,"accuracy":50})
        base = {"Origin":session.origin,"X-Location-Token":session.token,"Content-Type":"application/json"}
        for bad in ({}, {**base,"Origin":"https://evil.example"}, {**base,"X-Location-Token":"wrong"},
                    {**base,"Host":"evil.example"}):
            assert request_session(session,"POST",body,bad)[0] == 403
            assert session.result is None
        assert request_session(session,"POST","not-json",base)[0] == 400
        assert request_session(session,"POST","[]",base)[0] == 400
        assert request_session(session,"POST","null",base)[0] == 400
        assert request_session(session,"POST",json.dumps({"latitude":91,"longitude":0}),base)[0] == 400
        assert request_session(session,"POST",body,base)[0] == 200
        assert session.result == loc.Coordinates(12.345,45.678,50)
        assert request_session(session,"POST",body,base)[0] == 400
    finally:
        done.set()
        thread.join(3)
        session.close()
    assert not thread.is_alive()
    assert not capsys.readouterr().err


@pytest.mark.parametrize("code", ["denied", "unavailable", "timeout"])
def test_browser_permission_errors(code):
    writers = []
    def browser(url):
        from urllib.parse import urlsplit
        parts = urlsplit(url)
        def send():
            client = http.client.HTTPConnection(parts.netloc, timeout=3)
            try:
                client.request("POST",parts.path,json.dumps({"error":code}),headers={
                    "Origin":"http://"+parts.netloc,"X-Location-Token":parts.path[1:]})
                assert client.getresponse().status == 200
            finally: client.close()
        thread = threading.Thread(target=send)
        thread.start()
        writers.append(thread)
        return True
    with pytest.raises(loc.LocationError, match=code):
        loc.detect_current_location(open_browser=browser,timeout=3)
    for thread in writers: thread.join(3)


def test_detection_cancel_timeout_and_no_browser():
    event = threading.Event(); event.set()
    with pytest.raises(loc.LocationError, match="cancelled"):
        loc.detect_current_location(open_browser=lambda _url:True,cancel_event=event)
    with pytest.raises(loc.LocationError, match="timeout"):
        loc.detect_current_location(open_browser=lambda _url:True,timeout=0)
    with pytest.raises(loc.LocationError, match="unavailable"):
        loc.detect_current_location(open_browser=lambda _url:False)


def test_detection_success_and_expired_session():
    addresses, writers = [], []
    def browser(url):
        from urllib.parse import urlsplit
        parts = urlsplit(url)
        addresses.append(parts)
        def send():
            client = http.client.HTTPConnection(parts.netloc, timeout=3)
            try:
                client.request("POST",parts.path,json.dumps({"latitude":12.345,"longitude":45.678}),headers={
                    "Origin":"http://"+parts.netloc,"X-Location-Token":parts.path[1:]})
                assert client.getresponse().status == 200
            finally: client.close()
        thread = threading.Thread(target=send); thread.start(); writers.append(thread)
        return True
    assert loc.detect_current_location(open_browser=browser,timeout=3)==loc.Coordinates(12.345,45.678)
    for thread in writers: thread.join(3)
    client = http.client.HTTPConnection(addresses[0].netloc, timeout=1)
    try:
        with pytest.raises(OSError): client.request("GET",addresses[0].path)
    finally: client.close()


def test_export_coordinate_snapshot_and_invalid_location_guard(tmp_path):
    run_isolated(tmp_path, '''
from types import SimpleNamespace
import main
from main import App
from qr_utils import Clinic, Prescription
rx=Prescription(clinic=Clinic(latitude="12.345",longitude="45.678",include_location=True))
calls=[]
ui=SimpleNamespace(_export_busy=False,_closing=False,
    _prepare_full_document=lambda action:(rx,{}),_upload_export_link=lambda op:calls.append(op))
ui._set_export_busy=lambda value:setattr(ui,"_export_busy",value)
op=App._start_cloud_export(ui,"Export")
assert op["payload"]["latitude"]==12.345
rx.clinic.latitude="2"; rx.clinic.include_location=False
assert op["payload"]["latitude"]==12.345
assert App._start_cloud_export(ui,"Export") is None
ui._export_busy=False
rx.clinic.include_location=True; rx.clinic.latitude="nan"
messages=[]
main.messagebox.showerror=lambda *args,**kwargs:messages.append(args)
assert App._start_cloud_export(ui,"Export") is None
assert messages and not ui._export_busy and len(calls)==1
''')


def test_encrypted_persistence_optional_payload_and_snapshot(tmp_path):
    run_isolated(tmp_path, '''
import copy
from config import Config, config, CONFIG_PATH
from qr_utils import Clinic, Prescription, Doctor, DrugItem
from clinic_location import LocationError
config.set_clinic(name="Fictitious clinic",latitude="12.3450000",longitude="45.6780000",include_location=False)
rx=Prescription(clinic=Clinic(**config.get_clinic()),doctor=Doctor(name="طبيب تجريبي"),
                drugs=[DrugItem(brand_name="Fictitious medicine")])
assert "latitude" not in rx.to_cloud_payload() and "longitude" not in rx.to_cloud_payload()
assert "12.3450000" not in CONFIG_PATH.read_text()
assert "45.6780000" not in CONFIG_PATH.read_text()
reloaded=Config()
assert reloaded.get_clinic()["latitude"]=="12.3450000"
rx.clinic.include_location=True
snapshot=copy.deepcopy(rx.to_cloud_payload())
assert snapshot["latitude"]==12.345 and snapshot["longitude"]==45.678
assert "clinic" not in snapshot and "logo_path" not in str(snapshot)
assert "latitude" not in str(rx.to_qr_payload())
rx.clinic.latitude="1"
assert snapshot["latitude"]==12.345
rx.clinic.latitude="nan"
try: rx.to_cloud_payload()
except LocationError: pass
else: raise AssertionError("invalid enabled pin must fail")
''')


def test_settings_confirmation_async_staleness_and_close(tmp_path):
    run_isolated(tmp_path, '''
import main, clinic_location as loc, i18n as I
from config import config
app=main.App(); app.update()
errors=[]
app.report_callback_exception=lambda *args:errors.append(args)
window=main.SettingsWindow(app); app.update()
window._show_section("clinic"); app.update()
jobs=[]
app.submit_background=lambda worker,success,on_error=None,**kwargs:jobs.append((worker,success,on_error))
main.messagebox.askyesno=lambda *args,**kwargs:True
main.messagebox.showerror=lambda *args,**kwargs:errors.append(args)
main.messagebox.showinfo=lambda *args,**kwargs:None
main.webbrowser.open=lambda _url:True
window.location_input_var.set("12.345,45.678")
window.check_clinic_location()
assert window._location_busy and not window.include_location_var.get()
worker,success,failure=jobs.pop()
pin=worker(); success(pin)
assert not window._location_busy and window.latitude_var.get()=="12.3450000"
assert not window.include_location_var.get()
window.location_input_var.set("1,2"); window.check_clinic_location()
worker,success,failure=jobs.pop()
window.latitude_var.set("3"); success(worker())
assert window.latitude_var.get()=="3"  # no stale overwrite
window.remove_clinic_location()
assert not window.latitude_var.get() and not window.include_location_var.get()
window.location_input_var.set("1,2"); window.check_clinic_location()
worker,success,failure=jobs.pop()
event=window._location_cancel_event
window.destroy()
assert event.is_set()
success(loc.Coordinates(1,2)); failure(loc.LocationError("denied"))
assert not errors
window=main.SettingsWindow(app); app.update()
window.latitude_var.set("12.345"); window.longitude_var.set("45.678")
window.include_location_var.set(True)
window.save()
assert config.get_clinic()["include_location"] is True
rx=app.collect()
assert rx.clinic.latitude=="12.3450000" and rx.clinic.include_location
for language in ("en","ar"):
    I.set_lang(language)
    for key in ("location_confirm","location_denied","location_browser_detail"):
        assert I.t(key)!=key
app.destroy()
assert not errors
''')
