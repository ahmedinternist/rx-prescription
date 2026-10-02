"""Behavioral guards for 7.3 optimizations, with isolated application data."""
import json
from concurrent.futures import ThreadPoolExecutor
from unittest.mock import patch

import pytest
from patient_history import PatientHistory
from test_core import run_isolated


def test_name_wrapping_reserves_actions_and_save_feedback(tmp_path):
    run_isolated(tmp_path, '''
import tkinter as tk
import customtkinter as ctk
from types import SimpleNamespace
from main import wrap_card_labels, VisualButton, App, GOOD, DANGER
root=tk.Tk();root.geometry("420x400")
try:
    header=ctk.CTkFrame(root);header.pack(fill="x")
    primary=ctk.CTkLabel(header,text="Long medicine name "*12);primary.pack(side="left")
    scientific=ctk.CTkLabel(header,text="Scientific ingredient "*12);scientific.pack(side="left")
    actions=ctk.CTkFrame(header,width=40);actions.pack(side="right")
    button=VisualButton(actions,text="+");button.pack()
    wrap_card_labels(header);root.update()
    assert primary.cget("wraplength")<200
    assert primary.winfo_height()>28
    assert actions.winfo_rootx()+actions.winfo_width()<=header.winfo_rootx()+header.winfo_width()
    assert button.winfo_viewable()
    label=ctk.CTkLabel(root,text="")
    owner=SimpleNamespace(template_save_status=label)
    owner._set_save_feedback=lambda scope,state:App._set_save_feedback(owner,scope,state)
    assert App._save_with_feedback(owner,"template",lambda:"id")=="id"
    assert label.cget("text")=="Saved" and label.cget("text_color")==GOOD
    assert App._save_with_feedback(owner,"template",lambda:False) is False
    assert label.cget("text")=="Failed" and label.cget("text_color")==DANGER
    try:App._save_with_feedback(owner,"template",lambda:(_ for _ in ()).throw(OSError("test")))
    except OSError:pass
    else:raise AssertionError("Save errors must remain visible to callers")
    assert label.cget("text")=="Failed"
finally:root.destroy()
''')


def test_viewport_buffer_and_settled_glass_redraw():
    import time
    from types import SimpleNamespace
    from main import App, GlassFrame
    canvas = SimpleNamespace(winfo_rooty=lambda: 100, winfo_height=lambda: 500)
    owner = SimpleNamespace(scroll=SimpleNamespace(_parent_canvas=canvas))
    card = SimpleNamespace(winfo_ismapped=lambda: True, winfo_rooty=lambda: 750)
    assert App._card_near_viewport(owner, card)
    card.winfo_rooty = lambda: 770
    assert not App._card_near_viewport(owner, card)
    scheduled = []
    panel = SimpleNamespace(_glass_job=None, winfo_exists=lambda: True,
        winfo_toplevel=lambda: SimpleNamespace(_glass_resize_until=time.monotonic()+.2),
        after=lambda delay, callback: scheduled.append(delay) or "pending",
        _paint_glass=lambda: None)
    GlassFrame._paint_glass(panel)
    assert panel._glass_job == "pending" and 150 <= scheduled[0] <= 201


def test_shared_visual_styles_and_focus_keep_geometry(tmp_path):
    run_isolated(tmp_path, '''
import tkinter as tk
from main import VisualEntry, VisualComboBox, VisualButton, GlassFrame, CARD_RADIUS, ICON_BUTTON_SIZE, ACCENT, DANGER, LINE, _edit_icon
root=tk.Tk()
try:
    card=GlassFrame(root,border_width=1,corner_radius=9)
    assert card.cget("corner_radius")==CARD_RADIUS
    for field in (VisualEntry(root,width=240,height=36,border_color=LINE),
                  VisualComboBox(root,width=240,height=36,border_color=LINE)):
        geometry=(field.cget("width"),field.cget("height"),field.cget("border_width"))
        field._field_focused()
        assert field.cget("border_color")==ACCENT
        field._field_focused()
        field._field_blurred()
        assert field.cget("border_color")==LINE
        assert geometry==(field.cget("width"),field.cget("height"),field.cget("border_width"))
        field.configure(border_color=DANGER)
        field._field_focused();field._field_blurred()
        assert field.cget("border_color")==DANGER
    for symbol in ("+","🗑","★"):
        button=VisualButton(root,text=symbol,width=36,height=34)
        assert button.cget("width")==button.cget("height")==ICON_BUTTON_SIZE
        assert button.cget("fg_color")=="transparent" and button.cget("border_width")==0
        button.configure(text_color=DANGER)
        assert button.cget("text")==symbol
    edit=VisualButton(root,text="",image=_edit_icon(18),width=28,height=28)
    assert edit.cget("width")==edit.cget("height")==ICON_BUTTON_SIZE
finally:root.destroy()
''')


def test_values_only_template_summary_single_editor_and_inline_mapping(tmp_path):
    run_isolated(tmp_path, '''
import time
from main import App, MappingMedicineList, CARD, TEXT, LINE, ACCENT
assert App._treatment_regimen_summary({"dosage":"10 mg","frequency":"مرتان يومياً",
    "duration":"7 days","notes":"after food"})=="10 mg   ·   مرتان يومياً   ·   7 days   ·   after food"
assert App._treatment_regimen_summary({"dosage":""})==""
app=App()
try:
    end=time.monotonic()+15
    while app._database_loading and time.monotonic()<end:
        app.update();time.sleep(.005)
    first=app.rows[0]
    second=app.add_row()
    assert not first.expanded and second.expanded
    first.set_expanded(True)
    assert first.expanded and not second.expanded
    first.set_expanded(True)
    assert not second.details.winfo_manager()
    second.set_expanded(True)
    assert not first.expanded and second.expanded
    view=MappingMedicineList(app,bg=CARD,fg=TEXT,line=LINE,select_bg=ACCENT,select_fg="white")
    view.pack(fill="x")
    view.insert("end","Synthetic", "Class", trade_name="Trade",status="Confirmed")
    app.update();view.refresh()
    items=[item for item in view.canvas.find_all() if view.canvas.type(item)=="text"]
    displayed={view.canvas.itemcget(item,"text"):view.canvas.coords(item) for item in items}
    assert "Synthetic" in displayed and "Trade" in displayed
    assert displayed["Trade"][0]>displayed["Synthetic"][0]
    assert abs(view._trade_font.cget("size"))==22
    view.delete(0,"end")
    assert not view._trade_names
finally:app.destroy()
''')


def test_history_cache_isolation_external_replacement_and_concurrent_saves(tmp_path):
    path = tmp_path / "history.json"
    first, second = PatientHistory(path), PatientHistory(path)
    patient = {"name": "مريض تجريبي", "age": "40"}
    medicine = [{"generic_name": "Synthetic", "notes": "اختبار"}]
    record = first.save_prescription(patient, medicine)
    returned = first.get(record["id"])
    returned["prescriptions"][0]["drugs"][0]["notes"] = "changed externally in memory"
    assert first.get(record["id"])["prescriptions"][0]["drugs"][0]["notes"] == "اختبار"
    second.save_prescription(patient, medicine, record["id"])
    assert len(first.get(record["id"])["prescriptions"]) == 2
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [executor.submit((first if i % 2 else second).save_prescription,
                                   patient, medicine, record["id"]) for i in range(8)]
        for future in futures:
            future.result()
    assert len(PatientHistory(path).get(record["id"])["prescriptions"]) == 10
    assert json.loads(path.read_text())["format"] == "dpapi-v1"
    assert "Synthetic" not in path.read_text()
    path.unlink()
    assert first.get(record["id"]) is None


def test_failed_atomic_history_write_keeps_original_and_cache(tmp_path):
    history = PatientHistory(tmp_path / "history.json")
    saved = history.save_patient({"name": "Synthetic"})
    original = history.path.read_bytes()
    with patch("patient_history.os.replace", side_effect=OSError("synthetic disk failure")):
        with pytest.raises(OSError):
            history.save_patient({"name": "Changed"}, saved["id"])
    assert history.path.read_bytes() == original
    assert history.get(saved["id"])["name"] == "Synthetic"
    assert not list(tmp_path.glob("*.tmp"))


def test_deferred_cache_and_search_contract(tmp_path):
    run_isolated(tmp_path, '''
import csv, sqlite3, os
from pathlib import Path
from drug_db import DrugDatabase, _norm
source = Path(os.environ["RX_APP_DATA_DIR"]) / "drugs.csv"
source.write_text("generic_name,brand_name,strength,form,category\\n"
    "Paracetamol,Alpha,10mg,tablet,Test\\n"
    "باراسيتامول,ألفا,10mg,tablet,Test\\n"
    "Paracetamol,Beta,10mg,tablet,Test\\n"
    "Amoxicillin,Alpha,10mg,tablet,Test\\n", encoding="utf-8")
db = DrugDatabase(str(source), defer_load=True)
assert not db.cache_path.exists()
assert db.load() == 4
drugs = list(db.drugs)
for method, field in ((db.search, "combined"), (db.search_scientific, "scientific"),
                      (db.search_prescribable, "brand")):
    for query in ("a", "al", "ALPHA", "أ", "ألف", "para", "بارا", "paracatmol", " Beta "):
        q = _norm(query)
        indexed = []
        for i, drug in enumerate(drugs):
            generic, brand = _norm(drug.generic_name), _norm(drug.brand_name)
            value = (generic if field == "scientific" else brand or generic if field == "brand"
                else " ".join(filter(None, (generic, brand, _norm(drug.strength),
                                             _norm(drug.form), _norm(drug.category)))))
            indexed.append((value, i, drug))
        prefix = sorted(item for item in indexed if q <= item[0] < q+"\\uffff")
        rest = sorted(item for item in indexed if q in item[0] and not(q <= item[0] < q+"\\uffff"))
        assert [d.to_dict() for d in method(query, 20)] == [t[2].to_dict() for t in prefix+rest]
with db._connect() as connection:
    assert connection.execute("SELECT 1 FROM sqlite_master WHERE name='idx_drugs_search'").fetchone()
''')


def test_glass_slice_matches_full_image_and_reuses_cards(tmp_path):
    run_isolated(tmp_path, '''
from main import glass_panel_image, App
from PIL import ImageChops
import config as cfg, time
size = (240, 600, 12, 1)
full = glass_panel_image(size, "panel")
clip = (0, 100, 240, 350)
sliced = glass_panel_image(size, "panel", clip=clip)
assert sliced.size == (240, 250)
extrema = ImageChops.difference(full.crop(clip), sliced).getextrema()
assert max(v[1] for v in extrema) <= 2
cfg.config.data["treatment_templates"] = [{"id":"t1", "disease":"Synthetic", "medications":[
    {"generic_name":"Synthetic", "brand_name":"Test"}]}]
app = App()
try:
    end = time.monotonic()+15
    while app._database_loading and time.monotonic()<end:
        app.update(); time.sleep(.01)
    assert not app._database_loading
    app._ensure_page_built("treatment_templates")
    app.show_page("treatment_templates")
    app.refresh_saved_treatment_templates()
    first = app._treatment_saved_card_widgets[0]
    app.refresh_saved_treatment_templates()
    assert app._treatment_saved_card_widgets[0] is first
    app.show_page("medications")
    app.update()
    app.show_page("medications")
    assert app.active_page == "medications"
finally:
    app.destroy()
''')


def test_openfda_worker_uses_queue_not_tk(tmp_path):
    run_isolated(tmp_path, '''
import queue, threading
from types import SimpleNamespace
from main import App
delivered = []
app = SimpleNamespace(_closing=False, _background_callbacks=queue.SimpleQueue(),
    _lookup_openfda_with_cache=lambda medicines, force: (["synthetic"], [], "today"),
    _show_openfda_results=lambda *args: delivered.append(threading.get_ident()))
worker = threading.Thread(target=App._lookup_openfda_worker, args=(app,["Synthetic"]))
worker.start(); worker.join()
assert not delivered
app._background_callbacks.get_nowait()()
assert delivered == [threading.get_ident()]
app._closing=True
App._lookup_openfda_worker(app,["Synthetic"])
assert app._background_callbacks.empty()
''')


def test_async_save_captures_snapshot_and_ignores_newer_edits(tmp_path):
    run_isolated(tmp_path, '''
import concurrent.futures, queue, threading, time
from types import SimpleNamespace
from unittest.mock import patch
from main import App
from qr_utils import DrugItem
release = threading.Event()
stored = []
def save(patient, drugs, record_id):
    assert release.wait(5)
    stored.append((patient, drugs, record_id, threading.get_ident()))
    return {"id":"patient-one", "prescriptions":[]}
state = {"signature":"old"}
value = lambda text: SimpleNamespace(get=lambda: text)
medicine = DrugItem(generic_name="Synthetic", notes="original")
executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
app = SimpleNamespace(_history_save_future=None, _closing=False,
    _loaded_patient_id="patient-one", _draft_signature=lambda: state["signature"],
    patient_vars={"name":value("Synthetic patient"),"age":value("40"),"sex":value("M")},
    rows=[SimpleNamespace(get_data=lambda: medicine)], _persistence_executor=executor,
    patient_history=SimpleNamespace(save_prescription=save),
    _background_callbacks=queue.SimpleQueue(), refresh_patient_history=lambda: None)
try:
    with patch("main.messagebox.showinfo"):
        App.save_prescription_for_patient(app)
        App.save_prescription_for_patient(app)
        state["signature"]="new"
        medicine.notes="edited while saving"
        release.set()
        end = time.monotonic()+5
        while app._background_callbacks.empty() and time.monotonic()<end:
            time.sleep(.005)
        app._background_callbacks.get_nowait()()
    assert len(stored)==1
    assert stored[0][1][0]["notes"]=="original"
    assert stored[0][3]!=threading.get_ident()
    assert app._loaded_patient_id=="patient-one"
    assert app._history_save_future is None
finally:
    release.set(); executor.shutdown(wait=True)
''')


def test_template_batches_retain_all_values_order_and_collapse(tmp_path):
    run_isolated(tmp_path, '''
import time
from types import SimpleNamespace
from main import App
app=App()
try:
    end=time.monotonic()+15
    while app._database_loading and time.monotonic()<end:
        app.update(); time.sleep(.005)
    choices=[]
    for i in range(4):
        medicine={"generic_name":f"Synthetic {i}","brand_name":f"Trade {i}",
                  "dosage":"test", "frequency":"مرة واحدة يومياً", "duration":"٧ أيام", "notes":"اختبار"}
        choices.append(([medicine],SimpleNamespace(get=lambda:0),SimpleNamespace(get=lambda:True)))
    dialog=SimpleNamespace(grab_release=lambda:None,destroy=lambda:None)
    original_row=app.rows[0]
    App._apply_treatment_selection(app,choices,True,dialog)
    assert app._applying_template
    end=time.monotonic()+15
    while app._applying_template and time.monotonic()<end:
        app.update(); time.sleep(.005)
    assert not app._applying_template
    assert [row.get_data().generic_name for row in app.rows]==[f"Synthetic {i}" for i in range(4)]
    assert all(row.get_data().notes=="اختبار" for row in app.rows)
    assert [row.number_badge.cget("text") for row in app.rows]==["1.","2.","3.","4."]
    assert [row.expanded for row in app.rows]==[False,False,False,True]
    assert app.rows[0] is original_row
    reused=list(app.rows)
    App._apply_treatment_selection(app,choices[:2],True,dialog)
    end=time.monotonic()+15
    while app._applying_template and time.monotonic()<end:
        app.update(); time.sleep(.005)
    assert app.rows==reused[:2]
    assert not reused[2].winfo_exists()
    assert [row.get_data().brand_name for row in app.rows]==["Trade 0","Trade 1"]
finally:
    app.destroy()
''')


def test_favorite_card_names_only_preserves_editor_data(tmp_path):
    run_isolated(tmp_path, '''
import customtkinter as ctk
from main import App
app=App()
try:
    app.show_page("favorites")
    favorite={"id":"compact-test", "brand_name":"Synthetic Trade",
              "generic_name":"Synthetic Scientific", "category":"Synthetic Category",
              "dosage":"UNIQUE DOSE", "frequency":"UNIQUE FREQUENCY",
              "duration":"UNIQUE DURATION", "notes":"UNIQUE NOTE"}
    before=dict(favorite)
    app._render_favorite_card(0,0,favorite)
    card=app._favorite_card_widgets[favorite["id"]]["card"]
    def labels(widget):
        result=[]
        for child in widget.winfo_children():
            if isinstance(child,ctk.CTkLabel):
                result.append(child.cget("text"))
            result.extend(labels(child))
        return result
    text=" ".join(labels(card))
    assert "Synthetic Trade" in text and "Synthetic Scientific" in text
    assert "Synthetic Category" in text
    assert not any(favorite[key] in text for key in ("dosage","frequency","duration","notes"))
    assert favorite==before
    actions=[child for child in card.winfo_children()
             if child.winfo_manager()=="grid" and int(child.grid_info()["row"])==1][0]
    category=[child for child in actions.winfo_children()
              if isinstance(child,ctk.CTkLabel)][0]
    assert category.pack_info()["side"]=="right"
finally:
    app.destroy()
''')


def test_unreadable_history_blocks_writes(tmp_path):
    run_isolated(tmp_path, '''
import json
from pathlib import Path
import config
from patient_history import PatientHistory, PatientHistoryError
from security import protect
path=config.APP_DIR / "broken.json"
cases=[b"{truncated", b'{"format":"unknown"}', b'{"format":"dpapi-v1","data":"invalid"}',
       json.dumps({"format":"dpapi-v1","data":protect(b'{"records":null}')}).encode()]
for original in cases:
    path.write_bytes(original)
    history=PatientHistory(path)
    for action in (lambda:history.search(), lambda:history.save_patient({"name":"Synthetic"}),
                   lambda:history._save([]), lambda:history.backup_bytes()):
        try: action()
        except PatientHistoryError: pass
        else: raise AssertionError("Unreadable history accepted")
        assert path.read_bytes()==original
''')


def test_same_name_patients_remain_distinct(tmp_path):
    run_isolated(tmp_path, '''
import config
from patient_history import PatientHistory
h=PatientHistory(config.APP_DIR / "identity.json")
first=h.save_prescription({"name":"Same Synthetic","age":"70","sex":"M"},[{"brand_name":"Old Drug"}])
second=h.save_patient({"name":"Same Synthetic","age":"20","sex":"F"})
assert first["id"]!=second["id"] and len(h.search())==2
assert h.get(first["id"])["age"]=="70"
h.save_prescription({"name":"Same Synthetic","age":"20","sex":"F"},[{"brand_name":"New Drug"}],second["id"])
assert h.get(first["id"])["prescriptions"][0]["drugs"][0]["brand_name"]=="Old Drug"
assert h.get(second["id"])["prescriptions"][0]["drugs"][0]["brand_name"]=="New Drug"
try: h.save_patient({"name":"Synthetic"},"deleted-id")
except ValueError: pass
else: raise AssertionError("Missing ID silently merged")
''')


def test_backups_include_and_restore_patient_history(tmp_path):
    run_isolated(tmp_path, '''
import config,json,zipfile
from patient_history import PatientHistory,HISTORY_PATH,PatientHistoryError
from unittest.mock import patch
h=PatientHistory()
record=h.save_prescription({"name":"Backup Synthetic"},[{"brand_name":"Backup Drug"}])
original=HISTORY_PATH.read_bytes()
backup=config.APP_DIR / "full.rxbackup"
config.config.create_backup(str(backup))
with zipfile.ZipFile(backup) as z: assert z.read("patient_history.json")==original
h.save_patient({"name":"Later Synthetic"})
config.config.restore_backup(str(backup))
assert HISTORY_PATH.read_bytes()==original and len(h.search())==1
legacy=config.APP_DIR / "old.rxbackup"
broken=config.APP_DIR / "bad.rxbackup"
with zipfile.ZipFile(backup) as source:
    for target,history in ((legacy,None),(broken,b"{bad")):
        with zipfile.ZipFile(target,"w") as z:
            for name in source.namelist():
                if name!="patient_history.json": z.writestr(name,source.read(name))
            if history is not None: z.writestr("patient_history.json",history)
config.config.restore_backup(str(legacy))
assert HISTORY_PATH.read_bytes()==original
settings=config.CONFIG_PATH.read_bytes()
try: config.config.restore_backup(str(broken))
except PatientHistoryError: pass
else: raise AssertionError("Bad history restored")
assert HISTORY_PATH.read_bytes()==original and config.CONFIG_PATH.read_bytes()==settings
config.config.set("auto_backup_enabled",True)
config.config.set("last_backup_at","")
automatic=config.config.maybe_create_automatic_backup()
with zipfile.ZipFile(automatic) as z: assert z.read("patient_history.json")==original
# Fail once after history replacement and verify rollback of every replaced file.
h.save_patient({"name":"Must survive failed restore"})
before_history=HISTORY_PATH.read_bytes(); before_settings=config.CONFIG_PATH.read_bytes()
replace=config._replace_bytes
failed=[False]
def fail_once(path,data):
    if path==config.CONFIG_PATH and not failed[0]:
        failed[0]=True; raise OSError("Synthetic restore failure")
    replace(path,data)
with patch("config._replace_bytes",side_effect=fail_once):
    try: config.config.restore_backup(str(backup))
    except OSError: pass
    else: raise AssertionError("Restore did not fail")
assert HISTORY_PATH.read_bytes()==before_history and config.CONFIG_PATH.read_bytes()==before_settings
''')


def test_invalid_cache_recovery_and_serialized_database_mutations(tmp_path):
    run_isolated(tmp_path, '''
import config,sqlite3
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from drug_db import DrugDatabase
path=config.APP_DIR / "medicine.csv"
path.write_text("generic_name,brand_name\\nScientific,Trade\\n",encoding="utf-8")
db=DrugDatabase(str(path),defer_load=True)
db.cache_path.write_bytes(b"broken sqlite")
assert db.load()==1
db.cache_path.unlink()
c=sqlite3.connect(db.cache_path)
c.execute("CREATE TABLE unexpected(x)"); c.commit(); c.close()
assert db.load()==1
other=DrugDatabase(str(path))
def add(i):
    source=db if i%2 else other
    return source.add_confirmed_medicine("Synthetic "+str(i),"","")
# Use CSV merges to exercise full read/modify/write transactions across instances.
def merge(i):
    source=config.APP_DIR / (str(i)+".csv")
    source.write_text("generic_name,brand_name\\nScientific "+str(i)+",Trade "+str(i)+"\\n")
    return (db if i%2 else other).import_file(str(source),replace=False)
with ThreadPoolExecutor(max_workers=4) as pool: list(pool.map(merge,range(12)))
assert db.count()==13 and len(list(other.drugs))==13
assert not list(path.parent.glob(path.name+".*.tmp"))
''')


def test_startup_recovers_invalid_cache_and_import_guard(tmp_path):
    run_isolated(tmp_path, '''
import config,time
from unittest.mock import patch
from main import App
config.ensure_seed_db()
cache=config.DEFAULT_DB_PATH.with_suffix(".csv.sqlite3")
cache.write_bytes(b"invalid cache")
errors=[]
with patch("main.messagebox.showerror",side_effect=lambda *a,**k:errors.append(a)):
    app=App()
    try:
        end=time.monotonic()+15
        while app._database_loading and time.monotonic()<end:
            app.update(); time.sleep(.005)
        assert not app._database_loading and not errors
        assert app.db.count()>0
        app._database_import_busy=True
        with patch("main.filedialog.askopenfilename") as choose:
            app.import_db()
            choose.assert_not_called()
    finally: app.destroy()
''')


def test_settings_atomic_failure_and_unchanged_save(tmp_path):
    run_isolated(tmp_path, '''
import config, json
from unittest.mock import patch
settings=config.Config()
settings.set("language", "ar")
original=config.CONFIG_PATH.read_bytes()
mtime=config.CONFIG_PATH.stat().st_mtime_ns
settings.save()
assert config.CONFIG_PATH.stat().st_mtime_ns==mtime
settings.data["language"]="en"
with patch("config.os.replace", side_effect=OSError("synthetic failure")):
    try:
        settings.save()
    except OSError:
        pass
    else:
        raise AssertionError("disk failure must be reported")
assert config.CONFIG_PATH.read_bytes()==original
assert not list(config.CONFIG_PATH.parent.glob("config.json.*.tmp"))
settings.save()
assert config.Config().language=="en"
assert json.loads(config.CONFIG_PATH.read_text())["format"]=="dpapi-v1"
''')
