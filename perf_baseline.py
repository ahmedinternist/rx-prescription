"""Isolated synthetic v7.2 audit harness; no network calls or real user data.

Run seed, core, ui, startup or imports with --output under project/output/.
Generated fixtures/results are NOT production files or packaging changes.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import platform
import statistics
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
parser = argparse.ArgumentParser()
parser.add_argument("mode", choices=("seed", "core", "ui", "startup", "imports"))
parser.add_argument("--output", required=True)
args = parser.parse_args()
OUTPUT = Path(args.output).resolve()
if not OUTPUT.is_relative_to(ROOT / "output"):
    raise SystemExit("Audit output must be inside the project's output directory")
OUTPUT.mkdir(parents=True, exist_ok=True)
os.environ["RX_APP_DATA_DIR"] = str(OUTPUT / "isolated-data")
RESULTS = {}


def save_results(name):
    target = OUTPUT / (name + ".json")
    temporary = target.with_suffix(".tmp")
    with temporary.open("w", encoding="utf-8") as stream:
        json.dump(RESULTS, stream, ensure_ascii=False, indent=2)
        stream.flush()
        os.fsync(stream.fileno())
    os.replace(temporary, target)


def measure(name, function, count=5):
    samples = []
    result = None
    for _ in range(count):
        start = time.perf_counter()
        result = function()
        samples.append(round((time.perf_counter()-start)*1000, 3))
    ordered = sorted(samples)
    RESULTS[name] = {"samples_ms": samples, "median_ms": statistics.median(samples),
                     "max_ms": max(samples)}
    print(name, RESULTS[name]["median_ms"], "ms", flush=True)
    save_results(args.mode)
    return result


def seed():
    import config as cfg
    import drug_db as dbmod
    from patient_history import PatientHistory
    drug_rows = []
    for i in range(22000):
        arabic = i % 7 == 0
        drug_rows.append(dbmod.Drug(
            generic_name=("باراسيتامول تجريبي" if arabic else "PerfIngredient") + f" {i:05d}",
            brand_name=("اختبار" if arabic else "PerfBrand") + f" {i:05d}",
            strength="10 mg", form="tablet", therapeutic_group="gastrointestinal",
            detailed_class="PPI (Proton Pump Inhibitor)", mapping_status="confirmed"))
    cfg.DEFAULT_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with cfg.DEFAULT_DB_PATH.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=dbmod.DRUG_FIELDS)
        writer.writeheader()
        writer.writerows(drug.to_dict() for drug in drug_rows)
    dbmod.DrugDatabase(str(cfg.DEFAULT_DB_PATH))
    medication = {"generic_name": "PerfIngredient 00001", "brand_name": "PerfBrand 00001",
                  "dosage": "SYNTHETIC 10 mg", "frequency": "مرة واحدة يومياً",
                  "duration": "٧ أيام", "notes": "اختبار فقط — not for prescribing", "quantity": ""}
    cfg.config.data["treatment_templates"] = [
        {"id": f"template-{i}", "disease": f"Synthetic disease {i:04d}",
         "medications": [dict(medication, generic_name=f"PerfIngredient {j:05d}",
                              brand_name=f"PerfBrand {j:05d}") for j in range(1, 11)]}
        for i in range(550)]
    cfg.config.data["medication_favorites"] = [
        dict(medication, id=f"favorite-{i}", pinned=i % 5 == 0, category="Synthetic",
             generic_name=f"PerfIngredient {i:05d}", brand_name=f"PerfBrand {i:05d}")
        for i in range(600)]
    cfg.config.data["auto_backup_enabled"] = False
    cfg.config.save()
    records = []
    for i in range(2200):
        prescriptions = [
            {"id": f"rx-{i}-{j}", "saved_at": "2026-01-01T00:00:00+00:00",
             "drugs": [dict(medication) for _ in range(5)]}
            for j in range(300 if i == 0 else (3 if i < 1600 else 2))]
        records.append({"id": f"patient-{i}", "name": f"مريض تجريبي {i:04d}",
                        "age": "40", "sex": "M", "created_at": "2026-01-01",
                        "updated_at": "2026-01-01", "prescriptions": prescriptions})
    PatientHistory()._save(records)
    RESULTS["dataset"] = {"drugs": len(drug_rows), "templates": 550,
        "favorites": 600, "patients": len(records),
        "prescriptions": sum(len(r["prescriptions"]) for r in records),
        "history_bytes": (cfg.APP_DIR / "patient_history.json").stat().st_size,
        "settings_bytes": cfg.CONFIG_PATH.stat().st_size,
        "synthetic_only": True, "python": sys.version, "windows": platform.platform()}
    save_results("seed")


def core():
    import config as cfg
    import drug_db as dbmod
    import main
    from patient_history import PatientHistory
    from security import protect, unprotect
    db = dbmod.DrugDatabase(str(cfg.DEFAULT_DB_PATH))
    measure("db.warm_cache_load", db.load)
    # A separate CSV copy measures a cache-miss without deleting the existing cache.
    fresh = OUTPUT / "cold-source.csv"
    fresh.write_bytes(cfg.DEFAULT_DB_PATH.read_bytes())
    measure("db.cold_cache_build", lambda: dbmod.DrugDatabase(str(fresh)), count=1)
    all_drugs = list(db.drugs)
    cases = ("P", "Pe", "Perf", "اخت", "ا", "اخ", "PerfBrand 001", "PerfIngredient 001",
             "perfbrnad", "PARACETAMOL", "باراسيتامول", "باراسيتامول تجريبي", "  PERFBRAND 001  ")
    comparisons = []
    for method, field in (("search_prescribable", "prescribable_norm"),
                          ("search_scientific", "generic_norm"), ("search", "search_norm")):
        for index, query in enumerate(cases):
            rows = measure(f"search.{method}.case{index}", lambda q=query, m=method: getattr(db, m)(q, 20))
            q = dbmod._norm(query)
            def value(drug):
                generic, brand = dbmod._norm(drug.generic_name), dbmod._norm(drug.brand_name)
                if field == "generic_norm": return generic
                if field == "prescribable_norm": return brand or generic
                return " ".join(filter(None, (generic, brand, dbmod._norm(drug.strength),
                                                dbmod._norm(drug.form), dbmod._norm(drug.category))))
            indexed = [(value(drug), i, drug) for i, drug in enumerate(all_drugs)]
            prefix = sorted((item for item in indexed if q <= item[0] < q + "\uffff"))
            substring = sorted(item for item in indexed if q in item[0]
                               and not (q <= item[0] < q + "\uffff"))
            expected = [item[2].to_dict() for item in (prefix + substring)[:20]]
            actual = [drug.to_dict() for drug in rows]
            assert actual == expected
            comparisons.append({"method": method, "query": query, "identical": True,
                                "count": len(actual), "ordered_data_sha256": hashlib.sha256(
                                    json.dumps(actual, ensure_ascii=False).encode()).hexdigest()})
    RESULTS["search_contract"] = comparisons
    with db._connect() as connection:
        RESULTS["sqlite_plans"] = {
            "prefix": [tuple(r) for r in connection.execute("EXPLAIN QUERY PLAN SELECT id FROM drugs "
                "WHERE prescribable_norm>=? AND prescribable_norm<? ORDER BY prescribable_norm,id LIMIT 20",
                ("perf", "perf\uffff"))],
            "substring": [tuple(r) for r in connection.execute("EXPLAIN QUERY PLAN SELECT id FROM drugs "
                "WHERE instr(prescribable_norm,?)>0 ORDER BY prescribable_norm,id LIMIT 20", ("x",))]}
    history = PatientHistory()
    measure("history.decrypt_parse", history._load)
    measure("history.search_all", lambda: history.search(""))
    measure("history.search_arabic", lambda: history.search("تجريبي 012"))
    measure("history.get_one", lambda: history.get("patient-1200"))
    measure("history.similar_scan", lambda: history.find_similar("مريض تجريبي 1200", "40", sex="M"))
    payload = (cfg.APP_DIR / "patient_history.json").read_text(encoding="utf-8")
    decrypted = measure("DPAPI.history_decrypt_only", lambda: unprotect(json.loads(payload)["data"]))
    measure("DPAPI.history_encrypt_only", lambda: protect(decrypted))
    measure("history.save_prescription", lambda: history.save_prescription(
        {"name": "مريض تجريبي 1200", "age": "40", "sex": "M"},
        [{"generic_name": "PerfIngredient 00001", "frequency": "مرة واحدة يومياً"}], "patient-1200"), count=3)
    measure("config.templates_normalize", cfg.config.treatment_templates)
    measure("config.favorites_normalize", cfg.config.medication_favorites)
    measure("config.encrypt_write", cfg.config.save)
    measure("classification.full_index", lambda: main.App._build_classification_index(tuple(db.drugs)), count=3)
    save_results("core")


def ui(startup_only=False):
    import tracemalloc
    from unittest.mock import patch
    import main
    import config as cfg
    import qr_utils as qu
    import i18n as I
    started = time.perf_counter()
    app = main.App()
    app.geometry("1200x800+20+20")
    failures = []
    app.report_callback_exception = lambda kind, error, tb: failures.append(str(error))
    try:
        def pump(seconds=.1):
            end = time.perf_counter() + seconds
            while time.perf_counter() < end:
                app.update()
                time.sleep(.001)

        def until(predicate, seconds=20):
            end = time.perf_counter() + seconds
            while not predicate():
                app.update()
                if time.perf_counter() > end: raise TimeoutError("Synthetic UI wait expired")
                time.sleep(.001)
            app.update_idletasks()

        app.update()
        RESULTS["startup.App_to_first_update_ms"] = round((time.perf_counter()-started)*1000, 3)
        if startup_only:
            save_results("startup")
            return
        until(lambda: not getattr(app, "_database_loading", False)
              and app._classification_future is None and bool(app._patient_history_records))
        for page in ("favorites", "treatment_templates", "drug_classes", "reference", "interaction_review", "medications"):
            measure("page.first." + page, lambda key=page: (app.show_page(key), app.update()), count=1)
            pump(.25)
        measure("favorites.refresh", app.refresh_favorites_page, count=3)
        measure("templates.refresh", app.refresh_saved_treatment_templates, count=3)
        measure("history.render_patient_list", lambda: app._render_patient_history(
            app.patient_search_var.get(), app._latest_query_tokens["patient_history"], app._patient_history_records), count=3)
        measure("history.render_300_prescriptions", lambda: app.show_patient_prescriptions(
            app.patient_history.get("patient-0")), count=1)
        app.show_page("medications")
        pump(.3)
        for query in ("P", "Pe", "Per", "ا", "اخ", "اخت"):
            row = app.rows[0]
            row._hide_ac()
            row.trade_var.set(query)
            start = time.perf_counter()
            row._on_trade_type()
            until(lambda: getattr(row, "_ac_top", None) is not None)
            bucket = ("ar" if ord(query[0]) > 255 else "en") + "." + str(len(query))
            RESULTS["autocomplete.last_key_to_popup." + bucket] = round((time.perf_counter()-start)*1000, 3)
            row._hide_ac()
        item = qu.DrugItem(generic_name="PerfIngredient 00001", brand_name="PerfBrand 00001",
                           dosage="SYNTHETIC 10 mg", frequency="مرة واحدة يومياً", duration="٧ أيام", notes="اختبار")
        measure("editor.add_drug", lambda: app.add_row(item), count=5)
        measure("editor.on_change_6_rows", app.on_any_change, count=5)
        measure("editor.preview_6_rows", app.render_word_preview, count=3)
        template = cfg.config.treatment_templates()[0]
        class Variable:
            def __init__(self, value): self.value = value
            def get(self): return self.value
        class Dialog:
            def grab_release(self): pass
            def destroy(self): pass
        selections = [([drug], Variable(0), Variable(True)) for drug in template["medications"]]
        measure("template.apply_10_drugs", lambda: app._apply_treatment_selection(
            selections, True, Dialog(), template["id"]), count=1)
        if getattr(app, "_applying_template", False):
            apply_started = time.perf_counter()
            until(lambda: not app._applying_template)
            RESULTS["template.async_completion_ms"] = round((time.perf_counter()-apply_started)*1000, 3)
        pump(.3)
        app.patient_vars["name"].set("مريض تجريبي 1200")
        app._loaded_patient_id = "patient-1200"
        with patch("main.messagebox.askyesno", return_value=False), patch("main.messagebox.showinfo"):
            measure("history.UI_save_prescription", app.save_prescription_for_patient, count=1)
            if getattr(app, "_history_save_future", None) is not None:
                save_started = time.perf_counter()
                until(lambda: app._history_save_future is None)
                RESULTS["history.async_save_completion_ms"] = round((time.perf_counter()-save_started)*1000, 3)
        pump(.4)
        for language in ("en", "ar"):
            for paper in ("A5", "A4"):
                rx = app.collect()
                for kind, function, kw in (
                    ("pdf", main.pdfgen.generate_prescription_pdf, {}),
                    ("word-header", main.pdfgen.generate_prescription_docx, {"show_header": True}),
                    ("word-noheader", main.pdfgen.generate_medication_label_docx, {})):
                    path = OUTPUT / f"synthetic-{language}-{paper}-{kind}.{'pdf' if kind == 'pdf' else 'docx'}"
                    def export(f=function, p=path, kwargs=kw, lang=language, size=paper):
                        with I.document_language(lang):
                            return f(rx, str(p), paper_size=size, qr_pil_image=None, **kwargs)
                    measure(f"export.{language}.{paper}.{kind}", export, count=1)
        for page in ("patient", "medications", "favorites", "treatment_templates", "drug_classes"):
            measure("page.warm." + page, lambda key=page: (app.show_page(key), app.update()), count=3)
        measure("resize.sequence_10", lambda: [
            (app.geometry(f"{1000+i*12}x800+20+20"), app.update()) for i in range(10)], count=2)
        tracemalloc.start()
        pump(.3)
        before = tracemalloc.get_traced_memory()[0]
        for _ in range(10):
            for page in ("favorites", "treatment_templates", "drug_classes", "medications"):
                app.show_page(page)
                app.update()
        import gc
        gc.collect()
        RESULTS["memory.40_switches_bytes"] = tracemalloc.get_traced_memory()[0] - before
        RESULTS["memory.tracemalloc_peak_bytes"] = tracemalloc.get_traced_memory()[1]
        tracemalloc.stop()
        RESULTS["callback_errors"] = failures
        if os.environ.get("RX_PERF_DEBUG") == "1":
            import perf_probe
            RESULTS["probe"] = perf_probe.summary()
        save_results("ui")
    finally:
        app.destroy()


def imports():
    result = subprocess.run([sys.executable, "-X", "importtime", "-c", "import main"],
                            env=os.environ.copy(), cwd=ROOT, capture_output=True, text=True, check=True)
    (OUTPUT / "importtime.txt").write_text(result.stderr, encoding="utf-8")
    rows = []
    for line in result.stderr.splitlines():
        if not line.startswith("import time:") or "self [us]" in line: continue
        parts = line[len("import time:"):].split("|")
        if len(parts) == 3:
            try: rows.append({"self_ms": int(parts[0])/1000, "cumulative_ms": int(parts[1])/1000,
                              "module": parts[2].strip()})
            except ValueError: pass
    RESULTS["slowest_self"] = sorted(rows, key=lambda r: r["self_ms"], reverse=True)[:20]
    RESULTS["slowest_cumulative"] = sorted(rows, key=lambda r: r["cumulative_ms"], reverse=True)[:20]
    save_results("imports")


if args.mode == "seed": seed()
elif args.mode == "core": core()
elif args.mode == "ui": ui()
elif args.mode == "startup": ui(startup_only=True)
else: imports()
