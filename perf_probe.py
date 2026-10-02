"""Removable, opt-in diagnostic timings. Never record arguments or clinical data.

Set RX_PERF_DEBUG=1 before importing main. This does not change storage or search.
"""
from __future__ import annotations

import functools
import logging
import os
import statistics
import threading
import time
from collections import defaultdict, deque

ENABLED = os.environ.get("RX_PERF_DEBUG") == "1"
STARTED = time.perf_counter()
_samples = defaultdict(lambda: deque(maxlen=4096))
_lock = threading.Lock()
_logger = logging.getLogger("rx.performance.debug")
if ENABLED:
    _logger.setLevel(logging.DEBUG)
    _logger.addHandler(logging.StreamHandler())
    _logger.propagate = False


def record(operation, elapsed_ms):
    if not ENABLED:
        return
    with _lock:
        _samples[operation].append(round(elapsed_ms, 3))
    _logger.debug("RX_PERF %s %.3f ms", operation, elapsed_ms)


def summary():
    with _lock:
        result = {}
        for name, samples in _samples.items():
            ordered = sorted(samples)
            result[name] = {"count": len(ordered), "median_ms": statistics.median(ordered),
                            "p95_ms": ordered[min(len(ordered)-1, int(len(ordered)*.95))],
                            "max_ms": max(ordered)}
        return result


def _wrap(owner, name, label):
    original = getattr(owner, name, None)
    if original is None or getattr(original, "_rx_perf_wrapped", False):
        return

    @functools.wraps(original)
    def measured(*args, **kwargs):
        started = time.perf_counter()
        try:
            return original(*args, **kwargs)
        finally:
            record(label, (time.perf_counter()-started)*1000)

    measured._rx_perf_wrapped = True
    setattr(owner, name, measured)


def install(namespace):
    """Instrument fixed operations only, leaving default runs entirely untouched."""
    if not ENABLED:
        return
    app_class = namespace["App"]
    row_class = namespace["DrugRow"]
    db = namespace["dbmod"].DrugDatabase
    configuration = namespace["cfg"]
    import patient_history

    for owner, names in (
        (app_class, ("__init__", "_build_ui", "_ensure_page_built", "show_page",
                     "_load_page_data", "add_row", "_apply_treatment_selection",
                     "refresh_favorites_page", "refresh_patient_history",
                     "_render_patient_history", "show_patient_prescriptions",
                     "save_prescription_for_patient", "on_any_change", "render_word_preview")),
        (db, ("load", "_cache_is_current", "_rebuild_cache", "search", "search_prescribable",
              "search_scientific", "find_exact", "contains_name")),
        (configuration.Config, ("load", "save", "_write_config", "medication_favorites",
                               "treatment_templates")),
        (patient_history.PatientHistory, ("_load", "_save", "search", "get",
                                         "save_patient", "save_prescription", "find_similar")),
        (namespace["pdfgen"], ("generate_prescription_pdf", "generate_prescription_docx",
                              "generate_medication_label_docx")),
        (namespace["GlassFrame"], ("_paint_glass",)),
    ):
        owner_name = getattr(owner, "__name__", type(owner).__name__)
        for name in names:
            _wrap(owner, name, owner_name + "." + name)
    for owner in (configuration, patient_history):
        for name in ("protect", "unprotect"):
            _wrap(owner, name, "DPAPI." + name)

    original_schedule = row_class._schedule_autocomplete

    @functools.wraps(original_schedule)
    def schedule(row, query, mode):
        row._perf_key_time = time.perf_counter()
        length = "1" if len(query) == 1 else "2" if len(query) == 2 else "3plus"
        language = "ar" if any("\u0600" <= c <= "\u06ff" for c in query) else "en"
        row._perf_search_bucket = "autocomplete." + language + "." + length
        return original_schedule(row, query, mode)

    row_class._schedule_autocomplete = schedule
    original_show = row_class._show_ac

    @functools.wraps(original_show)
    def show(row, matches, mode):
        result = original_show(row, matches, mode)
        if hasattr(row, "_perf_key_time"):
            record(row._perf_search_bucket, (time.perf_counter()-row._perf_key_time)*1000)
        return result

    row_class._show_ac = show
    original_init = app_class.__init__

    @functools.wraps(original_init)
    def initialize(app, *args, **kwargs):
        original_init(app, *args, **kwargs)

        def usable():
            record("startup.module-to-first-idle", (time.perf_counter()-STARTED)*1000)
            expected = time.perf_counter() + .016

            def heartbeat():
                nonlocal expected
                if app._closing:
                    return
                now = time.perf_counter()
                record("event-loop.delay", max(0., (now-expected)*1000))
                expected = now + .016
                app.after(16, heartbeat)

            app.after(16, heartbeat)

        app.after_idle(usable)

    app_class.__init__ = initialize
