# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import datetime as dt
from PySide6.QtCore import QSettings

from .processing import default_department_dictionary, normalize_dictionary

def append_runtime_log(qs: QSettings, text: str):
    """Append a line to runtime logs stored in QSettings."""
    ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    prev = qs.value("runtime/logs", "")
    newlog = f"[{ts}] {text}"
    qs.setValue("runtime/logs", (prev + "\n" + newlog).strip())

def get_bool_setting(qs: QSettings, key: str, default: bool) -> bool:
    value = qs.value(key, default)
    if isinstance(value, str):
        return value.lower() == "true"
    return bool(value)

def set_bool_setting(qs: QSettings, key: str, value: bool):
    qs.setValue(key, bool(value))

def read_dictionary_setting(qs: QSettings):
    raw = qs.value("dictionary/json", "")
    if not raw:
        return default_department_dictionary()
    try:
        data = json.loads(raw)
    except Exception:
        return default_department_dictionary()
    return normalize_dictionary(data)

def save_dictionary_setting(qs: QSettings, data):
    normalized = normalize_dictionary(data)
    qs.setValue("dictionary/json", json.dumps(normalized, ensure_ascii=False, indent=2))
