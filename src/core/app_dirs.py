#core/app_dirs.py
from __future__ import annotations
from pathlib import Path
import os

APP_NAME = "AppCorreo"

def data_dir() ->Path:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    p = Path(base) / APP_NAME / "data"
    p.mkdir(parents=True, exist_ok=True)
    return p

def config_path() -> Path:
    return data_dir() / "config.json"

def responder_control_path()->Path:
    return data_dir() / "responder_control.xlsx"



CONFIG_PATH=config_path()
RESPONDER_CONTROL_PATH=responder_control_path()