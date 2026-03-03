#core/app_settings.py
import json
import os

CONFIG_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__),"..","data", "config.json"))

def _load_config() ->dict:
    if not os.path.exists(CONFIG_PATH):
        return {}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}
    
def _save_config(data:dict) -> None:
    os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
        
def set_attachments_folder(path:str) -> None:
    data =_load_config()
    data["responder_attachments_folder"] = os.path.abspath(path)
    _save_config(data)
    
def get_attachments_folder() ->str | None:
    data = _load_config()
    p = data.get("responder_attachments_folder")
    return os.path.abspath(p) if p else None
      