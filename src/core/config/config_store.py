#core/config_store.py

import json
from core.config.app_dirs import CONFIG_PATH

def load_config() -> dict:
    if not CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8")) or {}
    except json.JSONDecodeError:
        return {}
    
def save_config(data:dict) -> None:
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")