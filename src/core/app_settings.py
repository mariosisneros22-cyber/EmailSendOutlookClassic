#core/app_settings.py

import os
from core.config_store import load_config, save_config
        
def set_attachments_folder(path:str) -> None:
    data = load_config()
    data["responder_attachments_folder"] = os.path.abspath(path)
    save_config(data)
    
def get_attachments_folder() ->str | None:
    data = load_config()
    p = data.get("responder_attachments_folder")
    return os.path.abspath(p) if p else None
      