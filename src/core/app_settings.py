#core/app_settings.py

from pathlib import Path
from core.config_store import load_config, save_config
        
def set_attachments_folder(path: str) -> None:
    if not path:
        raise ValueError("La ruta de adjuntos no puede estar vacía.")
    data = load_config()
    data["responder_attachments_folder"] = str(Path(path).resolve())
    save_config(data)
    
def get_attachments_folder() -> str | None:
    data = load_config()
    return data.get("responder_attachments_folder")
