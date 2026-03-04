# app.py
import os
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1] 
PYCACHE_DIR =PROJECT_ROOT / ".pycache_global"
PYCACHE_DIR.mkdir(parents=True, exist_ok=True)
os.environ["PYTHONPYCACHEPREFIX"] = str(PYCACHE_DIR)


import customtkinter as ctk

from ui.bulk_tab import mount as mount_bulk
from ui.responder_tab import mount as mount_responder

def main():
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Automatización Outlook")
    root.geometry("1200x850")

    tabs = ctk.CTkTabview(root)
    tabs.pack(fill="both", expand=True, padx=10, pady=10)

    tab_bulk = tabs.add("Envío masivo")
    tab_resp = tabs.add("Responder")

    mount_bulk(tab_bulk)
    mount_responder(tab_resp)

    root.mainloop()

if __name__ == "__main__":
    main()