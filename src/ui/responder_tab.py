#ui/responder_tab.py

import os, subprocess
import customtkinter as ctk
from tkinter import messagebox

from core.outlook_folders import pick_outlook_folder_and_save, get_saved_outlook_folder
from core.responder_queue import update_control_from_outlook, CONTROL_PATH

PADY_SM = 6
PADY_MD = 10
 
FONT_LABEL = ("segoe UI", 13, "bold")
FONT_BUTTON = ("segoe UI", 13, "bold")

def _open_file(path: str):
    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    
    
    #Windows: abre con app por defecto
    os.startfile(path) # noqa: S606
    

def _safe_folder_name(folder):
    try:
        return str(folder.Name)
    except Exception:
        return  "(carpeta seleccionada)"
    
def mount(parent):
    """
    Tab: RESPONDER (cola / control)
    - Elegir carpeta Outlook
    - Mostrar estado carpeta guardada
    - Actualizar cola -> control.xlsx
    - Abrir control.xlsx
    """
    root = parent.winfo_toplevel()

    frame = ctk.CTkFrame(parent)
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    # ---------- Header ----------
    ctk.CTkLabel(frame, text="Responder (cola masiva)", font=("Segoe UI", 16, "bold")).pack(
        anchor="w", pady=(0, PADY_MD)
    )

    status_var = ctk.StringVar(value="Carpeta Outlook: (no seleccionada)")
    lbl_status = ctk.CTkLabel(frame, textvariable=status_var)
    lbl_status.pack(anchor="w", pady=(0, PADY_MD))

    # ---------- Acciones ----------
    actions = ctk.CTkFrame(frame, fg_color="transparent")
    actions.pack(fill="x", pady=(0, PADY_MD))

    actions.grid_columnconfigure(0, weight=1)
    actions.grid_columnconfigure(1, weight=1)
    actions.grid_columnconfigure(2, weight=1)

    def refresh_status():
        try:
            folder = get_saved_outlook_folder()
            if folder is None:
                status_var.set("Carpeta Outlook: (no seleccionada)")
            else:
                status_var.set(f"Carpeta Outlook: {_safe_folder_name(folder)}")
        except Exception:
            status_var.set("Carpeta Outlook: (error al leer configuración)")

    def on_pick_folder():
        try:
            folder = pick_outlook_folder_and_save()
            if folder is None:
                return
            refresh_status()
            messagebox.showinfo("Listo", "Carpeta guardada. Ya puedes actualizar la cola.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def on_update_queue():
        try:
            refresh_status()
            folder = get_saved_outlook_folder()
            if folder is None:
                messagebox.showerror("Error", "Primero elige la carpeta Outlook.")
                return

            n = update_control_from_outlook()
            messagebox.showinfo("Actualización completada", f"Se agregaron {n} conversaciones nuevas.\n\n{CONTROL_PATH}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def on_open_control():
        try:
            _open_file(CONTROL_PATH)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    btn_folder = ctk.CTkButton(
        actions, text="Elegir carpeta Outlook", command=on_pick_folder, font=FONT_BUTTON, height=40
    )
    btn_folder.grid(row=0, column=0, sticky="ew", padx=(0, 8), pady=6)

    btn_update = ctk.CTkButton(
        actions, text="Actualizar cola", command=on_update_queue, font=FONT_BUTTON, height=40
    )
    btn_update.grid(row=0, column=1, sticky="ew", padx=8, pady=6)

    btn_open = ctk.CTkButton(
        actions, text="Abrir control.xlsx", command=on_open_control, font=FONT_BUTTON, height=40
    )
    btn_open.grid(row=0, column=2, sticky="ew", padx=(8, 0), pady=6)

    # ---------- Info ----------
    info = (
        "Flujo:\n"
        "1) Elige la carpeta Outlook donde quedan los correos a responder.\n"
        "2) Actualiza cola: se llena/actualiza el archivo control.xlsx.\n"
        "3) Luego (siguiente fase) agregaremos el botón 'Responder' para ReplyAll + adjunto."
    )
    ctk.CTkLabel(frame, text=info, justify="left").pack(anchor="w", pady=(PADY_MD, 0))

    refresh_status()

    return frame