# ui/responder_tab.py
import os
import queue
import threading
from datetime import datetime
from tkinter import filedialog, messagebox

import customtkinter as ctk

from core.app_settings import get_attachments_folder, set_attachments_folder
from core.outlook_folders import pick_outlook_folder_and_save, get_saved_outlook_folder
from core.responder_queue import (
    update_control_from_outlook,
    process_pending_responses,
    fill_suggested_names,
    apply_suggested_to_nombre_archivo,
    CONTROL_PATH,
)

# -----------------------
# UI tokens (locales)
# -----------------------
PAD_OUTER = 10
PAD_INNER = 14
GAP_SM = 6
GAP_MD = 10

FONT_TITLE = ("Segoe UI", 18, "bold")
FONT_SECTION = ("Segoe UI", 14, "bold")
FONT_BODY = ("Segoe UI", 13)

FONT_BTN_PRIMARY = ("Segoe UI", 13, "bold")
FONT_BTN = ("Segoe UI", 12)
FONT_BTN_TOOL = ("Segoe UI", 11)

H_PRIMARY = 46
H_SECONDARY = 36
H_TOOL = 30

BTN_PRIMARY = dict(
    fg_color=("#2563eb", "#2563eb"),
    hover_color=("#1d4ed8", "#1d4ed8"),
    text_color="white",
)

BTN_SECONDARY = dict(
    fg_color=("#e5e7eb", "#2a2a2a"),
    hover_color=("#d1d5db", "#3a3a3a"),
    text_color=("#111111", "#eaeaea"),
)

BTN_TOOL = dict(
    fg_color="transparent",
    hover_color=("#f3f4f6", "#2b2b2b"),
    border_width=1,
    border_color=("#cfcfcf", "#3a3a3a"),
    text_color=("#111111", "#eaeaea"),
)

LABEL_COLOR = dict(text_color=("#111111", "#eaeaea"))


# -----------------------
# Helpers
# -----------------------
def _open_file(path: str) -> None:
    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    os.startfile(path)  # noqa: S606


def _safe_folder_name(folder) -> str:
    try:
        return str(folder.Name)
    except Exception:
        # COM a veces falla al leer Name
        return "(carpeta seleccionada)"


def _now_ts() -> str:
    return datetime.now().strftime("%H:%M:%S")


def _section(parent, title: str) -> ctk.CTkFrame:
    card = ctk.CTkFrame(
        parent,
        corner_radius=12,
        border_width=1,
        border_color=("#d0d0d0", "#3a3a3a"),
        fg_color=("#ffffff", "#1f1f1f"),
    )
    card.pack(fill="x", pady=(0, GAP_MD))

    header = ctk.CTkFrame(card, fg_color="transparent")
    header.pack(fill="x", padx=PAD_INNER, pady=(PAD_INNER, 6))

    ctk.CTkLabel(header, text=title, font=FONT_SECTION, **LABEL_COLOR).pack(anchor="w")

    body = ctk.CTkFrame(card, fg_color="transparent")
    body.pack(fill="x", padx=PAD_INNER, pady=(0, PAD_INNER))
    return body


# -----------------------
# Main mount
# -----------------------
def mount(parent):
    root = parent.winfo_toplevel()

    frame = ctk.CTkFrame(parent, fg_color="transparent")
    frame.pack(fill="both", expand=True, padx=PAD_OUTER, pady=PAD_OUTER)

    header = ctk.CTkFrame(frame, fg_color="transparent")
    header.pack(fill="x", pady=(0, GAP_MD))
    ctk.CTkLabel(header, text="Responder (cola masiva)", font=FONT_TITLE, **LABEL_COLOR).pack(anchor="w")

    status_var = ctk.StringVar(value="Carpeta Outlook: (no seleccionada)")
    attachments_var = ctk.StringVar(value="Carpeta adjuntos: (no seleccionada)")
    ui_status_var = ctk.StringVar(value="Estado: Listo")

    # -----------------------
    # Log (colapsable)
    # -----------------------
    log_visible = ctk.BooleanVar(value=False)

    log_card = ctk.CTkFrame(
        frame,
        corner_radius=12,
        border_width=1,
        border_color=("#d0d0d0", "#3a3a3a"),
        fg_color=("#ffffff", "#1f1f1f"),
    )
    log_hdr = ctk.CTkFrame(log_card, fg_color="transparent")
    log_body = ctk.CTkFrame(log_card, fg_color="transparent")
    log_box = ctk.CTkTextbox(log_body, height=180)

    def _log_clear():
        log_box.delete("1.0", "end")

    def _log_show():
        if log_visible.get():
            return
        log_visible.set(True)
        log_card.pack(fill="both", expand=True, pady=(0, 0))
        log_hdr.pack(fill="x", padx=PAD_INNER, pady=(PAD_INNER, 6))
        log_body.pack(fill="both", expand=True, padx=PAD_INNER, pady=(0, PAD_INNER))
        log_box.pack(fill="both", expand=True)

    def _log_hide():
        if not log_visible.get():
            return
        log_visible.set(False)
        log_card.pack_forget()

    def _log_toggle():
        _log_hide() if log_visible.get() else _log_show()

    def log(msg: str):
        s = str(msg).strip()
        is_error = s.lower().startswith("error") or s.lower().startswith("exception")
        if is_error:
            _log_show()
        try:
            log_box.insert("end", f"[{_now_ts()}] {s}\n")
            log_box.see("end")
            log_box.update_idletasks()
        except Exception:
            # UI-only: no romper procesos por error de render
            pass

    ctk.CTkLabel(log_hdr, text="Log (detalle)", font=FONT_SECTION, **LABEL_COLOR).pack(side="left")

    btn_hide = ctk.CTkButton(
        log_hdr, text="Ocultar", width=90, height=H_TOOL, font=FONT_BTN_TOOL, command=_log_hide, **BTN_TOOL
    )
    btn_hide.pack(side="right", padx=(GAP_SM, 0))

    btn_clear = ctk.CTkButton(
        log_hdr, text="Limpiar", width=90, height=H_TOOL, font=FONT_BTN_TOOL, command=_log_clear, **BTN_TOOL
    )
    btn_clear.pack(side="right", padx=(GAP_SM, 0))

    # -----------------------
    # Busy state (disable UI)
    # -----------------------
    busy = {"flag": False}

    def _set_busy(is_busy: bool):
        busy["flag"] = bool(is_busy)
        state = "disabled" if is_busy else "normal"
        tool_state = "disabled" if is_busy else "normal"

        # Botones principales
        btn_pick_outlook.configure(state=tool_state)
        btn_pick_attach.configure(state=tool_state)

        btn_update.configure(state=state)
        btn_open.configure(state=state)
        btn_process.configure(state=state)

        # Tools
        btn_suggest.configure(state=tool_state)
        btn_apply.configure(state=tool_state)
        btn_show_log.configure(state=tool_state)

    # -----------------------
    # Thread runner helper (no UI freeze)
    # -----------------------
    event_q: queue.Queue = queue.Queue()

    def _poll_events():
        try:
            while True:
                kind, payload = event_q.get_nowait()
                if kind == "log":
                    log(payload)
                elif kind == "status":
                    ui_status_var.set(payload)
                elif kind == "done_ok":
                    _set_busy(False)
                    ui_status_var.set(payload.get("ui_status", "Estado: Listo"))
                    msg = payload.get("messagebox_info")
                    if msg:
                        messagebox.showinfo(payload.get("title", "Listo"), msg)
                elif kind == "done_err":
                    _set_busy(False)
                    ui_status_var.set("Estado: Error")
                    log(f"ERROR: {payload}")
                    messagebox.showerror("Error", str(payload))
        except queue.Empty:
            pass

        if frame.winfo_exists():
            frame.after(100, _poll_events)

    _poll_events()

    def _run_in_thread(fn, *, ui_start_status: str, ok_title: str | None = None, ok_msg: str | None = None):
        if busy["flag"]:
            return

        _set_busy(True)
        ui_status_var.set(ui_start_status)

        def worker():
            try:
                result = fn()
                payload = {"ui_status": "Estado: Listo"}
                if ok_title or ok_msg:
                    payload["title"] = ok_title or "Listo"
                    payload["messagebox_info"] = ok_msg
                event_q.put(("done_ok", payload))
                return result
            except Exception as e:
                event_q.put(("done_err", e))
                return None

        threading.Thread(target=worker, daemon=True).start()

    # -----------------------
    # State refresh
    # -----------------------
    def refresh_status():
        # Outlook folder
        try:
            folder = get_saved_outlook_folder()
        except Exception:
            folder = None

        if folder is None:
            status_var.set("Carpeta Outlook: (no seleccionada)")
        else:
            status_var.set(f"Carpeta Outlook: {_safe_folder_name(folder)}")

        # Attachments folder
        try:
            p = get_attachments_folder()
        except Exception:
            p = None

        attachments_var.set(f"Carpeta adjuntos: {p}" if p else "Carpeta adjuntos: (no seleccionada)")

    # -----------------------
    # Actions (callbacks)
    # -----------------------
    def on_pick_folder():
        if busy["flag"]:
            return
        try:
            ui_status_var.set("Estado: Seleccionando carpeta Outlook…")
            folder = pick_outlook_folder_and_save()
            if folder is None:
                ui_status_var.set("Estado: Listo")
                return
            refresh_status()
            ui_status_var.set("Estado: Carpeta Outlook guardada")
            messagebox.showinfo("Listo", "Carpeta guardada. Ya puedes actualizar la cola.")
        except Exception as e:
            ui_status_var.set("Estado: Error")
            log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))

    def on_pick_attachments_folder():
        if busy["flag"]:
            return
        try:
            ui_status_var.set("Estado: Seleccionando carpeta de adjuntos…")
            path = filedialog.askdirectory(title="Selecciona la carpeta de adjuntos (PDF/archivos)")
            if not path:
                ui_status_var.set("Estado: Listo")
                return
            set_attachments_folder(path)
            refresh_status()
            ui_status_var.set("Estado: Carpeta de adjuntos guardada")
            messagebox.showinfo("Listo", "Carpeta de adjuntos guardada.")
        except Exception as e:
            ui_status_var.set("Estado: Error")
            log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))

    def on_update_queue():
        refresh_status()
        folder = get_saved_outlook_folder()
        if folder is None:
            messagebox.showerror("Error", "Primero elige la carpeta Outlook.")
            return

        def job():
            event_q.put(("status", "Estado: Actualizando cola…"))
            event_q.put(("log", "Actualizando cola desde Outlook…"))
            n = update_control_from_outlook()
            event_q.put(("log", f"Listo. Se agregaron {n} conversaciones nuevas."))
            event_q.put(("done_ok", {
                "ui_status": f"Estado: Cola actualizada (+{n})",
                "title": "Actualización completada",
                "messagebox_info": f"Se agregaron {n} conversaciones nuevas.\n\n{CONTROL_PATH}",
            }))

        _set_busy(True)
        threading.Thread(target=job, daemon=True).start()

    def on_open_control():
        if busy["flag"]:
            return
        try:
            ui_status_var.set("Estado: Abriendo control.xlsx…")
            _open_file(CONTROL_PATH)
            ui_status_var.set("Estado: Listo")
        except Exception as e:
            ui_status_var.set("Estado: Error")
            log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))

    def on_fill_suggested():
        def job():
            event_q.put(("status", "Estado: Generando sugeridos…"))
            event_q.put(("log", "Generando nombre_sugerido desde subject…"))
            n = fill_suggested_names(year_mode="current", only_if_empty=True)
            event_q.put(("log", f"Listo. Se generaron/actualizaron {n} sugeridos."))
            event_q.put(("done_ok", {
                "ui_status": f"Estado: Sugeridos listos ({n})",
                "title": "Listo",
                "messagebox_info": f"Sugeridos generados: {n}\n\n{CONTROL_PATH}",
            }))

        _set_busy(True)
        threading.Thread(target=job, daemon=True).start()

    def on_apply_suggested():
        def job():
            event_q.put(("status", "Estado: Aplicando sugeridos…"))
            event_q.put(("log", "Copiando nombre_sugerido → nombre_archivo…"))
            n = apply_suggested_to_nombre_archivo(only_if_empty=True, add_pdf_ext=False)
            event_q.put(("log", f"Listo. Se copiaron {n} valores a nombre_archivo."))
            event_q.put(("done_ok", {
                "ui_status": f"Estado: NombreArchivo actualizado ({n})",
                "title": "Listo",
                "messagebox_info": f"Copiados a nombre_archivo: {n}\n\n{CONTROL_PATH}",
            }))

        _set_busy(True)
        threading.Thread(target=job, daemon=True).start()

    def on_process_pending():
        refresh_status()

        folder = get_saved_outlook_folder()
        if folder is None:
            messagebox.showerror("Error", "Primero elige la carpeta Outlook.")
            return

        attach_dir = get_attachments_folder()
        if not attach_dir:
            messagebox.showerror("Error", "Primero elige la carpeta de adjuntos.")
            return

        def job():
            event_q.put(("status", "Estado: Procesando pendientes…"))
            event_q.put(("log", "Procesando pendientes…"))

            n = process_pending_responses(
                carpeta_archivos=attach_dir,
                html_body=None,
                delay_segundos=1.0,
                only_first_n=None,
                logger=lambda s: event_q.put(("log", s)),
            )

            event_q.put(("log", f"Listo. Procesados OK: {n}"))
            event_q.put(("done_ok", {
                "ui_status": f"Estado: Terminado (OK: {n})",
                "title": "Terminado",
                "messagebox_info": f"Procesados OK: {n}\n\nRevisa el Excel:\n{CONTROL_PATH}",
            }))

        _set_busy(True)
        threading.Thread(target=job, daemon=True).start()

    # -----------------------
    # UI sections
    # -----------------------

    cfg = _section(frame, "Configuración")
    cfg.grid_columnconfigure(0, weight=1)
    cfg.grid_columnconfigure(1, weight=0)

    ctk.CTkLabel(cfg, textvariable=status_var, font=FONT_BODY, **LABEL_COLOR).grid(
        row=0, column=0, sticky="w", pady=(0, GAP_SM)
    )
    ctk.CTkLabel(cfg, textvariable=attachments_var, font=FONT_BODY, **LABEL_COLOR).grid(
        row=1, column=0, sticky="w"
    )

    btn_pick_outlook = ctk.CTkButton(
        cfg, text="Outlook…", command=on_pick_folder,
        font=FONT_BTN_TOOL, height=H_TOOL, width=120, **BTN_TOOL
    )
    btn_pick_outlook.grid(row=0, column=1, sticky="e", padx=(GAP_MD, 0), pady=(0, GAP_SM))

    btn_pick_attach = ctk.CTkButton(
        cfg, text="Adjuntos…", command=on_pick_attachments_folder,
        font=FONT_BTN_TOOL, height=H_TOOL, width=120, **BTN_TOOL
    )
    btn_pick_attach.grid(row=1, column=1, sticky="e", padx=(GAP_MD, 0))

    act = _section(frame, "Acciones")
    act.grid_columnconfigure(0, weight=1)
    act.grid_columnconfigure(1, weight=1)
    act.grid_columnconfigure(2, weight=2)

    btn_update = ctk.CTkButton(
        act, text="Actualizar cola", command=on_update_queue,
        font=FONT_BTN, height=H_SECONDARY, **BTN_SECONDARY
    )
    btn_update.grid(row=0, column=0, sticky="ew", padx=(0, GAP_MD), pady=(0, GAP_SM))

    btn_open = ctk.CTkButton(
        act, text="Abrir control.xlsx", command=on_open_control,
        font=FONT_BTN, height=H_SECONDARY, **BTN_SECONDARY
    )
    btn_open.grid(row=0, column=1, sticky="ew", padx=(0, GAP_MD), pady=(0, GAP_SM))

    btn_process = ctk.CTkButton(
        act, text="Procesar pendientes", command=on_process_pending,
        font=FONT_BTN_PRIMARY, height=H_PRIMARY, **BTN_PRIMARY
    )
    btn_process.grid(row=0, column=2, sticky="ew", pady=(0, GAP_SM))

    ctk.CTkLabel(act, textvariable=ui_status_var, font=FONT_BODY, **LABEL_COLOR).grid(
        row=1, column=0, columnspan=3, sticky="w", pady=(6, 0)
    )

    tools_row = ctk.CTkFrame(act, fg_color="transparent")
    tools_row.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 0))

    tools_row.grid_columnconfigure(0, weight=1)
    tools_row.grid_columnconfigure(1, weight=1)
    tools_row.grid_columnconfigure(2, weight=1)

    btn_suggest = ctk.CTkButton(
        tools_row, text="Generar sugeridos", command=on_fill_suggested,
        font=FONT_BTN_TOOL, height=H_TOOL, **BTN_TOOL
    )
    btn_suggest.grid(row=0, column=0, sticky="ew", padx=(0, GAP_MD))

    btn_apply = ctk.CTkButton(
        tools_row, text="Sugerido → NombreArchivo", command=on_apply_suggested,
        font=FONT_BTN_TOOL, height=H_TOOL, **BTN_TOOL
    )
    btn_apply.grid(row=0, column=1, sticky="ew", padx=(0, GAP_MD))

    btn_show_log = ctk.CTkButton(
        tools_row, text="Mostrar log", command=_log_toggle,
        font=FONT_BTN_TOOL, height=H_TOOL, width=120, **BTN_TOOL
    )
    btn_show_log.grid(row=0, column=2, sticky="e")

    hint = (
        "Flujo:\n"
        "1) Configura carpeta Outlook y carpeta de adjuntos.\n"
        "2) Actualiza cola para llenar/actualizar control.xlsx.\n"
        "3) Procesa pendientes para responder (ReplyAll) y adjuntar archivos.\n"
        "4) (Opcional) Mostrar log para ver detalle/errores."
    )
    ctk.CTkLabel(frame, text=hint, justify="left", font=FONT_BODY, **LABEL_COLOR).pack(
        anchor="w", pady=(0, GAP_MD)
    )

    refresh_status()
    ui_status_var.set("Estado: Listo")

    return frame