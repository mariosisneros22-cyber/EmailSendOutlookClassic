# ui/modal_run.py
import queue
import threading
import tkinter as tk
import tkinter.messagebox as mb
import customtkinter as ctk


def run_with_modal(
    root,
    title: str,
    confirm_text: str,
    total: int,
    runner,  # callable que hace el trabajo
    runner_kwargs: dict,  # kwargs para runner
    set_ui_busy=None,  # callable(bool) para deshabilitar UI principal
):
    """
    runner(**runner_kwargs) debe soportar:
      - progress (CTkProgressBar o ttk)
      - root (para update_idletasks)
      - is_cancelled (callable)
      - on_progress (callable(i,total,estado,*extras))

    Este modal inyecta:
      - progress: None (UI se actualiza por cola de eventos)
      - root: None
      - is_cancelled: stop_event.is_set
      - on_progress: callback encolado y procesado en el hilo UI
    """
    stop_event = threading.Event()
    event_queue: queue.Queue = queue.Queue()

    modal = ctk.CTkToplevel(root)
    modal.title(title)
    modal.geometry("520x240")
    modal.resizable(False, False)
    modal.transient(root)
    modal.grab_set()

    lbl = ctk.CTkLabel(modal, text=confirm_text)
    lbl.pack(anchor="w", padx=20, pady=(20, 8))

    status_lbl = ctk.CTkLabel(modal, text="Listo para iniciar.")
    status_lbl.pack(anchor="w", padx=20, pady=(0, 8))

    modal_progress = ctk.CTkProgressBar(modal)
    modal_progress.set(0)
    modal_progress.pack(fill="x", padx=20, pady=(0, 8))

    count_lbl = ctk.CTkLabel(modal, text=f"0 / {total}")
    count_lbl.pack(anchor="w", padx=20, pady=(0, 12))

    btns = ctk.CTkFrame(modal, fg_color="transparent")
    btns.pack(fill="x", padx=20, pady=(0, 20))

    enviando = {"flag": False}

    def on_close():
        if enviando["flag"]:
            on_stop()          # solicita cancelación
            return             # no destruyas aún
        modal.destroy()

    modal.protocol("WM_DELETE_WINDOW", on_close)

    def on_stop():
        if stop_event.is_set():
            status_lbl.configure(text="Cancelando...")
            return

    def _on_progress(i, total, estado, *extras):
        count_lbl.configure(text=f"{i} / {total}")
        
        try:
            frac = float(i) / float(total or 1)
        except (TypeError, ValueError, ZeroDivisionError):
            frac = 0.0
            
        modal_progress.set(max(0.0, min(1.0, frac)))
        est = (str(estado) or "").strip().lower()
        if est.startswith("error"):
            status_lbl.configure(text=f"Error en {i}/{total}")
        else:
            status_lbl.configure(text=f"Procesando {i}/{total}...")

    def _finish_ok(result):
        enviando["flag"] = False
        btn_stop.configure(state="disabled")
        btn_cerrar.configure(state="normal")
        btn_iniciar.configure(state="normal")

        if callable(set_ui_busy):
            set_ui_busy(False)

        if isinstance(result, dict):
            msg = str(result.get("mensaje", "")).strip()
            if msg:
                mb.showinfo("Proceso finalizado", msg)
        status_lbl.configure(text="Proceso finalizado.")

    def _finish_error(exc):
        enviando["flag"] = False
        btn_stop.configure(state="disabled")
        btn_cerrar.configure(state="normal")
        btn_iniciar.configure(state="normal")

        if callable(set_ui_busy):
            set_ui_busy(False)

        mb.showerror("Error", str(exc))
        status_lbl.configure(text="Error.")

    def _poll_events(user_on_progress):
        if not modal.winfo_exists():
            return
        try:
            while True:
                kind, payload = event_queue.get_nowait()
                if kind == "progress":
                    i, total_i, estado, extras = payload
                    _on_progress(i, total_i, estado, *extras)
                    if callable(user_on_progress):
                        try:
                            user_on_progress(i, total_i, estado, *extras)
                        except Exception as e:
                            # opcional: no romper el envío, pero deja rastro
                            print(f"user_on_progress falló: {e}")
                elif kind == "done_ok":
                    _finish_ok(payload)
                    return
                elif kind == "done_err":
                    _finish_error(payload)
                    return
        except queue.Empty:
            pass

        if enviando["flag"] and modal.winfo_exists():
            modal.after(80, lambda: _poll_events(user_on_progress))

    def iniciar():
        btn_iniciar.configure(state="disabled")
        btn_cerrar.configure(state="disabled")
        btn_stop.configure(state="normal")

        stop_event.clear()
        enviando["flag"] = True

        if callable(set_ui_busy):
            set_ui_busy(True)

        status_lbl.configure(text="Procesando...")
        modal_progress.set(0)
        count_lbl.configure(text=f"0 / {total}")

        # inyectar args comunes
        kwargs = dict(runner_kwargs or {})
        kwargs["progress"] = None
        kwargs["root"] = None
        kwargs["is_cancelled"] = stop_event.is_set

        # Encadenar callback si el runner ya trae on_progress
        user_on_progress = kwargs.get("on_progress", None)

        def queued_progress(i, total_i, estado, *extras):
            event_queue.put(("progress", (i, total_i, estado, extras)))

        kwargs["on_progress"] = queued_progress

        def worker():
            try:
                result = runner(**kwargs)
                event_queue.put(("done_ok", result))
            except Exception as exc:
                event_queue.put(("done_err", exc))

        t = threading.Thread(target=worker, daemon=True)
        t.start()
        _poll_events(user_on_progress)

    btn_cerrar = ctk.CTkButton(btns, text="Cerrar", command=on_close)
    btn_cerrar.pack(side="right")

    btn_stop = ctk.CTkButton(btns, text="STOP", command=on_stop, state="disabled")
    btn_stop.pack(side="right", padx=(0, 10))

    btn_iniciar = ctk.CTkButton(btns, text="Iniciar", command=iniciar)
    btn_iniciar.pack(side="right", padx=(0, 10))

    return modal
