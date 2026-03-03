# ui/modal_run.py
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
      - progress: CTkProgressBar del modal
      - root: modal
      - is_cancelled: lambda cancel_var.get()
      - on_progress: callback que actualiza labels
    """
    cancel_var = tk.BooleanVar(master=root, value=False)

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
            return
        modal.destroy()

    modal.protocol("WM_DELETE_WINDOW", on_close)

    def on_stop():
        cancel_var.set(True)
        status_lbl.configure(text="Cancelando…")

    def _on_progress(i, total, estado, *extras):
        count_lbl.configure(text=f"{i} / {total}")
        est = (str(estado) or "").strip().lower()
        if est.startswith("error"):
            status_lbl.configure(text=f"Error en {i}/{total}")
        else:
            status_lbl.configure(text=f"Procesando {i}/{total}...")

    def iniciar():
        btn_iniciar.configure(state="disabled")
        btn_cerrar.configure(state="disabled")
        btn_stop.configure(state="normal")

        cancel_var.set(False)
        enviando["flag"] = True

        if callable(set_ui_busy):
            set_ui_busy(True)

        status_lbl.configure(text="Procesando…")

        # inyectar args comunes
        kwargs = dict(runner_kwargs or {})
        kwargs["progress"] = modal_progress
        kwargs["root"] = modal
        kwargs["is_cancelled"] = lambda: cancel_var.get()

        # Encadenar callback si el runner ya trae on_progress
        user_on_progress = kwargs.get("on_progress", None)

        def chained(i, total, estado, *extras):
            _on_progress(i, total, estado, *extras)
            if callable(user_on_progress):
                try:
                    user_on_progress(i, total, estado, *extras)
                except Exception:
                    pass

        kwargs["on_progress"] = chained

        try:
            runner(**kwargs)
            status_lbl.configure(text="Proceso finalizado.")
        except Exception as e:
            mb.showerror("Error", str(e))
            status_lbl.configure(text="Error.")
        finally:
            enviando["flag"] = False
            btn_stop.configure(state="disabled")
            btn_cerrar.configure(state="normal")

            if callable(set_ui_busy):
                set_ui_busy(False)

    btn_cerrar = ctk.CTkButton(btns, text="Cerrar", command=on_close)
    btn_cerrar.pack(side="right")

    btn_stop = ctk.CTkButton(btns, text="STOP", command=on_stop, state="disabled")
    btn_stop.pack(side="right", padx=(0, 10))

    btn_iniciar = ctk.CTkButton(btns, text="Iniciar", command=iniciar)
    btn_iniciar.pack(side="right", padx=(0, 10))

    return modal
