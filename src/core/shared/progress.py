def _progress_init(progress, total: int):
    """
    Soporta ttk.Progressbar (['maximum']/['value']) y customtkinter.CTkProgressBar (.set()).
    """
    if progress is None:
        return

    total = max(1, int(total))

    # CTkProgressBar: tiene método set()
    if hasattr(progress, "set") and callable(getattr(progress, "set")):
        # guardar total para _progress_set
        setattr(progress, "_total", total)
        try:
            progress.set(0)
        except Exception:
            pass
        return
    
    # ttk.Progressbar (dict-style)
    try:
        progress["maximum"] = total
        progress["value"] = 0
    except Exception:
        pass


def _progress_set(progress, value: int):
    """
    Actualiza progreso:
    - ttk: progress['value'] = value
    - CTk: progress.set(value/total)
    """
    if progress is None:
        return

    if hasattr(progress, "set") and callable(getattr(progress, "set")):
        total = getattr(progress, "_total", None)
        try:
            total = int(total) if total else 1
        except Exception:
            total = 1
        total = max(1, total)

        frac = float(value) / float(total)
        frac = max(0.0, min(1.0, frac))

        try:
            progress.set(frac)
        except Exception:
            pass
        return

    try:
        progress["value"] = value
    except Exception:
        pass
