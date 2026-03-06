# common.py
import os
import shutil
import tempfile

import pandas as pd


def _cell_text(v) -> str:
    """
    Convierte valores de Excel a texto:
    - NaN/None => ""
    - strings => strip()
    - otros => str(...).strip()
    """
    if pd.isna(v):
        return ""
    
    return str(v).strip()


def _asegurar_leible(path: str):
    try:
        with open(path, "rb") as f:
            f.read(1)
    except (FileNotFoundError, PermissionError, OSError) as e:
        raise FileNotFoundError(
            f"No se pudo acceder al archivo: {path}\n"
            "Si está en OneDrive, marca la carpeta/archivo como 'Mantener siempre en este dispositivo'."
        ) from e



def _copiar_a_temp_corto(ruta: str) -> str:
    tmp_dir = os.path.join(tempfile.gettempdir(), "app_correo_adjuntos")
    os.makedirs(tmp_dir, exist_ok=True)
    base_name = os.path.basename(ruta)
    stem, ext = os.path.splitext(base_name)
    safe_stem = (stem or "adjunto")[:40]

    fd, destino = tempfile.mkstemp(prefix=f"{safe_stem}_", suffix=ext, dir=tmp_dir)
    os.close(fd)
    shutil.copy2(ruta, destino)
    return destino


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


def safe_join_file(carpeta_base: str, nombre_archivo: str) -> str:
    if not nombre_archivo:
        raise ValueError("NombreArchivo vacío.")

    if os.path.isabs(nombre_archivo) or nombre_archivo.startswith("\\\\"):
        raise ValueError(f"NombreArchivo no puede ser ruta absoluta: {nombre_archivo}")

    base = os.path.abspath(carpeta_base)
    candidato = os.path.abspath(os.path.join(base, nombre_archivo))

    if not (candidato == base or candidato.startswith(base + os.sep)):
        raise ValueError(
            f"NombreArchivo apunta fuera de la carpeta base: {nombre_archivo}"
        )

    if not os.path.isfile(candidato):
        raise FileNotFoundError(f"No existe el archivo: {candidato}")

    return candidato
