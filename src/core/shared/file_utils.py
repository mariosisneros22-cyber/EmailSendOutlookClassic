#core.shared.file_utils
import tempfile
import os
import shutil

def _asegurar_leible(path: str):
    try:
        with open(path, "rb") as f:
            f.read(1)
    except (FileNotFoundError, PermissionError, OSError) as e:
        raise FileNotFoundError(
            f"No se pudo acceder al archivo: {path}\n"
            "Si está en OneDrive, marca la carpeta/archivo como 'Mantener siempre en este dispositivo'."
        ) from e

def crear_tmp_dir_adjuntos() -> str:
    return tempfile.mkdtemp(prefix="app_correo_adjuntos_")

def _copiar_a_temp_corto(ruta: str, tmp_dir: str | None = None) -> str:
    if not tmp_dir:
        tmp_dir = os.path.join(tempfile.gettempdir(), "app_correo_adjuntos")
        os.makedirs(tmp_dir, exist_ok=True)
        
    base_name = os.path.basename(ruta)
    stem, ext = os.path.splitext(base_name)
    
    invalid_chars = '<>:"/\\|?*'
    safe_stem = "".join("_" if c in invalid_chars else c for c in stem).strip()
    safe_stem = safe_stem[:100] or "adjunto"
    
    destino = os.path.join(tmp_dir, safe_stem + ext)
    shutil.copy2(ruta, destino)
    return destino
  
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
