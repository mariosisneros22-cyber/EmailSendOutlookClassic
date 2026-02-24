# mailer.py
from tkinter import messagebox
import pandas as pd
import win32com.client as win32
import os
import time
from datetime import datetime
from openpyxl.drawing.image import Image as XLImage

from reportes import PDF_DISPONIBLE, generar_pdf_registro

# Tamaño único para TODOS (email/excel/pdf)
LOGO_H_PX = 60

# Para PDF: conversión px->points (asumiendo 96dpi)
LOGO_H_PT = LOGO_H_PX * 72 / 96
def _logo_size_keep_height_px(path: str, target_h_px: int) -> tuple[int, int]:
    """
    Retorna (w_px, h_px) manteniendo proporción con altura fija target_h_px.
    Usa Pillow si está disponible; si no, cae a 200px de ancho por defecto.
    """
    try:
        from PIL import Image
        with Image.open(path) as im:
            w0, h0 = im.size
        if not w0 or not h0:
            return (200, target_h_px)
        w = int(round(target_h_px * (w0 / h0)))
        return (max(1, w), target_h_px)
    except Exception:
        # fallback si no hay Pillow
        return (200, target_h_px)
    
    
def _cell_text(v) -> str:
    """
    Convierte valores de Excel a texto:
    - NaN/None => ""
    - strings => strip()
    - otros => str(...).strip()
    """
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def _adjunto_seguro(carpeta_base: str, nombre_archivo: str) -> str:
    if not nombre_archivo:
        raise ValueError("NombreArchivo vacío.")

    if os.path.isabs(nombre_archivo) or nombre_archivo.startswith("\\\\"):
        raise ValueError(f"NombreArchivo no puede ser ruta absoluta: {nombre_archivo}")

    base = os.path.abspath(carpeta_base)
    candidato = os.path.abspath(os.path.join(base, nombre_archivo))

    if not (candidato == base or candidato.startswith(base + os.sep)):
        raise ValueError(f"NombreArchivo apunta fuera de la carpeta base: {nombre_archivo}")

    if not os.path.isfile(candidato):
        raise FileNotFoundError(f"No existe el archivo: {candidato}")

    return candidato


def _progress_init(progress, total: int):
    """
    Soporta ttk.Progressbar (['maximum']/['value']) y customtkinter.CTkProgressBar (.set()).
    """
    if progress is None:
        return

    # CTkProgressBar: tiene método set()
    if hasattr(progress, "set") and callable(getattr(progress, "set")):
        try:
            setattr(progress, "_total", max(1, int(total)))
        except Exception:
            pass
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


def enviar_correos(
    ruta_excel: str,
    carpeta_archivos: str,
    asunto: str,
    mensaje: str,
    progress,
    root,
    ruta_base_reporte: str | None = None,
    generar_excel: bool = False,
    generar_pdf: bool = False,
    max_por_ejecucion: int = 500,
    delay_segundos: float = 1.5,
    pedir_confirmacion: bool = True,
    is_cancelled=None,
    # NUEVO: hoja del Excel y header
    hoja_excel=0,
    header_idx: int = 0,
    # columnas elegidas
    col_nombre: str | None = None,
    col_correo: str | None = None,
    col_archivo: str | None = None,
    logo_path: str | None = None,   # <-- NUEVO
    on_progress=None,
):
    if not ruta_excel or not os.path.isfile(ruta_excel):
        raise FileNotFoundError("Selecciona un archivo Excel (.xlsx) válido.")

    if not carpeta_archivos or not os.path.isdir(carpeta_archivos):
        raise FileNotFoundError("Selecciona una carpeta válida donde estén los archivos a adjuntar.")

    if (generar_excel or generar_pdf) and (not ruta_base_reporte or not ruta_base_reporte.strip()):
        raise ValueError("Selecciona dónde guardar el informe (ruta base).")

    if generar_pdf and not PDF_DISPONIBLE:
        raise RuntimeError("Para generar PDF necesitas reportlab: pip install reportlab")

    # Leer la hoja seleccionada
    df = pd.read_excel(ruta_excel, sheet_name=hoja_excel, header=header_idx)

    if not col_nombre or not col_correo or not col_archivo:
        raise ValueError("Selecciona las columnas de Nombre/Correo/NombreArchivo antes de enviar.")

    for col in (col_nombre, col_correo, col_archivo):
        if col not in df.columns:
            raise KeyError(f"La columna seleccionada no existe en el Excel: {col}")

    total = len(df)

    if total > max_por_ejecucion:
        raise ValueError(
            f"El Excel tiene {total} filas, pero el máximo por ejecución es {max_por_ejecucion}. "
            "Ajusta el límite o divide el Excel."
        )

    # Confirmación opcional (en tu nuevo flujo normalmente viene False)
    if pedir_confirmacion:
        ok = messagebox.askyesno("Confirmar envío", f"Se enviarán {total} correos.\n\n¿Deseas continuar?")
        if not ok:
            return

    try:
        outlook = win32.gencache.EnsureDispatch("Outlook.Application")
    except Exception:
        outlook = win32.Dispatch("Outlook.Application")

    _progress_init(progress, total)

    registros = []
    cancelado = False

    for i, (_, fila) in enumerate(df.iterrows(), start=1):
        if callable(is_cancelled) and is_cancelled():
            cancelado = True
            break

        nombre = _cell_text(fila[col_nombre])
        correo = _cell_text(fila[col_correo]).lower()  # opcional: normalizar a minúsculas
        nombre_archivo = _cell_text(fila[col_archivo])
        
        ahora = datetime.now()
        fecha = ahora.strftime("%Y-%m-%d")
        hora = ahora.strftime("%H:%M:%S")

        estado = "Enviado"

        try:
            if not correo:
                raise ValueError("Correo vacío.")
            # Validación simple (rápida, sin regex)
            if "@" not in correo or " " in correo:
                raise ValueError(f"Correo inválido: {correo}")
            
            if callable(is_cancelled) and is_cancelled():
                cancelado = True
                break

            ruta_adj = _adjunto_seguro(carpeta_archivos, nombre_archivo)

            mail = outlook.CreateItem(0)
            mail.To = correo
            mail.Subject = asunto

            # Texto -> HTML simple
            body_txt = mensaje.replace("{nombre}", nombre)
            body_html = (
                "<html><body style='font-family:Segoe UI, Arial; font-size:11pt;'>"
                + body_txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br>")
            )

            # Adjuntar archivo principal
            mail.Attachments.Add(ruta_adj)

            # Footer con logo (si existe)
            if logo_path and os.path.isfile(logo_path):
                cid = "logo_footer"
                att = mail.Attachments.Add(os.path.abspath(logo_path))

                pa = att.PropertyAccessor
                pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)

                body_html += (
                    "<br><br>"
                    f"<img src='cid:{cid}' style='height:{LOGO_H_PX}px; width:auto; display:block;'/>"
                    "</body></html>"
                )
            else:
                body_html += "</body></html>"

            mail.HTMLBody = body_html

            if callable(is_cancelled) and is_cancelled():
                cancelado = True
                break

            mail.Display()
            time.sleep(delay_segundos)

        except Exception as e:
            estado = f"Error: {str(e)}"

        registros.append({
            "Nombre": nombre,
            "Correo": correo,
            "NombreArchivo": nombre_archivo,
            "Fecha": fecha,
            "Hora": hora,
            "Estado": estado,
        })

        if callable(on_progress):
            try:
                on_progress(i, total, estado, nombre, correo)
            except Exception:
                pass
            
            
        _progress_set(progress, i)
        if root is not None:
            root.update_idletasks()

    # Guardar reporte con lo procesado hasta ahora
    rutas_generadas = []
    if (generar_excel or generar_pdf) and registros:
        df_reg = pd.DataFrame(registros)

        # Separación Enviados vs Errores (todo lo que NO sea "Enviado" se va a Errores)
        mask_enviado = df_reg["Estado"].astype(str).str.strip().str.lower().eq("enviado")
        df_enviados = df_reg[mask_enviado].copy()
        df_errores = df_reg[~mask_enviado].copy()

        if generar_excel:
            ruta_excel_log = ruta_base_reporte.strip() + ".xlsx"
            startrow = 5  # deja espacio arriba para el logo (ajustable)
        
            with pd.ExcelWriter(ruta_excel_log, engine="openpyxl") as writer:
                df_enviados.to_excel(writer, sheet_name="Enviados", index=False, startrow=startrow)
                df_errores.to_excel(writer, sheet_name="Errores", index=False, startrow=startrow)

                if logo_path and os.path.isfile(logo_path):
                    wb = writer.book
                    w_px, h_px = _logo_size_keep_height_px(os.path.abspath(logo_path), LOGO_H_PX)

                    for sheet_name in ("Enviados", "Errores"):
                        ws = wb[sheet_name]

                        img = XLImage(os.path.abspath(logo_path))
                        img.height = h_px
                        img.width = w_px

                        ws.add_image(img, "A1")
                        ws.row_dimensions[1].height = 45  # opcional
                        
            rutas_generadas.append(ruta_excel_log)

        if generar_pdf:
            ruta_pdf_log = ruta_base_reporte.strip() + ".pdf"
            # El PDF lo generamos con la lista original (reportes.py separa adentro)
            generar_pdf_registro(registros, ruta_pdf_log, logo_path=logo_path)
            rutas_generadas.append(ruta_pdf_log)

    procesados = len(registros)

    if cancelado:
        msg = f"Envío CANCELADO.\n\nProcesados: {procesados} de {total}."
        if rutas_generadas:
            msg += "\n\nInformes generados:\n" + "\n".join(rutas_generadas)
        messagebox.showwarning("Cancelado", msg)
        return

    if rutas_generadas:
        messagebox.showinfo(
            "Proceso finalizado",
            "Envío terminado.\n\nInformes generados:\n" + "\n".join(rutas_generadas)
        )
    else:
        messagebox.showinfo("Proceso finalizado", "Envío terminado (sin generar informe).")