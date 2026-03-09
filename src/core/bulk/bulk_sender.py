# core/bulk_sender.py
from tkinter import messagebox
import pandas as pd
import os, time, tempfile, shutil
from datetime import datetime
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment, Font

from core.shared.file_utils import (
    _asegurar_leible, 
    _copiar_a_temp_corto, 
    safe_join_file, 
    crear_tmp_dir_adjuntos,
)
from core.shared.progress import _progress_init, _progress_set
from core.shared.text_utils import _cell_text
    
from core.outlook.outlook_client import get_outlook_app

from reports.report_pdf import generar_pdfs_registro

# Tamaño único para TODOS (email/excel/pdf)
LOGO_H_PX = 60


def _logo_size_keep_height_px(path: str, target_h_px: int) -> tuple[int, int]:
    """
    Retorna (w_px, h_px) manteniendo proporción con altura fija target_h_px.
    Usa Pillow si está disponible; si no, cae a 200px de ancho por defecto.
    """
    from PIL import Image

    with Image.open(path) as im:
        w0, h0 = im.size
    if not w0 or not h0:
        return (200, target_h_px)
    w = int(round(target_h_px * (w0 / h0)))
    return (max(1, w), target_h_px)

def _prepare_logo_fixed_height(logo_path: str, target_h_px: int) -> str:
    """
    Crea una copia temporal del logo con altura fija (px) manteniendo proporción.
    Retorna la ruta del archivo temporal (png).
    Requiere Pillow; si no existe, devuelve el original.
    """
    try:
        from PIL import Image
    except ImportError:
        return logo_path  # fallback (no garantiza tamaño)

    with Image.open(logo_path) as im:
        im = im.convert("RGBA")
        w0, h0 = im.size
        if not w0 or not h0:
            return logo_path
        new_w = int(round(target_h_px * (w0 / h0)))
        resized = im.resize((max(1, new_w), target_h_px), Image.LANCZOS)

    fd, tmp_path = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    resized.save(tmp_path, format="PNG")
    return tmp_path


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
    df_preloaded=None,
    # columnas elegidas
    col_nombre: str | None = None,
    col_correo: str | None = None,
    col_archivo: str | None = None,
    logo_path: str | None = None,  # <-- NUEVO
    report_title: str = "Registro de envío de correos",
    on_progress=None,
    show_summary_messagebox: bool = True,
):
    if not ruta_excel or not os.path.isfile(ruta_excel):
        raise FileNotFoundError("Selecciona un archivo Excel (.xlsx) válido.")

    if not carpeta_archivos or not os.path.isdir(carpeta_archivos):
        raise FileNotFoundError(
            "Selecciona una carpeta válida donde estén los archivos a adjuntar."
        )

    if (generar_excel or generar_pdf) and (
        not ruta_base_reporte or not ruta_base_reporte.strip()
    ):
        raise ValueError("Selecciona dónde guardar el informe (ruta base).")

    # Reusar dataframe precargado cuando viene desde UI para evitar doble lectura de Excel.
    if df_preloaded is not None:
        df = df_preloaded.copy()
    else:
        _asegurar_leible(ruta_excel)
        df = pd.read_excel(ruta_excel, sheet_name=hoja_excel, header=header_idx)

    if not col_nombre or not col_correo or not col_archivo:
        raise ValueError(
            "Selecciona las columnas de Nombre/Correo/NombreArchivo antes de enviar."
        )

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
        ok = messagebox.askyesno(
            "Confirmar envío", f"Se enviarán {total} correos.\n\n¿Deseas continuar?"
        )
        if not ok:
            return

    logo_for_use = None
    logo_tmp_to_cleanup = None
    temp_adjuntos = []

    try:
        if logo_path and os.path.isfile(logo_path):
            fixed = _prepare_logo_fixed_height(os.path.abspath(logo_path), LOGO_H_PX)
            logo_for_use = fixed
            if os.path.abspath(fixed) != os.path.abspath(logo_path):
                logo_tmp_to_cleanup = fixed

        outlook = get_outlook_app()
        if outlook is None:
            raise RuntimeError("No se pudo acceder a Outlook.")

        _progress_init(progress, total)

        registros = []
        cancelado = False
        tmp_adjuntos_dir = crear_tmp_dir_adjuntos()

        for i, (raw_nombre, raw_correo, raw_archivo) in enumerate(
            df[[col_nombre, col_correo, col_archivo]].itertuples(index=False, name=None), start=1
        ):
            if callable(is_cancelled) and is_cancelled():
                cancelado = True
                break

            nombre = _cell_text(raw_nombre)
            correo = _cell_text(raw_correo).lower()  # opcional: normalizar a minúsculas
            nombre_archivo = _cell_text(raw_archivo)

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

                ruta_adj = safe_join_file(carpeta_archivos, nombre_archivo)

                mail = outlook.CreateItem(0)
                mail.To = correo
                mail.Subject = asunto

                # Texto -> HTML simple
                body_txt = mensaje.replace("{nombre}", nombre)
                body_html = (
                    "<html><body style='font-family:Segoe UI, Arial; font-size:11pt;'>"
                    + body_txt.replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                    .replace("\n", "<br>")
                )

                # Adjuntar archivo principal

                _asegurar_leible(ruta_adj)

                ruta_para_adjuntar = _copiar_a_temp_corto( ruta_adj, tmp_adjuntos_dir)  # evita rutas largas/locks
                temp_adjuntos.append(ruta_para_adjuntar)
                mail.Attachments.Add(ruta_para_adjuntar)
                
                # Footer con logo (si existe)

                if logo_for_use and os.path.isfile(logo_for_use):
                    cid = f"logo_footer_{i}"  # único por correo
                    att = mail.Attachments.Add(os.path.abspath(logo_for_use))
                    att.PropertyAccessor.SetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid
                    )

                    body_html += (
                        "<br><br>"
                        f"<img src='cid:{cid}' height='{LOGO_H_PX}' style='width:auto;display:block;border:0;'/>"
                        "</body></html>"
                    )
                else:
                    body_html += "</body></html>"

                mail.HTMLBody = body_html

                if callable(is_cancelled) and is_cancelled():
                    cancelado = True
                    break

                mail.Send()
                time.sleep(delay_segundos)

            except Exception as e:
                import traceback
                estado = f"Error: {str(e)}"
                print(f"[ERROR] Fila {i} - {correo} - {nombre_archivo}")
                print(traceback.format_exc())

            registros.append(
                {
                    "Nombre": nombre,
                    "Correo": correo,
                    "NombreArchivo": nombre_archivo,
                    "Fecha": fecha,
                    "Hora": hora,
                    "Estado": estado,
                }
            )

            if callable(on_progress):
                try:
                    on_progress(i, total, estado, nombre, correo)
                except Exception as progress_err:
                    print(f"on_progress falló: {progress_err}")

            _progress_set(progress, i)
            if root is not None:
                root.update_idletasks()

        # Guardar reporte con lo procesado hasta ahora
        rutas_generadas = []
        if (generar_excel or generar_pdf) and registros:
            df_reg = pd.DataFrame(registros)

            # Separación Enviados vs Errores (todo lo que NO sea "Enviado" se va a Errores)
            mask_enviado = (
                df_reg["Estado"].astype(str).str.strip().str.lower().eq("enviado")
            )
            df_enviados = df_reg[mask_enviado].copy()
            df_errores = df_reg[~mask_enviado].copy()

            for df_tmp in (df_enviados, df_errores):
                if "NombreArchivo" in df_tmp.columns:
                    df_tmp.drop(columns=["NombreArchivo"], inplace=True)

                df_tmp.insert(0, "N°", range(1, len(df_tmp) + 1))

            if generar_excel:
                ruta_excel_log = ruta_base_reporte.strip() + ".xlsx"
                titulo_txt = (
                    report_title or ""
                ).strip() or "Registro de envío de correos"

                startrow = 3  # tabla comienza en fila 4 (recomendado)

                with pd.ExcelWriter(ruta_excel_log, engine="openpyxl") as writer:
                    df_enviados.to_excel(
                        writer, sheet_name="Enviados", index=False, startrow=startrow
                    )
                    df_errores.to_excel(
                        writer, sheet_name="Errores", index=False, startrow=startrow
                    )
                    header_row = (
                        startrow + 1
                    )  # fila donde Excel escribe los headers del DF
                    start_col = 1  # A

                    def _add_excel_table(ws, df, table_name: str):
                        if df is None or df.empty:
                            return

                        nrows = len(df)
                        ncols = len(df.columns)

                        end_row = header_row + nrows  # header + data
                        end_col = start_col + ncols - 1  # A + ncols - 1

                        # Convertir número de columna a letra (A, B, C...)
                        from openpyxl.utils import get_column_letter

                        end_col_letter = get_column_letter(end_col)

                        table_ref = f"A{header_row}:{end_col_letter}{end_row}"

                        tab = Table(displayName=table_name, ref=table_ref)

                        style = TableStyleInfo(
                            name="TableStyleMedium9",
                            showFirstColumn=False,
                            showLastColumn=False,
                            showRowStripes=True,
                            showColumnStripes=False,
                        )
                        tab.tableStyleInfo = style
                        ws.add_table(tab)

                    wb = writer.book

                    ws_env = wb["Enviados"]
                    ws_err = wb["Errores"]

                    _add_excel_table(ws_env, df_enviados, "TablaEnviados")
                    _add_excel_table(ws_err, df_errores, "TablaErrores")

                    for sheet_name, df_sheet in (
                        ("Enviados", df_enviados),
                        ("Errores", df_errores),
                    ):
                        ws = wb[sheet_name]

                        # --- rango dinámico para centrar título ---
                        ncols = max(1, len(df_sheet.columns))
                        start_col = 4  # B (A queda para el logo)
                        end_col = start_col + ncols - 1  # hasta donde llegue la tabla

                        ws.merge_cells(
                            start_row=1,
                            start_column=start_col,
                            end_row=1,
                            end_column=end_col,
                        )
                        cell = ws.cell(row=1, column=start_col)
                        cell.value = titulo_txt
                        cell.alignment = Alignment(
                            horizontal="center", vertical="center"
                        )
                        cell.font = Font(bold=True, size=14)

                        ws.row_dimensions[1].height = 45

                        # --- logo en A1 ---
                        if logo_for_use and os.path.isfile(logo_for_use):
                            w_px, h_px = _logo_size_keep_height_px(
                                os.path.abspath(logo_for_use), LOGO_H_PX
                            )
                            img = XLImage(os.path.abspath(logo_for_use))
                            img.height = h_px
                            img.width = w_px
                            ws.add_image(img, "A1")

                rutas_generadas.append(ruta_excel_log)
            if generar_pdf:
                
                base_pdf = ruta_base_reporte.strip()
            
                ruta_pdf_enviados = base_pdf + "_enviados.pdf"
                ruta_pdf_errores = base_pdf + "_errores.pdf"
                
                generar_pdfs_registro(
                    registros=registros,
                    ruta_pdf_enviados=ruta_pdf_enviados,
                    ruta_pdf_errores=ruta_pdf_errores,
                    logo_path = logo_for_use,
                    report_title = f"{report_title} - Enviados",
                )
                
                rutas_generadas.append(ruta_pdf_enviados)
                rutas_generadas.append(ruta_pdf_errores)
            
        procesados = len(registros)

        if cancelado:
            msg = f"Envío CANCELADO.\n\nProcesados: {procesados} de {total}."
            if rutas_generadas:
                msg += "\n\nInformes generados:\n" + "\n".join(rutas_generadas)
            if show_summary_messagebox:
                messagebox.showwarning("Cancelado", msg)
            return {
                "cancelado": True,
                "procesados": procesados,
                "total": total,
                "rutas_generadas": rutas_generadas,
                "mensaje": msg,
            }

        if rutas_generadas:
            final_msg = "Envío terminado.\n\nInformes generados:\n" + "\n".join(
                rutas_generadas
            )
        else:
            final_msg = "Envío terminado (sin generar informe)."

        if show_summary_messagebox:
            messagebox.showinfo("Proceso finalizado", final_msg)

        return {
            "cancelado": False,
            "procesados": procesados,
            "total": total,
            "rutas_generadas": rutas_generadas,
            "mensaje": final_msg,
        }
    finally:
        # limpia logo temporal
        if logo_tmp_to_cleanup and os.path.isfile(logo_tmp_to_cleanup):
            try:
                os.remove(logo_tmp_to_cleanup)
            except OSError:
                pass

        # limpia adjuntos temporales copiados
        for p in temp_adjuntos:
            if p and os.path.isfile(p):
                try:
                    os.remove(p)
                except OSError:
                    pass
                
        if tmp_adjuntos_dir and os.path.isdir(tmp_adjuntos_dir):
                try:
                    shutil.rmtree(tmp_adjuntos_dir, ignore_errors = True)
                except OSError:
                    pass
