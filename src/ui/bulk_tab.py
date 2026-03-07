#ui/bulk_tab.py
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import pandas as pd
import customtkinter as ctk
import os, shutil

from core.bulk.bulk_sender import enviar_correos
from core.shared.text_utils import _cell_text
from ui.modal_run import run_with_modal
from ui.theme import (
    CARD_BORDER_COLOR,
    CARD_FG_COLOR,
    CHECKBOX_STYLE,
    FONT_LABEL,
    PADY_MD,
    PADY_SM,
)
from ui.widgets import (
    make_primary_button,
    make_secondary_button,
    make_tool_button,
)

PREVIEW_N = 5
preview_df = None

LOGOS_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "assets", "logo"))
LOGO_NONE_LABEL = "Ninguno"
logo_map = {}              # {"Nombre bonito": "ruta"}
selected_logo_path = None  # ruta final a usar al enviar


def seleccionar_excel():
    path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    if path:
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, path)
        cargar_hojas_excel()
        cargar_columnas_excel()  # auto-cargar headers
        


def seleccionar_carpeta():
    path = filedialog.askdirectory()
    if path:
        carpeta_entry.delete(0, tk.END)
        carpeta_entry.insert(0, path)


def seleccionar_ruta_reporte():
    sugerido = f"registro_envios_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    path = filedialog.asksaveasfilename(
        defaultextension="",
        initialfile=sugerido,
        filetypes=[("Base (sin extensión)", "*.*")]
    )
    if path:
        base, ext = os.path.splitext(path)
        if ext.lower() in (".xlsx", ".pdf", ".docx"):
            path = base
        reporte_entry.delete(0, tk.END)
        reporte_entry.insert(0, path)

def _set_excel_controls_enabled(enabled: bool):
    state_cb = "readonly" if enabled else "disabled"

    for cb in (cb_hoja, cb_nombre, cb_correo, cb_archivo):
        cb.configure(state=state_cb)
    
def _toggle_reporte_ui():
    activo = (gen_excel_var.get() or gen_pdf_var.get())
    state = "normal" if activo else "disabled"
    reporte_entry.configure(state=state)
    btn_reporte.configure(state=state)

def _get_header_index() -> int:
    """
    Devuelve el índice 0-based de la fila de encabezados para pandas (header=...).
    UI usa 1-based (1 = primera fila).
    """
    try:
        n = int(header_row_entry.get().strip())
        if n <= 0:
            raise ValueError
        return n - 1
    except (ValueError, TypeError):
        return 0
    
def cargar_hojas_excel():
    ruta = excel_entry.get().strip()
    if not ruta:
        return
    try:
        xls = pd.ExcelFile(ruta)
        hojas = list(xls.sheet_names)
        cb_hoja.configure(values=hojas)
        _set_excel_controls_enabled(True)
        
        # Selección por defecto: heurística simple
        def pick_sheet(preferencias):
            lower_map = {str(s).strip().lower(): s for s in hojas}
            for p in preferencias:
                if p in lower_map:
                    return lower_map[p]
            for s in hojas:
                sl = str(s).strip().lower()
                for p in preferencias:
                    if p in sl:
                        return s
            return hojas[0] if hojas else ""

        cb_hoja.set(pick_sheet(["data", "base", "envios", "envío", "hoja3", "sheet3", "sheet1", "hoja1"]))
    except Exception as e:
        _set_excel_controls_enabled(False)
        messagebox.showerror("Error", f"No se pudieron leer las hojas del Excel:\n{e}")
        
def cargar_columnas_excel():
    ruta = excel_entry.get().strip()
    if not ruta:
        messagebox.showerror("Error", "Selecciona primero un archivo Excel.")
        return

    try:
        sheet = cb_hoja.get().strip() if cb_hoja.get().strip() else 0
        header_idx = _get_header_index()
        df0 = pd.read_excel(ruta, sheet_name=sheet, header=header_idx, nrows=0)
        cols = list(df0.columns)

        if not cols:
            raise ValueError("El Excel no tiene columnas.")

        cb_nombre.configure(values=cols)
        cb_correo.configure(values=cols)
        cb_archivo.configure(values=cols)

        # Autoselección simple (heurística)
        def pick(preferencias):
            lower_map = {str(c).strip().lower(): c for c in cols}
            for p in preferencias:
                if p in lower_map:
                    return lower_map[p]
            for c in cols:
                cl = str(c).strip().lower()
                for p in preferencias:
                    if p in cl:
                        return c
            return ""

        cb_nombre.set(pick(["nombre", "nombres", "name"]))
        cb_correo.set(pick(["correo", "email", "e-mail", "mail"]))
        cb_archivo.set(pick(["nombrearchivo", "nombre_archivo", "archivo", "adjunto", "file", "documento"]))

        _load_preview_df(PREVIEW_N)
        actualizar_vista_previa()

        messagebox.showinfo("Listo", "Columnas cargadas. Revisa/ajusta el mapeo y luego envía.")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el Excel:\n{e}")


# -------- VISTA PREVIA --------

def _load_preview_df(n=PREVIEW_N):
    global preview_df
    ruta = excel_entry.get().strip()
    preview_df = None

    if not ruta:
        return

    try:
        sheet = cb_hoja.get().strip() if cb_hoja.get().strip() else 0
        header_idx = _get_header_index()
        dfp = pd.read_excel(ruta, sheet_name=sheet, header=header_idx, nrows=n)
        if dfp.empty:
            preview_df = None
            return
        preview_df = dfp
    except Exception as e:
        preview_df = None
        messagebox.showerror("Error", f"No se pudo cargar la vista previa:\n{e}")


def actualizar_tabla_preview():
    if preview_box is None:
        return

    preview_box.configure(state="normal")
    preview_box.delete("1.0", "end")

    if preview_df is None or preview_df.empty:
        preview_box.insert("1.0", "(sin datos para previsualizar)")
        preview_box.configure(state="disabled")
        return

    col_n = cb_nombre.get().strip()
    col_c = cb_correo.get().strip()
    col_a = cb_archivo.get().strip()

    if not all(c and c in preview_df.columns for c in (col_n, col_c, col_a)):
        preview_box.insert("1.0", "Selecciona columnas validas para la vista previa.")
        preview_box.configure(state="disabled")
        return

    def _clip(text: str, n: int) -> str:
        text = str(text)
        return text if len(text) <= n else text[: n - 3] + "..."

    header = f"{'#':>2} | {'Nombre':<26} | {'Correo':<30} | {'NombreArchivo':<30}"
    sep = "-" * len(header)
    lines = [header, sep]

    for idx, (v_n_raw, v_c_raw, v_a_raw) in enumerate(
        preview_df[[col_n, col_c, col_a]].itertuples(index=False, name=None), start=1
    ):
        v_n = _cell_text(v_n_raw)
        v_c = _cell_text(v_c_raw)
        v_a = _cell_text(v_a_raw)
        lines.append(
            f"{idx:>2} | {_clip(v_n, 26):<26} | {_clip(v_c, 30):<30} | {_clip(v_a, 30):<30}"
        )

    preview_box.insert("1.0", "\n".join(lines))
    preview_box.configure(state="disabled")


def actualizar_vista_previa(event=None):
    col_n = cb_nombre.get().strip()
    col_c = cb_correo.get().strip()
    col_a = cb_archivo.get().strip()

    actualizar_tabla_preview()

def cargar_logos_predeterminados():
    """
    Carga logos desde ./logo (png/jpg/jpeg/gif).
    Siempre incluye la opción 'Ninguno' como primera opción.
    """
    global logo_map
    logo_map = {}

    opciones = [LOGO_NONE_LABEL]

    if os.path.isdir(LOGOS_DIR):
        exts = (".png", ".jpg", ".jpeg", ".gif")
        files = [f for f in os.listdir(LOGOS_DIR) if f.lower().endswith(exts)]
        files.sort()

        for f in files:
            name = os.path.splitext(f)[0]
            logo_map[name] = os.path.join(LOGOS_DIR, f)
            opciones.append(name)

    cb_logo.configure(values=opciones)

    # Mantener selección si existe; si no, volver a Ninguno
    actual = cb_logo.get().strip()
    if actual in opciones:
        cb_logo.set(actual)
    else:
        cb_logo.set(LOGO_NONE_LABEL)
        
def seleccionar_logo():
    cargar_logos_predeterminados()
    # opcional: feedback
    # lbl_logo_estado.configure(text="Logos actualizados.")
    
def insertar_logo():
    """
    Permite elegir un archivo de imagen y lo copia a ./logo para que quede como predeterminado.
    Luego refresca el combobox y lo deja seleccionado.
    """
    global selected_logo_path

    file_path = filedialog.askopenfilename(
        title="Seleccionar logo",
        filetypes=[("Imágenes", "*.png;*.jpg;*.jpeg;*.gif")]
    )
    if not file_path:
        return

    # Crear carpeta si no existe
    os.makedirs(LOGOS_DIR, exist_ok=True)

    base = os.path.basename(file_path)
    name, ext = os.path.splitext(base)

    # Evitar sobreescribir: si existe, agrega sufijo _1, _2...
    dest = os.path.join(LOGOS_DIR, base)
    k = 1
    while os.path.exists(dest):
        dest = os.path.join(LOGOS_DIR, f"{name}_{k}{ext}")
        k += 1

    shutil.copy2(file_path, dest)

    # Refrescar lista y seleccionar el nuevo
    cargar_logos_predeterminados()
    new_display_name = os.path.splitext(os.path.basename(dest))[0]
    if new_display_name in logo_map:
        # aplicar selección inmediatamente
        cb_logo.set(new_display_name)
        aplicar_logo_seleccionado()
    else:
        cb_logo.set(LOGO_NONE_LABEL)
        selected_logo_path = None
        lbl_logo_estado.configure(text="Logo: (ninguno)")


def aplicar_logo_seleccionado(event=None):
    """
    Aplica el logo elegido en el combobox como logo activo.
    """
    global selected_logo_path
    name = cb_logo.get().strip()

    if not name or name == LOGO_NONE_LABEL:
        selected_logo_path = None
        lbl_logo_estado.configure(text="Logo: (ninguno)")
        return

    if name not in logo_map:
        selected_logo_path = None
        lbl_logo_estado.configure(text="Logo: (ninguno)")
        return

    selected_logo_path = logo_map[name]
    lbl_logo_estado.configure(text=f"Logo: {name}")
    
    
def _set_ui_enviando(enviando: bool):
    state_inputs = "disabled" if enviando else "normal"

    btn_excel.configure(state=state_inputs)
    btn_carpeta.configure(state=state_inputs)
    excel_entry.configure(state=state_inputs)
    carpeta_entry.configure(state=state_inputs)
    asunto_entry.configure(state=state_inputs)
    mensaje_text.configure(state=state_inputs)

    cb_nombre.configure(state="disabled" if enviando else "readonly")
    cb_correo.configure(state="disabled" if enviando else "readonly")
    cb_archivo.configure(state="disabled" if enviando else "readonly")
    btn_cargar_cols.configure(state=state_inputs)

    chk_excel.configure(state=state_inputs)
    chk_pdf.configure(state=state_inputs)

    reporte_entry.configure(state=state_inputs if (gen_excel_var.get() or gen_pdf_var.get()) else "disabled")
    btn_reporte.configure(state=state_inputs if (gen_excel_var.get() or gen_pdf_var.get()) else "disabled")

    max_entry.configure(state=state_inputs)
    delay_entry.configure(state=state_inputs)

    btn_enviar.configure(state="disabled" if enviando else "normal")



def abrir_modal_envio_y_ejecutar():
    """
    Modal que confirma y luego muestra progreso.
    - NO usa messagebox.askyesno
    - Usa progressbar dentro del modal
    - Usa STOP con cancelación cooperativa
    """
    ruta_excel = excel_entry.get().strip()
    carpeta = carpeta_entry.get().strip()
    asunto = asunto_entry.get().strip()
    mensaje = mensaje_text.get("1.0", tk.END).rstrip("\n")

    col_nombre = cb_nombre.get().strip()
    col_correo = cb_correo.get().strip()
    col_archivo = cb_archivo.get().strip()

    if not ruta_excel:
        messagebox.showerror("Error", "Selecciona un archivo Excel.")
        return

    if not carpeta:
        messagebox.showerror("Error", "Selecciona una carpeta de archivos.")
        return

    if not col_nombre or not col_correo or not col_archivo:
        messagebox.showerror("Error", "Selecciona las 3 columnas: Nombre, Correo y NombreArchivo.")
        return

    # ---- parse numéricos: captura específica ----
    try:
        max_envios = int(max_entry.get().strip())
    except (ValueError, TypeError):
        messagebox.showerror("Error", "El 'Máximo por ejecución' debe ser un entero > 0.")
        return
    if max_envios <= 0:
        messagebox.showerror("Error", "El 'Máximo por ejecución' debe ser un entero > 0.")
        return

    try:
        delay = float(delay_entry.get().strip())
    except (ValueError, TypeError):
        messagebox.showerror("Error", "El 'Delay (segundos)' debe ser un número >= 0.")
        return
    if delay < 0:
        messagebox.showerror("Error", "El 'Delay (segundos)' debe ser un número >= 0.")
        return

    hoja_excel = cb_hoja.get().strip() if cb_hoja.get().strip() else 0
    header_idx = _get_header_index()

    # ---- si reporte está activo, exige ruta ----
    gen_excel = bool(gen_excel_var.get())
    gen_pdf = bool(gen_pdf_var.get())
    ruta_base_reporte = reporte_entry.get().strip() if (gen_excel or gen_pdf) else None

    if (gen_excel or gen_pdf) and not ruta_base_reporte:
        messagebox.showerror("Error", "Elige la 'Ruta base del informe (sin extensión)'.")
        return

    report_title = titulo_reporte_entry.get().strip()

    # ---- calcular total sin cargar todo el Excel ----
    try:
        df_send = pd.read_excel(
            ruta_excel,
            sheet_name=hoja_excel,
            header=header_idx,
            usecols=[col_nombre, col_correo, col_archivo],
        )
    except ValueError as e:
        # ValueError típico: hoja no existe / usecols no coincide
        messagebox.showerror("Error", f"No se pudo leer el Excel (hoja/columnas):\n{e}")
        return
    except FileNotFoundError:
        messagebox.showerror("Error", "No se encontró el archivo Excel seleccionado.")
        return
    except PermissionError:
        messagebox.showerror("Error", "No se pudo abrir el Excel (archivo en uso o sin permisos).")
        return
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el Excel:\n{e}")
        return

    total = len(df_send)
    if total <= 0:
        messagebox.showerror("Error", "El Excel no tiene filas.")
        return

    if total > max_envios:
        messagebox.showerror(
            "Error",
            f"El Excel tiene {total} filas, pero el máximo por ejecución es {max_envios}."
        )
        return

    runner_kwargs = dict(
        ruta_excel=ruta_excel,
        carpeta_archivos=carpeta,
        asunto=asunto,
        mensaje=mensaje,
        ruta_base_reporte=ruta_base_reporte,
        generar_excel=gen_excel,
        generar_pdf=gen_pdf,
        max_por_ejecucion=max_envios,
        delay_segundos=delay,
        pedir_confirmacion=False,
        col_nombre=col_nombre,
        col_correo=col_correo,
        col_archivo=col_archivo,
        hoja_excel=hoja_excel,
        header_idx=header_idx,  # usa el ya calculado
        df_preloaded=df_send,
        logo_path=selected_logo_path,
        report_title=report_title,
        show_summary_messagebox=False,
    )

    try:
        run_with_modal(
            root=root,
            title="Confirmar envío",
            confirm_text=f"Se enviarán {total} correos.\n¿Deseas continuar?",
            total=total,
            runner=enviar_correos,
            runner_kwargs=runner_kwargs,
            set_ui_busy=_set_ui_enviando,
        )
    except Exception as e:
        messagebox.showerror("Error", str(e))
        

def mount(parent):
    """
    Construye la UI de envío masivo dentro de `parent` (CTkFrame o tab).
    """
    global root
    global excel_entry, carpeta_entry, asunto_entry, mensaje_text
    global cb_hoja, cb_nombre, cb_correo, cb_archivo
    global btn_excel, btn_carpeta, btn_cargar_cols, btn_enviar
    global preview_box, cb_logo, lbl_logo_estado
    global gen_excel_var, gen_pdf_var, chk_excel, chk_pdf
    global reporte_entry, btn_reporte, titulo_reporte_entry
    global max_entry, delay_entry, header_row_entry
    root = parent.winfo_toplevel()

    frame_main = ctk.CTkFrame(parent, fg_color="transparent")
    frame_main.pack(fill="both", expand=True, padx=10, pady=10)
    frame_main.rowconfigure(0, weight=1)
    frame_main.columnconfigure(0, weight=1)
    frame_main.columnconfigure(1, weight=1)

    frame_left = ctk.CTkFrame(
        frame_main,
        corner_radius=12,
        fg_color=CARD_FG_COLOR,
        border_width=1,
        border_color=CARD_BORDER_COLOR,
    )
    frame_left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
    frame_right = ctk.CTkFrame(
        frame_main,
        corner_radius=12,
        fg_color=CARD_FG_COLOR,
        border_width=1,
        border_color=CARD_BORDER_COLOR,
    )
    frame_right.grid(row=0, column=1, sticky="nsew")

    # Contenedores internos con padding global (gap)
    left_content = ctk.CTkFrame(frame_left, fg_color="transparent")
    left_content.pack(fill="both", expand=True, padx=20, pady=20)

    right_content = ctk.CTkFrame(frame_right, fg_color="transparent")
    right_content.pack(fill="both", expand=True, padx=20, pady=20)

    def _simple_section(parent, title: str):
        ctk.CTkLabel(parent, text=title, font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
        body = ctk.CTkFrame(parent, fg_color="transparent")
        body.pack(fill="x", pady=(0, PADY_MD))
        return body

    sec_left_input = _simple_section(left_content, "Datos de entrada")
    sec_left_preview = _simple_section(left_content, f"Vista previa ({PREVIEW_N} filas)")

    sec_right_brand = _simple_section(right_content, "Mensaje y marca")
    sec_right_send = _simple_section(right_content, "Parametros de envio")
    sec_right_report = _simple_section(right_content, "Informe / Registro")
    sec_right_action = ctk.CTkFrame(right_content, fg_color="transparent")
    sec_right_action.pack(fill="x", pady=(0, PADY_MD))

    # Excel
    ctk.CTkLabel(sec_left_input, text="Archivo Excel", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
    row_excel = ctk.CTkFrame(sec_left_input, fg_color="transparent")
    row_excel.pack(fill="x", pady=(0, PADY_MD))
    row_excel.grid_columnconfigure(0, weight=1)
    excel_entry = ctk.CTkEntry(row_excel)
    excel_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
    btn_excel = make_secondary_button(
        row_excel, text="Seleccionar Excel", command=seleccionar_excel, width=170
    )
    btn_excel.grid(row=0, column=1, sticky="e")

    # --- HOJA + MAPEO (layout 2 columnas: Título | Celda) ---
    excel_map_frame = ctk.CTkFrame(sec_left_input, fg_color="transparent")
    excel_map_frame.pack(fill="x", pady=(0, PADY_MD))

    excel_map_frame.grid_columnconfigure(0, weight=0)  # títulos
    excel_map_frame.grid_columnconfigure(1, weight=1)  # celdas (se estiran)

    ROW_PADY = 6

    # Hoja del Excel
    ctk.CTkLabel(excel_map_frame, text="Hoja del Excel", font=FONT_LABEL).grid(
        row=0, column=0, sticky="w", padx=(0, 12), pady=ROW_PADY
    )
    cb_hoja = ctk.CTkComboBox(
        excel_map_frame,
        values=[],
        state="disabled",
        command=lambda _v: cargar_columnas_excel(),
    )
    cb_hoja.grid(row=0, column=1, sticky="ew", pady=ROW_PADY)

    # (Opcional) Fila de encabezados (si la implementas)
    ctk.CTkLabel(excel_map_frame, text="Fila de encabezados", font=FONT_LABEL).grid(
        row=1, column=0, sticky="w", padx=(0, 12), pady=ROW_PADY
    )
    header_row_entry = ctk.CTkEntry(excel_map_frame, width=120)
    header_row_entry.insert(0, "1")
    header_row_entry.grid(row=1, column=1, sticky="ew", pady=ROW_PADY)

    # Columna Nombre
    ctk.CTkLabel(excel_map_frame, text="Columna Nombre", font=FONT_LABEL).grid(
        row=2, column=0, sticky="w", padx=(0, 12), pady=ROW_PADY
    )
    cb_nombre = ctk.CTkComboBox(
        excel_map_frame,
        values=[],
        state="disabled",
        command=actualizar_vista_previa,
    )
    cb_nombre.grid(row=2, column=1, sticky="ew", pady=ROW_PADY)

    # Columna Correo
    ctk.CTkLabel(excel_map_frame, text="Columna Correo", font=FONT_LABEL).grid(
        row=3, column=0, sticky="w", padx=(0, 12), pady=ROW_PADY
    )
    cb_correo = ctk.CTkComboBox(
        excel_map_frame,
        values=[],
        state="disabled",
        command=actualizar_vista_previa,
    )
    cb_correo.grid(row=3, column=1, sticky="ew", pady=ROW_PADY)

    # Columna NombreArchivo
    ctk.CTkLabel(excel_map_frame, text="Columna NombreArchivo", font=FONT_LABEL).grid(
        row=4, column=0, sticky="w", padx=(0, 12), pady=ROW_PADY
    )
    cb_archivo = ctk.CTkComboBox(
        excel_map_frame,
        values=[],
        state="disabled",
        command=actualizar_vista_previa,
    )
    cb_archivo.grid(row=4, column=1, sticky="ew", pady=ROW_PADY)

    # Botón cargar columnas (alineado a la derecha)
    btn_cargar_cols = make_secondary_button(
        excel_map_frame,
        text="Cargar columnas desde Excel",
        command=cargar_columnas_excel,
    )
    btn_cargar_cols.grid(row=5, column=1, sticky="e", pady=(ROW_PADY, 0))

    # Vista previa (CTkTextbox)
    frame_table = ctk.CTkFrame(sec_left_preview)
    frame_table.pack(fill="both", expand=True)
    preview_box = ctk.CTkTextbox(frame_table, height=170)
    preview_box.pack(fill="both", expand=True, padx=10, pady=10)
    preview_box.insert("1.0", "(sin datos para previsualizar)")
    preview_box.configure(state="disabled")

    # Carpeta
    ctk.CTkLabel(sec_left_input, text="Carpeta donde estan TODOS los archivos", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
    row_carpeta = ctk.CTkFrame(sec_left_input, fg_color="transparent")
    row_carpeta.pack(fill="x", pady=(0, PADY_MD))
    row_carpeta.grid_columnconfigure(0, weight=1)
    carpeta_entry = ctk.CTkEntry(row_carpeta)
    carpeta_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
    btn_carpeta = make_secondary_button(
        row_carpeta, text="Seleccionar carpeta", command=seleccionar_carpeta, width=170
    )
    btn_carpeta.grid(row=0, column=1, sticky="e")



    # --- Logos (arriba de Asunto) ---
    frame_logo = ctk.CTkFrame(sec_right_brand, fg_color="transparent")
    frame_logo.pack(fill="x", pady=(0, PADY_MD))

    frame_logo.grid_columnconfigure(0, weight=0)  # "Logo"
    frame_logo.grid_columnconfigure(1, weight=1)  # combobox
    frame_logo.grid_columnconfigure(2, weight=0)  # Seleccionar logo
    frame_logo.grid_columnconfigure(3, weight=0)  # Insertar logo

    ctk.CTkLabel(frame_logo, text="Logo", font=FONT_LABEL).grid(
        row=0, column=0, sticky="w", padx=(0, 12), pady=6
    )

    cb_logo = ctk.CTkComboBox(
        frame_logo,
        values=[LOGO_NONE_LABEL],
        state="readonly",
        command=aplicar_logo_seleccionado,
    )
    cb_logo.grid(row=0, column=1, sticky="ew", pady=6)

    btn_sel_logo = make_tool_button(
        frame_logo,
        text="Actualizar lista",
        command=seleccionar_logo,
        width=150,
    )
    btn_sel_logo.grid(row=0, column=2, sticky="e", padx=(12, 0), pady=6)

    btn_ins_logo = make_tool_button(
        frame_logo,
        text="Insertar logo",
        command=insertar_logo,
        width=140,
    )
    btn_ins_logo.grid(row=0, column=3, sticky="e", padx=(12, 0), pady=6)

    lbl_logo_estado = ctk.CTkLabel(sec_right_brand, text="Logo: (ninguno)")
    lbl_logo_estado.pack(anchor="w", pady=(0, PADY_MD))

    # Asunto
    ctk.CTkLabel(sec_right_brand, text="Asunto", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
    asunto_entry = ctk.CTkEntry(sec_right_brand)
    asunto_entry.pack(fill="x", pady=(0, PADY_MD))

    # Mensaje
    ctk.CTkLabel(sec_right_brand, text="Mensaje", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
    mensaje_text = ctk.CTkTextbox(sec_right_brand, height=200)
    mensaje_text.insert(
        "1.0",
        "Hola,\n\n"
        "Te envío el documento solicitado.\n"
        "Cualquier duda, quedo atento.\n\n"
        "Saludos,\n"
    )
    mensaje_text.pack(fill="both", expand=True, pady=(0, PADY_MD))

    # Controles: límite y delay
    frame_envio = ctk.CTkFrame(sec_right_send, fg_color="transparent")
    frame_envio.pack(fill="x")
    frame_envio.columnconfigure(1, weight=0)
    frame_envio.columnconfigure(3, weight=0)

    ctk.CTkLabel(frame_envio, text="Máximo por ejecución", font=FONT_LABEL).grid(row=0, column=0, sticky="w", pady=4)
    max_entry = ctk.CTkEntry(frame_envio, width=80)
    max_entry.insert(0, "500")
    max_entry.grid(row=0, column=1, padx=(10, 20), pady=4)

    ctk.CTkLabel(frame_envio, text="Delay (segundos) entre envíos", font=FONT_LABEL).grid(row=0, column=2, sticky="w", pady=4)
    delay_entry = ctk.CTkEntry(frame_envio, width=80)
    delay_entry.insert(0, "0")
    delay_entry.grid(row=0, column=3, padx=(10, 0), pady=4)

    frame_checks = ctk.CTkFrame(sec_right_report, fg_color="transparent")
    frame_checks.pack(fill="x", pady=(0, PADY_MD))
    frame_checks.grid_columnconfigure(0, weight=1)
    frame_checks.grid_columnconfigure(1, weight=1)

    gen_excel_var = tk.BooleanVar(master=root, value=True)
    gen_pdf_var = tk.BooleanVar(master=root, value=False)

    chk_excel = ctk.CTkCheckBox(
        frame_checks,
        text="Generar registro en Excel (.xlsx)",
        variable=gen_excel_var,
        command=_toggle_reporte_ui,
        **CHECKBOX_STYLE,
    )
    chk_excel.grid(row=0, column=0, sticky="w", pady=(0, 6), padx=(0, 12))

    chk_pdf = ctk.CTkCheckBox(
        frame_checks,
        text="Generar registro en PDF (.pdf)",
        variable=gen_pdf_var,
        command=_toggle_reporte_ui,
        **CHECKBOX_STYLE,
    )
    chk_pdf.grid(row=0, column=1, sticky="w", pady=(0, 6))

    ctk.CTkLabel(sec_right_report, text="Ruta base del informe (sin extension)", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
    row_reporte = ctk.CTkFrame(sec_right_report, fg_color="transparent")
    row_reporte.pack(fill="x", pady=(0, PADY_SM))
    row_reporte.grid_columnconfigure(0, weight=1)
    reporte_entry = ctk.CTkEntry(row_reporte)
    reporte_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

    btn_reporte = make_secondary_button(
        row_reporte,
        text="Elegir donde guardar informe",
        command=seleccionar_ruta_reporte,
        width=220,
    )
    btn_reporte.grid(row=0, column=1, sticky="e")

    _toggle_reporte_ui()


    # --- Título del reporte (antes de ENVIAR) ---
    ctk.CTkLabel(sec_right_report, text="Titulo del reporte", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
    titulo_reporte_entry = ctk.CTkEntry(sec_right_report)
    titulo_reporte_entry.insert(0, "Registro de envío de correos")  # default
    titulo_reporte_entry.pack(fill="x", pady=(0, PADY_MD))

    # Botones
    btn_enviar = make_primary_button(
        sec_right_action,
        text="ENVIAR CORREOS",
        command=abrir_modal_envio_y_ejecutar,
        height=40,
    )
    btn_enviar.pack(fill="x")


    _set_excel_controls_enabled(False)
    cargar_logos_predeterminados()
    aplicar_logo_seleccionado()
