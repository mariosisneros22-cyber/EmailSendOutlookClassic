#ui/bulk_tab.py
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import pandas as pd
import customtkinter as ctk
import os, shutil

from mailer import enviar_correos
from core.common import _cell_text
from ui.modal_run import run_with_modal

PREVIEW_N = 5
preview_df = None
preview_row = None

# Espaciado consistente (simple)
PADY_SM = 6
PADY_MD = 10

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
        base, ext = path.rsplit(".", 1) if "." in path and len(path.split(".")) > 1 else (path, "")
        if ext.lower() in ("xlsx", "pdf", "docx"):
            path = base
        reporte_entry.delete(0, tk.END)
        reporte_entry.insert(0, path)

def _set_excel_controls_enabled(enabled: bool):
    state_cb = "readonly" if enabled else "disabled"
    style_name = "Enabled.TCombobox" if enabled else "Disabled.TCombobox"

    for cb in (cb_hoja, cb_nombre, cb_correo, cb_archivo):
        cb.configure(state=state_cb)
        cb.configure(style=style_name)
    
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
    except Exception:
        # fallback seguro
        return 0
    
def cargar_hojas_excel():
    ruta = excel_entry.get().strip()
    if not ruta:
        return
    try:
        xls = pd.ExcelFile(ruta)
        hojas = list(xls.sheet_names)
        cb_hoja["values"] = hojas
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

        cb_nombre["values"] = cols
        cb_correo["values"] = cols
        cb_archivo["values"] = cols

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
    global preview_df, preview_row
    ruta = excel_entry.get().strip()
    preview_df = None
    preview_row = None

    if not ruta:
        return

    try:
        sheet = cb_hoja.get().strip() if cb_hoja.get().strip() else 0
        header_idx = _get_header_index()
        dfp = pd.read_excel(ruta, sheet_name=sheet,header=header_idx, nrows=n)
        if dfp.empty:
            preview_df = None
            preview_row = None
            return
        preview_df = dfp
        preview_row = dfp.iloc[0].to_dict()
    except Exception:
        preview_df = None
        preview_row = None


def actualizar_tabla_preview():
    for item in preview_tree.get_children():
        preview_tree.delete(item)

    if preview_df is None or preview_df.empty:
        return

    col_n = cb_nombre.get().strip()
    col_c = cb_correo.get().strip()
    col_a = cb_archivo.get().strip()

    for idx, (_, row) in enumerate(preview_df.iterrows(), start=1):
        v_n = _cell_text(row[col_n]) if col_n and col_n in preview_df.columns else ""
        v_c = _cell_text(row[col_c]) if col_c and col_c in preview_df.columns else ""
        v_a = _cell_text(row[col_a]) if col_a and col_a in preview_df.columns else ""
        preview_tree.insert("", "end", values=(idx, v_n, v_c, v_a))


def actualizar_vista_previa(event=None):
    col_n = cb_nombre.get().strip()
    col_c = cb_correo.get().strip()
    col_a = cb_archivo.get().strip()

    _ = _cell_text(preview_row.get(col_n, "")) if col_n and preview_row else ""
    _ = _cell_text(preview_row.get(col_c, "")) if col_c and preview_row else ""
    _ = _cell_text(preview_row.get(col_a, "")) if col_a and preview_row else ""

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

    cb_logo["values"] = opciones

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
        cb_logo.set(new_display_name)
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
    # ---------- Validaciones rápidas (reusa lo que ya tienes) ----------
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

    try:
        max_envios = int(max_entry.get().strip())
        if max_envios <= 0:
            raise ValueError
    except Exception:
        messagebox.showerror("Error", "El 'Máximo por ejecución' debe ser un entero > 0.")
        return

    try:
        delay = float(delay_entry.get().strip())
        if delay < 0:
            raise ValueError
    except Exception:
        messagebox.showerror("Error", "El 'Delay (segundos)' debe ser un número >= 0.")
        return

    # Si todavía NO implementaste selector de hoja, deja hoja_excel = 0
    hoja_excel = cb_hoja.get().strip() if cb_hoja.get().strip() else 0

    # Calcular total para mostrar en el modal (y validar el límite)
    try:
        header_idx = _get_header_index()
        df = pd.read_excel(ruta_excel, sheet_name=hoja_excel, header=header_idx)
        total = len(df)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el Excel:\n{e}")
        return

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
        ruta_base_reporte=reporte_entry.get().strip() if (gen_excel_var.get() or gen_pdf_var.get()) else None,
        generar_excel=gen_excel_var.get(),
        generar_pdf=gen_pdf_var.get(),
        max_por_ejecucion=max_envios,
        delay_segundos=delay,
        pedir_confirmacion=False,
        col_nombre=col_nombre,
        col_correo=col_correo,
        col_archivo=col_archivo,
        hoja_excel=hoja_excel,
        header_idx=_get_header_index(),
        logo_path=selected_logo_path,
        report_title=titulo_reporte_entry.get().strip(),
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
        

FONT_LABEL = ("Segoe UI", 13, "bold")
FONT_BUTTON = ("Segoe UI", 13, "bold")




def mount(parent):
    """
    Construye la UI de envío masivo dentro de `parent` (CTkFrame o tab).
    """
    global root
    global excel_entry, carpeta_entry, asunto_entry, mensaje_text
    global cb_hoja, cb_nombre, cb_correo, cb_archivo
    global btn_excel, btn_carpeta, btn_cargar_cols, btn_enviar
    global preview_tree, cb_logo, lbl_logo_estado
    global gen_excel_var, gen_pdf_var, chk_excel, chk_pdf
    global reporte_entry, btn_reporte, titulo_reporte_entry
    global max_entry, delay_entry, header_row_entry
    root = parent.winfo_toplevel()

    # Estilo ttk (AHORA aquí adentro)
    style = ttk.Style()
    try:
        style.theme_use("clam")
        style.configure("Enabled.TCombobox", fieldbackground="white", background="white", foreground="black")
        style.map("Enabled.TCombobox",
                  fieldbackground=[("readonly", "white"), ("!disabled", "white")],
                  foreground=[("readonly", "black"), ("!disabled", "black")])

        style.configure("Disabled.TCombobox", fieldbackground="#e6e6e6", background="#e6e6e6", foreground="#7a7a7a")
        style.map("Disabled.TCombobox",
                  fieldbackground=[("disabled", "#e6e6e6")],
                  foreground=[("disabled", "#7a7a7a")])
    except Exception:
        pass

    style.configure("Treeview", rowheight=24)
    style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
    style.configure("TCombobox", padding=4)

    frame_main = ctk.CTkFrame(parent)
    frame_main.pack(fill="both", expand=True, padx=10, pady=10)
    frame_main.columnconfigure(0, weight=1)
    frame_main.columnconfigure(1, weight=1)
    frame_left = ctk.CTkFrame(frame_main, corner_radius=12)
    frame_left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
    frame_right = ctk.CTkFrame(frame_main, corner_radius=12)
    frame_right.grid(row=0, column=1, sticky="nsew")

    # Contenedores internos con padding global (gap)
    left_content = ctk.CTkFrame(frame_left, fg_color="transparent")
    left_content.pack(fill="both", expand=True, padx=20, pady=20)

    right_content = ctk.CTkFrame(frame_right, fg_color="transparent")
    right_content.pack(fill="both", expand=True, padx=20, pady=20)

    # Excel
    ctk.CTkLabel(left_content, text="Archivo Excel", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
    excel_entry = ctk.CTkEntry(left_content)
    excel_entry.pack(fill="x", pady=(0, PADY_MD))
    btn_excel = ctk.CTkButton(left_content, text="Seleccionar Excel", command=seleccionar_excel, font=FONT_BUTTON)
    btn_excel.pack(fill="x", pady=(0, PADY_MD))

    # --- HOJA + MAPEO (layout 2 columnas: Título | Celda) ---
    excel_map_frame = ctk.CTkFrame(left_content, fg_color="transparent")
    excel_map_frame.pack(fill="x", pady=(0, PADY_MD))

    excel_map_frame.grid_columnconfigure(0, weight=0)  # títulos
    excel_map_frame.grid_columnconfigure(1, weight=1)  # celdas (se estiran)

    ROW_PADY = 6

    # Hoja del Excel
    ctk.CTkLabel(excel_map_frame, text="Hoja del Excel", font=FONT_LABEL).grid(
        row=0, column=0, sticky="w", padx=(0, 12), pady=ROW_PADY
    )
    cb_hoja = ttk.Combobox(excel_map_frame, state="disabled", style="Disabled.TCombobox")
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
    cb_nombre = ttk.Combobox(excel_map_frame, state="disabled", style="Disabled.TCombobox")
    cb_nombre.grid(row=2, column=1, sticky="ew", pady=ROW_PADY)

    # Columna Correo
    ctk.CTkLabel(excel_map_frame, text="Columna Correo", font=FONT_LABEL).grid(
        row=3, column=0, sticky="w", padx=(0, 12), pady=ROW_PADY
    )
    cb_correo = ttk.Combobox(excel_map_frame, state="disabled", style="Disabled.TCombobox")
    cb_correo.grid(row=3, column=1, sticky="ew", pady=ROW_PADY)

    # Columna NombreArchivo
    ctk.CTkLabel(excel_map_frame, text="Columna NombreArchivo", font=FONT_LABEL).grid(
        row=4, column=0, sticky="w", padx=(0, 12), pady=ROW_PADY
    )
    cb_archivo = ttk.Combobox(excel_map_frame, state="disabled", style="Disabled.TCombobox")
    cb_archivo.grid(row=4, column=1, sticky="ew", pady=ROW_PADY)

    # Botón cargar columnas (alineado a la derecha)
    btn_cargar_cols = ctk.CTkButton(
        excel_map_frame,
        text="Cargar columnas desde Excel",
        command=cargar_columnas_excel,
        font=FONT_BUTTON
    )
    btn_cargar_cols.grid(row=5, column=1, sticky="e", pady=(ROW_PADY, 0))

    # Vista previa (Treeview ttk)
    ctk.CTkLabel(left_content, text=f"Vista previa ({PREVIEW_N} filas)", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
    # Nota: frame_table NO transparente para que ttk no choque con fondos
    frame_table = ctk.CTkFrame(left_content)
    frame_table.pack(fill="both", expand=True, pady=(0, PADY_MD))

    columns = ("#", "Nombre", "Correo", "NombreArchivo")
    preview_tree = ttk.Treeview(frame_table, columns=columns, show="headings", height=PREVIEW_N)
    preview_tree.heading("#", text="#")
    preview_tree.heading("Nombre", text="Nombre (según mapeo)")
    preview_tree.heading("Correo", text="Correo (según mapeo)")
    preview_tree.heading("NombreArchivo", text="NombreArchivo (según mapeo)")

    preview_tree.column("#", width=40, anchor="center", stretch=False)
    preview_tree.column("Nombre", width=150, anchor="w", stretch=True)
    preview_tree.column("Correo", width=170, anchor="w", stretch=True)
    preview_tree.column("NombreArchivo", width=170, anchor="w", stretch=True)

    scroll_y = ttk.Scrollbar(frame_table, orient="vertical", command=preview_tree.yview)
    preview_tree.configure(yscrollcommand=scroll_y.set)

    preview_tree.pack(side="left", fill="both", expand=True, padx=(10, 6), pady=10)
    scroll_y.pack(side="right", fill="y", padx=(0, 10), pady=10)

    # Carpeta
    ctk.CTkLabel(left_content, text="Carpeta donde están TODOS los archivos", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
    carpeta_entry = ctk.CTkEntry(left_content)
    carpeta_entry.pack(fill="x", pady=(0, PADY_MD))
    btn_carpeta = ctk.CTkButton(left_content, text="Seleccionar carpeta", command=seleccionar_carpeta, font=FONT_BUTTON)
    btn_carpeta.pack(fill="x", pady=(0, PADY_MD))



    # --- Logos (arriba de Asunto) ---
    frame_logo = ctk.CTkFrame(right_content, fg_color="transparent")
    frame_logo.pack(fill="x", pady=(0, PADY_MD))

    frame_logo.grid_columnconfigure(0, weight=0)  # "Logo"
    frame_logo.grid_columnconfigure(1, weight=1)  # combobox
    frame_logo.grid_columnconfigure(2, weight=0)  # Seleccionar logo
    frame_logo.grid_columnconfigure(3, weight=0)  # Insertar logo

    ctk.CTkLabel(frame_logo, text="Logo", font=FONT_LABEL).grid(
        row=0, column=0, sticky="w", padx=(0, 12), pady=6
    )

    cb_logo = ttk.Combobox(frame_logo, state="readonly", style="Enabled.TCombobox")
    cb_logo.grid(row=0, column=1, sticky="ew", pady=6)

    btn_sel_logo = ctk.CTkButton(
        frame_logo,
        text="Actualizar lista",
        command=seleccionar_logo,
        font=FONT_BUTTON,
        width=150
    )
    btn_sel_logo.grid(row=0, column=2, sticky="e", padx=(12, 0), pady=6)

    btn_ins_logo = ctk.CTkButton(
        frame_logo,
        text="Insertar logo",
        command=insertar_logo,     # <-- IMPORTANTE: usar insertar_logo
        font=FONT_BUTTON,
        width=140
    )
    btn_ins_logo.grid(row=0, column=3, sticky="e", padx=(12, 0), pady=6)

    lbl_logo_estado = ctk.CTkLabel(right_content, text="Logo: (ninguno)")
    lbl_logo_estado.pack(anchor="w", pady=(0, PADY_MD))

    # aplicar selección al cambiar el combobox
    cb_logo.bind("<<ComboboxSelected>>", aplicar_logo_seleccionado)


    # Asunto
    ctk.CTkLabel(right_content, text="Asunto", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
    asunto_entry = ctk.CTkEntry(right_content)
    asunto_entry.pack(fill="x", pady=(0, PADY_MD))

    # Mensaje
    ctk.CTkLabel(right_content, text="Mensaje", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
    mensaje_text = ctk.CTkTextbox(right_content, height=200)
    mensaje_text.insert(
        "1.0",
        "Hola,\n\n"
        "Te envío el documento solicitado.\n"
        "Cualquier duda, quedo atento.\n\n"
        "Saludos,\n"
    )
    mensaje_text.pack(fill="both", expand=True, pady=(0, PADY_MD))

    # Controles: límite y delay
    frame_envio = ctk.CTkFrame(right_content, fg_color="transparent")
    frame_envio.pack(fill="x", pady=(0, PADY_MD))
    frame_envio.columnconfigure(1, weight=0)
    frame_envio.columnconfigure(3, weight=0)

    ctk.CTkLabel(frame_envio, text="Máximo por ejecución", font=FONT_LABEL).grid(row=0, column=0, sticky="w", pady=4)
    max_entry = ctk.CTkEntry(frame_envio, width=80)
    max_entry.insert(0, "500")
    max_entry.grid(row=0, column=1, padx=(10, 20), pady=4)

    ctk.CTkLabel(frame_envio, text="Delay (segundos) entre envíos", font=FONT_LABEL).grid(row=0, column=2, sticky="w", pady=4)
    delay_entry = ctk.CTkEntry(frame_envio, width=80)
    delay_entry.insert(0, "1.5")
    delay_entry.grid(row=0, column=3, padx=(10, 0), pady=4)

    # Sección Reporte
    ctk.CTkLabel(right_content, text="Informe / Registro", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))

    frame_checks = ctk.CTkFrame(right_content, fg_color="transparent")
    frame_checks.pack(fill="x", pady=(0, PADY_MD))

    gen_excel_var = tk.BooleanVar(master=root, value=True)
    gen_pdf_var = tk.BooleanVar(master=root, value=False)

    chk_excel = ctk.CTkCheckBox(frame_checks, text="Generar registro en Excel (.xlsx)", variable=gen_excel_var, command=_toggle_reporte_ui)
    chk_excel.pack(anchor="w", pady=(0, 6))

    chk_pdf = ctk.CTkCheckBox(frame_checks, text="Generar registro en PDF (.pdf)", variable=gen_pdf_var, command=_toggle_reporte_ui)
    chk_pdf.pack(anchor="w")

    ctk.CTkLabel(right_content, text="Ruta base del informe (sin extensión)", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
    reporte_entry = ctk.CTkEntry(right_content)
    reporte_entry.pack(fill="x", pady=(0, PADY_SM))

    btn_reporte = ctk.CTkButton(right_content, text="Elegir dónde guardar informe", command=seleccionar_ruta_reporte, font=FONT_BUTTON)
    btn_reporte.pack(fill="x", pady=(0, PADY_MD))

    _toggle_reporte_ui()


    # --- Título del reporte (antes de ENVIAR) ---
    ctk.CTkLabel(right_content, text="Título del reporte", font=FONT_LABEL).pack(anchor="w", pady=(0, PADY_SM))
    titulo_reporte_entry = ctk.CTkEntry(right_content)
    titulo_reporte_entry.insert(0, "Registro de envío de correos")  # default
    titulo_reporte_entry.pack(fill="x", pady=(0, PADY_MD))

    # Botones
    btn_enviar = ctk.CTkButton(
        right_content,
        text="ENVIAR CORREOS",
        command=abrir_modal_envio_y_ejecutar,
        height=40,
        font=FONT_BUTTON
    )
    btn_enviar.pack(fill="x", pady=(0, PADY_SM))


    # Refresco automático del preview al cambiar combobox:
    cb_nombre.bind("<<ComboboxSelected>>", actualizar_vista_previa)
    cb_correo.bind("<<ComboboxSelected>>", actualizar_vista_previa)
    cb_archivo.bind("<<ComboboxSelected>>", actualizar_vista_previa)
    cb_hoja.bind("<<ComboboxSelected>>", lambda e: cargar_columnas_excel())

    _set_excel_controls_enabled(False)
    cargar_logos_predeterminados()
    aplicar_logo_seleccionado()
