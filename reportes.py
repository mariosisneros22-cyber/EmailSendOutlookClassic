# reportes.py
try:
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepInFrame, PageBreak
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.colors import black, lightgrey
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import Image as RLImage
    from reportlab.lib.utils import ImageReader
    from reportlab.lib.enums import TA_CENTER
    import os
    PDF_DISPONIBLE = True
except Exception:
    PDF_DISPONIBLE = False

def _logo_size_keep_height_pt(path: str, target_h_pt: float) -> tuple[float, float]:
    """
    Retorna (w_pt, h_pt) manteniendo proporción con altura fija target_h_pt.
    """
    ir = ImageReader(path)
    w0, h0 = ir.getSize()
    if not w0 or not h0:
        return (target_h_pt * 3, target_h_pt)
    w_pt = target_h_pt * (w0 / h0)
    return (max(1.0, w_pt), target_h_pt)


def generar_pdf_registro(registros, ruta_pdf, logo_path: str | None = None, report_title: str = "Registro de envío de correos"):
    if not PDF_DISPONIBLE:
        raise RuntimeError("No está instalado reportlab. Instala con: pip install reportlab")

    enviados = [r for r in registros if str(r.get("Estado", "")).strip().lower() == "enviado"]
    errores = [r for r in registros if str(r.get("Estado", "")).strip().lower() != "enviado"]

    doc = SimpleDocTemplate(
        ruta_pdf,
        pagesize=landscape(A4),
        leftMargin=24,
        rightMargin=24,
        topMargin=24,
        bottomMargin=24,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "title_center",
        parent=styles["Title"],
        alignment=TA_CENTER,
    )
    section_style = styles["Heading2"]

    cell_style = ParagraphStyle(
        "cell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        wordWrap="CJK",
    )

    def P(x):
        s = "" if x is None else str(x)
        s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return Paragraph(s, cell_style)

    headers = ["N°","Nombre", "Correo", "Fecha", "Hora", "Estado"]

    def build_table(rows):
        data = [headers]
        for idx, r in enumerate(rows, start=1):
            data.append([
                P(idx),
                P(r.get("Nombre", "")),
                P(r.get("Correo", "")),
                P(r.get("Fecha", "")),
                P(r.get("Hora", "")),
                P(r.get("Estado", "")),
            ])

        available_width = doc.width
        col_widths = [
            available_width * 0.06,  # N°
            available_width * 0.14,  # Nombre
            available_width * 0.20,  # Correo
            available_width * 0.10,  # Fecha
            available_width * 0.08,  # Hora
            available_width * 0.22,  # Estado
        ]

        t = Table(data, repeatRows=1, colWidths=col_widths)
        t.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, black),
            ("BACKGROUND", (0, 0), (-1, 0), lightgrey),
            ("FONT", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (3, 1), (4, -1), "CENTER"),
        ]))
        return t

    elementos = []
    # Tamaño único (mismo que mailer.py)
    LOGO_H_PX = 60
    LOGO_H_PT = LOGO_H_PX * 72 / 96  # 96dpi -> points

    # --- Header: logo izquierda + título centrado (misma fila) ---
    header_h = LOGO_H_PT  # altura del header según el logo
    logo_flowable = Spacer(1, header_h)  # placeholder si no hay logo

    if logo_path and os.path.isfile(logo_path):
        w_pt, h_pt = _logo_size_keep_height_pt(logo_path, LOGO_H_PT)
        logo_flowable = RLImage(logo_path, width=w_pt, height=h_pt)

    titulo_txt = (report_title or "").strip() or "Registro de envío de correos"
    titulo = Paragraph(titulo_txt, title_style)

    # Mantener el título dentro del ancho disponible y centrado
    # (evita que se “desplace” raro si el título es largo)
    titulo_box = KeepInFrame(doc.width, header_h, [titulo], hAlign="CENTER", vAlign="MIDDLE")

    # Anchos: col logo = ancho real del logo (o un mínimo), col título = resto
    logo_col_w = (w_pt if (logo_path and os.path.isfile(logo_path)) else header_h * 2.5)
    logo_col_w = max(80, float(logo_col_w))  # mínimo razonable

    header_tbl = Table(
        [[logo_flowable, titulo_box]],
        colWidths=[logo_col_w, doc.width - logo_col_w],
        rowHeights=[header_h],
    )
    header_tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, 0), "LEFT"),    # logo
        ("ALIGN", (1, 0), (1, 0), "CENTER"),  # título
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))

    elementos.append(header_tbl)
    elementos.append(Spacer(1, 10))

    elementos.append(Paragraph(f"Enviados ({len(enviados)})", section_style))
    elementos.append(Spacer(1, 4))
    if enviados:
        elementos.append(build_table(enviados))
    else:
        elementos.append(Paragraph("(sin registros)", styles["Normal"]))

    elementos.append(PageBreak()) #romper pagina

    elementos.append(Paragraph(f"Errores ({len(errores)})", section_style))
    elementos.append(Spacer(1, 4))
    if errores:
        elementos.append(build_table(errores))
    else:
        elementos.append(Paragraph("(sin registros)", styles["Normal"]))

    doc.build(elementos)