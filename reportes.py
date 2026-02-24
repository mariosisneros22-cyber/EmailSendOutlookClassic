# reportes.py
try:
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.colors import black, lightgrey
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    PDF_DISPONIBLE = True
except Exception:
    PDF_DISPONIBLE = False


def generar_pdf_registro(registros, ruta_pdf):
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
    title_style = styles["Title"]
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

    headers = ["Nombre", "Correo", "NombreArchivo", "Fecha", "Hora", "Estado"]

    def build_table(rows):
        data = [headers]
        for r in rows:
            data.append([
                P(r.get("Nombre", "")),
                P(r.get("Correo", "")),
                P(r.get("NombreArchivo", "")),
                P(r.get("Fecha", "")),
                P(r.get("Hora", "")),
                P(r.get("Estado", "")),
            ])

        available_width = doc.width
        col_widths = [
            available_width * 0.14,  # Nombre
            available_width * 0.20,  # Correo
            available_width * 0.26,  # NombreArchivo
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
    elementos.append(Paragraph("Registro de envío de correos", title_style))
    elementos.append(Spacer(1, 10))

    elementos.append(Paragraph(f"Enviados ({len(enviados)})", section_style))
    elementos.append(Spacer(1, 6))
    if enviados:
        elementos.append(build_table(enviados))
    else:
        elementos.append(Paragraph("(sin registros)", styles["Normal"]))

    elementos.append(Spacer(1, 14))

    elementos.append(Paragraph(f"Errores ({len(errores)})", section_style))
    elementos.append(Spacer(1, 6))
    if errores:
        elementos.append(build_table(errores))
    else:
        elementos.append(Paragraph("(sin registros)", styles["Normal"]))

    doc.build(elementos)