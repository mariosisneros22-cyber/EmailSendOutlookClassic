import os

from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    KeepInFrame,
)
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.colors import black, lightgrey
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Image as RLImage
from reportlab.lib.utils import ImageReader
from reportlab.lib.enums import TA_CENTER


def _logo_size_keep_height_pt(path: str, target_h_pt: float) -> tuple[float, float]:
    ir = ImageReader(path)
    w0, h0 = ir.getSize()
    if not w0 or not h0:
        return (target_h_pt * 3, target_h_pt)
    w_pt = target_h_pt * (w0 / h0)
    return (max(1.0, w_pt), target_h_pt)


def _generar_pdf_tabla(
    registros,
    ruta_pdf: str,
    titulo_pdf: str,
    logo_path: str | None = None,
):
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

    headers = ["N°", "Nombre", "Correo", "Fecha", "Hora", "Estado"]

    data = [headers]
    for idx, r in enumerate(registros, start=1):
        data.append(
            [
                P(idx),
                P(r.get("Nombre", "")),
                P(r.get("Correo", "")),
                P(r.get("Fecha", "")),
                P(r.get("Hora", "")),
                P(r.get("Estado", "")),
            ]
        )

    available_width = doc.width
    col_widths = [
        available_width * 0.08,
        available_width * 0.18,
        available_width * 0.28,
        available_width * 0.12,
        available_width * 0.10,
        available_width * 0.24,
    ]

    elementos = []

    LOGO_H_PX = 60
    LOGO_H_PT = LOGO_H_PX * 72 / 96

    header_h = LOGO_H_PT
    has_logo = bool(logo_path) and os.path.isfile(logo_path)

    logo_flowable = Spacer(1, header_h)
    w_pt = header_h * 2.5

    if has_logo:
        w_pt, h_pt = _logo_size_keep_height_pt(logo_path, LOGO_H_PT)
        logo_flowable = RLImage(logo_path, width=w_pt, height=h_pt)

    titulo = Paragraph(titulo_pdf.strip(), title_style)
    titulo_box = KeepInFrame(
        doc.width, header_h, [titulo], hAlign="CENTER", vAlign="MIDDLE"
    )

    logo_col_w = max(80.0, float(w_pt))
    header_tbl = Table(
        [[logo_flowable, titulo_box]],
        colWidths=[logo_col_w, doc.width - logo_col_w],
        rowHeights=[header_h],
    )
    header_tbl.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (0, 0), "LEFT"),
                ("ALIGN", (1, 0), (1, 0), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )

    elementos.append(header_tbl)
    elementos.append(Spacer(1, 10))

    if registros:
        tabla = Table(data, repeatRows=1, colWidths=col_widths)
        tabla.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.5, black),
                    ("BACKGROUND", (0, 0), (-1, 0), lightgrey),
                    ("FONT", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 9),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("ALIGN", (0, 1), (0, -1), "CENTER"),
                    ("ALIGN", (3, 1), (4, -1), "CENTER"),
                ]
            )
        )
        elementos.append(tabla)
    else:
        elementos.append(Paragraph("(sin registros)", styles["Normal"]))

    doc.build(elementos)


def generar_pdfs_registro(
    registros,
    ruta_pdf_enviados: str,
    ruta_pdf_errores: str,
    logo_path: str | None = None,
    report_title: str = "Registro de envío de correos"
):
    def estado_texto(r):
        return str(r.get("Estado", "")).strip().lower()

    enviados = [r for r in registros if estado_texto(r) == "enviado"]
    errores = [r for r in registros if estado_texto(r) != "enviado"]

    _generar_pdf_tabla(
        registros=enviados,
        ruta_pdf=ruta_pdf_enviados,
        titulo_pdf=f"{report_title} ({len(enviados)})",
        logo_path=logo_path,
    )

    _generar_pdf_tabla(
        registros=errores,
        ruta_pdf=ruta_pdf_errores,
        titulo_pdf=f"{report_title} ({len(errores)})",
        logo_path=logo_path,
    )