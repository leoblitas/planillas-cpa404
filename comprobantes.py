"""
comprobantes.py
===============
Impresión (PDF carta y Excel) de los asientos contables de la planilla:

* ``pdf_asiento_privado`` / ``excel_asiento_privado``: asiento para empresa privada.
* ``pdf_comprobante_publico`` / ``excel_comprobante_publico``: comprobante
  patrimonial para el sector público.

Los datos vienen de ``asientos.py``. ``meta`` es el mismo diccionario que usa
``documentos.py`` y, para el sector público, además ``meta["comprobante"]``.
"""

from __future__ import annotations

import io
from datetime import date
from decimal import Decimal
from typing import Any, Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (KeepTogether, Paragraph, SimpleDocTemplate, Spacer,
                                Table, TableStyle)

from documentos import (COLOR_HEADER, COLOR_TOTAL, _firmas_pdf, _fmt_fecha, _fmt_monto,
                        _pie)

NEGRO = colors.black
GRIS = colors.HexColor("#" + COLOR_TOTAL)

COMPROBANTE_POR_DEFECTO: dict[str, Any] = {
    "numero": "",
    "areas": ["Presupuesto", "Patrimonio"],
    "imputacion": "CON IMPUTACIÓN PRESUPUESTARIA",
    "operacion": None,            # Balance Inicial / Traspaso Caja-Banco / Ajuste Contable
    "momento_recursos": None,     # Estimación / Devengado / Percibido
    "momento_gastos": "DEVENGADO",  # Apropiación / Preventivo / Compromiso / Devengado / Pagado
    "beneficiario": "",
    "respaldo": "Planilla de sueldos y salarios",
    "apertura": "---/---/---", "ff": "", "of": "", "et": "",
}
AREAS = ["Presupuesto", "Patrimonio", "Tesorería"]
IMPUTACIONES = ["CON IMPUTACIÓN PRESUPUESTARIA", "SIN IMPUTACIÓN PRESUPUESTARIA"]
OPERACIONES = ["Balance Inicial", "Traspaso Caja-Banco", "Ajuste Contable"]
MOMENTOS_RECURSOS = ["ESTIMACIÓN", "DEVENGADO", "PERCIBIDO"]
MOMENTOS_GASTOS = ["APROPIACIÓN", "PREVENTIVO", "COMPROMISO", "DEVENGADO", "PAGADO"]


def _glosa(meta: dict) -> str:
    return meta.get("glosa") or (
        f"Registro de la planilla de sueldos y salarios correspondiente a {meta['periodo']}.")


def _monto(v: Decimal | None) -> str:
    return "" if v is None else _fmt_monto(v)


def _doc(buf, apaisado: bool) -> SimpleDocTemplate:
    tam = landscape(letter) if apaisado else letter
    m = 22 if apaisado else 36
    return SimpleDocTemplate(buf, pagesize=tam, leftMargin=m, rightMargin=m,
                             topMargin=m, bottomMargin=m)


# --------------------------------------------------------------------------- #
# Número a letras
# --------------------------------------------------------------------------- #
_UNIDADES = ["", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve",
             "diez", "once", "doce", "trece", "catorce", "quince", "dieciséis",
             "diecisiete", "dieciocho", "diecinueve", "veinte", "veintiuno", "veintidós",
             "veintitrés", "veinticuatro", "veinticinco", "veintiséis", "veintisiete",
             "veintiocho", "veintinueve"]
_DECENAS = ["", "", "", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta",
            "noventa"]
_CENTENAS = ["", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos",
             "seiscientos", "setecientos", "ochocientos", "novecientos"]


def _menor_mil(n: int) -> str:
    if n == 100:
        return "cien"
    c, resto = divmod(n, 100)
    partes = [_CENTENAS[c]] if c else []
    if resto < 30:
        if resto:
            partes.append(_UNIDADES[resto])
    else:
        d, u = divmod(resto, 10)
        partes.append(_DECENAS[d] + (f" y {_UNIDADES[u]}" if u else ""))
    return " ".join(partes)


def _apocopar(n: int) -> str:
    """Letras de n antes de 'mil' o 'millones' (uno -> un, veintiuno -> veintiún)."""
    t = _menor_mil(n)
    if t.endswith("veintiuno"):
        return t[:-9] + "veintiún"
    return t[:-3] + "un" if t.endswith("uno") else t


def numero_a_letras(monto: Decimal | float | int) -> str:
    """80743.11 -> 'Ochenta mil setecientos cuarenta y tres 11/100 bolivianos'."""
    total = Decimal(str(monto)).quantize(Decimal("0.01"))
    entero, centavos = int(total), int((total - int(total)) * 100)
    if entero == 0:
        letras = "cero"
    else:
        millones, resto = divmod(entero, 1_000_000)
        miles, unidades = divmod(resto, 1000)
        partes = []
        if millones:
            partes.append("un millón" if millones == 1 else f"{_apocopar(millones)} millones")
        if miles:
            partes.append("mil" if miles == 1 else f"{_apocopar(miles)} mil")
        if unidades:
            partes.append(_menor_mil(unidades))
        letras = " ".join(partes)
    return f"{letras.capitalize()} {centavos:02d}/100 bolivianos"


# --------------------------------------------------------------------------- #
# Empresa privada
# --------------------------------------------------------------------------- #
_SANGRIA = {0: 0, 1: 16, 2: 8, 3: 28}
AZUL = colors.HexColor("#" + COLOR_HEADER)
AZUL_CLARO = colors.HexColor("#E8EEF5")
LINEA = colors.HexColor("#D5DBE3")


def pdf_asiento_privado(asiento: dict, meta: dict) -> bytes:
    """Asiento contable de la planilla (empresa privada), carta vertical."""
    buf = io.BytesIO()
    doc = _doc(buf, apaisado=False)
    W = letter[0] - 72
    fecha = _fmt_fecha(meta.get("fecha_asiento") or meta["fecha_corte"])
    est = lambda n, **k: ParagraphStyle(n, fontName=k.pop("f", "Helvetica"), **k)

    # --- Cabecera: empresa a la izquierda, datos del asiento a la derecha ----
    izq = [Paragraph(meta["empresa"], est("emp", f="Helvetica-Bold", fontSize=15, leading=18,
                                          textColor=AZUL))]
    if meta.get("nit"):
        izq.append(Paragraph(f"NIT: {meta['nit']}", est("nit", fontSize=8.5, leading=11,
                                                       textColor=colors.HexColor("#555555"))))
    cab_der = Table(
        [[Paragraph("ASIENTO CONTABLE", est("ac", f="Helvetica-Bold", fontSize=9.5, leading=12,
                                            textColor=colors.white, alignment=1)), ""],
         [Paragraph("N°", est("l1", f="Helvetica-Bold", fontSize=8, leading=10)),
          Paragraph(str(meta.get("numero_asiento") or ""), est("v1", fontSize=9, leading=11))],
         [Paragraph("Fecha", est("l2", f="Helvetica-Bold", fontSize=8, leading=10)),
          Paragraph(fecha, est("v2", fontSize=9, leading=11))]],
        colWidths=[46, 134])
    cab_der.setStyle(TableStyle([
        ("SPAN", (0, 0), (1, 0)), ("BACKGROUND", (0, 0), (1, 0), AZUL),
        ("BOX", (0, 0), (-1, -1), 0.8, AZUL), ("LINEBELOW", (0, 1), (-1, 1), 0.4, LINEA),
        ("BACKGROUND", (0, 1), (0, -1), AZUL_CLARO),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    cabecera = Table([[izq, cab_der]], colWidths=[W - 180, 180])
    cabecera.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                  ("LEFTPADDING", (0, 0), (0, 0), 0),
                                  ("RIGHTPADDING", (1, 0), (1, 0), 0)]))

    # --- Título -------------------------------------------------------------------
    titulo = Table(
        [[Paragraph("REGISTRO DE LA PLANILLA DE SUELDOS Y SALARIOS",
                    est("tt", f="Helvetica-Bold", fontSize=11, leading=14, alignment=1))],
         [Paragraph(f"Periodo: {meta['periodo']}  &nbsp;·&nbsp;  Moneda: Bolivianos (Bs)",
                    est("ts", fontSize=8.5, leading=11, alignment=1,
                        textColor=colors.HexColor("#555555")))]], colWidths=[W])
    titulo.setStyle(TableStyle([("LINEABOVE", (0, 0), (-1, 0), 1.2, AZUL),
                                ("LINEBELOW", (0, -1), (-1, -1), 0.4, LINEA),
                                ("TOPPADDING", (0, 0), (-1, -1), 4),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))

    # --- Cuerpo del asiento ----------------------------------------------------------
    enc = est("enc", f="Helvetica-Bold", fontSize=8.5, leading=10, textColor=colors.white,
              alignment=1)
    datos: list[list[Any]] = [[Paragraph("CUENTA", enc), Paragraph("PARCIAL", enc),
                               Paragraph("DEBE", enc), Paragraph("HABER", enc)]]
    estilos: list[tuple] = []
    primer_haber = None
    for i, l in enumerate(asiento["lineas"], start=1):
        principal = l["debe"] is not None or l["haber"] is not None
        nombre = l["cuenta"]
        if l["haber"] is not None and l["nivel"] == 2:
            nombre = f'<font color="#777777"><i>a</i></font>&nbsp;&nbsp;{nombre}'
            if primer_haber is None:
                primer_haber = i
        sty = est(f"c{i}", f="Helvetica-Bold" if principal else "Helvetica", fontSize=9,
                  leading=11, leftIndent=_SANGRIA[l["nivel"]])
        datos.append([Paragraph(nombre, sty), _monto(l["parcial"]), _monto(l["debe"]),
                      _monto(l["haber"])])
        if l["debe"] is not None and l["nivel"] == 0:
            estilos.append(("BACKGROUND", (0, i), (-1, i), AZUL_CLARO))
        if principal:
            estilos.append(("FONT", (1, i), (-1, i), "Helvetica-Bold", 9))
    n = len(datos)
    datos.append([Paragraph("TOTAL", est("tot", f="Helvetica-Bold", fontSize=9.5, leading=12,
                                          alignment=2)),
                  "", _monto(asiento["total_debe"]), _monto(asiento["total_haber"])])
    if primer_haber:
        estilos.append(("LINEABOVE", (0, primer_haber), (-1, primer_haber), 0.9, AZUL))

    cuerpo = Table(datos, colWidths=[W - 285, 95, 95, 95], repeatRows=1)
    cuerpo.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("FONT", (1, 1), (-1, -1), "Helvetica", 9),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
        ("RIGHTPADDING", (1, 0), (-1, -1), 8),
        ("LINEBELOW", (0, 1), (-1, n - 1), 0.4, LINEA),
        ("LINEBEFORE", (1, 0), (-1, -1), 0.4, LINEA),
        ("BOX", (0, 0), (-1, -1), 0.9, AZUL),
        ("BACKGROUND", (0, n), (-1, n), AZUL_CLARO),
        ("FONT", (2, n), (-1, n), "Helvetica-Bold", 9.5),
        ("LINEABOVE", (0, n), (-1, n), 1.2, AZUL),
        ("SPAN", (0, n), (1, n)),
    ] + estilos))

    est_n = est("n", fontSize=9, leading=12)
    son = Table([[Paragraph(f"<b>Son:</b> {numero_a_letras(asiento['total_debe'])}", est_n)]],
                colWidths=[W])
    son.setStyle(TableStyle([("TOPPADDING", (0, 0), (-1, -1), 6),
                             ("LEFTPADDING", (0, 0), (-1, -1), 2)]))
    glosa = Table([[Paragraph("<b>GLOSA</b>", est("gl", f="Helvetica-Bold", fontSize=8,
                                                   leading=10, textColor=AZUL))],
                   [Paragraph(_glosa(meta), est_n)]], colWidths=[W])
    glosa.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, AZUL),
                               ("BACKGROUND", (0, 0), (-1, 0), AZUL_CLARO),
                               ("TOPPADDING", (0, 0), (-1, -1), 4),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))

    def firma(rol: str) -> Table:
        t = Table([[""], [Paragraph(f"<b>{rol}</b>", est("fr", fontSize=8, leading=10,
                                                         alignment=1))],
                   [Paragraph("Nombre y firma", est("fs", fontSize=7, leading=9, alignment=1,
                                                    textColor=colors.HexColor("#777777")))]],
                  colWidths=[(W - 40) / 3], rowHeights=[38, 11, 10])
        t.setStyle(TableStyle([("LINEABOVE", (0, 1), (-1, 1), 0.7, colors.black),
                               ("TOPPADDING", (0, 0), (-1, -1), 1),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
        return t

    firmas = Table([[firma("Elaborado por"), "", firma("Revisado por"), "",
                     firma("Aprobado por")]],
                   colWidths=[(W - 40) / 3, 20, (W - 40) / 3, 20, (W - 40) / 3])
    firmas.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0),
                                ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))

    historia: list = [cabecera, Spacer(1, 8), titulo, Spacer(1, 12), cuerpo, son,
                      Spacer(1, 8), glosa]
    if not asiento["cuadra"]:
        historia.append(Paragraph("<b>ATENCIÓN: el asiento no cuadra.</b>", est_n))
    historia += [Spacer(1, 18), KeepTogether(firmas)]
    doc.build(historia, onFirstPage=_pie, onLaterPages=_pie)
    return buf.getvalue()


def _hoja_encabezado(ws, meta: dict, titulo: str) -> None:
    ws["A1"] = meta["empresa"]
    ws["A1"].font = Font(bold=True, size=14, color=COLOR_HEADER)
    ws["A2"] = f"NIT: {meta['nit']}" if meta.get("nit") else ""
    ws["A3"] = titulo
    ws["A3"].font = Font(bold=True, size=12)
    ws["A4"] = (f"Periodo: {meta['periodo']}  |  Fecha: "
                f"{_fmt_fecha(meta.get('fecha_asiento') or meta['fecha_corte'])}")


def excel_asiento_privado(asiento: dict, meta: dict) -> bytes:
    """Mismo asiento en Excel, con el diseño del PDF."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Asiento"
    fecha = _fmt_fecha(meta.get("fecha_asiento") or meta["fecha_corte"])
    azul, claro = COLOR_HEADER, COLOR_TOTAL
    fino = Side(style="thin", color="D5DBE3")
    medio = Side(style="medium", color=azul)

    ws["A1"] = meta["empresa"]
    ws["A1"].font = Font(bold=True, size=15, color=azul)
    if meta.get("nit"):
        ws["A2"] = f"NIT: {meta['nit']}"
        ws["A2"].font = Font(size=9, color="555555")
    ws.merge_cells("C1:D1")
    ws["C1"] = "ASIENTO CONTABLE"
    ws["C1"].font = Font(bold=True, color="FFFFFF")
    ws["C1"].fill = PatternFill("solid", fgColor=azul)
    ws["C1"].alignment = Alignment(horizontal="center")
    ws["D1"].fill = PatternFill("solid", fgColor=azul)
    ws["C2"], ws["D2"] = "N°", str(meta.get("numero_asiento") or "")
    ws["C3"], ws["D3"] = "Fecha", fecha
    for r in (2, 3):
        ws.cell(r, 3).font = Font(bold=True)
        ws.cell(r, 3).fill = PatternFill("solid", fgColor=claro)
    ws.merge_cells("A5:D5")
    ws["A5"] = "REGISTRO DE LA PLANILLA DE SUELDOS Y SALARIOS"
    ws["A5"].font = Font(bold=True, size=12)
    ws["A5"].alignment = Alignment(horizontal="center")
    ws["A5"].border = Border(top=medio)
    ws.merge_cells("A6:D6")
    ws["A6"] = f"Periodo: {meta['periodo']}  ·  Moneda: Bolivianos (Bs)"
    ws["A6"].font = Font(size=9, color="555555")
    ws["A6"].alignment = Alignment(horizontal="center")

    fila_enc = 8
    for j, h in enumerate(["CUENTA", "PARCIAL", "DEBE", "HABER"], start=1):
        c = ws.cell(fila_enc, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=azul)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(left=fino, right=fino)
    r = fila_enc + 1
    primer_haber_visto = False
    for l in asiento["lineas"]:
        principal = l["debe"] is not None or l["haber"] is not None
        nombre = ("a   " + l["cuenta"]) if (l["haber"] is not None and l["nivel"] == 2) \
            else l["cuenta"]
        c = ws.cell(r, 1, nombre)
        c.font = Font(bold=principal)
        c.alignment = Alignment(horizontal="left", indent=_SANGRIA[l["nivel"]] // 5)
        for j, k in ((2, "parcial"), (3, "debe"), (4, "haber")):
            cc = ws.cell(r, j)
            if l[k] is not None:
                cc.value = float(l[k])
                cc.number_format = "#,##0.00"
            cc.font = Font(bold=principal and j > 1)
            cc.alignment = Alignment(horizontal="right")
        for j in range(1, 5):
            cell = ws.cell(r, j)
            top = medio if (l["haber"] is not None and l["nivel"] == 2
                            and not primer_haber_visto) else None
            cell.border = Border(bottom=fino, left=fino if j > 1 else None, right=fino if j > 1 else None,
                                 top=top)
            if l["debe"] is not None and l["nivel"] == 0:
                cell.fill = PatternFill("solid", fgColor=claro)
        if l["haber"] is not None and l["nivel"] == 2:
            primer_haber_visto = True
        r += 1
    ws.cell(r, 1, "TOTAL").alignment = Alignment(horizontal="right")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws.cell(r, 3, f"=SUM(C{fila_enc + 1}:C{r - 1})").number_format = "#,##0.00"
    ws.cell(r, 4, f"=SUM(D{fila_enc + 1}:D{r - 1})").number_format = "#,##0.00"
    for j in range(1, 5):
        c = ws.cell(r, j)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor=claro)
        c.border = Border(top=medio, bottom=medio)
    ws.cell(r + 1, 1, f"Son: {numero_a_letras(asiento['total_debe'])}").font = Font(bold=True)
    ws.merge_cells(start_row=r + 3, start_column=1, end_row=r + 3, end_column=4)
    ws.cell(r + 3, 1, f"Glosa: {_glosa(meta)}")
    ws.cell(r + 3, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r + 3].height = 32
    for k, rol in enumerate(["Elaborado por", "Revisado por", "Aprobado por"]):
        col = (1, 3, 4)[k]
        c = ws.cell(r + 7, col, rol)
        c.alignment = Alignment(horizontal="center")
        c.border = Border(top=Side(style="thin"))
        c.font = Font(bold=True, size=9)
    for col, w in zip("ABCD", (54, 16, 17, 17)):
        ws.column_dimensions[col].width = w
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.showGridLines = False
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# Sector público: comprobante patrimonial
# --------------------------------------------------------------------------- #
_P_PEQ = ParagraphStyle("pq", fontName="Helvetica", fontSize=6.5, leading=7.6)
_P_NEG = ParagraphStyle("pn", fontName="Helvetica-Bold", fontSize=7, leading=8.4)


def _opciones(opciones: Sequence[str], marcadas: Sequence[str], ancho_etq: float) -> Table:
    """Lista de opciones con casilla; la marcada lleva una X."""
    marc = {m.upper() for m in marcadas if m}
    filas = [[Paragraph(o, ParagraphStyle("o", parent=_P_PEQ, alignment=2)),
              "X" if o.upper() in marc else ""] for o in opciones]
    t = Table(filas, colWidths=[ancho_etq, 24], rowHeights=[11] * len(filas))
    t.setStyle(TableStyle([
        ("BOX", (1, 0), (1, -1), 0.8, NEGRO), ("INNERGRID", (1, 0), (1, -1), 0.5, NEGRO),
        ("FONT", (1, 0), (1, -1), "Helvetica-Bold", 7), ("ALIGN", (1, 0), (1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return t


def pdf_comprobante_publico(asiento: dict, meta: dict) -> bytes:
    """Comprobante patrimonial (sector público), carta apaisado."""
    cmp_ = dict(COMPROBANTE_POR_DEFECTO)
    cmp_.update(meta.get("comprobante") or {})
    buf = io.BytesIO()
    doc = _doc(buf, apaisado=True)
    W = landscape(letter)[0] - 44
    fecha = _fmt_fecha(meta.get("fecha_asiento") or meta["fecha_corte"])

    def caja(datos, **kw):
        t = Table(datos, **kw)
        return t

    # 1. Título
    titulo = Table([[Paragraph(f"<b>{meta['empresa'].upper()}</b>",
                               ParagraphStyle("x", fontName="Helvetica-Bold", fontSize=11,
                                              alignment=1))],
                    [Paragraph("<b>COMPROBANTE PATRIMONIAL</b>",
                               ParagraphStyle("y", fontName="Helvetica-Bold", fontSize=10,
                                              alignment=1))]], colWidths=[W])
    titulo.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, NEGRO)]))

    # 2. Tipo / imputación / fecha
    c1 = [Paragraph(f"<b>Comprobante N°:</b> {cmp_['numero']}", _P_PEQ),
          _opciones(AREAS, cmp_["areas"], 80)]
    c2 = [Spacer(1, 12), _opciones(IMPUTACIONES, [cmp_["imputacion"]], 170)]
    c3 = [Paragraph(f"<b>Fecha:</b> {fecha}", _P_PEQ),
          _opciones(OPERACIONES, [cmp_["operacion"] or ""], 100)]
    info = Table([[c1, c2, c3]], colWidths=[W * 0.3, W * 0.4, W * 0.3])
    info.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, NEGRO),
                              ("VALIGN", (0, 0), (-1, -1), "TOP")]))

    # 3. Momentos de registro
    rec = [Paragraph("<b>MOMENTOS DE REGISTRO DE RECURSOS</b>", _P_NEG),
           _opciones(MOMENTOS_RECURSOS, [cmp_["momento_recursos"] or ""], 90)]
    gas = [Paragraph("<b>MOMENTOS DE REGISTRO DE GASTOS</b>", _P_NEG),
           _opciones(MOMENTOS_GASTOS, [cmp_["momento_gastos"] or ""], 90)]
    momentos = Table([[rec, gas]], colWidths=[W / 2, W / 2])
    momentos.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, NEGRO),
                                  ("LINEAFTER", (0, 0), (0, 0), 0.8, NEGRO),
                                  ("VALIGN", (0, 0), (-1, -1), "TOP")]))

    # 4. Beneficiario y respaldo
    benef = Table([[Paragraph(f"<b>BENEFICIARIO:</b> {cmp_['beneficiario']}", _P_PEQ)],
                   [Paragraph(f"<b>DOCUMENTO DE RESPALDO:</b> {cmp_['respaldo']}", _P_PEQ)]],
                  colWidths=[W])
    benef.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, NEGRO)]))

    # 5. Detalle
    anchos = [72, 46, 22, 22, 22, 140, 52, 176, 62, 62, 62]
    anchos[5] += W - sum(anchos)
    enc = ["APERTURA\nPROGRAMÁTICA", "RUBRO /\nPARTIDA", "FF", "OF", "ET",
           "DESCRIPCIÓN PARTIDA PRESUPUESTARIA", "CÓDIGO\nCONTABLE",
           "DESCRIPCIÓN CUENTA CONTABLE", "PARCIAL", "DEBE", "HABER"]
    filas: list[list[Any]] = [[Paragraph(e.replace("\n", "<br/>"),
                                         ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=6.5,
                                                        leading=7.5, alignment=1)) for e in enc]]
    primera = True
    for l in asiento["lineas"]:
        con_ppto = bool(l["rubro"])
        filas.append([
            cmp_["apertura"] if (con_ppto and primera) else "",
            l["rubro"], cmp_["ff"] if con_ppto else "", cmp_["of"] if con_ppto else "",
            cmp_["et"] if con_ppto else "", l["partida"], l["codigo"], l["cuenta"],
            _monto(l["parcial"]), _monto(l["debe"]), _monto(l["haber"])])
        if con_ppto:
            primera = False
    n_det = len(filas) - 1
    filas.append(["", "", "", "", "", "", "", "TOTAL", "", _monto(asiento["total_debe"]),
                  _monto(asiento["total_haber"])])
    det = Table(filas, colWidths=anchos, repeatRows=1)
    det.setStyle(TableStyle([
        ("FONT", (0, 1), (-1, -1), "Helvetica", 6.5),
        ("BOX", (0, 0), (-1, -1), 0.8, NEGRO), ("GRID", (0, 0), (-1, 0), 0.8, NEGRO),
        ("LINEBEFORE", (1, 1), (-1, -2), 0.5, NEGRO),
        ("ALIGN", (0, 1), (1, -1), "CENTER"), ("ALIGN", (2, 1), (4, -1), "CENTER"),
        ("ALIGN", (6, 1), (6, -1), "CENTER"), ("ALIGN", (8, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 1), (-1, -1), 1.1), ("BOTTOMPADDING", (0, 1), (-1, -1), 1.1),
        ("BOX", (0, -1), (-1, -1), 0.8, NEGRO),
        ("FONT", (7, -1), (-1, -1), "Helvetica-Bold", 7), ("ALIGN", (7, -1), (7, -1), "RIGHT"),
        ("BOX", (9, -1), (9, -1), 0.8, NEGRO), ("BOX", (10, -1), (10, -1), 0.8, NEGRO),
    ]))
    glosa = Table([[Paragraph(f"<b>GLOSA:</b> {_glosa(meta)}", _P_PEQ)]], colWidths=[W],
                  rowHeights=[26])
    glosa.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, NEGRO), ("VALIGN", (0, 0), (-1, -1), "TOP")]))

    def firma(etq, cargo):
        t = Table([[Paragraph(f"<b>{etq}</b>", _P_PEQ)], [""], [Paragraph(f"<b>{cargo}</b>",
                  ParagraphStyle("c", fontName="Helvetica-Bold", fontSize=7, alignment=1))]],
                  colWidths=[W * 0.36], rowHeights=[14, 26, 16])
        t.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, NEGRO),
                               ("LINEABOVE", (0, 2), (-1, 2), 0.8, NEGRO)]))
        return t

    firmas = Table([[firma("ELABORADO POR:", "CONTADOR"), "",
                     firma("REVISADO POR:", "GERENTE FINANCIERO")]],
                   colWidths=[W * 0.36, W * 0.12, W * 0.36])
    firmas.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER")]))

    historia = [titulo, info, momentos, benef, Spacer(1, 6), det, glosa, Spacer(1, 8),
                KeepTogether(firmas)]
    if not asiento["cuadra"]:
        historia.append(Paragraph("<b>ATENCIÓN: el comprobante no cuadra.</b>", _P_PEQ))
    doc.build(historia, onFirstPage=_pie, onLaterPages=_pie)
    return buf.getvalue()


def excel_comprobante_publico(asiento: dict, meta: dict) -> bytes:
    """Comprobante patrimonial en Excel."""
    cmp_ = dict(COMPROBANTE_POR_DEFECTO)
    cmp_.update(meta.get("comprobante") or {})
    wb = Workbook()
    ws = wb.active
    ws.title = "Comprobante"
    _hoja_encabezado(ws, meta, "COMPROBANTE PATRIMONIAL")
    ws["A2"] = ""
    ws["A5"] = (f"Comprobante N°: {cmp_['numero']}  |  Áreas: {', '.join(cmp_['areas'])}  |  "
                f"{cmp_['imputacion']}")
    ws["A6"] = (f"Momento de registro de gastos: {cmp_['momento_gastos'] or '-'}  |  "
                f"Momento de registro de recursos: {cmp_['momento_recursos'] or '-'}  |  "
                f"Operación: {cmp_['operacion'] or '-'}")
    ws["A7"] = f"Beneficiario: {cmp_['beneficiario']}"
    ws["A8"] = f"Documento de respaldo: {cmp_['respaldo']}"
    enc = ["Apertura programática", "Rubro/Partida", "FF", "OF", "ET", "Partida presupuestaria",
           "Código contable", "Descripción cuenta contable", "PARCIAL", "DEBE", "HABER"]
    for j, h in enumerate(enc, start=1):
        c = ws.cell(10, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=COLOR_HEADER)
        c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    ws.row_dimensions[10].height = 30
    r, primera = 11, True
    for l in asiento["lineas"]:
        con = bool(l["rubro"])
        vals = [cmp_["apertura"] if (con and primera) else "", l["rubro"] or None,
                cmp_["ff"] if con else "", cmp_["of"] if con else "", cmp_["et"] if con else "",
                l["partida"], l["codigo"], l["cuenta"]]
        for j, v in enumerate(vals, start=1):
            ws.cell(r, j, v)
        for j, k in ((9, "parcial"), (10, "debe"), (11, "haber")):
            if l[k] is not None:
                ws.cell(r, j, float(l[k])).number_format = "#,##0.00"
        primera = primera and not con
        r += 1
    ws.cell(r, 8, "TOTAL").font = Font(bold=True)
    ws.cell(r, 10, f"=SUM(J11:J{r - 1})").number_format = "#,##0.00"
    ws.cell(r, 11, f"=SUM(K11:K{r - 1})").number_format = "#,##0.00"
    for j in range(8, 12):
        ws.cell(r, j).font = Font(bold=True)
        ws.cell(r, j).fill = PatternFill("solid", fgColor=COLOR_TOTAL)
    ws.cell(r + 2, 1, f"GLOSA: {_glosa(meta)}")
    ws.cell(r + 5, 2, "ELABORADO POR: CONTADOR")
    ws.cell(r + 5, 8, "REVISADO POR: GERENTE FINANCIERO")
    for col, w in zip("ABCDEFGHIJK", (16, 12, 6, 6, 6, 40, 12, 44, 14, 14, 14)):
        ws.column_dimensions[col].width = w
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.showGridLines = False
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
