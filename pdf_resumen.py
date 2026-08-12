# -*- coding: utf-8 -*-
"""
Genera el PDF del resumen de cuenta, filtrado por vendedor.

Existe por dos razones:
  1. Si el usuario carga un EXCEL no hay PDF para adjuntar al WhatsApp.
  2. Cada vendedor deberia recibir SOLO sus clientes, no la cartera completa
     de la empresa.

Las fuentes base de reportlab (Helvetica) no manejan bien algunos caracteres,
asi que todo el texto que viene de los datos pasa por modelos.sin_acentos().
"""

from __future__ import annotations

import os
import re
from datetime import date
from typing import Any, Iterable

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import (
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from config import CARPETA_SALIDA, hora_arg, log, log_error
from modelos import (
    Cliente,
    ETIQUETA_ESTADO,
    Reporte,
    VENCIDO,
    formatear_fecha,
    formatear_pesos,
    sin_acentos,
)

# --------------------------------------------------------------------------
# PALETA
# --------------------------------------------------------------------------
AZUL = colors.HexColor("#1F3864")
AZUL_CLARO = colors.HexColor("#D6E0F0")
ROJO = colors.HexColor("#9C0006")
ROJO_CLARO = colors.HexColor("#FDE3E1")
AMBAR = colors.HexColor("#7F6000")
AMBAR_CLARO = colors.HexColor("#FFF2CC")
GRIS = colors.HexColor("#808080")
GRIS_CLARO = colors.HexColor("#F2F2F2")


def _estilos() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "titulo": ParagraphStyle(
            "wtTitulo", parent=base["Title"], fontSize=18, leading=21,
            textColor=AZUL, spaceAfter=2, alignment=TA_LEFT,
        ),
        "subtitulo": ParagraphStyle(
            "wtSubtitulo", parent=base["Normal"], fontSize=11, leading=14,
            textColor=colors.black, spaceAfter=1,
        ),
        "meta": ParagraphStyle(
            "wtMeta", parent=base["Normal"], fontSize=9, leading=11, textColor=GRIS,
        ),
        "seccion": ParagraphStyle(
            "wtSeccion", parent=base["Heading2"], fontSize=12, leading=14,
            textColor=AZUL, spaceBefore=10, spaceAfter=4,
        ),
        "celda": ParagraphStyle(
            "wtCelda", parent=base["Normal"], fontSize=8, leading=9.5,
        ),
        "celda_neg": ParagraphStyle(
            "wtCeldaNeg", parent=base["Normal"], fontSize=8, leading=9.5,
            fontName="Helvetica-Bold",
        ),
        "celda_der": ParagraphStyle(
            "wtCeldaDer", parent=base["Normal"], fontSize=8, leading=9.5,
            alignment=TA_RIGHT,
        ),
        "cab": ParagraphStyle(
            "wtCab", parent=base["Normal"], fontSize=8, leading=9.5,
            fontName="Helvetica-Bold", textColor=colors.white,
        ),
        "nota": ParagraphStyle(
            "wtNota", parent=base["Normal"], fontSize=7.5, leading=9, textColor=GRIS,
        ),
    }


def _t(texto: Any) -> str:
    """Texto seguro para el PDF: sin acentos y con los & < > escapados."""
    s = sin_acentos(str(texto if texto is not None else ""))
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# --------------------------------------------------------------------------
# PIE DE PAGINA CON "Pagina X de Y"
# --------------------------------------------------------------------------
class _LienzoNumerado(rl_canvas.Canvas):
    """Canvas de dos pasadas para poder escribir el total de paginas."""

    def __init__(self, *args: Any, **kwargs: Any):
        self._pie_izq: str = kwargs.pop("pie_izquierda", "")
        super().__init__(*args, **kwargs)
        self._paginas: list[dict[str, Any]] = []

    def showPage(self) -> None:  # noqa: N802 (nombre impuesto por reportlab)
        self._paginas.append(dict(self.__dict__))
        self._startPage()

    def save(self) -> None:
        total = len(self._paginas)
        for estado in self._paginas:
            self.__dict__.update(estado)
            self._dibujar_pie(total)
            super().showPage()
        super().save()

    def _dibujar_pie(self, total: int) -> None:
        ancho, _alto = A4
        self.saveState()
        self.setFont("Helvetica", 7)
        self.setFillColor(GRIS)
        self.setStrokeColor(colors.HexColor("#CCCCCC"))
        self.line(1.5 * cm, 1.15 * cm, ancho - 1.5 * cm, 1.15 * cm)
        if self._pie_izq:
            self.drawString(1.5 * cm, 0.8 * cm, self._pie_izq)
        self.drawRightString(
            ancho - 1.5 * cm, 0.8 * cm, f"Pagina {self._pageNumber} de {total}"
        )
        self.restoreState()


# --------------------------------------------------------------------------
# NOMBRE DE ARCHIVO
# --------------------------------------------------------------------------
_INVALIDOS_WINDOWS = r'\/:*?"<>|'


def nombre_archivo_sugerido(nombre_vendedor: str, fecha: date | None = None) -> str:
    """'Resumen_de_cuenta_Juan_Perez_2026-08-10.pdf', sin caracteres invalidos."""
    base = sin_acentos(nombre_vendedor or "").strip()
    for ch in _INVALIDOS_WINDOWS:
        base = base.replace(ch, "")
    base = re.sub(r"\s+", "_", base)
    base = re.sub(r"[^A-Za-z0-9_\-\.]", "", base).strip("_.")

    if isinstance(fecha, date):
        sufijo = fecha.strftime("%Y-%m-%d")
    else:
        sufijo = hora_arg().strftime("%Y-%m-%d")

    partes = ["Resumen_de_cuenta"]
    if base:
        partes.append(base)
    partes.append(sufijo)
    return "_".join(partes) + ".pdf"


# --------------------------------------------------------------------------
# ORDENAMIENTO
# --------------------------------------------------------------------------
def _ordenar_vencidos(clientes: Iterable[Cliente]) -> list[Cliente]:
    return sorted(clientes, key=lambda c: (-c.dias_max_vencido, -c.saldo_vencido))


def _ordenar_a_vencer(clientes: Iterable[Cliente]) -> list[Cliente]:
    def clave(c: Cliente):
        d = c.dias_hasta_proximo_vto
        return (999999 if d is None else d, -c.saldo_a_vencer)

    return sorted(clientes, key=clave)


# --------------------------------------------------------------------------
# BLOQUES DEL DOCUMENTO
# --------------------------------------------------------------------------
def _encabezado(
    est: dict[str, ParagraphStyle],
    reporte: Reporte,
    nombre_vendedor: str,
    codigo_vendedor: str,
) -> list[Any]:
    empresa = reporte.empresa or "WOOD TOOLS SRL"
    bloques: list[Any] = [
        Paragraph(_t(empresa), est["titulo"]),
        Paragraph(
            "Resumen de cuenta - Clientes vencidos y a vencer", est["subtitulo"]
        ),
    ]
    if nombre_vendedor:
        etiqueta = nombre_vendedor
        if codigo_vendedor:
            etiqueta = f"{nombre_vendedor} ({codigo_vendedor})"
        bloques.append(Paragraph(f"Vendedor: {_t(etiqueta)}", est["subtitulo"]))
    if reporte.fecha_reporte:
        bloques.append(
            Paragraph(
                f"Reporte al {formatear_fecha(reporte.fecha_reporte)}", est["meta"]
            )
        )
    bloques.append(Spacer(1, 0.35 * cm))
    return bloques


def _recuadro_totales(
    est: dict[str, ParagraphStyle], clientes: list[Cliente]
) -> list[Any]:
    cantidad = len(clientes)
    vencido = round(sum(c.saldo_vencido for c in clientes), 2)
    a_vencer = round(sum(c.saldo_a_vencer for c in clientes), 2)
    total = round(sum(c.saldo_total for c in clientes), 2)
    cant_vencidos = sum(1 for c in clientes if c.estado == VENCIDO)

    datos = [
        [
            Paragraph("CLIENTES", est["cab"]),
            Paragraph("VENCIDO", est["cab"]),
            Paragraph("A VENCER", est["cab"]),
            Paragraph("TOTAL A COBRAR", est["cab"]),
        ],
        [
            Paragraph(
                f"{cantidad} ({cant_vencidos} vencidos)", est["celda_neg"]
            ),
            Paragraph(formatear_pesos(vencido), est["celda_neg"]),
            Paragraph(formatear_pesos(a_vencer), est["celda_neg"]),
            Paragraph(formatear_pesos(total), est["celda_neg"]),
        ],
    ]
    tabla = Table(datos, colWidths=[4.2 * cm, 4.4 * cm, 4.4 * cm, 4.9 * cm])
    tabla.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), AZUL),
                ("BACKGROUND", (0, 1), (-1, 1), AZUL_CLARO),
                ("TEXTCOLOR", (1, 1), (1, 1), ROJO),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.white),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return [tabla, Spacer(1, 0.4 * cm)]


_ANCHOS_CLIENTES = [
    1.7 * cm,   # codigo
    5.3 * cm,   # cliente
    1.0 * cm,   # cantidad de comprobantes
    2.6 * cm,   # vencido
    2.6 * cm,   # a vencer
    2.6 * cm,   # total
    2.2 * cm,   # estado
]


def _tabla_clientes(
    est: dict[str, ParagraphStyle],
    clientes: list[Cliente],
    es_vencidos: bool,
    incluir_telefonos: bool,
) -> Table:
    """Tabla de clientes con el desglose vencido / a vencer / total.

    El total de la fila SUBTOTAL suma el saldo total de los clientes listados,
    de manera que las dos tablas sumen exactamente el total a cobrar.
    """
    encabezado = [
        Paragraph("Codigo", est["cab"]),
        Paragraph("Cliente", est["cab"]),
        Paragraph("Cp.", est["cab"]),
        Paragraph("Vencido", est["cab"]),
        Paragraph("A vencer", est["cab"]),
        Paragraph("Saldo total", est["cab"]),
        Paragraph("Situacion", est["cab"]),
    ]
    filas: list[list[Any]] = [encabezado]

    for cli in clientes:
        nombre = _t(cli.nombre)
        if cli.nombre_fantasia:
            nombre += f"<br/><font size=7 color='#808080'>{_t(cli.nombre_fantasia)}</font>"
        if incluir_telefonos and cli.telefonos:
            tels = ", ".join(cli.telefonos[:2])
            nombre += f"<br/><font size=7 color='#808080'>Tel: {_t(tels)}</font>"
        elif incluir_telefonos and cli.telefonos_raw:
            nombre += (
                f"<br/><font size=7 color='#808080'>Tel: {_t(cli.telefonos_raw[:60])}</font>"
            )

        filas.append(
            [
                Paragraph(_t(cli.codigo_corto), est["celda"]),
                Paragraph(nombre, est["celda"]),
                Paragraph(str(cli.cantidad_comprobantes), est["celda"]),
                Paragraph(
                    formatear_pesos(cli.saldo_vencido) if cli.saldo_vencido else "-",
                    est["celda_der"],
                ),
                Paragraph(
                    formatear_pesos(cli.saldo_a_vencer) if cli.saldo_a_vencer else "-",
                    est["celda_der"],
                ),
                Paragraph(formatear_pesos(cli.saldo_total), est["celda_der"]),
                Paragraph(_t(cli.descripcion_tiempo()), est["celda"]),
            ]
        )

    # Fila de subtotal
    sub_venc = round(sum(c.saldo_vencido for c in clientes), 2)
    sub_avenc = round(sum(c.saldo_a_vencer for c in clientes), 2)
    sub_total = round(sum(c.saldo_total for c in clientes), 2)
    filas.append(
        [
            Paragraph("", est["celda"]),
            Paragraph("SUBTOTAL", est["celda_neg"]),
            Paragraph("", est["celda"]),
            Paragraph(formatear_pesos(sub_venc), est["celda_neg"]),
            Paragraph(formatear_pesos(sub_avenc), est["celda_neg"]),
            Paragraph(formatear_pesos(sub_total), est["celda_neg"]),
            Paragraph("", est["celda"]),
        ]
    )

    color_cab = ROJO if es_vencidos else AMBAR
    color_zebra = ROJO_CLARO if es_vencidos else AMBAR_CLARO
    ultima = len(filas) - 1

    estilo: list[Any] = [
        ("BACKGROUND", (0, 0), (-1, 0), color_cab),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#BFBFBF")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (2, 1), (2, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND", (0, ultima), (-1, ultima), GRIS_CLARO),
        ("TEXTCOLOR", (3, 1), (3, -1), ROJO),
    ]
    for i in range(1, ultima):
        if i % 2 == 0:
            estilo.append(("BACKGROUND", (0, i), (-1, i), color_zebra))

    tabla = Table(filas, colWidths=_ANCHOS_CLIENTES, repeatRows=1)
    tabla.setStyle(TableStyle(estilo))
    return tabla


_ANCHOS_COMP = [4.2 * cm, 2.2 * cm, 3.4 * cm, 3.4 * cm, 1.8 * cm, 2.9 * cm]


def _tabla_comprobantes(est: dict[str, ParagraphStyle], cli: Cliente) -> list[Any]:
    """Sub-tabla con el saldo parcial de cada comprobante del cliente."""
    if not cli.comprobantes:
        return []

    filas: list[list[Any]] = [
        [
            Paragraph("Comprobante", est["cab"]),
            Paragraph("F. Vto", est["cab"]),
            Paragraph("Imp. original", est["cab"]),
            Paragraph("Saldo", est["cab"]),
            Paragraph("Dias", est["cab"]),
            Paragraph("Estado", est["cab"]),
        ]
    ]
    ordenados = sorted(
        cli.comprobantes,
        key=lambda k: (-(k.dias_venc if k.dias_venc is not None else -99999)),
    )
    for comp in ordenados:
        dias = "-" if comp.dias_venc is None else str(comp.dias_venc)
        filas.append(
            [
                Paragraph(_t(comp.comprobante_str), est["celda"]),
                Paragraph(formatear_fecha(comp.fecha_vto), est["celda"]),
                Paragraph(formatear_pesos(comp.importe_original), est["celda_der"]),
                Paragraph(formatear_pesos(comp.saldo), est["celda_der"]),
                Paragraph(dias, est["celda_der"]),
                Paragraph(_t(ETIQUETA_ESTADO.get(comp.estado, comp.estado)), est["celda"]),
            ]
        )

    estilo = [
        ("BACKGROUND", (0, 0), (-1, 0), GRIS),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9D9D9")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]
    for i in range(1, len(filas)):
        if i % 2 == 0:
            estilo.append(("BACKGROUND", (0, i), (-1, i), GRIS_CLARO))

    tabla = Table(filas, colWidths=_ANCHOS_COMP, repeatRows=1)
    tabla.setStyle(TableStyle(estilo))

    titulo = Paragraph(
        f"<b>{_t(cli.nombre)}</b> (Cod. {_t(cli.codigo_corto)}) - "
        f"total {formatear_pesos(cli.saldo_total)}",
        est["celda"],
    )
    return [KeepTogether([Spacer(1, 0.22 * cm), titulo, Spacer(1, 0.08 * cm), tabla])]


# --------------------------------------------------------------------------
# GENERACION
# --------------------------------------------------------------------------
def generar_pdf(
    reporte: Reporte,
    clientes: list[Cliente],
    ruta_salida: str,
    nombre_vendedor: str = "",
    codigo_vendedor: str = "",
    incluir_comprobantes: bool = True,
    incluir_telefonos: bool = False,
) -> tuple[bool, str]:
    """Genera el PDF del resumen. Devuelve (ok, ruta) o (False, motivo)."""
    if not clientes:
        return False, "No hay clientes para incluir en el PDF."

    try:
        carpeta = os.path.dirname(os.path.abspath(ruta_salida))
        os.makedirs(carpeta, exist_ok=True)
    except Exception as e:
        log_error(f"No se pudo crear la carpeta del PDF: {e}")
        return False, f"No se pudo crear la carpeta de destino: {e}"

    est = _estilos()
    vencidos = _ordenar_vencidos([c for c in clientes if c.estado == VENCIDO])
    a_vencer = _ordenar_a_vencer([c for c in clientes if c.estado != VENCIDO])

    historia: list[Any] = []
    historia += _encabezado(est, reporte, nombre_vendedor, codigo_vendedor)
    historia += _recuadro_totales(est, clientes)

    if vencidos:
        historia.append(Paragraph("VENCIDOS", est["seccion"]))
        historia.append(_tabla_clientes(est, vencidos, True, incluir_telefonos))
    if a_vencer:
        historia.append(Paragraph("A VENCER", est["seccion"]))
        historia.append(_tabla_clientes(est, a_vencer, False, incluir_telefonos))

    if incluir_comprobantes:
        historia.append(PageBreak())
        historia.append(Paragraph("DETALLE DE COMPROBANTES", est["seccion"]))
        historia.append(
            Paragraph(
                "Saldo parcial de cada comprobante pendiente, por cliente.",
                est["nota"],
            )
        )
        for cli in vencidos + a_vencer:
            historia += _tabla_comprobantes(est, cli)

    historia.append(Spacer(1, 0.5 * cm))
    historia.append(
        Paragraph(
            "Documento generado automaticamente desde el reporte "
            "'A cobrar por cliente' de Gestion Comercial. Ante cualquier "
            "diferencia, vale el sistema de administracion.",
            est["nota"],
        )
    )

    pie = (
        f"{sin_acentos(reporte.empresa or 'Wood Tools SRL')} - generado el "
        f"{hora_arg():%d/%m/%Y %H:%M}"
    )

    def hacer_lienzo(*args: Any, **kwargs: Any) -> _LienzoNumerado:
        kwargs["pie_izquierda"] = pie
        return _LienzoNumerado(*args, **kwargs)

    try:
        doc = SimpleDocTemplate(
            ruta_salida,
            pagesize=A4,
            leftMargin=1.5 * cm,
            rightMargin=1.5 * cm,
            topMargin=1.5 * cm,
            bottomMargin=1.6 * cm,
            title=sin_acentos(f"Resumen de cuenta {nombre_vendedor}".strip()),
            author=sin_acentos(reporte.empresa or "Wood Tools SRL"),
            subject="Clientes vencidos y a vencer",
        )
        doc.build(historia, canvasmaker=hacer_lienzo)
    except PermissionError:
        log_error(f"PDF en uso: {ruta_salida}")
        return False, (
            f"No se pudo escribir el PDF porque el archivo esta abierto:\n"
            f"{os.path.basename(ruta_salida)}\n\nCerralo e intenta de nuevo."
        )
    except Exception as e:
        log_error(f"Error generando el PDF {ruta_salida}: {e}")
        return False, f"Error generando el PDF: {e}"

    log(f"PDF generado: {ruta_salida} ({len(clientes)} clientes)")
    return True, ruta_salida


def generar_pdf_en_carpeta_salida(
    reporte: Reporte,
    clientes: list[Cliente],
    nombre_vendedor: str = "",
    codigo_vendedor: str = "",
    incluir_comprobantes: bool = True,
    incluir_telefonos: bool = False,
) -> tuple[bool, str]:
    """Igual que generar_pdf pero eligiendo solo el nombre, en Resumenes/."""
    nombre = nombre_archivo_sugerido(nombre_vendedor, reporte.fecha_reporte)
    ruta = os.path.join(CARPETA_SALIDA, nombre)
    return generar_pdf(
        reporte,
        clientes,
        ruta,
        nombre_vendedor=nombre_vendedor,
        codigo_vendedor=codigo_vendedor,
        incluir_comprobantes=incluir_comprobantes,
        incluir_telefonos=incluir_telefonos,
    )
