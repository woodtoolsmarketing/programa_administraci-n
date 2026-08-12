# -*- coding: utf-8 -*-
"""
Armado del texto del resumen de cuenta que se manda por WhatsApp.

Usa el markdown de WhatsApp: *negrita*, _cursiva_.
Respeta el limite de 4096 caracteres partiendo por cliente, nunca al medio.
"""

from __future__ import annotations

from typing import Iterable

from config import CONFIG
from modelos import (
    A_VENCER,
    Cliente,
    Reporte,
    VENCE_HOY,
    VENCIDO,
    formatear_fecha,
    formatear_pesos,
)

SEPARADOR = "─" * 24   # linea fina


# ==========================================================================
# ORDENAMIENTO
# ==========================================================================
def _ordenar_vencidos(clientes: Iterable[Cliente]) -> list[Cliente]:
    """Del mas vencido al menos vencido; a igual atraso, mayor deuda primero."""
    return sorted(
        clientes, key=lambda c: (-c.dias_max_vencido, -c.saldo_vencido, c.nombre)
    )


def _ordenar_a_vencer(clientes: Iterable[Cliente]) -> list[Cliente]:
    """Por vencimiento mas proximo; los sin fecha al final."""
    def clave(c: Cliente):
        d = c.dias_hasta_proximo_vto
        return (999999 if d is None else d, -c.saldo_a_vencer, c.nombre)

    return sorted(clientes, key=clave)


# ==========================================================================
# PIEZAS DEL MENSAJE
# ==========================================================================
def resumen_titulo(reporte: Reporte, nombre_vendedor: str = "") -> str:
    """Cabecera del mensaje: titulo, vendedor y fecha del reporte."""
    lineas = ["*RESUMEN DE CUENTA - WOOD TOOLS*"]
    if nombre_vendedor:
        lineas.append(f"Vendedor: {nombre_vendedor}")
    if reporte.fecha_reporte:
        lineas.append(f"Reporte al {formatear_fecha(reporte.fecha_reporte)}")
    return "\n".join(lineas)


def linea_cliente(cliente: Cliente, indice: int | None = None) -> str:
    """Una linea (o dos) por cliente: nombre, codigo, deuda y tiempo.

    El monto principal es SIEMPRE el saldo total del cliente, para que los
    subtotales de cada seccion sumen exactamente el total a cobrar. Cuando el
    cliente tiene plata vencida Y plata por vencer, se aclara el desglose.
    """
    prefijo = f"{indice}) " if indice is not None else "- "
    detalle = (
        f"   {formatear_pesos(cliente.saldo_total)} - {cliente.descripcion_tiempo()}"
    )

    extras: list[str] = []
    if cliente.saldo_vencido > 0 and cliente.saldo_a_vencer > 0:
        extras.append(
            f"vencido {formatear_pesos(cliente.saldo_vencido)} "
            f"+ a vencer {formatear_pesos(cliente.saldo_a_vencer)}"
        )
    if cliente.cantidad_comprobantes > 1:
        extras.append(f"{cliente.cantidad_comprobantes} comprobantes")
    if cliente.estado == VENCIDO and cliente.vto_mas_antiguo:
        extras.append(f"vto. mas viejo {formatear_fecha(cliente.vto_mas_antiguo)}")
    elif cliente.estado != VENCIDO and cliente.proximo_vto:
        extras.append(f"vto. {formatear_fecha(cliente.proximo_vto)}")

    if extras:
        detalle += f" ({', '.join(extras)})"

    return f"{prefijo}*{cliente.nombre}* (Cod. {cliente.codigo_corto})\n{detalle}"


def _bloque_comprobantes(cliente: Cliente) -> str:
    """Detalle comprobante por comprobante (saldo parcial), indentado."""
    filas: list[str] = []
    for comp in cliente.comprobantes:
        filas.append(
            f"      • {comp.comprobante_str} | vto {formatear_fecha(comp.fecha_vto)} "
            f"| {formatear_pesos(comp.saldo)} | {comp.descripcion_tiempo()}"
        )
    return "\n".join(filas)


def _seccion(
    titulo: str,
    clientes: list[Cliente],
    total: float,
    incluir_comprobantes: bool,
) -> list[str]:
    """Devuelve la lista de bloques de una seccion (uno por cliente)."""
    if not clientes:
        return []
    cantidad = len(clientes)
    palabra = "cliente" if cantidad == 1 else "clientes"
    bloques = [f"*{titulo}* ({cantidad} {palabra} - {formatear_pesos(total)})"]
    for i, cli in enumerate(clientes, start=1):
        texto = linea_cliente(cli, i)
        if incluir_comprobantes and cli.cantidad_comprobantes > 0:
            texto += "\n" + _bloque_comprobantes(cli)
        bloques.append(texto)
    return bloques


# ==========================================================================
# MENSAJE COMPLETO
# ==========================================================================
def construir_mensaje(
    reporte: Reporte,
    clientes: list[Cliente],
    nombre_vendedor: str = "",
    codigo_vendedor: str = "",
    incluir_comprobantes: bool = False,
) -> str:
    """Arma el texto completo del resumen para un vendedor.

    `clientes` es el subconjunto que le toca a ese vendedor (ya filtrado).
    """
    etiqueta = nombre_vendedor
    if nombre_vendedor and codigo_vendedor:
        etiqueta = f"{nombre_vendedor} ({codigo_vendedor})"

    if not clientes:
        cabecera = resumen_titulo(reporte, etiqueta)
        return (
            f"{cabecera}\n\n"
            "No hay clientes con saldo pendiente para informar.\n\n"
            f"_{CONFIG.get('firma_mensaje', '')}_"
        )

    vencidos = _ordenar_vencidos([c for c in clientes if c.estado == VENCIDO])
    a_vencer = _ordenar_a_vencer([c for c in clientes if c.estado != VENCIDO])

    # Subtotal de cada seccion = saldo total de los clientes QUE SE LISTAN ahi,
    # asi los dos subtotales suman exactamente el TOTAL A COBRAR.
    subtotal_vencidos = round(sum(c.saldo_total for c in vencidos), 2)
    subtotal_a_vencer = round(sum(c.saldo_total for c in a_vencer), 2)
    # Desglose real del dinero, que puede cruzar secciones (un cliente vencido
    # tambien puede tener comprobantes por vencer).
    total_vencido = round(sum(c.saldo_vencido for c in clientes), 2)
    total_a_vencer = round(sum(c.saldo_a_vencer for c in clientes), 2)
    total = round(sum(c.saldo_total for c in clientes), 2)

    partes: list[str] = [resumen_titulo(reporte, etiqueta)]

    bloques_vencidos = _seccion(
        "CLIENTES CON DEUDA VENCIDA", vencidos, subtotal_vencidos, incluir_comprobantes
    )
    if bloques_vencidos:
        partes.append(SEPARADOR)
        partes.extend(bloques_vencidos)

    bloques_a_vencer = _seccion(
        "CLIENTES A VENCER", a_vencer, subtotal_a_vencer, incluir_comprobantes
    )
    if bloques_a_vencer:
        partes.append(SEPARADOR)
        partes.extend(bloques_a_vencer)

    partes.append(SEPARADOR)
    resumen_final = [f"*TOTAL A COBRAR: {formatear_pesos(total)}*"]
    if total_vencido and total_a_vencer:
        resumen_final.append(
            f"Del total: {formatear_pesos(total_vencido)} ya vencido y "
            f"{formatear_pesos(total_a_vencer)} por vencer."
        )
    partes.append("\n".join(resumen_final))

    firma = str(CONFIG.get("firma_mensaje", "") or "").strip()
    if firma:
        partes.append(f"_{firma}_")

    return "\n\n".join(partes)


# ==========================================================================
# PARTIDO EN VARIOS MENSAJES
# ==========================================================================
def partir_mensaje(texto: str, max_caracteres: int | None = None) -> list[str]:
    """Parte el texto en mensajes de WhatsApp sin cortar un cliente al medio.

    Numera las partes como "(1/3)" cuando hay mas de una.
    """
    if max_caracteres is None:
        max_caracteres = int(CONFIG.get("max_caracteres_mensaje", 3800))
    max_caracteres = max(500, int(max_caracteres))

    texto = (texto or "").strip()
    if len(texto) <= max_caracteres:
        return [texto] if texto else []

    # Reservamos lugar para el sufijo " (99/99)"
    limite = max_caracteres - 12
    bloques = texto.split("\n\n")

    partes: list[str] = []
    actual = ""
    for bloque in bloques:
        # Un bloque solo mas grande que el limite: hay que trocearlo por lineas
        if len(bloque) > limite:
            if actual:
                partes.append(actual)
                actual = ""
            for trozo in _trocear_bloque(bloque, limite):
                partes.append(trozo)
            continue

        candidato = f"{actual}\n\n{bloque}" if actual else bloque
        if len(candidato) <= limite:
            actual = candidato
        else:
            if actual:
                partes.append(actual)
            actual = bloque
    if actual:
        partes.append(actual)

    if len(partes) <= 1:
        return partes

    total = len(partes)
    return [f"{p}\n\n({i}/{total})" for i, p in enumerate(partes, start=1)]


def _trocear_bloque(bloque: str, limite: int) -> list[str]:
    """Ultimo recurso: parte un bloque muy largo por lineas."""
    salida: list[str] = []
    actual = ""
    for linea in bloque.split("\n"):
        if len(linea) > limite:
            if actual:
                salida.append(actual)
                actual = ""
            for i in range(0, len(linea), limite):
                salida.append(linea[i : i + limite])
            continue
        candidato = f"{actual}\n{linea}" if actual else linea
        if len(candidato) <= limite:
            actual = candidato
        else:
            salida.append(actual)
            actual = linea
    if actual:
        salida.append(actual)
    return salida


# ==========================================================================
# VARIANTE PARA PLANTILLA DE META
# ==========================================================================
def construir_mensaje_plantilla(
    reporte: Reporte, clientes: list[Cliente], nombre_vendedor: str = ""
) -> str:
    """Version de una sola linea para un parametro de plantilla de Meta.

    Meta RECHAZA los parametros de plantilla que contengan saltos de linea o
    tabulaciones, asi que aca todo va en una linea separado por ' | ' y
    truncado a 1000 caracteres.
    """
    items: list[str] = []
    vencidos = _ordenar_vencidos([c for c in clientes if c.estado == VENCIDO])
    a_vencer = _ordenar_a_vencer([c for c in clientes if c.estado != VENCIDO])

    for cli in vencidos + a_vencer:
        items.append(
            f"{cli.nombre} ({cli.codigo_corto}) {formatear_pesos(cli.saldo_total)} "
            f"{cli.descripcion_tiempo()}"
        )

    total = round(sum(c.saldo_total for c in clientes), 2)
    cuerpo = " | ".join(items)
    cierre = f" | TOTAL {formatear_pesos(total)}"

    disponible = 1000 - len(cierre) - 4
    if len(cuerpo) > disponible:
        cuerpo = cuerpo[:disponible].rstrip() + "..."

    texto = cuerpo + cierre
    # Blindaje: ni saltos de linea ni tabs, aunque vengan en un nombre
    return texto.replace("\n", " ").replace("\r", " ").replace("\t", " ")[:1000]


# ==========================================================================
# TEXTO CORTO PARA LA VISTA PREVIA / EPIGRAFE DEL PDF
# ==========================================================================
def epigrafe_pdf(reporte: Reporte, clientes: list[Cliente], nombre_vendedor: str = "") -> str:
    """Texto breve que acompana al PDF adjunto (maximo ~900 caracteres)."""
    cantidad = len(clientes)
    vencidos = sum(1 for c in clientes if c.estado == VENCIDO)
    total = round(sum(c.saldo_total for c in clientes), 2)
    partes = ["Resumen de cuenta - Wood Tools"]
    if nombre_vendedor:
        partes.append(f"Vendedor: {nombre_vendedor}")
    if reporte.fecha_reporte:
        partes.append(f"Al {formatear_fecha(reporte.fecha_reporte)}")
    partes.append(
        f"{cantidad} cliente{'s' if cantidad != 1 else ''} "
        f"({vencidos} vencido{'s' if vencidos != 1 else ''}) - "
        f"Total {formatear_pesos(total)}"
    )
    return " | ".join(partes)[:900]
