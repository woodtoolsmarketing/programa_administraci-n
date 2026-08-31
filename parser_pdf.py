# -*- coding: utf-8 -*-
"""
Lectura del reporte desde un PDF.

Dos caminos:
  - extraer_texto() / volcar_para_claude(): extraccion confiable para
    alimentar a la API de Claude, que es el camino recomendado para PDFs.
  - parsear_pdf(): intento deterministico por expresiones regulares. Si no
    reconoce el formato, devuelve el Reporte vacio con un aviso claro; NUNCA
    inventa datos.
"""

from __future__ import annotations

import os
import re
from datetime import date
from typing import Any

from config import log, log_error
from modelos import (
    Cliente,
    Comprobante,
    Reporte,
    telefonos_de_texto,
    sin_acentos,
)


class ErrorPDF(Exception):
    """Problema al leer el PDF, con mensaje ya legible."""


# ==========================================================================
# UTILIDADES
# ==========================================================================
def es_pdf(ruta: str) -> bool:
    """True si el archivo existe y empieza con la firma %PDF."""
    if not ruta or not os.path.isfile(ruta):
        return False
    if not ruta.lower().endswith(".pdf"):
        return False
    try:
        with open(ruta, "rb") as fh:
            return fh.read(5) == b"%PDF-"
    except OSError:
        return False


def contar_paginas(ruta: str) -> int:
    """Cantidad de paginas. 0 si no se pudo leer."""
    try:
        import pdfplumber

        with pdfplumber.open(ruta) as pdf:
            return len(pdf.pages)
    except Exception:
        pass
    try:
        from pypdf import PdfReader

        return len(PdfReader(ruta).pages)
    except Exception as e:
        log_error(f"No se pudo contar las paginas de {ruta}: {e}")
        return 0


def extraer_texto(ruta: str, max_paginas: int | None = None) -> str:
    """Texto del PDF, con cada pagina precedida por '### PAGINA n'.

    Usa pdfplumber con layout=True para conservar las columnas; si falla, cae a
    pypdf.
    """
    if not os.path.exists(ruta):
        raise ErrorPDF(f"No se encuentra el archivo: {ruta}")

    partes: list[str] = []
    try:
        import pdfplumber

        with pdfplumber.open(ruta) as pdf:
            total = len(pdf.pages)
            limite = total if max_paginas is None else min(total, max_paginas)
            for i in range(limite):
                try:
                    texto = pdf.pages[i].extract_text(layout=True) or ""
                except Exception:
                    try:
                        texto = pdf.pages[i].extract_text() or ""
                    except Exception as e:
                        log_error(f"Pagina {i + 1} ilegible: {e}")
                        texto = ""
                partes.append(f"### PAGINA {i + 1}\n{texto}")
        if any(p.strip().splitlines()[1:] for p in partes):
            return "\n".join(partes)
    except Exception as e:
        log_error(f"pdfplumber fallo con {ruta}: {e}")

    # Respaldo: pypdf
    try:
        from pypdf import PdfReader

        lector = PdfReader(ruta)
        total = len(lector.pages)
        limite = total if max_paginas is None else min(total, max_paginas)
        partes = []
        for i in range(limite):
            try:
                texto = lector.pages[i].extract_text() or ""
            except Exception:
                texto = ""
            partes.append(f"### PAGINA {i + 1}\n{texto}")
        return "\n".join(partes)
    except Exception as e:
        log_error(f"pypdf tambien fallo con {ruta}: {e}")
        raise ErrorPDF(
            "No se pudo leer el PDF. Puede estar danado o protegido con "
            "contrasena."
        )


def extraer_lineas(ruta: str) -> list[str]:
    """Lineas de texto del PDF, sin vacias y sin los marcadores de pagina."""
    texto = extraer_texto(ruta)
    salida: list[str] = []
    for linea in texto.splitlines():
        limpia = linea.rstrip()
        if not limpia.strip():
            continue
        if limpia.startswith("### PAGINA"):
            continue
        salida.append(limpia)
    return salida


def esta_escaneado(ruta: str) -> bool:
    """True si casi no hay texto extraible (PDF de imagen)."""
    try:
        texto = extraer_texto(ruta, max_paginas=3)
    except ErrorPDF:
        return True
    sin_marcas = re.sub(r"### PAGINA \d+", "", texto)
    return len(sin_marcas.strip()) < 120


def volcar_para_claude(ruta: str, max_paginas: int | None = None) -> str:
    """Texto del PDF con una cabecera de contexto para mandarselo a Claude."""
    paginas = contar_paginas(ruta)
    encabezado = [
        f"# ARCHIVO: {os.path.basename(ruta)}",
        f"# PAGINAS: {paginas}",
    ]
    if esta_escaneado(ruta):
        encabezado.append(
            "# AVISO: este PDF parece un escaneo (imagen). Si el texto de abajo "
            "sale vacio o incompleto, leelo directamente de la imagen del "
            "documento."
        )
    return "\n".join(encabezado) + "\n\n" + extraer_texto(ruta, max_paginas)


# ==========================================================================
# CONVERSION DE NUMEROS EN FORMATO ARGENTINO
# ==========================================================================
_RE_NUM_AR = re.compile(r"^-?\(?-?\$?\s*[\d\.]*\d(?:,\d+)?\)?-?$")


def _a_float_ar(texto: str) -> float | None:
    """'37.165,39' -> 37165.39 ; '(1.234,00)' -> -1234.0 ; '1.234-' -> -1234.0."""
    if texto is None:
        return None
    s = str(texto).strip().replace("$", "").replace(" ", "").replace("\xa0", "")
    if not s:
        return None
    if not _RE_NUM_AR.match(s):
        return None

    negativo = False
    if s.startswith("(") and s.endswith(")"):
        negativo = True
        s = s[1:-1]
    if s.endswith("-"):
        negativo = True
        s = s[:-1]
    if s.startswith("-"):
        negativo = True
        s = s[1:]
    s = s.strip("()")

    if "," in s:
        # coma decimal, punto de miles
        s = s.replace(".", "").replace(",", ".")
    elif s.count(".") == 1:
        entero, dec = s.split(".")
        # "37.165" es miles; "37.16" seria decimal, pero el ERP no lo usa asi
        if len(dec) == 3 and len(entero) <= 3:
            s = entero + dec
    else:
        s = s.replace(".", "")

    try:
        n = float(s)
    except ValueError:
        return None
    return -n if negativo else n


# ==========================================================================
# CLASIFICACION DE LINEAS
# ==========================================================================
_RE_FECHA = re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b")
_RE_FECHA_REPORTE = re.compile(r"Fecha\s*:\s*(\d{1,2}/\d{1,2}/\d{2,4})", re.IGNORECASE)
_RE_CAB_CLIENTE = re.compile(r"^\s*(\d{6,})\s+(\S.*)$")
_RE_TOTAL_CLIENTE = re.compile(
    r"total\s+cliente\s+(\d+)\s*(.*?)\s*$", re.IGNORECASE
)
_RE_TOTAL_GENERAL = re.compile(r"total\s+general\s*:?", re.IGNORECASE)
_RE_LOCALIDAD = re.compile(r"\(\s*(\d{3,5})\s*\)\s*(.+?)\s*$")
# La localidad en cualquier parte de la linea: sirve de ancla para separar la
# razon social de lo que viene despues, porque el ERP siempre la imprime como
# "(codigo postal) NOMBRE".
_RE_LOCALIDAD_INLINE = re.compile(r"\(\s*\d{3,5}\s*\)")
# Arranque de un tramo de telefonos: 7 caracteres o mas de digitos, espacios,
# puntos y guiones. El look-behind evita cortar adentro de una fecha.
_RE_TRAMO_TELEFONOS = re.compile(r"(?<![\d/])\d[\d\s.\-]{6,}")
# Nombre de contacto pegado justo antes del numero ("CAPITAL FEDERAL Marcos").
_RE_NOMBRE_PEGADO = re.compile(r"\s+([A-ZÁÉÍÓÚÑ][a-záéíóúñ]{2,})\s*$")
# Importes al final de la fila "Total cliente" ("... SRL  57.165,39 52.165,39").
_RE_COLA_IMPORTES = re.compile(r"(?:\s+-?\(?\$?[\d.]+,\d{2}\)?-?)+\s*$")
_RE_TIPO_COMPROBANTE = re.compile(
    r"\b(FE?\s?[ABCEM]|FAC?\s?[ABCEM]|N[CD]\s?[ABCEM]|REC|TIC?\s?[ABC]|"
    r"FCE\s?[ABC]|LIQ)\b",
    re.IGNORECASE,
)
_RE_NUMEROS_EN_LINEA = re.compile(r"-?\(?\$?\s*[\d\.]*\d(?:,\d+)?\)?-?")
_RE_VENDEDOR = re.compile(r"Vendedor\s*:\s*(\d+)\s*-\s*(\d+)", re.IGNORECASE)
_RE_COBRADOR = re.compile(r"Cobrador\s*:\s*(\d+)", re.IGNORECASE)

_RUIDO = (
    "GESTION COMERCIAL",
    "A COBRAR POR CLIENTE",
    "CLIENTE:",
    "CLIENTE COMPRADOR:",
    "COND.VTA",
    "EXCLUYE",
    "ORDENADO POR",
    "HOJA:",
    "FECHA:",
    "CONDICION DE PAGO",
    "COMPROBANTE",
    "IMPORTE",
    "ORIGINAL",
    "SALDO",
    "DIAS",
    "VENC.",
    "EMISION",
    "VTO.",
    "FC CREDITO",
    "N DE",
    "CUOTA",
    "CONTINUA EN LA HOJA",
    "PAGINA",
)


def _es_ruido(linea: str) -> bool:
    t = sin_acentos(linea).upper().strip()
    if not t:
        return True
    for pref in _RUIDO:
        if t.startswith(pref):
            return True
    # Fila de titulos: varias etiquetas en la misma linea
    etiquetas = sum(
        1
        for e in ("COMPROBANTE", "CONDICION DE PAGO", "SALDO", "IMPORTE", "DIAS", "VTO")
        if e in t
    )
    return etiquetas >= 2


def _fecha_de(texto: str) -> date | None:
    m = _RE_FECHA.search(texto or "")
    if not m:
        return None
    d, mes, a = (int(x) for x in m.groups())
    if a < 100:
        a += 2000
    try:
        return date(a, mes, d)
    except ValueError:
        return None


# ==========================================================================
# PARSER DETERMINISTICO
# ==========================================================================
def parsear_pdf(ruta: str, area_telefono_defecto: str = "11") -> Reporte:
    """Intenta leer el reporte del PDF por patrones de texto.

    Si no reconoce clientes, devuelve el Reporte vacio con un aviso pidiendo
    usar la interpretacion con Claude.
    """
    reporte = Reporte(origen_archivo=ruta, motor="pdf")

    try:
        lineas = extraer_lineas(ruta)
    except ErrorPDF as e:
        reporte.avisos.append(str(e))
        return reporte

    if not lineas:
        reporte.avisos.append(
            "El PDF no tiene texto extraible (parece un escaneo). Usa el boton "
            "'Interpretar con Claude', que lee la imagen del documento."
        )
        return reporte

    # --- cabecera del reporte ---
    for linea in lineas[:40]:
        if not reporte.empresa and not _es_ruido(linea):
            reporte.empresa = linea.strip()
        if not reporte.titulo and sin_acentos(linea).upper().strip().startswith("A COBRAR"):
            reporte.titulo = linea.strip()
        m = _RE_FECHA_REPORTE.search(linea)
        if m and reporte.fecha_reporte is None:
            reporte.fecha_reporte = _fecha_de(m.group(1))
        if "Vendedor" in linea or "Cobrador" in linea or "Cliente:" in linea:
            reporte.filtros = (reporte.filtros + " " + linea.strip()).strip()

    m = _RE_VENDEDOR.search(reporte.filtros)
    if m:
        reporte.filtro_vendedor_desde, reporte.filtro_vendedor_hasta = m.group(1), m.group(2)
    m = _RE_COBRADOR.search(reporte.filtros)
    if m:
        reporte.filtro_cobrador = m.group(1)

    # --- largo esperado del codigo de cliente ---
    # Las filas "Total cliente <codigo>" son inequivocas, asi que sirven de
    # referencia para saber cuantos digitos tiene un codigo de cliente en este
    # reporte. Sin esto, un telefono suelto como "45246075" se confunde con la
    # cabecera de un cliente y se inventa un cliente que no existe.
    largos_validos = _largos_de_codigo(lineas)

    # --- recorrido principal ---
    actual: Cliente | None = None
    info_pendiente: list[str] = []
    vio_comprobante = False

    def cerrar() -> None:
        nonlocal actual, info_pendiente, vio_comprobante
        if actual is not None:
            _completar_info(actual, info_pendiente, area_telefono_defecto)
            reporte.clientes.append(actual)
        actual = None
        info_pendiente = []
        vio_comprobante = False

    for linea in lineas:
        cruda = linea.strip()

        # total general
        if _RE_TOTAL_GENERAL.search(cruda):
            cerrar()
            nums = _numeros_de_linea(cruda)
            if len(nums) >= 2:
                reporte.total_general_original = nums[-2]
                reporte.total_general_saldo = nums[-1]
            elif len(nums) == 1:
                reporte.total_general_saldo = nums[-1]
            continue

        # total por cliente
        m_total = _RE_TOTAL_CLIENTE.search(sin_acentos(cruda))
        if m_total:
            nums = _numeros_de_linea(cruda)
            if actual is None:
                actual = Cliente(
                    codigo=m_total.group(1),
                    # sin los importes, que en el PDF quedan pegados al nombre
                    razon_social=_RE_COLA_IMPORTES.sub(
                        "", m_total.group(2) or ""
                    ).strip(),
                )
                reporte.avisos.append(
                    f"Cliente {m_total.group(1)}: se encontro el total sin haber "
                    "reconocido su cabecera."
                )
            if len(nums) >= 2:
                actual.total_original_reporte = nums[-2]
                actual.total_saldo_reporte = nums[-1]
            elif len(nums) == 1:
                actual.total_saldo_reporte = nums[-1]
            cerrar()
            continue

        if _es_ruido(cruda):
            continue

        # comprobante: tiene tipo, al menos una fecha y dos importes
        if actual is not None and _parece_comprobante(cruda):
            comp = _armar_comprobante(cruda, reporte.fecha_reporte)
            if comp is not None:
                actual.comprobantes.append(comp)
                vio_comprobante = True
                continue

        # cabecera de cliente
        m_cab = _RE_CAB_CLIENTE.match(cruda)
        if (
            m_cab
            and not _parece_comprobante(cruda)
            and len(m_cab.group(1)) in largos_validos
        ):
            codigo, resto = m_cab.group(1), m_cab.group(2).strip()
            cerrar()
            # En el PDF no hay columnas: la razon social, la localidad y los
            # telefonos vienen pegados en el mismo renglon y hay que separarlos
            # aca, o el nombre del cliente llega asi al WhatsApp del vendedor.
            razon, localidad, telefonos = _partir_cabecera_cliente(resto)
            actual = Cliente(
                codigo=codigo,
                razon_social=razon,
                localidad=localidad,
                telefonos_raw=telefonos,
            )
            info_pendiente = []
            vio_comprobante = False
            continue

        # lineas de informacion del cliente
        if actual is not None and not vio_comprobante:
            info_pendiente.append(cruda)

    cerrar()

    # --- limpieza de razon social con telefonos pegados ---
    for cli in reporte.clientes:
        _separar_telefonos_de_razon(cli, area_telefono_defecto)

    if not reporte.clientes:
        reporte.avisos.append(
            "No se reconocio ningun cliente en el PDF. El formato puede ser "
            "distinto al esperado: usa el boton 'Interpretar con Claude'."
        )
    else:
        log(
            f"Parser PDF: {len(reporte.clientes)} clientes en "
            f"{os.path.basename(ruta)}"
        )

    problemas = reporte.validar()
    for p in problemas[:10]:
        reporte.avisos.append(f"Revisar: {p}")
    if problemas:
        reporte.avisos.append(
            "Los totales no cierran del todo. Conviene verificar con Claude "
            "antes de mandar los mensajes."
        )

    return reporte


def _largos_de_codigo(lineas: list[str]) -> set[int]:
    """Cuantos digitos tiene un codigo de cliente en este reporte.

    Se deduce de las filas "Total cliente <codigo>", que no admiten confusion.
    Si no hay ninguna, se usa el largo mas frecuente entre las cabeceras
    candidatas. Como ultimo recurso se aceptan codigos de 9 digitos o mas, que
    es el formato del ERP (11 digitos con ceros a la izquierda).
    """
    from collections import Counter

    desde_totales: Counter[int] = Counter()
    for linea in lineas:
        m = _RE_TOTAL_CLIENTE.search(sin_acentos(linea))
        if m:
            desde_totales[len(m.group(1))] += 1
    if desde_totales:
        return {largo for largo, _ in desde_totales.most_common(2)}

    candidatos: Counter[int] = Counter()
    for linea in lineas:
        m = _RE_CAB_CLIENTE.match(linea.strip())
        if m and not _parece_comprobante(linea):
            candidatos[len(m.group(1))] += 1
    if candidatos:
        return {candidatos.most_common(1)[0][0]}

    return {9, 10, 11, 12}


def _numeros_de_linea(linea: str) -> list[float]:
    """Todos los numeros en formato argentino de la linea, en orden."""
    salida: list[float] = []
    for m in _RE_NUMEROS_EN_LINEA.finditer(linea):
        n = _a_float_ar(m.group(0))
        if n is not None:
            salida.append(n)
    return salida


def _parece_comprobante(linea: str) -> bool:
    """Tipo de comprobante + al menos una fecha + al menos dos numeros."""
    if not _RE_TIPO_COMPROBANTE.search(linea):
        return False
    if not _RE_FECHA.search(linea):
        return False
    return len(_numeros_de_linea(linea)) >= 2


def _armar_comprobante(linea: str, fecha_reporte: date | None) -> Comprobante | None:
    """Extrae un comprobante de una linea de texto del PDF."""
    fechas: list[date] = []
    for m in _RE_FECHA.finditer(linea):
        f = _fecha_de(m.group(0))
        if f:
            fechas.append(f)

    m_tipo = _RE_TIPO_COMPROBANTE.search(linea)
    tipo = m_tipo.group(0).strip().upper() if m_tipo else ""

    # Todo lo que esta antes del tipo suele ser el talonario
    talonario = ""
    if m_tipo:
        antes = linea[: m_tipo.start()].strip()
        m_tal = re.search(r"(\d{1,5})\s*$", antes)
        if m_tal:
            talonario = m_tal.group(1)

    # Numero de comprobante: primer bloque largo de digitos despues del tipo
    numero = ""
    if m_tipo:
        despues = linea[m_tipo.end():]
        m_num = re.search(r"\b(\d{5,})\b", despues)
        if m_num:
            numero = m_num.group(1)

    # Cuota tipo "01/01"
    cuota = ""
    m_cuota = re.search(r"\b(\d{2}/\d{2})\b(?!/)", linea)
    if m_cuota:
        cuota = m_cuota.group(1)

    # Importes y dias: los ultimos numeros de la linea.
    # El ERP los imprime como: importe original, saldo, dias.
    candidatos: list[float] = []
    for m in _RE_NUMEROS_EN_LINEA.finditer(linea):
        crudo = m.group(0).strip()
        # Descartamos los tramos que forman parte de una fecha
        inicio, fin = m.span()
        contexto = linea[max(0, inicio - 1): min(len(linea), fin + 1)]
        if "/" in contexto:
            continue
        n = _a_float_ar(crudo)
        if n is not None:
            candidatos.append(n)

    if len(candidatos) < 2:
        return None

    dias: int | None = None
    ultimo = candidatos[-1]
    # El campo dias es un entero chico; si el ultimo numero lo parece, lo usamos
    if abs(ultimo) < 100000 and float(ultimo).is_integer() and len(candidatos) >= 3:
        dias = int(ultimo)
        importes = candidatos[-3:-1]
    else:
        importes = candidatos[-2:]

    original, saldo = importes[0], importes[1]

    vto = fechas[1] if len(fechas) >= 2 else (fechas[0] if fechas else None)
    emision = fechas[0] if len(fechas) >= 2 else None

    if dias is None and vto and fecha_reporte:
        dias = (fecha_reporte - vto).days

    # Condicion de pago: texto entre el numero de comprobante y la primera fecha
    condicion = ""
    if numero and fechas:
        try:
            desde = linea.index(numero) + len(numero)
            m_f = _RE_FECHA.search(linea, desde)
            if m_f:
                condicion = linea[desde: m_f.start()].strip(" .-")
        except ValueError:
            condicion = ""

    return Comprobante(
        talonario=talonario,
        tipo=tipo,
        numero=numero,
        condicion_pago=condicion,
        fecha_emision=emision,
        fecha_vto=vto,
        cuota=cuota,
        importe_original=round(float(original), 2),
        saldo=round(float(saldo), 2),
        dias_venc=dias,
    )


def _completar_info(cliente: Cliente, lineas: list[str], area: str) -> None:
    """Reparte las lineas sueltas del bloque del cliente."""
    textos: list[str] = []
    for linea in lineas:
        m_loc = _RE_LOCALIDAD.search(linea)
        if m_loc and not cliente.localidad:
            cliente.localidad = f"({m_loc.group(1)}) {m_loc.group(2).strip()}"
            resto = linea[: m_loc.start()].strip()
            if resto:
                textos.append(resto)
            continue
        textos.append(linea.strip())

    # numeros de telefono que aparezcan en estas lineas
    crudos: list[str] = []
    for t in textos:
        if telefonos_de_texto(t, area):
            crudos.append(t)

    limpios = [t for t in textos if t not in crudos]
    if limpios:
        cliente.direccion = limpios[-1]
        if len(limpios) > 1:
            cliente.nombre_fantasia = " / ".join(limpios[:-1])

    if crudos:
        cliente.telefonos_raw = " / ".join(
            dict.fromkeys(([cliente.telefonos_raw] if cliente.telefonos_raw else []) + crudos)
        )
    cliente.telefonos = telefonos_de_texto(cliente.telefonos_raw, area)


def _partir_cabecera_cliente(resto: str) -> tuple[str, str, str]:
    """Separa la linea de cabecera del cliente en sus tres datos.

    En el PDF no hay columnas de verdad: la razon social, la localidad y los
    telefonos quedan en un solo renglon, por ejemplo

        ACME HERRAMIENTAS SRL (1047) CAPITAL FEDERAL Marcos 11 3900 9488

    La localidad es el ancla confiable, porque el ERP siempre la imprime como
    "(codigo postal) NOMBRE". Si no esta, se corta en el primer tramo que
    parezca un telefono.

    Devuelve (razon_social, localidad, telefonos_raw); lo que no se puede
    identificar vuelve como cadena vacia. Es idempotente: volver a pasarle su
    propia salida no cambia el resultado.
    """
    texto = (resto or "").strip()
    if not texto:
        return "", "", ""

    m_loc = _RE_LOCALIDAD_INLINE.search(texto)
    if m_loc:
        razon = texto[: m_loc.start()].strip()
        cola = texto[m_loc.start():].strip()
    else:
        razon, cola = texto, ""

    # El tramo de telefonos arranca en la primera secuencia larga de digitos,
    # pero solo se corta si de ahi sale al menos un numero valido: asi una
    # razon social con numeros ("FERRETERIA 2000 SA") no se rompe.
    fuente = cola if cola else razon
    telefonos = ""
    m_tel = _RE_TRAMO_TELEFONOS.search(fuente)
    if m_tel and telefonos_de_texto(fuente[m_tel.start():]):
        telefonos = fuente[m_tel.start():].strip()
        cabeza = fuente[: m_tel.start()]
        if cola:
            # Un nombre de contacto pegado al numero. Solo se busca despues de
            # la localidad, que el ERP imprime en mayusculas: dentro de la
            # razon social una palabra en Titulo bien puede ser un apellido.
            m_nombre = _RE_NOMBRE_PEGADO.search(cabeza)
            if m_nombre:
                telefonos = f"{m_nombre.group(1)} {telefonos}".strip()
                cabeza = cabeza[: m_nombre.start()]
            cola = cabeza
        else:
            razon = cabeza

    return razon.strip(), cola.strip(), telefonos.strip()


def _separar_telefonos_de_razon(cliente: Cliente, area: str) -> None:
    """Repaso final por si la razon social todavia arrastra localidad o telefonos.

    La cabecera ya se separa al leerla; esto cubre a los clientes que se
    reconstruyeron desde la fila "Total cliente" porque no se reconocio su
    cabecera.
    """
    razon, localidad, telefonos = _partir_cabecera_cliente(cliente.razon_social)
    if razon:
        cliente.razon_social = razon
    if localidad and not cliente.localidad:
        cliente.localidad = localidad
    if telefonos:
        cliente.telefonos_raw = " / ".join(
            x for x in (cliente.telefonos_raw, telefonos) if x
        ).strip(" /")
    cliente.telefonos = telefonos_de_texto(cliente.telefonos_raw, area)
