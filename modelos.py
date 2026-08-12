# -*- coding: utf-8 -*-
"""
Modelos de datos del resumen de cuenta (reporte "A cobrar por cliente").

Este modulo es el CONTRATO comun entre los parsers (Excel / PDF / Claude),
la generacion de mensajes, el PDF por vendedor y la interfaz.
No importa nada del proyecto: es la base de la que dependen los demas modulos.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field, asdict
from datetime import date
from typing import Any

# ==========================================================================
# ESTADOS DE UN COMPROBANTE
# ==========================================================================
VENCIDO = "VENCIDO"
VENCE_HOY = "VENCE_HOY"
A_VENCER = "A_VENCER"

ETIQUETA_ESTADO = {
    VENCIDO: "Vencido",
    VENCE_HOY: "Vence hoy",
    A_VENCER: "A vencer",
}


def clasificar_dias(dias: int | None) -> str:
    """dias > 0 => VENCIDO ; dias == 0 => VENCE_HOY ; dias < 0 => A_VENCER.

    La columna "Dias Venc." del reporte ya viene calculada como
    (fecha del reporte - fecha de vencimiento): positiva = atrasado.
    """
    if dias is None:
        return A_VENCER
    if dias > 0:
        return VENCIDO
    if dias == 0:
        return VENCE_HOY
    return A_VENCER


# ==========================================================================
# COMPROBANTE (una factura / cuota) -> "saldo parcial"
# ==========================================================================
@dataclass
class Comprobante:
    talonario: str = ""              # col A  (ej "71")
    tipo: str = ""                   # col B  (ej "FE A")
    numero: str = ""                 # col C  (ej "00072065")
    condicion_pago: str = ""         # col D  (ej "30 dias.ff.CONTRA ENTREGA")
    fecha_emision: date | None = None   # col E
    fecha_vto: date | None = None        # col F
    cuota: str = ""                  # col G  (ej "01/01")
    importe_original: float = 0.0    # col H
    saldo: float = 0.0               # col I  <- SALDO PARCIAL
    dias_venc: int | None = None     # col J  (positivo = vencido)
    fc_credito: str = ""             # col K
    resaltado: bool = False          # alguna celda de la fila venia en amarillo
    fila_origen: int | None = None   # fila del Excel (1-based), para auditoria

    @property
    def estado(self) -> str:
        return clasificar_dias(self.dias_venc)

    @property
    def esta_vencido(self) -> bool:
        return self.estado == VENCIDO

    @property
    def comprobante_str(self) -> str:
        """Identificacion legible: 'FE A 71-00072065'."""
        partes = [p for p in (self.tipo, f"{self.talonario}-{self.numero}") if p.strip(" -")]
        return " ".join(partes).strip()

    def descripcion_tiempo(self) -> str:
        """'vencido hace 10 dias' / 'vence hoy' / 'faltan 5 dias'."""
        d = self.dias_venc
        if d is None:
            return "sin fecha de vencimiento"
        if d > 0:
            return f"vencido hace {d} {'dia' if d == 1 else 'dias'}"
        if d == 0:
            return "vence hoy"
        falta = -d
        return f"faltan {falta} {'dia' if falta == 1 else 'dias'}"


# ==========================================================================
# CLIENTE
# ==========================================================================
@dataclass
class Cliente:
    codigo: str = ""                 # col A tal cual: "00000002761"
    razon_social: str = ""           # col C
    nombre_fantasia: str = ""
    direccion: str = ""
    localidad: str = ""              # ej "(1047) CAPITAL FEDERAL"
    contacto: str = ""               # ej "Jonathan"
    telefonos_raw: str = ""          # col G, texto crudo con nombres y numeros
    telefonos: list[str] = field(default_factory=list)   # normalizados E.164 sin '+'
    comprobantes: list[Comprobante] = field(default_factory=list)
    total_original_reporte: float | None = None   # col H de "Total cliente"
    total_saldo_reporte: float | None = None      # col I de "Total cliente"  <- SALDO TOTAL
    resaltado: bool = False          # el cliente venia marcado en amarillo
    codigo_vendedor: str = ""        # si el reporte lo trae por cliente
    fila_origen: int | None = None

    # ---------------- identificacion ----------------
    @property
    def codigo_corto(self) -> str:
        """'00000002761' -> '2761' (lo que el vendedor reconoce)."""
        c = (self.codigo or "").strip()
        limpio = c.lstrip("0")
        return limpio or c or ""

    @property
    def nombre(self) -> str:
        return (self.razon_social or self.nombre_fantasia or "").strip()

    # ---------------- saldos ----------------
    @property
    def saldo_total(self) -> float:
        """Suma de saldos parciales. Es la fuente de verdad calculada."""
        return round(sum(c.saldo for c in self.comprobantes), 2)

    @property
    def saldo_vencido(self) -> float:
        return round(sum(c.saldo for c in self.comprobantes if c.estado == VENCIDO), 2)

    @property
    def saldo_a_vencer(self) -> float:
        return round(
            sum(c.saldo for c in self.comprobantes if c.estado in (VENCE_HOY, A_VENCER)), 2
        )

    @property
    def cantidad_comprobantes(self) -> int:
        return len(self.comprobantes)

    # ---------------- estado ----------------
    @property
    def estado(self) -> str:
        """VENCIDO si tiene al menos un comprobante vencido."""
        if any(c.estado == VENCIDO for c in self.comprobantes):
            return VENCIDO
        if any(c.estado == VENCE_HOY for c in self.comprobantes):
            return VENCE_HOY
        return A_VENCER

    @property
    def dias_max_vencido(self) -> int:
        vencidos = [c.dias_venc for c in self.comprobantes if c.estado == VENCIDO and c.dias_venc]
        return max(vencidos) if vencidos else 0

    @property
    def dias_hasta_proximo_vto(self) -> int | None:
        """Dias que faltan para el vencimiento mas cercano (solo no vencidos)."""
        pend = [
            -c.dias_venc
            for c in self.comprobantes
            if c.dias_venc is not None and c.dias_venc <= 0
        ]
        return min(pend) if pend else None

    @property
    def proximo_vto(self) -> date | None:
        fechas = [
            c.fecha_vto
            for c in self.comprobantes
            if c.fecha_vto and c.dias_venc is not None and c.dias_venc <= 0
        ]
        return min(fechas) if fechas else None

    @property
    def vto_mas_antiguo(self) -> date | None:
        fechas = [c.fecha_vto for c in self.comprobantes if c.fecha_vto and c.esta_vencido]
        return min(fechas) if fechas else None

    # ---------------- consistencia ----------------
    def descuadre(self) -> float:
        """Diferencia entre el total del reporte y la suma de comprobantes."""
        if self.total_saldo_reporte is None:
            return 0.0
        return round(self.saldo_total - self.total_saldo_reporte, 2)

    def esta_consistente(self, tolerancia: float = 0.05) -> bool:
        return abs(self.descuadre()) <= tolerancia

    def descripcion_tiempo(self) -> str:
        """Texto de tiempo a nivel cliente, para el mensaje de WhatsApp."""
        if self.estado == VENCIDO:
            d = self.dias_max_vencido
            return f"vencido hace {d} {'dia' if d == 1 else 'dias'}"
        if self.estado == VENCE_HOY:
            return "vence hoy"
        d = self.dias_hasta_proximo_vto
        if d is None:
            return "sin vencimiento informado"
        return f"faltan {d} {'dia' if d == 1 else 'dias'} para el vencimiento"


# ==========================================================================
# REPORTE COMPLETO
# ==========================================================================
@dataclass
class Reporte:
    empresa: str = ""
    titulo: str = ""
    fecha_reporte: date | None = None
    filtros: str = ""
    filtro_vendedor_desde: str = ""
    filtro_vendedor_hasta: str = ""
    filtro_cobrador: str = ""
    clientes: list[Cliente] = field(default_factory=list)
    total_general_original: float | None = None
    total_general_saldo: float | None = None
    origen_archivo: str = ""
    motor: str = ""                   # "excel", "pdf", "claude-pdf", "claude-texto"
    avisos: list[str] = field(default_factory=list)

    # ---------------- totales ----------------
    @property
    def saldo_total(self) -> float:
        return round(sum(c.saldo_total for c in self.clientes), 2)

    @property
    def saldo_vencido(self) -> float:
        return round(sum(c.saldo_vencido for c in self.clientes), 2)

    @property
    def saldo_a_vencer(self) -> float:
        return round(sum(c.saldo_a_vencer for c in self.clientes), 2)

    @property
    def clientes_vencidos(self) -> list[Cliente]:
        return [c for c in self.clientes if c.estado == VENCIDO]

    @property
    def clientes_a_vencer(self) -> list[Cliente]:
        return [c for c in self.clientes if c.estado != VENCIDO]

    @property
    def clientes_resaltados(self) -> list[Cliente]:
        return [c for c in self.clientes if c.resaltado]

    def cliente_por_codigo(self, codigo: str) -> Cliente | None:
        codigo = (codigo or "").strip()
        for c in self.clientes:
            if c.codigo == codigo or c.codigo_corto == codigo.lstrip("0"):
                return c
        return None

    # ---------------- consistencia ----------------
    def descuadre_general(self) -> float:
        if self.total_general_saldo is None:
            return 0.0
        return round(self.saldo_total - self.total_general_saldo, 2)

    def validar(self, tolerancia: float = 0.05) -> list[str]:
        """Chequeos de integridad. Devuelve lista de problemas legibles."""
        problemas: list[str] = []
        if not self.clientes:
            problemas.append("No se detectaron clientes en el archivo.")
        for c in self.clientes:
            if not c.esta_consistente(tolerancia):
                problemas.append(
                    f"{c.nombre} ({c.codigo_corto}): la suma de comprobantes "
                    f"({formatear_pesos(c.saldo_total)}) no coincide con el total del "
                    f"reporte ({formatear_pesos(c.total_saldo_reporte or 0)})."
                )
            if not c.comprobantes:
                problemas.append(f"{c.nombre} ({c.codigo_corto}): sin comprobantes.")
        if self.total_general_saldo is not None and abs(self.descuadre_general()) > tolerancia:
            problemas.append(
                f"Total general: calculado {formatear_pesos(self.saldo_total)} vs "
                f"reporte {formatear_pesos(self.total_general_saldo)} "
                f"(diferencia {formatear_pesos(self.descuadre_general())})."
            )
        return problemas

    def resumen_corto(self) -> str:
        return (
            f"{len(self.clientes)} clientes | "
            f"vencido {formatear_pesos(self.saldo_vencido)} | "
            f"a vencer {formatear_pesos(self.saldo_a_vencer)} | "
            f"total {formatear_pesos(self.saldo_total)}"
        )

    def a_dict(self) -> dict[str, Any]:
        return asdict(self)


# ==========================================================================
# UTILIDADES DE FORMATO Y NORMALIZACION
# ==========================================================================
def formatear_pesos(valor: float | None) -> str:
    """1234567.5 -> '$1.234.567,50' (formato argentino)."""
    if valor is None:
        return "$0,00"
    negativo = valor < 0
    txt = f"{abs(float(valor)):,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    return ("-$" if negativo else "$") + txt


def formatear_fecha(f: date | None) -> str:
    return f.strftime("%d/%m/%Y") if f else "-"


def sin_acentos(txt: str) -> str:
    """Quita tildes/dieresis. Util para comparar y para PDF con fuentes basicas."""
    if not txt:
        return ""
    norm = unicodedata.normalize("NFKD", str(txt))
    return "".join(ch for ch in norm if not unicodedata.combining(ch))


# --- telefonos argentinos -------------------------------------------------
_RE_NO_DIGITO = re.compile(r"\D+")


def extraer_telefonos_crudos(texto: str) -> list[str]:
    """Saca secuencias de numeros de un texto libre tipo
    'Marcos 11 3900 9488 -4313 0466 // 4312 0'.

    Agrupa digitos separados por espacios/guiones cuando en conjunto forman un
    telefono plausible (8 a 13 digitos).
    """
    if not texto:
        return []
    candidatos: list[str] = []
    # Trozos separados por / , ; o palabras: nos quedamos con los tramos numericos
    for tramo in re.split(r"[/,;|]+", str(texto)):
        # secuencias de digitos, espacios, guiones y puntos
        for m in re.finditer(r"[\d][\d\s\.\-]{6,}", tramo):
            solo = _RE_NO_DIGITO.sub("", m.group(0))
            if 8 <= len(solo) <= 13:
                candidatos.append(solo)
            elif len(solo) > 13:
                # dos numeros pegados ("11 3900 9488 -4313 0466" -> 18 digitos):
                # cortamos bloques de 10 y conservamos el ultimo tramo si sirve
                resto = solo
                while len(resto) > 13:
                    candidatos.append(resto[:10])
                    resto = resto[10:]
                if 8 <= len(resto) <= 13:
                    candidatos.append(resto)
    # sin duplicados, en orden
    vistos: set[str] = set()
    salida: list[str] = []
    for c in candidatos:
        if c not in vistos:
            vistos.add(c)
            salida.append(c)
    return salida


def normalizar_telefono_ar(numero: str, area_defecto: str = "11") -> str | None:
    """Normaliza a formato WhatsApp AR: 549 + area + abonado (sin '+').

    Devuelve None si no parece un numero movil valido.
    Reglas aplicadas:
      - quita 0 de area y 15 de abonado
      - si ya empieza con 54, respeta e inserta el 9 si falta
      - CABA: 11 + 8 digitos
    """
    if not numero:
        return None
    d = _RE_NO_DIGITO.sub("", str(numero))
    if not d:
        return None

    # Ya viene internacional
    if d.startswith("54"):
        resto = d[2:]
        if resto.startswith("9"):
            resto = resto[1:]
        d = resto
    if d.startswith("00"):
        d = d[2:]

    # 0 inicial de area (0341 -> 341)
    if d.startswith("0"):
        d = d[1:]

    # 15 despues del area (11 15 4064 7300 -> 11 4064 7300)
    m = re.match(r"^(11|\d{3}|\d{4})15(\d{6,8})$", d)
    if m:
        d = m.group(1) + m.group(2)

    # 15 sin area, formato viejo ("15-3359-5605") -> area por defecto
    # Ningun codigo de area argentino empieza en 15, asi que no hay ambiguedad.
    if d.startswith("15") and len(d) == 10:
        d = area_defecto + d[2:]

    # Sin area: 8 digitos -> asumimos area por defecto
    if len(d) == 8:
        d = area_defecto + d
    # 7 digitos: interior sin area, no podemos adivinar con confianza
    if len(d) < 10 or len(d) > 11:
        return None

    return "549" + d


def es_probable_movil_ar(numero_normalizado: str) -> bool:
    """Heuristica: en CABA (11) los fijos historicos arrancan en 4 o 5.

    No es exacta (hay moviles 11-4xxx). Se usa solo para ordenar candidatos,
    nunca para descartar: el usuario siempre puede elegir el numero a mano.
    """
    if not numero_normalizado or not numero_normalizado.startswith("549"):
        return False
    nac = numero_normalizado[3:]
    if nac.startswith("11") and len(nac) == 10:
        return nac[2] not in ("4", "5")
    return True


def telefonos_de_texto(texto: str, area_defecto: str = "11") -> list[str]:
    """Pipeline completo: texto libre -> lista de numeros WhatsApp normalizados."""
    normalizados: list[str] = []
    for crudo in extraer_telefonos_crudos(texto):
        n = normalizar_telefono_ar(crudo, area_defecto)
        if n and n not in normalizados:
            normalizados.append(n)
    # Los probables moviles primero
    normalizados.sort(key=lambda n: (not es_probable_movil_ar(n),))
    return normalizados
