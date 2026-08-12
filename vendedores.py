# -*- coding: utf-8 -*-
"""
Alta/baja de vendedores y asignacion de clientes a vendedores.

Persistencia en JSON al lado del ejecutable (vendedores.json / asignaciones.json),
con escritura atomica para que un corte de luz no deje el archivo a medio grabar.
"""

from __future__ import annotations

import json
import os
import uuid
from dataclasses import dataclass, asdict, field
from typing import Any, Iterable

from config import (
    ARCHIVO_ASIGNACIONES,
    ARCHIVO_VENDEDORES,
    CONFIG,
    log,
    log_error,
)
from modelos import normalizar_telefono_ar


# ==========================================================================
# ESCRITURA ATOMICA
# ==========================================================================
def _guardar_json(ruta: str, datos: Any) -> bool:
    tmp = ruta + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(datos, f, indent=2, ensure_ascii=False)
        os.replace(tmp, ruta)
        return True
    except Exception as e:
        log_error(f"No se pudo guardar {os.path.basename(ruta)}: {e}")
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        return False


def _leer_json(ruta: str, defecto: Any) -> Any:
    if not os.path.exists(ruta):
        return defecto
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        log_error(f"No se pudo leer {os.path.basename(ruta)}: {e}")
        return defecto


# ==========================================================================
# VENDEDOR
# ==========================================================================
@dataclass
class Vendedor:
    id: str = ""
    codigo: str = ""          # codigo del ERP, ej "002"
    nombre: str = ""
    telefono: str = ""        # formato WhatsApp: 549XXXXXXXXXX, sin '+'
    activo: bool = True
    notas: str = ""

    def __post_init__(self) -> None:
        if not self.id:
            self.id = uuid.uuid4().hex[:12]

    def telefono_valido(self) -> bool:
        t = "".join(ch for ch in str(self.telefono) if ch.isdigit())
        return 10 <= len(t) <= 15

    def etiqueta(self) -> str:
        """'Juan Perez (002)' o solo el nombre si no tiene codigo."""
        nombre = (self.nombre or "sin nombre").strip()
        cod = (self.codigo or "").strip()
        return f"{nombre} ({cod})" if cod else nombre

    def descripcion(self) -> str:
        """Etiqueta + telefono, para listas y combos."""
        estado = "" if self.activo else "  [inactivo]"
        return f"{self.etiqueta()} - {self.telefono}{estado}"

    def a_dict(self) -> dict[str, Any]:
        return asdict(self)

    @staticmethod
    def desde_dict(d: dict[str, Any]) -> "Vendedor":
        return Vendedor(
            id=str(d.get("id", "") or ""),
            codigo=str(d.get("codigo", "") or ""),
            nombre=str(d.get("nombre", "") or ""),
            telefono=str(d.get("telefono", "") or ""),
            activo=bool(d.get("activo", True)),
            notas=str(d.get("notas", "") or ""),
        )


# ==========================================================================
# CRUD DE VENDEDORES
# ==========================================================================
def cargar_vendedores() -> list[Vendedor]:
    """Lee vendedores.json. Devuelve lista vacia si no existe o esta roto."""
    crudo = _leer_json(ARCHIVO_VENDEDORES, [])
    if not isinstance(crudo, list):
        log_error("vendedores.json no contiene una lista; se ignora.")
        return []
    salida: list[Vendedor] = []
    for item in crudo:
        if isinstance(item, dict):
            try:
                salida.append(Vendedor.desde_dict(item))
            except Exception as e:
                log_error(f"Vendedor invalido en el archivo: {e}")
    return salida


def guardar_vendedores(lista: Iterable[Vendedor]) -> bool:
    """Graba la lista completa de vendedores."""
    return _guardar_json(ARCHIVO_VENDEDORES, [v.a_dict() for v in lista])


def agregar_vendedor(
    nombre: str, telefono: str, codigo: str = "", notas: str = ""
) -> tuple[Vendedor | None, str]:
    """Crea un vendedor normalizando el telefono.

    Devuelve (vendedor, "") si salio bien, o (None, motivo) si no.
    """
    nombre = (nombre or "").strip()
    if not nombre:
        return None, "Falta el nombre del vendedor."

    area = str(CONFIG.get("area_telefono_defecto", "11"))
    normalizado = normalizar_telefono_ar(telefono, area)
    if not normalizado:
        return None, (
            f"El telefono '{telefono}' no parece un celular argentino valido. "
            "Escribilo con codigo de area, por ejemplo 11 3900-9488."
        )

    lista = cargar_vendedores()
    for v in lista:
        if v.telefono == normalizado:
            return None, f"Ese numero ya esta cargado para {v.etiqueta()}."
        if codigo and v.codigo and v.codigo.strip() == codigo.strip():
            return None, f"El codigo {codigo} ya lo usa {v.etiqueta()}."

    nuevo = Vendedor(
        codigo=(codigo or "").strip(),
        nombre=nombre,
        telefono=normalizado,
        notas=(notas or "").strip(),
    )
    lista.append(nuevo)
    if not guardar_vendedores(lista):
        return None, "No se pudo grabar el archivo de vendedores."
    log(f"Vendedor agregado: {nuevo.etiqueta()} {nuevo.telefono}")
    return nuevo, ""


def actualizar_vendedor(id_vendedor: str, **campos: Any) -> tuple[bool, str]:
    """Modifica campos de un vendedor existente. (ok, motivo_si_falla)."""
    lista = cargar_vendedores()
    objetivo = None
    for v in lista:
        if v.id == id_vendedor:
            objetivo = v
            break
    if objetivo is None:
        return False, "No se encontro el vendedor."

    if "telefono" in campos:
        area = str(CONFIG.get("area_telefono_defecto", "11"))
        normalizado = normalizar_telefono_ar(campos["telefono"], area)
        if not normalizado:
            return False, f"El telefono '{campos['telefono']}' no es valido."
        for otro in lista:
            if otro.id != id_vendedor and otro.telefono == normalizado:
                return False, f"Ese numero ya lo usa {otro.etiqueta()}."
        campos["telefono"] = normalizado

    if "codigo" in campos and campos["codigo"]:
        cod = str(campos["codigo"]).strip()
        for otro in lista:
            if otro.id != id_vendedor and otro.codigo.strip() == cod:
                return False, f"El codigo {cod} ya lo usa {otro.etiqueta()}."
        campos["codigo"] = cod

    if "nombre" in campos and not str(campos["nombre"]).strip():
        return False, "El nombre no puede quedar vacio."

    for clave, valor in campos.items():
        if hasattr(objetivo, clave):
            setattr(objetivo, clave, valor)

    if not guardar_vendedores(lista):
        return False, "No se pudo grabar el archivo de vendedores."
    log(f"Vendedor actualizado: {objetivo.etiqueta()}")
    return True, ""


def eliminar_vendedor(id_vendedor: str) -> bool:
    """Borra el vendedor y deja sin asignar a sus clientes."""
    lista = cargar_vendedores()
    quedan = [v for v in lista if v.id != id_vendedor]
    if len(quedan) == len(lista):
        return False
    if not guardar_vendedores(quedan):
        return False

    asignaciones = cargar_asignaciones()
    limpio = {k: v for k, v in asignaciones.items() if v != id_vendedor}
    if len(limpio) != len(asignaciones):
        guardar_asignaciones(limpio)
    log(f"Vendedor eliminado: {id_vendedor}")
    return True


def vendedor_por_id(id_vendedor: str) -> Vendedor | None:
    if not id_vendedor:
        return None
    for v in cargar_vendedores():
        if v.id == id_vendedor:
            return v
    return None


def vendedor_por_codigo(codigo: str) -> Vendedor | None:
    """Busca por codigo del ERP, tolerando ceros a la izquierda ('2' == '002')."""
    if not codigo:
        return None
    buscado = str(codigo).strip().lstrip("0") or "0"
    for v in cargar_vendedores():
        propio = (v.codigo or "").strip().lstrip("0") or "0"
        if propio == buscado:
            return v
    return None


def vendedores_activos() -> list[Vendedor]:
    return [v for v in cargar_vendedores() if v.activo]


# ==========================================================================
# ASIGNACION CLIENTE -> VENDEDOR
# ==========================================================================
def _clave_cliente(codigo: str) -> str:
    """Normaliza el codigo de cliente: sin ceros a la izquierda ni no-digitos."""
    solo = "".join(ch for ch in str(codigo or "") if ch.isdigit())
    return solo.lstrip("0") or solo


def cargar_asignaciones() -> dict[str, str]:
    """Devuelve {codigo_cliente_normalizado: id_vendedor}."""
    crudo = _leer_json(ARCHIVO_ASIGNACIONES, {})
    if not isinstance(crudo, dict):
        log_error("asignaciones.json no contiene un diccionario; se ignora.")
        return {}
    return {
        _clave_cliente(k): str(v)
        for k, v in crudo.items()
        if _clave_cliente(k) and v
    }


def guardar_asignaciones(d: dict[str, str]) -> bool:
    return _guardar_json(ARCHIVO_ASIGNACIONES, d)


def asignar_cliente(codigo_cliente: str, id_vendedor: str) -> bool:
    """Asocia un cliente a un vendedor y lo persiste."""
    clave = _clave_cliente(codigo_cliente)
    if not clave or not id_vendedor:
        return False
    d = cargar_asignaciones()
    d[clave] = id_vendedor
    return guardar_asignaciones(d)


def asignar_varios(codigos_clientes: Iterable[str], id_vendedor: str) -> int:
    """Asigna muchos clientes de una sola escritura. Devuelve cuantos asigno."""
    if not id_vendedor:
        return 0
    d = cargar_asignaciones()
    n = 0
    for cod in codigos_clientes:
        clave = _clave_cliente(cod)
        if clave:
            d[clave] = id_vendedor
            n += 1
    if n and guardar_asignaciones(d):
        return n
    return 0


def desasignar_cliente(codigo_cliente: str) -> bool:
    clave = _clave_cliente(codigo_cliente)
    d = cargar_asignaciones()
    if clave in d:
        del d[clave]
        return guardar_asignaciones(d)
    return False


def vendedor_de_cliente(codigo_cliente: str) -> Vendedor | None:
    """Vendedor asignado a ese cliente, o None si no tiene."""
    d = cargar_asignaciones()
    id_v = d.get(_clave_cliente(codigo_cliente), "")
    return vendedor_por_id(id_v) if id_v else None


SIN_ASIGNAR = ""


def agrupar_por_vendedor(clientes: Iterable[Any]) -> dict[str, list[Any]]:
    """Agrupa clientes por id de vendedor segun las asignaciones guardadas.

    Los clientes sin asignar quedan bajo la clave "" (SIN_ASIGNAR).
    """
    d = cargar_asignaciones()
    grupos: dict[str, list[Any]] = {}
    for cli in clientes:
        codigo = getattr(cli, "codigo", "") or getattr(cli, "codigo_corto", "")
        id_v = d.get(_clave_cliente(codigo), SIN_ASIGNAR)
        grupos.setdefault(id_v, []).append(cli)
    return grupos
