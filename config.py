# -*- coding: utf-8 -*-
"""
Configuracion, rutas persistentes y log de errores.

Las credenciales NO van hardcodeadas: se guardan en `config.json` al lado del
ejecutable (o del .py en desarrollo). Ese archivo esta en .gitignore.
Tambien se pueden pasar por variables de entorno, que tienen prioridad.
"""

from __future__ import annotations

import json
import os
import sys
import threading
from datetime import datetime, timedelta, timezone
from typing import Any

APP_NOMBRE = "Gestor de Cobranzas Wood Tools"
APP_VERSION = "1.0.0"

# --------------------------------------------------------------------------
# RUTAS (compatible con PyInstaller --onefile, igual que WhatsAppMessage)
# --------------------------------------------------------------------------
def ruta_recurso(relativa: str) -> str:
    """Archivos empaquetados dentro del .exe (iconos, imagenes)."""
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relativa)


def ruta_persistente(nombre: str) -> str:
    """Archivos que el usuario modifica y deben sobrevivir a las updates."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, nombre)


ARCHIVO_CONFIG = ruta_persistente("config.json")
ARCHIVO_VENDEDORES = ruta_persistente("vendedores.json")
ARCHIVO_ASIGNACIONES = ruta_persistente("asignaciones.json")
ARCHIVO_LOG = ruta_persistente("errores_log.txt")
CARPETA_SALIDA = ruta_persistente("Resumenes")
CARPETA_HISTORIAL = ruta_persistente("Historial")


def hora_arg() -> datetime:
    """Hora de Argentina (UTC-3), sin depender de la zona del sistema."""
    return datetime.now(timezone.utc).astimezone(timezone(timedelta(hours=-3)))


# --------------------------------------------------------------------------
# LOG (nunca usar print(): congela los .exe con --noconsole)
# --------------------------------------------------------------------------
_lock_log = threading.Lock()


def log(mensaje: str, nivel: str = "INFO") -> None:
    try:
        with _lock_log:
            with open(ARCHIVO_LOG, "a", encoding="utf-8") as f:
                f.write(f"[{hora_arg():%Y-%m-%d %H:%M:%S}] {nivel}: {mensaje}\n")
    except Exception:
        pass


def log_error(mensaje: str) -> None:
    log(mensaje, "ERROR")


# --------------------------------------------------------------------------
# CONFIGURACION
# --------------------------------------------------------------------------
CONFIG_POR_DEFECTO: dict[str, Any] = {
    # --- Claude API ---
    "anthropic_api_key": "",
    "modelo_claude": "claude-opus-5",
    "esfuerzo_claude": "high",              # low | medium | high | xhigh | max
    "usar_claude_en_excel": False,          # el parser propio ya alcanza; True = verificar con Claude
    "usar_claude_en_pdf": True,             # los PDF conviene interpretarlos con Claude
    "usar_claude_para_telefonos": True,     # separar nombres de numeros en la col. Telefonos

    # --- WhatsApp Cloud API (Meta) ---
    "whatsapp_token": "",
    "whatsapp_phone_number_id": "",
    "whatsapp_version": "v21.0",
    "modo_envio": "libre",                  # "libre" (ventana 24h) | "plantilla"
    "plantilla_nombre": "",                 # solo si modo_envio == "plantilla"
    "plantilla_idioma": "es",
    "pausa_entre_envios_seg": 1.5,

    # --- Criterios de negocio ---
    "usar_fecha_reporte": True,             # False = recalcular dias contra HOY
    "dias_aviso_previo": 7,                 # a vencer dentro de N dias entra al mensaje
    "incluir_a_vencer": True,
    "solo_resaltados": False,               # usar unicamente los clientes marcados en amarillo
    "saldo_minimo": 0.0,                    # ignorar clientes con saldo menor a este
    "area_telefono_defecto": "11",

    # --- Mensajes / PDF ---
    "generar_pdf_por_vendedor": True,       # PDF filtrado con solo los clientes de ese vendedor
    "adjuntar_pdf_original": False,          # adjuntar el PDF tal como se cargo
    "max_caracteres_mensaje": 3800,          # WhatsApp corta en 4096
    "firma_mensaje": "Administracion Wood Tools",
}


class Config:
    """Acceso tipo diccionario a la configuracion, con persistencia en JSON."""

    # Variables de entorno que sobreescriben el JSON
    ENV = {
        "anthropic_api_key": "ANTHROPIC_API_KEY",
        "whatsapp_token": "WHATSAPP_TOKEN",
        "whatsapp_phone_number_id": "WHATSAPP_PHONE_NUMBER_ID",
    }

    def __init__(self, ruta: str = ARCHIVO_CONFIG):
        self.ruta = ruta
        self.datos: dict[str, Any] = dict(CONFIG_POR_DEFECTO)
        self.cargar()

    # ---------------- persistencia ----------------
    def cargar(self) -> None:
        if os.path.exists(self.ruta):
            try:
                with open(self.ruta, "r", encoding="utf-8") as f:
                    guardado = json.load(f)
                if isinstance(guardado, dict):
                    # solo claves conocidas, para no arrastrar basura
                    for k, v in guardado.items():
                        self.datos[k] = v
            except Exception as e:
                log_error(f"No se pudo leer config.json: {e}")
        # las variables de entorno mandan
        for clave, env in self.ENV.items():
            valor = os.environ.get(env)
            if valor:
                self.datos[clave] = valor.strip()

    def guardar(self) -> bool:
        try:
            with open(self.ruta, "w", encoding="utf-8") as f:
                json.dump(self.datos, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            log_error(f"No se pudo guardar config.json: {e}")
            return False

    # ---------------- acceso ----------------
    def get(self, clave: str, defecto: Any = None) -> Any:
        if clave in self.datos:
            return self.datos[clave]
        if clave in CONFIG_POR_DEFECTO:
            return CONFIG_POR_DEFECTO[clave]
        return defecto

    def set(self, clave: str, valor: Any) -> None:
        self.datos[clave] = valor

    def __getitem__(self, clave: str) -> Any:
        return self.get(clave)

    def __setitem__(self, clave: str, valor: Any) -> None:
        self.set(clave, valor)

    # ---------------- estado de credenciales ----------------
    @property
    def tiene_claude(self) -> bool:
        return bool(str(self.get("anthropic_api_key", "")).strip())

    @property
    def tiene_whatsapp(self) -> bool:
        return bool(
            str(self.get("whatsapp_token", "")).strip()
            and str(self.get("whatsapp_phone_number_id", "")).strip()
        )

    @property
    def base_url_whatsapp(self) -> str:
        version = self.get("whatsapp_version", "v21.0")
        pid = self.get("whatsapp_phone_number_id", "")
        return f"https://graph.facebook.com/{version}/{pid}"


def asegurar_carpetas() -> None:
    for carpeta in (CARPETA_SALIDA, CARPETA_HISTORIAL):
        try:
            os.makedirs(carpeta, exist_ok=True)
        except Exception as e:
            log_error(f"No se pudo crear la carpeta {carpeta}: {e}")


# Instancia global que usan los demas modulos
CONFIG = Config()
