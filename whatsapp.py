# -*- coding: utf-8 -*-
"""
Envio por WhatsApp Cloud API de Meta.

Mismo enfoque que WhatsAppMessage (la app de marketing que ya usa la empresa),
pero sin credenciales hardcodeadas: todo sale de CONFIG.

IMPORTANTE - VENTANA DE 24 HORAS:
Meta solo permite mensajes de texto/documento LIBRES si el destinatario le
escribio al numero de la empresa en las ultimas 24 horas. Si no, hay que usar
una plantilla aprobada. Los errores 131047 / 131026 son exactamente eso, y esta
clase los traduce a un mensaje entendible.
"""

from __future__ import annotations

import os
import time
import urllib.parse
from dataclasses import dataclass, field
from typing import Any, Callable

import requests

from config import CONFIG, log, log_error

TIMEOUT_MENSAJE = 20
TIMEOUT_SUBIDA = 60
TIMEOUT_DIAGNOSTICO = 15
MAX_REINTENTOS = 2


# ==========================================================================
# RESULTADO
# ==========================================================================
@dataclass
class ResultadoEnvio:
    ok: bool
    mensaje: str
    id_mensaje: str = ""
    codigo_error: str = ""
    detalle: str = ""
    telefono: str = ""

    def __bool__(self) -> bool:
        return self.ok


# ==========================================================================
# TRADUCCION DE ERRORES DE META
# ==========================================================================
_ERRORES_META: dict[str, str] = {
    "131047": (
        "La ventana de 24 horas esta cerrada. WhatsApp no permite mandar mensajes "
        "libres si la persona no te escribio en las ultimas 24 h.\n"
        "Solucion: pedile al vendedor que mande cualquier mensaje al numero de la "
        "empresa (con eso se abre la ventana por 24 h), o cambia a modo plantilla "
        "en Configuracion."
    ),
    "131026": (
        "El mensaje no se pudo entregar. Puede ser que la ventana de 24 horas este "
        "cerrada, que el numero no tenga WhatsApp, o que el telefono este mal escrito."
    ),
    "131030": (
        "El numero no esta en la lista de destinatarios permitidos. La cuenta de "
        "WhatsApp esta en modo de prueba: hay que agregar el numero en Meta o pasar "
        "la cuenta a produccion."
    ),
    "131051": "Tipo de mensaje no soportado por el destinatario.",
    "190": (
        "El token de WhatsApp vencio o es invalido. Hay que generar uno nuevo en "
        "Meta for Developers y pegarlo en Configuracion."
    ),
    "401": (
        "Meta rechazo la autenticacion: el token es invalido o no tiene permisos "
        "sobre este numero."
    ),
    "132000": "La plantilla existe pero la cantidad de parametros no coincide.",
    "132001": (
        "La plantilla no existe con ese nombre/idioma. Revisa el nombre exacto en "
        "Meta (respeta minusculas y guiones bajos)."
    ),
    "132005": "El texto de un parametro de la plantilla es demasiado largo.",
    "132007": (
        "El contenido del parametro viola el formato de la plantilla. Recorda que "
        "los parametros no pueden tener saltos de linea ni tabulaciones."
    ),
    "132012": "El formato de los parametros de la plantilla no es valido.",
    "132015": "La plantilla esta pausada por mala calidad.",
    "132016": "La plantilla fue deshabilitada por Meta.",
    "133010": (
        "El numero de la empresa (phone_number_id) no esta registrado en la Cloud API."
    ),
    "100": (
        "Parametro invalido en la llamada a Meta. Verifica el phone_number_id y la "
        "version de la API en Configuracion."
    ),
    "4": "Superaste el limite de velocidad de la API. Espera unos segundos.",
    "80007": "Superaste el limite de velocidad de la API. Espera unos segundos.",
    "368": "El numero de la empresa esta restringido temporalmente por Meta.",
    "133016": "El numero esta bloqueado temporalmente por Meta (demasiados intentos).",
}

_CODIGOS_REINTENTABLES = {"4", "80007", "131056", "500", "503"}


def _traducir_error(codigo: Any, mensaje_meta: str) -> str:
    cod = str(codigo or "").strip()
    if cod in _ERRORES_META:
        return _ERRORES_META[cod]
    if mensaje_meta and "re-engagement" in mensaje_meta.lower():
        return _ERRORES_META["131047"]
    if cod:
        return f"Meta rechazo el mensaje (codigo {cod}): {mensaje_meta or 'sin detalle'}"
    return f"Meta rechazo el mensaje: {mensaje_meta or 'sin detalle'}"


# ==========================================================================
# HELPERS
# ==========================================================================
def esta_configurado() -> bool:
    """True si hay token y phone_number_id cargados."""
    return CONFIG.tiene_whatsapp


def _headers_json() -> dict[str, str]:
    return {
        "Authorization": f"Bearer {CONFIG.get('whatsapp_token', '')}",
        "Content-Type": "application/json",
    }


def _headers_auth() -> dict[str, str]:
    return {"Authorization": f"Bearer {CONFIG.get('whatsapp_token', '')}"}


def normalizar_destino(telefono: str) -> tuple[str | None, str]:
    """Valida el telefono destino. Devuelve (solo_digitos, motivo_si_falla)."""
    solo = "".join(ch for ch in str(telefono or "") if ch.isdigit())
    if not solo:
        return None, "El telefono esta vacio."
    if len(solo) < 10 or len(solo) > 15:
        return None, (
            f"El telefono '{telefono}' no es valido ({len(solo)} digitos). "
            "Tiene que estar en formato internacional, por ejemplo 5491139009488."
        )
    return solo, ""


def _extraer_error(res: requests.Response) -> tuple[str, str]:
    """Saca (codigo, mensaje) del cuerpo de error de Meta."""
    try:
        cuerpo = res.json()
        err = cuerpo.get("error", {}) if isinstance(cuerpo, dict) else {}
        codigo = str(err.get("code", "") or res.status_code)
        mensaje = str(err.get("message", "") or "")
        detalles = err.get("error_data") or {}
        if isinstance(detalles, dict) and detalles.get("details"):
            mensaje = f"{mensaje} ({detalles['details']})"
        return codigo, mensaje
    except Exception:
        return str(res.status_code), (res.text or "")[:400]


# ==========================================================================
# LLAMADA BASE A /messages
# ==========================================================================
def _postear_mensaje(datos: dict[str, Any], telefono: str = "") -> ResultadoEnvio:
    """POST a /messages con reintentos solo en fallas transitorias."""
    if not esta_configurado():
        return ResultadoEnvio(
            False,
            "Falta configurar el token y el ID de numero de WhatsApp en Configuracion.",
            telefono=telefono,
        )

    url = f"{CONFIG.base_url_whatsapp}/messages"
    espera = 2

    for intento in range(1, MAX_REINTENTOS + 2):
        try:
            res = requests.post(
                url, headers=_headers_json(), json=datos, timeout=TIMEOUT_MENSAJE
            )
        except requests.exceptions.Timeout:
            log_error(f"WhatsApp timeout (intento {intento}) hacia {telefono}")
            if intento <= MAX_REINTENTOS:
                time.sleep(espera)
                espera *= 2
                continue
            return ResultadoEnvio(
                False, "Meta tardo demasiado en responder. Intenta de nuevo.",
                codigo_error="timeout", telefono=telefono,
            )
        except requests.exceptions.RequestException as e:
            log_error(f"WhatsApp error de red (intento {intento}): {e}")
            if intento <= MAX_REINTENTOS:
                time.sleep(espera)
                espera *= 2
                continue
            return ResultadoEnvio(
                False, "No hay conexion con Meta. Revisa internet.",
                codigo_error="red", detalle=str(e), telefono=telefono,
            )

        if res.status_code == 200:
            id_msg = ""
            try:
                cuerpo = res.json()
                mensajes = cuerpo.get("messages") or []
                if mensajes:
                    id_msg = str(mensajes[0].get("id", ""))
            except Exception:
                pass
            return ResultadoEnvio(True, "Enviado", id_mensaje=id_msg, telefono=telefono)

        codigo, mensaje_meta = _extraer_error(res)
        log_error(
            f"WhatsApp {res.status_code} hacia {telefono} | codigo={codigo} | "
            f"{mensaje_meta} | cuerpo={(res.text or '')[:500]}"
        )

        transitorio = res.status_code >= 500 or codigo in _CODIGOS_REINTENTABLES
        if transitorio and intento <= MAX_REINTENTOS:
            time.sleep(espera)
            espera *= 2
            continue

        return ResultadoEnvio(
            False,
            _traducir_error(codigo, mensaje_meta),
            codigo_error=codigo,
            detalle=(res.text or "")[:500],
            telefono=telefono,
        )

    return ResultadoEnvio(False, "No se pudo enviar el mensaje.", telefono=telefono)


# ==========================================================================
# DIAGNOSTICO
# ==========================================================================
def diagnosticar() -> ResultadoEnvio:
    """Verifica que el token y el phone_number_id sean validos."""
    if not esta_configurado():
        return ResultadoEnvio(
            False, "Falta el token o el ID de numero de WhatsApp en Configuracion."
        )
    version = CONFIG.get("whatsapp_version", "v21.0")
    pid = CONFIG.get("whatsapp_phone_number_id", "")
    url = f"https://graph.facebook.com/{version}/{pid}"
    try:
        res = requests.get(
            url,
            headers=_headers_auth(),
            params={"fields": "display_phone_number,verified_name,quality_rating"},
            timeout=TIMEOUT_DIAGNOSTICO,
        )
    except requests.exceptions.RequestException as e:
        return ResultadoEnvio(
            False, "No hay conexion con Meta. Revisa internet.", detalle=str(e)
        )

    if res.status_code == 200:
        try:
            d = res.json()
        except Exception:
            d = {}
        numero = d.get("display_phone_number", "?")
        nombre = d.get("verified_name", "?")
        calidad = d.get("quality_rating", "?")
        return ResultadoEnvio(
            True,
            f"Conexion OK.\nNumero: {numero}\nNombre verificado: {nombre}\n"
            f"Calidad segun Meta: {calidad}",
            detalle=res.text[:400],
        )

    codigo, mensaje_meta = _extraer_error(res)
    log_error(f"Diagnostico WhatsApp fallo: {codigo} {mensaje_meta}")
    return ResultadoEnvio(
        False, _traducir_error(codigo, mensaje_meta), codigo_error=codigo,
        detalle=(res.text or "")[:400],
    )


# ==========================================================================
# SUBIDA DE DOCUMENTOS
# ==========================================================================
def subir_documento(ruta: str) -> tuple[str | None, str]:
    """Sube un PDF a Meta y devuelve (media_id, mensaje)."""
    if not esta_configurado():
        return None, "Falta configurar WhatsApp."
    if not ruta or not os.path.exists(ruta):
        return None, f"No se encuentra el archivo: {ruta}"

    tamanio = os.path.getsize(ruta)
    if tamanio > 95 * 1024 * 1024:
        return None, "El PDF supera el limite de 100 MB de WhatsApp."
    if tamanio == 0:
        return None, "El PDF esta vacio."

    url = f"{CONFIG.base_url_whatsapp}/media"
    try:
        with open(ruta, "rb") as fh:
            archivos = {
                "file": (os.path.basename(ruta), fh, "application/pdf"),
            }
            datos = {"messaging_product": "whatsapp", "type": "application/pdf"}
            res = requests.post(
                url,
                headers=_headers_auth(),
                files=archivos,
                data=datos,
                timeout=TIMEOUT_SUBIDA,
            )
    except requests.exceptions.Timeout:
        return None, "Meta tardo demasiado subiendo el PDF. Intenta de nuevo."
    except requests.exceptions.RequestException as e:
        log_error(f"Error de red subiendo PDF: {e}")
        return None, "No hay conexion con Meta para subir el PDF."
    except OSError as e:
        log_error(f"No se pudo abrir el PDF {ruta}: {e}")
        return None, f"No se pudo leer el PDF: {e}"

    if res.status_code == 200:
        try:
            media_id = str(res.json().get("id", ""))
        except Exception:
            media_id = ""
        if media_id:
            log(f"PDF subido a Meta: {os.path.basename(ruta)} -> {media_id}")
            return media_id, "PDF subido"
        return None, "Meta no devolvio el identificador del PDF."

    codigo, mensaje_meta = _extraer_error(res)
    log_error(f"Error subiendo PDF: {codigo} {mensaje_meta}")
    return None, _traducir_error(codigo, mensaje_meta)


# ==========================================================================
# ENVIOS
# ==========================================================================
def enviar_texto(telefono: str, texto: str) -> ResultadoEnvio:
    """Manda un mensaje de texto libre (requiere ventana de 24 h abierta)."""
    destino, motivo = normalizar_destino(telefono)
    if destino is None:
        return ResultadoEnvio(False, motivo, telefono=str(telefono))
    if not (texto or "").strip():
        return ResultadoEnvio(False, "El mensaje esta vacio.", telefono=destino)

    return _postear_mensaje(
        {
            "messaging_product": "whatsapp",
            "recipient_type": "individual",
            "to": destino,
            "type": "text",
            "text": {"preview_url": False, "body": texto[:4096]},
        },
        destino,
    )


def enviar_documento(
    telefono: str, media_id: str, nombre_archivo: str, epigrafe: str = ""
) -> ResultadoEnvio:
    """Manda un PDF ya subido (requiere ventana de 24 h abierta)."""
    destino, motivo = normalizar_destino(telefono)
    if destino is None:
        return ResultadoEnvio(False, motivo, telefono=str(telefono))
    if not media_id:
        return ResultadoEnvio(False, "Falta el identificador del PDF.", telefono=destino)

    documento: dict[str, Any] = {"id": media_id}
    if nombre_archivo:
        documento["filename"] = nombre_archivo[:240]
    if epigrafe:
        documento["caption"] = epigrafe[:1000]

    return _postear_mensaje(
        {
            "messaging_product": "whatsapp",
            "recipient_type": "individual",
            "to": destino,
            "type": "document",
            "document": documento,
        },
        destino,
    )


def enviar_plantilla(
    telefono: str,
    nombre: str,
    parametros_cuerpo: list[str],
    media_id_documento: str | None = None,
) -> ResultadoEnvio:
    """Manda una plantilla aprobada (funciona con la ventana de 24 h cerrada)."""
    destino, motivo = normalizar_destino(telefono)
    if destino is None:
        return ResultadoEnvio(False, motivo, telefono=str(telefono))
    if not nombre:
        return ResultadoEnvio(
            False, "Falta el nombre de la plantilla en Configuracion.", telefono=destino
        )

    componentes: list[dict[str, Any]] = []
    if media_id_documento:
        componentes.append(
            {
                "type": "header",
                "parameters": [
                    {"type": "document", "document": {"id": media_id_documento}}
                ],
            }
        )
    if parametros_cuerpo:
        limpios = [
            str(p).replace("\n", " ").replace("\r", " ").replace("\t", " ")[:1000]
            for p in parametros_cuerpo
        ]
        componentes.append(
            {
                "type": "body",
                "parameters": [{"type": "text", "text": p} for p in limpios],
            }
        )

    plantilla: dict[str, Any] = {
        "name": nombre,
        "language": {"code": CONFIG.get("plantilla_idioma", "es")},
    }
    if componentes:
        plantilla["components"] = componentes

    return _postear_mensaje(
        {
            "messaging_product": "whatsapp",
            "recipient_type": "individual",
            "to": destino,
            "type": "template",
            "template": plantilla,
        },
        destino,
    )


def link_wa_me(telefono: str, texto: str) -> str:
    """Link wa.me con el texto pre-cargado, para enviar a mano si la API falla."""
    solo = "".join(ch for ch in str(telefono or "") if ch.isdigit())
    codificado = urllib.parse.quote(texto or "", safe="")
    return f"https://wa.me/{solo}?text={codificado}"


# ==========================================================================
# ORQUESTADOR: TODO EL RESUMEN A UN VENDEDOR
# ==========================================================================
def enviar_resumen(
    telefono: str,
    partes_texto: list[str],
    ruta_pdf: str | None = None,
    nombre_pdf: str = "resumen.pdf",
    epigrafe_pdf: str = "",
    progreso: Callable[[str], None] | None = None,
) -> list[ResultadoEnvio]:
    """Manda el PDF (si hay) y despues cada parte del texto, en orden.

    Si un envio falla, corta y devuelve los resultados obtenidos hasta ahi.
    """
    def avisar(txt: str) -> None:
        if progreso:
            try:
                progreso(txt)
            except Exception:
                pass

    resultados: list[ResultadoEnvio] = []
    modo_plantilla = str(CONFIG.get("modo_envio", "libre")) == "plantilla"
    pausa = float(CONFIG.get("pausa_entre_envios_seg", 1.5) or 0)

    # --- PDF ---
    media_id: str | None = None
    if ruta_pdf:
        avisar("Subiendo el PDF a WhatsApp...")
        media_id, msg_subida = subir_documento(ruta_pdf)
        if media_id is None:
            resultados.append(ResultadoEnvio(False, msg_subida, telefono=telefono))
            return resultados

    # --- modo plantilla: un solo mensaje ---
    if modo_plantilla:
        avisar("Enviando la plantilla...")
        texto_plano = " | ".join(p.replace("\n", " ") for p in partes_texto)[:1000]
        res = enviar_plantilla(
            telefono,
            str(CONFIG.get("plantilla_nombre", "")),
            [texto_plano],
            media_id,
        )
        resultados.append(res)
        return resultados

    # --- modo libre: PDF primero, despues el texto ---
    if media_id:
        avisar("Enviando el PDF...")
        res = enviar_documento(telefono, media_id, nombre_pdf, epigrafe_pdf)
        resultados.append(res)
        if not res.ok:
            return resultados
        if pausa:
            time.sleep(pausa)

    total = len(partes_texto)
    for i, parte in enumerate(partes_texto, start=1):
        avisar(f"Enviando mensaje {i} de {total}...")
        res = enviar_texto(telefono, parte)
        resultados.append(res)
        if not res.ok:
            return resultados
        if i < total and pausa:
            time.sleep(pausa)

    return resultados
