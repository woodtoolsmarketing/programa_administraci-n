# -*- coding: utf-8 -*-
"""
Integracion con la API de Claude para interpretar los resumenes de cuenta.

Para que sirve:
  1. Interpretar un PDF del reporte (Claude lo lee como documento, incluso si es
     un escaneo) y devolver un Reporte poblado.
  2. Interpretar el volcado de un Excel cuyo formato cambio y el parser
     deterministico no pudo leer. El volcado trae marcadas las celdas
     resaltadas en amarillo con el sufijo [AMARILLO].
  3. Auditar un Reporte ya parseado (verificar que los totales cierren).
  4. Separar nombres de contacto y numeros de telefono del texto libre que trae
     la columna Telefonos del ERP.

Notas de la API que hay que respetar (claude-opus-5):
  - NO se envian temperature / top_p / top_k / thinking.budget_tokens: dan 400.
  - No se configura `thinking`: en claude-opus-5 el pensamiento ya viene activo
    por defecto, y omitirlo es lo mas compatible con otros modelos.
  - `effort` va DENTRO de output_config, junto con `format`.
  - El bloque `document` del PDF va ANTES del bloque de texto.
  - Se usa streaming para no cortar por timeout HTTP en entradas largas.
"""

from __future__ import annotations

import base64
import json
import os
import re
from datetime import date, datetime
from typing import Any, Callable

from config import CONFIG, log, log_error
from modelos import (
    Cliente,
    Comprobante,
    Reporte,
    clasificar_dias,
    telefonos_de_texto,
)

# Precios de claude-opus-5, en USD por millon de tokens
PRECIO_ENTRADA = 5.0
PRECIO_SALIDA = 25.0

MAX_TOKENS_STREAM = 64000
BETA_FALLBACK = "server-side-fallback-2026-07-01"

# Clientes por llamada al ordenar telefonos. Con carteras grandes, una sola
# llamada se corta por max_tokens.
TAMANIO_LOTE_TELEFONOS = 60


class ErrorClaude(Exception):
    """Cualquier problema al hablar con la API de Claude, ya traducido."""


# ==========================================================================
# ESQUEMAS DE SALIDA ESTRUCTURADA
# ==========================================================================
_ESQUEMA_COMPROBANTE: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "required": [
        "talonario", "tipo", "numero", "condicion_pago", "fecha_emision",
        "fecha_vto", "cuota", "importe_original", "saldo", "dias_venc",
        "resaltado",
    ],
    "properties": {
        "talonario": {"type": "string", "description": "Numero de talonario, ej '71'"},
        "tipo": {"type": "string", "description": "Tipo de comprobante, ej 'FE A'"},
        "numero": {"type": "string", "description": "Numero, ej '00072065'"},
        "condicion_pago": {"type": "string"},
        "fecha_emision": {
            "type": ["string", "null"], "description": "DD/MM/AAAA o null",
        },
        "fecha_vto": {
            "type": ["string", "null"], "description": "DD/MM/AAAA o null",
        },
        "cuota": {"type": "string", "description": "ej '01/01'"},
        "importe_original": {
            "type": "number", "description": "Numero, sin separador de miles",
        },
        "saldo": {
            "type": "number",
            "description": "Saldo parcial pendiente de este comprobante",
        },
        "dias_venc": {
            "type": ["integer", "null"],
            "description": "Positivo = vencido, 0 = vence hoy, negativo = a vencer",
        },
        "resaltado": {
            "type": "boolean", "description": "true si alguna celda venia [AMARILLO]",
        },
    },
}

_ESQUEMA_CLIENTE: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "required": [
        "codigo", "razon_social", "nombre_fantasia", "direccion", "localidad",
        "contacto", "telefonos_raw", "total_original_reporte",
        "total_saldo_reporte", "resaltado", "comprobantes",
    ],
    "properties": {
        "codigo": {
            "type": "string",
            "description": "Codigo de cliente tal cual figura, ej '00000002761'",
        },
        "razon_social": {"type": "string"},
        "nombre_fantasia": {"type": "string"},
        "direccion": {"type": "string"},
        "localidad": {"type": "string", "description": "ej '(1047) CAPITAL FEDERAL'"},
        "contacto": {"type": "string", "description": "Nombre de la persona de contacto"},
        "telefonos_raw": {
            "type": "string",
            "description": "El texto de telefonos tal cual aparece, sin limpiar",
        },
        "total_original_reporte": {
            "type": ["number", "null"],
            "description": "Importe original de la fila 'Total cliente'",
        },
        "total_saldo_reporte": {
            "type": ["number", "null"],
            "description": "SALDO TOTAL de la fila 'Total cliente'",
        },
        "resaltado": {"type": "boolean"},
        "comprobantes": {"type": "array", "items": _ESQUEMA_COMPROBANTE},
    },
}

ESQUEMA_REPORTE: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "required": [
        "empresa", "titulo", "fecha_reporte", "filtros",
        "total_general_original", "total_general_saldo", "clientes", "avisos",
    ],
    "properties": {
        "empresa": {"type": "string"},
        "titulo": {"type": "string"},
        "fecha_reporte": {
            "type": ["string", "null"],
            "description": "Fecha del reporte en DD/MM/AAAA",
        },
        "filtros": {
            "type": "string",
            "description": "Linea de filtros aplicados (Cliente, Cobrador, Vendedor...)",
        },
        "total_general_original": {"type": ["number", "null"]},
        "total_general_saldo": {"type": ["number", "null"]},
        "clientes": {"type": "array", "items": _ESQUEMA_CLIENTE},
        "avisos": {
            "type": "array",
            "items": {"type": "string"},
            "description": "Cosas dudosas o ilegibles que encontraste",
        },
    },
}

ESQUEMA_VERIFICACION: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "required": [
        "coincide", "observaciones", "clientes_faltantes", "diferencias",
    ],
    "properties": {
        "coincide": {
            "type": "boolean",
            "description": "true si los datos parseados representan fielmente el original",
        },
        "observaciones": {"type": "array", "items": {"type": "string"}},
        "clientes_faltantes": {
            "type": "array",
            "items": {"type": "string"},
            "description": "Clientes que estan en el original y faltan en lo parseado",
        },
        "diferencias": {
            "type": "array",
            "items": {"type": "string"},
            "description": "Diferencias concretas de importes o de cantidad de comprobantes",
        },
    },
}

ESQUEMA_TELEFONOS: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "required": ["clientes"],
    "properties": {
        "clientes": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["codigo", "contactos", "telefonos"],
                "properties": {
                    "codigo": {"type": "string"},
                    "contactos": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Nombres de personas encontrados en el texto",
                    },
                    "telefonos": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "required": ["numero", "es_movil", "contacto"],
                            "properties": {
                                "numero": {
                                    "type": "string",
                                    "description": "549 + area + abonado, solo digitos",
                                },
                                "es_movil": {"type": "boolean"},
                                "contacto": {
                                    "type": "string",
                                    "description": "Nombre asociado, o cadena vacia",
                                },
                            },
                        },
                    },
                },
            },
        }
    },
}


# ==========================================================================
# PROMPTS
# ==========================================================================
PROMPT_SISTEMA_EXTRACCION = """\
Sos un asistente de administracion de Wood Tools SRL. Tu trabajo es leer el \
reporte "A COBRAR POR CLIENTE EXPRESADO EN $" que exporta el ERP Gestion \
Comercial y devolverlo como datos estructurados, con precision contable.

ESTRUCTURA DEL REPORTE

Cabecera: nombre de la empresa, el titulo, "Fecha:" con la fecha del reporte, \
"Hoja:" con el numero de pagina, y una linea de filtros aplicados que puede \
incluir "Cliente:", "Cobrador:NNN-NNN", "Vendedor:NNN-NNN", rangos de emision y \
de vencimiento. Estas lineas se repiten en cada pagina: no las confundas con \
datos de clientes.

Por cada cliente, en este orden:
1. Una linea de cabecera con el CODIGO DE CLIENTE (numero de 11 digitos con \
ceros a la izquierda, por ejemplo 00000002761) y la RAZON SOCIAL. En esa misma \
linea o en las siguientes pueden venir los TELEFONOS como texto libre, mezclados \
con nombres de personas (por ejemplo "Marcos 11 3900 9488 -4313 0466" o \
"45246075 LAUTARO 1140647300 KAREN PAGO").
2. Una o dos lineas mas con el NOMBRE DE FANTASIA, la LOCALIDAD con el formato \
"(1047) CAPITAL FEDERAL", la DIRECCION y a veces un nombre de CONTACTO.
   Regla practica: la ultima de esas lineas de texto suele ser la direccion; las \
anteriores, el nombre de fantasia.
3. La lista de COMPROBANTES. Cada uno trae: talonario, tipo (por ejemplo \
"FE A"), numero, condicion de pago (por ejemplo "30 dias.ff.CONTRA ENTREGA"), \
fecha de emision, fecha de vencimiento, numero de cuota (por ejemplo "01/01"), \
IMPORTE ORIGINAL, SALDO (este es el saldo parcial, lo que todavia se debe de ese \
comprobante) y DIAS VENC.
4. Una fila "Total cliente <codigo> <razon social>" con el importe original total \
y el SALDO TOTAL del cliente.

Al final del reporte, una fila "Total general:" con los totales de todo.

REGLA CLAVE DE LOS DIAS
"Dias Venc." = fecha del reporte menos la fecha de vencimiento del comprobante.
  - Positivo  => el comprobante esta VENCIDO (hace esa cantidad de dias).
  - Cero      => vence hoy.
  - Negativo  => todavia no vencio (falta esa cantidad de dias).
Si en el documento el numero de dias aparece sin signo pero la fecha de \
vencimiento es POSTERIOR a la fecha del reporte, entonces el comprobante NO esta \
vencido y tenes que devolver dias_venc en NEGATIVO.

CELDAS RESALTADAS
Cuando la entrada sea un volcado de Excel, las celdas que la persona de \
administracion marco a mano vienen anotadas con el sufijo [AMARILLO]. Ese \
resaltado marca los datos importantes: el codigo y la razon social del cliente, \
y la columna Saldo. Devolve resaltado=true en el cliente y en el comprobante \
correspondiente cuando veas esa marca. El resaltado es solo una senal de \
atencion: NO cambia la clasificacion de vencido o a vencer, que sale siempre de \
las fechas.

COMO DEVOLVER LOS DATOS
- Los importes van como NUMEROS. El reporte usa formato argentino, con punto de \
miles y coma decimal: "37.165,39" se devuelve como 37165.39. "1.234" se devuelve \
como 1234.
- Las fechas van como texto "DD/MM/AAAA".
- El codigo de cliente se devuelve tal cual aparece, con los ceros a la izquierda.
- telefonos_raw se devuelve tal cual, sin limpiar ni reordenar.
- Si un dato no esta, devolve cadena vacia para los textos y null para los \
numeros y las fechas. No inventes NADA.
- Si algo esta ilegible, cortado o te genera dudas, no lo adivines: describilo en \
el array "avisos".

VERIFICACION OBLIGATORIA ANTES DE RESPONDER
Para cada cliente, la suma de los saldos de sus comprobantes tiene que dar el \
saldo total de su fila "Total cliente". Y la suma de los totales de todos los \
clientes tiene que dar el "Total general". Si algo no cierra, revisa si te \
salteaste o duplicaste un comprobante. Si despues de revisar sigue sin cerrar, \
dejalo asentado en "avisos" con los numeros concretos.
"""

PROMPT_SISTEMA_AUDITORIA = """\
Sos auditor de datos de cobranzas en Wood Tools SRL. NO extraes datos: revisas
el trabajo de otro programa.

Te paso dos cosas: (1) lo que un parser propio leyo de un reporte de cuentas a
cobrar, y (2) el documento original. Tu unico trabajo es decir si lo leido
representa fielmente al original.

Que tenes que revisar, en este orden de importancia:
1. CLIENTES FALTANTES O SOBRANTES. Conta los clientes del original y comparalos
uno por uno con la lista que te paso, por codigo. Un cliente que falta es el
error mas grave, porque nadie lo va a llamar a cobrar.
2. IMPORTES. Para cada cliente, el saldo total leido tiene que coincidir con el
de su fila "Total cliente" en el original, y la suma de todos con el
"Total general". Cita los numeros concretos cuando no coincidan.
3. CANTIDAD DE COMPROBANTES por cliente.
4. EL RESALTADO AMARILLO: si un cliente venia marcado con [AMARILLO] en el
original, tiene que figurar como resaltado.

Reglas:
- Se concreto y verificable: cita codigos de cliente e importes exactos, no
"parece que falta algo".
- Una diferencia de menos de un peso por redondeo no es un problema: no la
reportes.
- Si todo esta bien, devolve coincide=true y los arrays vacios. No inventes
observaciones para parecer util.
- Si no podes verificar algo porque el documento esta ilegible o incompleto,
decilo en observaciones y devolve coincide=false. Es mejor admitir que no
pudiste revisar que dar por bueno algo que no viste.
"""

PROMPT_SISTEMA_TELEFONOS = """\
Sos un asistente que ordena datos de contacto de clientes argentinos.

Te paso, por cada cliente, el codigo y el contenido crudo de la columna \
"Telefonos" del ERP. Ese texto viene sucio: nombres de personas pegados a los \
numeros, varios telefonos separados por barras, numeros incompletos y sobrantes.

Tu trabajo es separar los nombres de los numeros y normalizar cada numero al \
formato que usa WhatsApp en Argentina:
    549 + codigo de area (sin el 0) + numero de abonado (sin el 15)
solo digitos, sin el signo mas, sin espacios ni guiones.

Ejemplos:
  "Marcos 11 3900 9488 -4313 0466"  -> contacto "Marcos";
      5491139009488 (movil) y 5491143130466 (fijo)
  "4382-1590 / Jonathan153-359-5605" -> contacto "Jonathan";
      5491143821590 (fijo) y 5491133595605 (movil, tenia el 15 viejo)
  "0341 15 456 7890"                 -> 5493414567890 (movil de Rosario)
  "114085-5781"                      -> 5491140855781

Reglas:
- Si el numero no tiene codigo de area y tiene 8 digitos, asumi el area 11 \
(Ciudad de Buenos Aires), que es donde esta la mayoria de los clientes.
- El viejo prefijo 15 de los celulares se elimina: "11 15 4064 7300" es \
5491140647300.
- Si un numero quedo cortado o tiene menos de 8 digitos utiles, DESCARTALO. Es \
mejor no tener un numero que tener uno que no existe.
- Marca es_movil=true solo si es razonablemente un celular. En el area 11, los \
numeros que empiezan con 4 o 5 suelen ser fijos.
- Ordena los telefonos poniendo primero los moviles, porque son los que sirven \
para WhatsApp.
- Si un nombre esta pegado a un numero en particular, asocialo en el campo \
contacto de ese telefono. Si no se puede saber a quien corresponde, deja \
contacto vacio y pone el nombre en el array contactos del cliente.
- No inventes numeros ni nombres.
"""


# ==========================================================================
# CLIENTE DE LA API
# ==========================================================================
def _credencial_en_entorno() -> bool:
    """True si el sistema ya tiene credenciales de Anthropic fuera del config.

    Cubre la variable de entorno ANTHROPIC_API_KEY (o ANTHROPIC_AUTH_TOKEN) y el
    perfil que deja el comando `ant auth login`.
    """
    if os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN"):
        return True
    for base in (
        os.environ.get("ANTHROPIC_CONFIG_DIR"),
        os.path.join(os.environ.get("APPDATA", ""), "Anthropic"),
        os.path.join(os.path.expanduser("~"), ".config", "anthropic"),
    ):
        if base and os.path.isdir(os.path.join(base, "credentials")):
            return True
    return False


def esta_disponible() -> bool:
    """True si se puede usar la API (clave en config, en entorno o perfil)."""
    return CONFIG.tiene_claude or _credencial_en_entorno()


def _cliente() -> Any:
    """Devuelve el cliente de anthropic, o levanta ErrorClaude."""
    try:
        import anthropic
    except ImportError:
        raise ErrorClaude(
            "No esta instalada la libreria 'anthropic'.\n"
            "Instalala con: pip install anthropic"
        )

    clave = str(CONFIG.get("anthropic_api_key", "") or "").strip()
    if not clave and not _credencial_en_entorno():
        raise ErrorClaude(
            "Falta la API key de Claude.\n\n"
            "Anda a la pestana Configuracion y pega tu clave de "
            "console.anthropic.com (empieza con 'sk-ant-')."
        )

    try:
        # Sin clave explicita, el SDK resuelve la credencial del entorno.
        return anthropic.Anthropic(api_key=clave) if clave else anthropic.Anthropic()
    except Exception as e:
        log_error(f"No se pudo crear el cliente de Anthropic: {e}")
        raise ErrorClaude(f"No se pudo inicializar la API de Claude: {e}")


def _sin_credito(texto: str) -> bool:
    """Detecta el error de saldo agotado, que Anthropic devuelve como HTTP 400.

    Ojo: viene como invalid_request_error, no como 401 ni 429, asi que si no se
    mira el texto se confunde con un problema del pedido.
    """
    t = (texto or "").lower()
    marcas = (
        "credit balance is too low",
        "plans & billing",
        "insufficient credit",
        "billing",
        "purchase credits",
    )
    return any(m in t for m in marcas)


def _traducir_excepcion(e: Exception) -> ErrorClaude:
    """Convierte una excepcion del SDK en un ErrorClaude legible."""
    try:
        import anthropic
    except ImportError:
        return ErrorClaude(f"Error llamando a Claude: {e}")

    # El saldo agotado llega como 400: hay que mirarlo antes que nada, porque
    # si no se confunde con "el pedido esta mal armado".
    if _sin_credito(str(e)):
        return ErrorClaude(
            "La cuenta de Anthropic no tiene credito disponible.\n\n"
            "La API key es valida, pero hay que cargar saldo para poder usarla:\n"
            "  1. Entra a console.anthropic.com\n"
            "  2. Anda a 'Plans & Billing'\n"
            "  3. Carga credito (con USD 5 alcanza para cientos de reportes)\n\n"
            "Mientras tanto el programa funciona igual con archivos Excel: el "
            "lector propio no usa la API."
        )

    if isinstance(e, anthropic.AuthenticationError):
        return ErrorClaude(
            "La API key de Claude es invalida o fue revocada.\n"
            "Genera una nueva en console.anthropic.com y actualizala en "
            "Configuracion."
        )
    if isinstance(e, anthropic.PermissionDeniedError):
        return ErrorClaude(
            "La API key no tiene permiso para usar este modelo. Revisa el plan "
            "de tu cuenta en console.anthropic.com."
        )
    if isinstance(e, anthropic.NotFoundError):
        return ErrorClaude(
            f"El modelo '{CONFIG.get('modelo_claude')}' no existe o no esta "
            "disponible para tu cuenta. Revisa el nombre en Configuracion."
        )
    if isinstance(e, anthropic.RateLimitError):
        return ErrorClaude(
            "Superaste el limite de uso de la API de Claude (o te quedaste sin "
            "credito). Espera unos minutos y volve a intentar, o revisa el saldo "
            "en console.anthropic.com."
        )
    if isinstance(e, anthropic.BadRequestError):
        return ErrorClaude(
            f"Claude rechazo el pedido: {e}\n\n"
            "Puede ser que el archivo sea demasiado grande. Proba con un reporte "
            "de menos paginas."
        )
    if isinstance(e, anthropic.APIConnectionError):
        return ErrorClaude(
            "No hay conexion con la API de Claude. Revisa internet y volve a "
            "intentar."
        )
    if isinstance(e, anthropic.APIStatusError):
        codigo = getattr(e, "status_code", "?")
        if isinstance(codigo, int) and codigo >= 500:
            return ErrorClaude(
                f"La API de Claude tuvo un error interno ({codigo}). Volve a "
                "intentar en un rato."
            )
        return ErrorClaude(f"Error de la API de Claude ({codigo}): {e}")
    return ErrorClaude(f"Error inesperado llamando a Claude: {e}")


def _avisar(progreso: Callable[[str], None] | None, texto: str) -> None:
    if progreso:
        try:
            progreso(texto)
        except Exception:
            pass


# ==========================================================================
# LLAMADA BASE
# ==========================================================================
def _llamar(
    sistema: str,
    contenido_usuario: list[dict[str, Any]],
    esquema: dict[str, Any],
    progreso: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    """Hace la llamada con salida estructurada y devuelve el dict ya parseado.

    Intenta primero el endpoint beta con `fallbacks="default"` (si Claude
    rechaza el pedido por politica, la API lo reintenta sola con otro modelo).
    Si esa combinacion no esta habilitada, reintenta lo mismo sin betas.
    """
    cliente = _cliente()
    modelo = str(CONFIG.get("modelo_claude", "claude-opus-5"))
    esfuerzo = str(CONFIG.get("esfuerzo_claude", "high"))

    parametros: dict[str, Any] = {
        "model": modelo,
        "max_tokens": MAX_TOKENS_STREAM,
        # Explicito a proposito: en claude-opus-5 el pensamiento ya viene activo,
        # pero en claude-opus-4-8 (que tambien esta en la lista de Configuracion)
        # omitirlo significa SIN pensamiento, y se notaria en la calidad de la
        # lectura sin que el usuario sepa por que.
        "thinking": {"type": "adaptive"},
        "system": [
            {
                "type": "text",
                "text": sistema,
                "cache_control": {"type": "ephemeral"},
            }
        ],
        "messages": [{"role": "user", "content": contenido_usuario}],
        "output_config": {
            "effort": esfuerzo,
            "format": {"type": "json_schema", "schema": esquema},
        },
    }

    try:
        import anthropic
    except ImportError:
        raise ErrorClaude("Falta la libreria 'anthropic'. Instalala con pip.")

    respuesta = None
    # --- camino 1: beta con fallback del lado del servidor ---
    try:
        with cliente.beta.messages.stream(
            betas=[BETA_FALLBACK], fallbacks="default", **parametros
        ) as flujo:
            respuesta = flujo.get_final_message()
        log("Claude: llamada por el endpoint beta con fallbacks")
    except anthropic.BadRequestError as e:
        log(f"Claude: el endpoint beta rechazo el pedido ({e}); reintento sin betas")
        respuesta = None
    except Exception as e:
        raise _traducir_excepcion(e)

    # --- camino 2: endpoint estandar ---
    if respuesta is None:
        try:
            with cliente.messages.stream(**parametros) as flujo:
                respuesta = flujo.get_final_message()
            log("Claude: llamada por el endpoint estandar")
        except Exception as e:
            raise _traducir_excepcion(e)

    # --- rechazo por politica ---
    if getattr(respuesta, "stop_reason", "") == "refusal":
        detalle = getattr(respuesta, "stop_details", None)
        motivo = getattr(detalle, "explanation", "") if detalle else ""
        raise ErrorClaude(
            "Claude se nego a procesar el contenido"
            + (f": {motivo}" if motivo else ".")
            + "\nSi el archivo es un reporte de cobranzas normal, avisale al "
            "administrador del sistema."
        )

    # Respuesta cortada: el JSON queda incompleto, asi que no tiene sentido
    # seguir. Reintentar tampoco sirve: el problema es el tamano del reporte.
    if getattr(respuesta, "stop_reason", "") == "max_tokens":
        log_error("Claude corto la respuesta por max_tokens")
        raise ErrorClaude(
            "El reporte es demasiado grande: la respuesta se corto por la mitad "
            f"(limite de {MAX_TOKENS_STREAM:,} tokens de salida).\n\n"
            "Reintentar no lo soluciona. Que hacer:\n"
            "  - Exporta el reporte del ERP en partes (por ejemplo por rango de "
            "vendedor o de cliente) y carga una parte a la vez.\n"
            "  - Si el archivo es un Excel, usa 'Leer el archivo' en lugar de "
            "Claude: el lector propio no tiene este limite."
        )

    # --- extraer el JSON ---
    texto = ""
    for bloque in getattr(respuesta, "content", []) or []:
        if getattr(bloque, "type", "") == "text":
            texto = getattr(bloque, "text", "") or ""
            break

    if not texto.strip():
        raise ErrorClaude(
            "Claude no devolvio datos. Puede que el archivo este vacio o no sea "
            "un reporte de cuentas a cobrar."
        )

    try:
        datos = json.loads(texto)
    except json.JSONDecodeError as e:
        log_error(f"Claude devolvio un JSON invalido: {e} | inicio: {texto[:300]}")
        raise ErrorClaude(
            "Claude devolvio una respuesta que no se pudo interpretar. "
            "Volve a intentar."
        )

    uso = getattr(respuesta, "usage", None)
    if uso is not None:
        log(
            "Claude uso: entrada="
            f"{getattr(uso, 'input_tokens', 0)} "
            f"cache_lectura={getattr(uso, 'cache_read_input_tokens', 0)} "
            f"salida={getattr(uso, 'output_tokens', 0)}"
        )
    if not isinstance(datos, dict):
        raise ErrorClaude("Claude devolvio un formato inesperado.")
    return datos


# ==========================================================================
# CONVERSION JSON -> REPORTE
# ==========================================================================
def _a_fecha(texto: Any) -> date | None:
    """'10/08/2026' -> date. Tolera '-' y fechas ISO."""
    if isinstance(texto, (date, datetime)):
        return texto.date() if isinstance(texto, datetime) else texto
    s = str(texto or "").strip()
    if not s or s in ("-", "null", "None"):
        return None
    m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$", s)
    if m:
        d, mes, a = (int(x) for x in m.groups())
        if a < 100:
            a += 2000
        try:
            return date(a, mes, d)
        except ValueError:
            return None
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        a, mes, d = (int(x) for x in m.groups())
        try:
            return date(a, mes, d)
        except ValueError:
            return None
    return None


def _a_float(valor: Any) -> float | None:
    """Convierte a float aceptando tambien '37.165,39' por si acaso."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return float(valor)
    s = str(valor).strip().replace("$", "").replace(" ", "")
    if not s:
        return None
    negativo = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        n = float(s)
    except ValueError:
        return None
    return -n if negativo else n


def _a_int(valor: Any) -> int | None:
    f = _a_float(valor)
    return int(round(f)) if f is not None else None


def _a_reporte(datos: dict[str, Any], motor: str, origen: str) -> Reporte:
    """Convierte el JSON de Claude en objetos del modulo modelos.

    Nunca levanta excepcion por un dato faltante: acumula los problemas en
    reporte.avisos.
    """
    reporte = Reporte(origen_archivo=origen, motor=motor)
    reporte.empresa = str(datos.get("empresa", "") or "")
    reporte.titulo = str(datos.get("titulo", "") or "")
    reporte.fecha_reporte = _a_fecha(datos.get("fecha_reporte"))
    reporte.filtros = str(datos.get("filtros", "") or "")
    reporte.total_general_original = _a_float(datos.get("total_general_original"))
    reporte.total_general_saldo = _a_float(datos.get("total_general_saldo"))

    for aviso in datos.get("avisos") or []:
        texto = str(aviso or "").strip()
        if texto:
            reporte.avisos.append(f"Claude: {texto}")

    # Filtros de vendedor / cobrador
    m = re.search(r"Vendedor\s*:\s*(\d+)\s*-\s*(\d+)", reporte.filtros, re.IGNORECASE)
    if m:
        reporte.filtro_vendedor_desde, reporte.filtro_vendedor_hasta = m.group(1), m.group(2)
    m = re.search(r"Cobrador\s*:\s*(\d+)", reporte.filtros, re.IGNORECASE)
    if m:
        reporte.filtro_cobrador = m.group(1)

    area = str(CONFIG.get("area_telefono_defecto", "11"))
    crudos = datos.get("clientes")
    if not isinstance(crudos, list):
        reporte.avisos.append("Claude no devolvio la lista de clientes.")
        return reporte

    for i, cd in enumerate(crudos):
        if not isinstance(cd, dict):
            reporte.avisos.append(f"El cliente numero {i + 1} vino con formato invalido.")
            continue
        try:
            cli = Cliente(
                codigo=str(cd.get("codigo", "") or "").strip(),
                razon_social=str(cd.get("razon_social", "") or "").strip(),
                nombre_fantasia=str(cd.get("nombre_fantasia", "") or "").strip(),
                direccion=str(cd.get("direccion", "") or "").strip(),
                localidad=str(cd.get("localidad", "") or "").strip(),
                contacto=str(cd.get("contacto", "") or "").strip(),
                telefonos_raw=str(cd.get("telefonos_raw", "") or "").strip(),
                total_original_reporte=_a_float(cd.get("total_original_reporte")),
                total_saldo_reporte=_a_float(cd.get("total_saldo_reporte")),
                resaltado=bool(cd.get("resaltado", False)),
            )
            cli.telefonos = telefonos_de_texto(cli.telefonos_raw, area)

            for j, kd in enumerate(cd.get("comprobantes") or []):
                if not isinstance(kd, dict):
                    continue
                vto = _a_fecha(kd.get("fecha_vto"))
                dias = _a_int(kd.get("dias_venc"))

                # Los dias SIEMPRE se recalculan desde las fechas cuando estan
                # las dos. Las fechas son verificables una por una; el numero de
                # dias es un dato derivado, y si Claude se equivoca en el SIGNO
                # un cliente que no debe nada aparece como vencido y el vendedor
                # sale a reclamar una deuda que no existe.
                if vto and reporte.fecha_reporte:
                    calculado = (reporte.fecha_reporte - vto).days
                    if dias is not None and abs(calculado - dias) > 1:
                        reporte.avisos.append(
                            f"{cli.nombre or cli.codigo}, comprobante "
                            f"{kd.get('tipo', '')} {kd.get('numero', '')}: Claude "
                            f"informo {dias} dias pero por las fechas "
                            f"(vto {vto.strftime('%d/%m/%Y')} contra reporte del "
                            f"{reporte.fecha_reporte.strftime('%d/%m/%Y')}) "
                            f"corresponden {calculado}. Se usaron {calculado}."
                        )
                    dias = calculado
                original = _a_float(kd.get("importe_original")) or 0.0
                saldo = _a_float(kd.get("saldo"))
                if saldo is None:
                    saldo = original
                    reporte.avisos.append(
                        f"{cli.nombre or cli.codigo}: un comprobante no traia saldo; "
                        "se uso el importe original."
                    )
                cli.comprobantes.append(
                    Comprobante(
                        talonario=str(kd.get("talonario", "") or "").strip(),
                        tipo=str(kd.get("tipo", "") or "").strip(),
                        numero=str(kd.get("numero", "") or "").strip(),
                        condicion_pago=str(kd.get("condicion_pago", "") or "").strip(),
                        fecha_emision=_a_fecha(kd.get("fecha_emision")),
                        fecha_vto=vto,
                        cuota=str(kd.get("cuota", "") or "").strip(),
                        importe_original=round(original, 2),
                        saldo=round(saldo, 2),
                        dias_venc=dias,
                        resaltado=bool(kd.get("resaltado", False)),
                    )
                )

            if not cli.comprobantes and cli.total_saldo_reporte:
                reporte.avisos.append(
                    f"{cli.nombre or cli.codigo}: se leyo el total "
                    "pero no los comprobantes."
                )
            reporte.clientes.append(cli)
        except Exception as e:
            log_error(f"Error convirtiendo el cliente {i + 1}: {e}")
            reporte.avisos.append(
                f"No se pudo interpretar el cliente numero {i + 1}: {e}"
            )

    return reporte


# ==========================================================================
# INTERPRETAR UN PDF
# ==========================================================================
def interpretar_pdf(
    ruta_pdf: str,
    *,
    fecha_referencia: date | None = None,
    progreso: Callable[[str], None] | None = None,
) -> Reporte:
    """Manda el PDF a Claude y devuelve el Reporte interpretado."""
    if not ruta_pdf or not os.path.exists(ruta_pdf):
        raise ErrorClaude(f"No se encuentra el archivo: {ruta_pdf}")

    tamanio = os.path.getsize(ruta_pdf)
    if tamanio == 0:
        raise ErrorClaude("El PDF esta vacio.")
    if tamanio > 30 * 1024 * 1024:
        raise ErrorClaude(
            "El PDF pesa mas de 30 MB, que es el limite de la API. "
            "Proba exportar el reporte en partes."
        )

    _avisar(progreso, "Leyendo el PDF...")
    try:
        with open(ruta_pdf, "rb") as fh:
            datos_b64 = base64.standard_b64encode(fh.read()).decode("ascii")
    except OSError as e:
        raise ErrorClaude(f"No se pudo leer el PDF: {e}")

    instruccion = (
        "Interpreta este reporte de cuentas a cobrar y devolvelo como datos "
        "estructurados, siguiendo exactamente las reglas del sistema."
    )
    if fecha_referencia:
        instruccion += (
            f"\n\nSi el documento no indica la fecha del reporte, usa "
            f"{fecha_referencia.strftime('%d/%m/%Y')} para calcular los dias "
            "de vencimiento."
        )

    contenido = [
        {
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": "application/pdf",
                "data": datos_b64,
            },
        },
        {"type": "text", "text": instruccion},
    ]

    _avisar(progreso, "Claude esta interpretando el PDF (puede tardar un minuto)...")
    datos = _llamar(PROMPT_SISTEMA_EXTRACCION, contenido, ESQUEMA_REPORTE, progreso)

    _avisar(progreso, "Armando el listado de clientes...")
    reporte = _a_reporte(datos, "claude-pdf", ruta_pdf)
    if fecha_referencia and reporte.fecha_reporte is None:
        reporte.fecha_reporte = fecha_referencia
    log(
        f"Claude interpreto el PDF {os.path.basename(ruta_pdf)}: "
        f"{len(reporte.clientes)} clientes"
    )
    return reporte


# ==========================================================================
# INTERPRETAR UN VOLCADO DE EXCEL
# ==========================================================================
def interpretar_volcado_excel(
    volcado: str,
    *,
    nombre_archivo: str = "",
    progreso: Callable[[str], None] | None = None,
) -> Reporte:
    """Interpreta el texto que produce parser_excel.volcar_para_claude()."""
    if not (volcado or "").strip():
        raise ErrorClaude("El volcado del Excel esta vacio.")

    # Un recorte silencioso hace desaparecer clientes sin que nadie se entere,
    # asi que se avisa arriba de todo en el reporte resultante.
    aviso_recorte = ""
    if len(volcado) > 1_500_000:
        filas_totales = volcado.count("\n") + 1
        volcado = volcado[:1_500_000]
        filas_enviadas = volcado.count("\n") + 1
        aviso_recorte = (
            f"ATENCION: el archivo es muy grande y se envio incompleto a Claude "
            f"(se mandaron {filas_enviadas:,} de {filas_totales:,} filas). "
            f"FALTAN CLIENTES en este resultado. Exporta el reporte en partes."
        )
        log_error(aviso_recorte)

    instruccion = (
        "Abajo esta el volcado celda por celda de un Excel con el reporte de "
        "cuentas a cobrar. El formato de cada linea es:\n"
        "  F<numero de fila> | <letra de columna>=<contenido> | ...\n"
        "Las celdas que la persona de administracion marco a mano vienen con el "
        "sufijo [AMARILLO].\n\n"
        "Interpretalo y devolvelo como datos estructurados.\n\n"
        f"{volcado}"
    )
    contenido = [{"type": "text", "text": instruccion}]

    _avisar(progreso, "Claude esta interpretando el Excel...")
    datos = _llamar(PROMPT_SISTEMA_EXTRACCION, contenido, ESQUEMA_REPORTE, progreso)
    reporte = _a_reporte(datos, "claude-texto", nombre_archivo)
    if aviso_recorte:
        reporte.avisos.insert(0, aviso_recorte)
    log(f"Claude interpreto el Excel: {len(reporte.clientes)} clientes")
    return reporte


# ==========================================================================
# AUDITORIA DE UN REPORTE YA PARSEADO
# ==========================================================================
def verificar_reporte(
    reporte: Reporte,
    volcado: str,
    progreso: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    """Le pide a Claude auditar lo que leyo el parser propio contra el original.

    Devuelve {"coincide": bool, "observaciones": [...], "clientes_faltantes":
    [...], "diferencias": [...]}. No modifica el reporte.
    """
    resumen: list[str] = []
    for cli in reporte.clientes:
        resumen.append(
            f"- codigo={cli.codigo} nombre={cli.nombre!r} "
            f"comprobantes={cli.cantidad_comprobantes} "
            f"suma_saldos={cli.saldo_total:.2f} "
            f"total_reporte={cli.total_saldo_reporte if cli.total_saldo_reporte is not None else 'null'} "
            f"resaltado={cli.resaltado}"
        )

    # Si se recorta el documento, la auditoria solo vio una parte: decirlo, o un
    # "coincide" da falsa tranquilidad.
    aviso_recorte = ""
    if len(volcado) > 800_000:
        volcado = volcado[:800_000]
        aviso_recorte = (
            "La auditoria se hizo sobre una PARTE del documento (es muy grande y "
            "se recorto). El resultado no cubre todo el reporte."
        )
        log_error(aviso_recorte)

    instruccion = (
        "Tengo un programa propio que ya interpreto este reporte. Necesito que "
        "audites su resultado contra el documento original y me digas si esta "
        "fiel.\n\n"
        "=== LO QUE LEYO MI PROGRAMA ===\n"
        f"Fecha del reporte: {reporte.fecha_reporte}\n"
        f"Clientes detectados: {len(reporte.clientes)}\n"
        f"Total general segun el documento: {reporte.total_general_saldo}\n"
        f"Suma de los saldos que calculo mi programa: {reporte.saldo_total:.2f}\n"
        + "\n".join(resumen)
        + "\n\n=== DOCUMENTO ORIGINAL ===\n"
        + volcado
        + "\n\nRevisa: (1) si falta algun cliente, (2) si algun importe o "
        "cantidad de comprobantes no coincide, (3) si el resaltado amarillo se "
        "detecto bien. Se concreto y cita codigos e importes. Si todo esta bien, "
        "devolve coincide=true y los arrays vacios."
    )

    _avisar(progreso, "Claude esta auditando los datos...")
    datos = _llamar(
        PROMPT_SISTEMA_AUDITORIA,
        [{"type": "text", "text": instruccion}],
        ESQUEMA_VERIFICACION,
        progreso,
    )
    observaciones = [str(x) for x in (datos.get("observaciones") or [])]
    coincide = bool(datos.get("coincide", False))
    if aviso_recorte:
        observaciones.insert(0, aviso_recorte)
        coincide = False  # no se puede afirmar que coincide si no se vio todo
    return {
        "coincide": coincide,
        "observaciones": observaciones,
        "clientes_faltantes": [str(x) for x in (datos.get("clientes_faltantes") or [])],
        "diferencias": [str(x) for x in (datos.get("diferencias") or [])],
    }


# ==========================================================================
# NORMALIZACION DE TELEFONOS
# ==========================================================================
def normalizar_telefonos(
    reporte: Reporte,
    progreso: Callable[[str], None] | None = None,
) -> int:
    """Separa contactos y numeros de la columna Telefonos, en una sola llamada.

    Escribe cliente.telefonos (moviles primero) y completa cliente.contacto.
    Devuelve cuantos clientes se actualizaron.
    """
    pendientes = [c for c in reporte.clientes if (c.telefonos_raw or "").strip()]
    if not pendientes:
        return 0

    # Se procesa en lotes: con una cartera de cientos de clientes, una sola
    # llamada corre riesgo de cortarse por max_tokens y perder la mitad.
    por_codigo: dict[str, dict[str, Any]] = {}
    total_lotes = (len(pendientes) + TAMANIO_LOTE_TELEFONOS - 1) // TAMANIO_LOTE_TELEFONOS

    for nro, inicio in enumerate(
        range(0, len(pendientes), TAMANIO_LOTE_TELEFONOS), start=1
    ):
        lote = pendientes[inicio : inicio + TAMANIO_LOTE_TELEFONOS]
        lineas = [
            f"- codigo={c.codigo} | cliente={c.nombre!r} | telefonos_raw={c.telefonos_raw!r}"
            for c in lote
        ]
        instruccion = (
            "Ordena los datos de contacto de estos clientes. Devolve una entrada "
            "por cada codigo que te paso, con el mismo codigo tal cual.\n\n"
            + "\n".join(lineas)
        )
        etiqueta = (
            f"Claude esta ordenando los telefonos ({len(lote)} clientes"
            + (f", lote {nro} de {total_lotes}" if total_lotes > 1 else "")
            + ")..."
        )
        _avisar(progreso, etiqueta)
        datos = _llamar(
            PROMPT_SISTEMA_TELEFONOS,
            [{"type": "text", "text": instruccion}],
            ESQUEMA_TELEFONOS,
            progreso,
        )
        for item in datos.get("clientes") or []:
            if isinstance(item, dict):
                cod = str(item.get("codigo", "") or "").strip().lstrip("0")
                if cod:
                    por_codigo[cod] = item

    actualizados = 0
    for cli in pendientes:
        item = por_codigo.get(cli.codigo.lstrip("0"))
        if not item:
            continue

        numeros: list[str] = []
        contactos_por_numero: list[str] = []
        for tel in item.get("telefonos") or []:
            if not isinstance(tel, dict):
                continue
            numero = "".join(ch for ch in str(tel.get("numero", "")) if ch.isdigit())
            if len(numero) < 12 or len(numero) > 15:
                continue
            if numero not in numeros:
                numeros.append(numero)
                contactos_por_numero.append(str(tel.get("contacto", "") or "").strip())

        if numeros:
            cli.telefonos = numeros
            actualizados += 1

        nombres = [str(x).strip() for x in (item.get("contactos") or []) if str(x).strip()]
        nombres += [c for c in contactos_por_numero if c]
        if nombres and not cli.contacto:
            cli.contacto = " / ".join(dict.fromkeys(nombres))

    log(f"Claude normalizo los telefonos de {actualizados} clientes")
    return actualizados


# ==========================================================================
# REDACCION ASISTIDA DEL MENSAJE
# ==========================================================================
_ESQUEMA_MENSAJE: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "required": ["mensaje"],
    "properties": {
        "mensaje": {
            "type": "string",
            "description": "El texto listo para pegar en WhatsApp",
        }
    },
}

PROMPT_SISTEMA_MENSAJE = """\
Escribis mensajes de WhatsApp para el area de administracion de Wood Tools SRL, \
una empresa argentina que vende herramientas para carpinteria.

El destinatario es un VENDEDOR de la empresa, no el cliente deudor. El mensaje \
le informa que clientes de su cartera tienen saldo vencido o por vencer, para \
que los llame.

Pautas:
- Espanol rioplatense, tuteo argentino ("vos", "tenes", "fijate").
- Tono cordial y profesional, de companero de trabajo. Sin sonar a robot ni a \
intimacion legal.
- Breve y escaneable. El vendedor lo lee del celular entre visitas.
- Usa el markdown de WhatsApp: *negrita* para lo importante.
- Como maximo dos emojis en todo el mensaje, y solo si suman.
- Los importes van en formato argentino con el signo pesos, por ejemplo \
$173.186,42. No cambies ni redondees ningun numero que te pasen.
- No inventes datos que no esten en la informacion que te doy.
"""


def redactar_mensaje(
    resumen_datos: dict[str, Any],
    progreso: Callable[[str], None] | None = None,
) -> str:
    """Le pide a Claude redactar el texto del WhatsApp. Devuelve el mensaje."""
    _avisar(progreso, "Claude esta redactando el mensaje...")
    instruccion = (
        "Redacta el mensaje de WhatsApp con estos datos:\n\n"
        + json.dumps(resumen_datos, ensure_ascii=False, indent=2)
    )
    datos = _llamar(
        PROMPT_SISTEMA_MENSAJE,
        [{"type": "text", "text": instruccion}],
        _ESQUEMA_MENSAJE,
        progreso,
    )
    return str(datos.get("mensaje", "") or "").strip()


# ==========================================================================
# ESTIMACION DE COSTO
# ==========================================================================
def estimar_costo(
    texto_o_ruta: str, *, es_pdf: bool = False, salida_estimada: int = 8000
) -> str:
    """Estima el costo de interpretar un archivo. Devuelve un texto legible.

    Levanta ErrorClaude si la API no responde, para que el motivo real (saldo,
    clave invalida, sin internet) llegue al usuario en lugar de un "no se pudo".
    """
    cliente = _cliente()
    modelo = str(CONFIG.get("modelo_claude", "claude-opus-5"))
    try:
        if es_pdf:
            with open(texto_o_ruta, "rb") as fh:
                b64 = base64.standard_b64encode(fh.read()).decode("ascii")
            contenido: list[dict[str, Any]] = [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": b64,
                    },
                },
                {"type": "text", "text": "Interpreta este reporte."},
            ]
        else:
            contenido = [{"type": "text", "text": texto_o_ruta}]

        resultado = cliente.messages.count_tokens(
            model=modelo,
            system=PROMPT_SISTEMA_EXTRACCION,
            messages=[{"role": "user", "content": contenido}],
        )
        entrada = int(getattr(resultado, "input_tokens", 0))
    except OSError as e:
        raise ErrorClaude(f"No se pudo leer el archivo para estimar el costo: {e}")
    except Exception as e:
        log_error(f"No se pudo contar tokens: {e}")
        raise _traducir_excepcion(e)

    costo = (entrada / 1_000_000) * PRECIO_ENTRADA + (
        salida_estimada / 1_000_000
    ) * PRECIO_SALIDA
    entrada_txt = f"{entrada:,}".replace(",", ".")
    costo_txt = f"{costo:.2f}".replace(".", ",")
    return f"~{entrada_txt} tokens de entrada (aprox. USD {costo_txt})"
