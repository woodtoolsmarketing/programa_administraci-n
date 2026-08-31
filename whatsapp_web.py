# -*- coding: utf-8 -*-
"""
Envio ASISTIDO por WhatsApp Web (el de escanear el QR).

A diferencia de whatsapp.py, que habla con la Cloud API de Meta, este modulo
usa la sesion normal de WhatsApp: la misma que se abre en web.whatsapp.com
escaneando el codigo QR con el telefono. No hay token, no hay plantillas que
aprobar, no se paga por conversacion y -lo mas importante- NO existe la
ventana de 24 horas: se le puede escribir a cualquiera, cuando sea.

POR QUE ES ASISTIDO Y NO AUTOMATICO
El programa abre el chat del vendedor, deja el mensaje escrito y el PDF
adjunto, y ahi se detiene. El boton Enviar lo aprieta la persona. Automatizar
tambien el envio va contra los terminos de servicio de WhatsApp y expone el
numero de la empresa a un bloqueo. Preparar el mensaje y que un humano lo
mande es la parte que ahorra el trabajo sin correr ese riesgo.

LA SESION SE GUARDA
Chrome arranca con una carpeta de perfil propia del programa (PerfilChrome,
al lado del ejecutable), asi que el QR se escanea una sola vez y la sesion
sobrevive a cerrar y volver a abrir.

FRAGILIDAD
WhatsApp Web cambia su HTML cada tanto y eso rompe los selectores. Por eso
cada elemento se busca con varias alternativas y, cuando ninguna funciona, el
error dice exactamente en que paso se corto para poder terminarlo a mano.
"""

from __future__ import annotations

import os
import time
import urllib.parse
from dataclasses import dataclass
from typing import Any, Callable

from config import log, log_error, ruta_persistente

URL_BASE = "https://web.whatsapp.com"
CARPETA_PERFIL = ruta_persistente("PerfilChrome")

ESPERA_CARGA = 60        # segundos para que cargue WhatsApp Web
ESPERA_QR = 180          # segundos que se le dan a la persona para escanear
ESPERA_ELEMENTO = 25     # segundos para encontrar un elemento de la pagina


class ErrorWhatsAppWeb(Exception):
    """Problema manejando WhatsApp Web, con el mensaje ya legible."""


# ==========================================================================
# DISPONIBILIDAD
# ==========================================================================
def disponible() -> tuple[bool, str]:
    """(se puede usar, motivo si no). No abre nada: solo revisa el entorno."""
    try:
        import selenium  # noqa: F401
    except ImportError:
        return False, (
            "Falta la libreria 'selenium', que es la que maneja el navegador.\n"
            "Instalala con:  pip install selenium\n"
            "(o volve a correr INSTALAR.bat)"
        )
    if not _ruta_chrome():
        return False, (
            "No se encontro Google Chrome instalado.\n"
            "WhatsApp Web asistido necesita Chrome. Instalalo desde "
            "google.com/chrome y volve a intentar."
        )
    return True, ""


def _ruta_chrome() -> str:
    for base in (
        os.environ.get("PROGRAMFILES", r"C:\Program Files"),
        os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), ""),
    ):
        ruta = os.path.join(base, "Google", "Chrome", "Application", "chrome.exe")
        if os.path.isfile(ruta):
            return ruta
    return ""


def _acepta_documentos(acepta: str) -> bool:
    """True si ese input[type=file] sirve para mandar un PDF.

    El input de fotos declara accept="image/*": si se le pasa un PDF, WhatsApp
    lo descarta sin decir nada. El de documentos declara accept="*".
    """
    a = (acepta or "").strip().lower()
    if not a or a in ("*", "*/*", "*.*"):
        return True
    if "image" in a or "video" in a or "audio" in a:
        return False
    return "pdf" in a or "application" in a


# ==========================================================================
# SELECTORES (varios por elemento, porque WhatsApp Web los cambia)
# ==========================================================================
def _selectores() -> dict[str, list[tuple[Any, str]]]:
    from selenium.webdriver.common.by import By

    return {
        # el QR de vinculacion
        "qr": [
            (By.CSS_SELECTOR, "canvas[aria-label*='scan' i]"),
            (By.CSS_SELECTOR, "div[data-ref]"),
            (By.XPATH, "//canvas[ancestor::div[contains(@class,'landing')]]"),
        ],
        # la lista de chats: si esta, la sesion ya esta iniciada
        "sesion_lista": [
            (By.ID, "pane-side"),
            (By.CSS_SELECTOR, "div[aria-label='Lista de chats']"),
            (By.CSS_SELECTOR, "div[data-testid='chat-list']"),
        ],
        # la caja donde se escribe el mensaje
        "caja_texto": [
            (By.CSS_SELECTOR, "footer div[contenteditable='true']"),
            (By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='10']"),
            (By.XPATH, "//footer//div[@role='textbox']"),
        ],
        # el input al que se le pasa la ruta del archivo
        "input_archivo": [
            (By.CSS_SELECTOR, "input[type='file']"),
        ],
        # el boton de adjuntar, por si hay que abrirlo para que aparezca el input
        "adjuntar": [
            (By.CSS_SELECTOR, "button[title='Adjuntar']"),
            (By.CSS_SELECTOR, "button[aria-label='Adjuntar']"),
            (By.CSS_SELECTOR, "div[title='Adjuntar']"),
            (By.CSS_SELECTOR, "button[title='Attach']"),
            (By.CSS_SELECTOR, "button[aria-label='Attach']"),
            (By.CSS_SELECTOR, "span[data-icon='ic-attach-file']"),
            (By.CSS_SELECTOR, "span[data-icon='plus-rounded']"),
            (By.CSS_SELECTOR, "span[data-icon='attach-menu-plus']"),
            (By.CSS_SELECTOR, "span[data-icon='plus']"),
            (By.CSS_SELECTOR, "span[data-icon='clip']"),
            (By.CSS_SELECTOR, "span[data-icon='attach']"),
        ],
        # La opcion "Documento" del menu de adjuntar. Es la que importa: recien
        # al apretarla WhatsApp crea el input[type=file] con accept="*". Antes
        # de eso el unico input de la pagina es el de fotos (accept="image/*"),
        # que descarta los PDF sin avisar.
        "menu_documento": [
            (By.CSS_SELECTOR, "button[role='menuitem'][aria-label='Documento']"),
            (By.CSS_SELECTOR, "button[role='menuitem'][aria-label='Document']"),
            (By.XPATH, "//*[@role='menuitem'][contains(@aria-label,'ocument')]"),
            (By.XPATH, "//*[@role='menuitem'][contains(., 'ocument')]"),
            (By.XPATH, "//*[normalize-space(text())='Documento']"),
        ],
        # Con el archivo ya cargado, la vista previa muestra estos controles.
        # Sirven para confirmar que el adjunto entro de verdad: sin esto, un
        # fallo silencioso terminaria mandando el texto sin el PDF.
        "adjunto_listo": [
            (By.CSS_SELECTOR, "[aria-label='Quitar archivo adjunto']"),
            (By.CSS_SELECTOR, "[aria-label='Remove attachment']"),
            (By.CSS_SELECTOR, "span[data-icon='wds-ic-send-filled']"),
            (By.XPATH, "//*[contains(@aria-label,'Enviar')]"),
            (By.XPATH, "//*[contains(@aria-label,'Send')]"),
        ],
        # la caja de epigrafe del documento adjunto (otra que la del chat)
        "caja_epigrafe": [
            (By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='undefined']"),
            (By.XPATH, "//div[@role='textbox'][@contenteditable='true']"),
        ],
        # el cartel de numero inexistente
        "numero_invalido": [
            (By.XPATH, "//*[contains(text(),'no esta en WhatsApp')]"),
            (By.XPATH, "//*[contains(text(),'no está en WhatsApp')]"),
            (By.XPATH, "//*[contains(text(),'is not on WhatsApp')]"),
        ],
    }


# ==========================================================================
# UN PASO DE LA COLA DE ENVIO
# ==========================================================================
@dataclass
class Paso:
    """Un mensaje listo para que la persona apriete Enviar."""

    vendedor: str
    telefono: str
    tipo: str            # "documento" | "texto"
    descripcion: str     # lo que se le muestra al usuario
    texto: str = ""
    ruta_pdf: str = ""
    epigrafe: str = ""


def armar_pasos(
    vendedor: str,
    telefono: str,
    partes_texto: list[str],
    ruta_pdf: str | None = None,
    epigrafe: str = "",
) -> list[Paso]:
    """Arma la cola de mensajes de un vendedor: primero el PDF, despues el texto.

    Es el mismo orden que usa el envio por la Cloud API, para que el vendedor
    reciba siempre lo mismo por los dos caminos.
    """
    pasos: list[Paso] = []
    if ruta_pdf:
        pasos.append(
            Paso(
                vendedor=vendedor,
                telefono=telefono,
                tipo="documento",
                descripcion=f"PDF del resumen ({os.path.basename(ruta_pdf)})",
                ruta_pdf=ruta_pdf,
                epigrafe=epigrafe,
            )
        )
    total = len(partes_texto)
    for i, parte in enumerate(partes_texto, start=1):
        etiqueta = "Mensaje" if total == 1 else f"Mensaje {i} de {total}"
        pasos.append(
            Paso(
                vendedor=vendedor,
                telefono=telefono,
                tipo="texto",
                descripcion=f"{etiqueta} ({len(parte)} caracteres)",
                texto=parte,
            )
        )
    return pasos


def url_chat(telefono: str, texto: str = "") -> str:
    """URL de WhatsApp Web que abre el chat con el texto ya cargado.

    Se usa el parametro 'text' en vez de tipear: asi los saltos de linea se
    respetan sin tocar el teclado (en WhatsApp Web, Enter manda el mensaje).
    """
    solo = "".join(ch for ch in str(telefono or "") if ch.isdigit())
    base = f"{URL_BASE}/send?phone={solo}"
    if texto:
        base += "&text=" + urllib.parse.quote(texto, safe="")
    return base


# ==========================================================================
# SESION
# ==========================================================================
class Sesion:
    """Maneja el Chrome con la sesion de WhatsApp Web."""

    def __init__(self) -> None:
        self.driver: Any = None

    # ---------------- apertura ----------------
    def abrir(self, progreso: Callable[[str], None] | None = None) -> None:
        ok, motivo = disponible()
        if not ok:
            raise ErrorWhatsAppWeb(motivo)

        from selenium import webdriver
        from selenium.common.exceptions import WebDriverException

        self._avisar(progreso, "Abriendo Chrome...")
        try:
            os.makedirs(CARPETA_PERFIL, exist_ok=True)
        except OSError as e:
            raise ErrorWhatsAppWeb(f"No se pudo crear la carpeta del perfil: {e}")

        opciones = webdriver.ChromeOptions()
        # Perfil propio del programa: la sesion sobrevive a cerrar Chrome, y no
        # se mezcla con el Chrome personal del usuario.
        opciones.add_argument(f"--user-data-dir={CARPETA_PERFIL}")
        opciones.add_argument("--profile-directory=Default")
        opciones.add_argument("--start-maximized")
        opciones.add_experimental_option("excludeSwitches", ["enable-automation"])
        opciones.add_experimental_option("useAutomationExtension", False)

        try:
            self.driver = webdriver.Chrome(options=opciones)
        except WebDriverException as e:
            detalle = str(e)
            if "user data directory is already in use" in detalle.lower():
                raise ErrorWhatsAppWeb(
                    "Ya hay una ventana de Chrome abierta con el perfil del "
                    "programa.\nCerrala y volve a intentar."
                )
            raise ErrorWhatsAppWeb(
                "No se pudo abrir Chrome.\n\n"
                "Si es la primera vez, el programa necesita internet para bajar "
                "el controlador del navegador.\n\n"
                f"Detalle: {detalle[:300]}"
            )

        self._avisar(progreso, "Cargando WhatsApp Web...")
        try:
            self.driver.get(URL_BASE)
        except WebDriverException as e:
            raise ErrorWhatsAppWeb(f"No se pudo abrir WhatsApp Web: {e}")
        log("WhatsApp Web: Chrome abierto con el perfil del programa")

    # ---------------- sesion iniciada ----------------
    def hay_sesion(self) -> bool:
        """True si ya estan los chats a la vista (no hace falta escanear)."""
        return self._buscar("sesion_lista", espera=3) is not None

    def hay_qr(self) -> bool:
        return self._buscar("qr", espera=3) is not None

    def esperar_sesion(
        self, progreso: Callable[[str], None] | None = None, espera: int = ESPERA_QR
    ) -> None:
        """Espera a que la sesion este lista, dando tiempo a escanear el QR."""
        # Restaurar una sesion guardada tarda entre 5 y 10 segundos. Si se
        # pregunta antes, parece que no hay sesion y se le pide el QR a alguien
        # que no lo necesita.
        self._avisar(progreso, "Buscando la sesion guardada...")
        if self._buscar("sesion_lista", espera=25) is not None:
            self._avisar(progreso, "Sesion de WhatsApp Web ya iniciada.")
            return

        self._avisar(
            progreso,
            "Escanea el codigo QR con el telefono (WhatsApp > Dispositivos "
            "vinculados). La sesion queda guardada para la proxima.",
        )
        limite = time.time() + espera
        while time.time() < limite:
            if self._buscar("sesion_lista", espera=2) is not None:
                self._avisar(progreso, "Sesion iniciada.")
                log("WhatsApp Web: sesion iniciada")
                return
            time.sleep(1)
        raise ErrorWhatsAppWeb(
            "Se agoto el tiempo esperando el escaneo del codigo QR.\n"
            "Volve a intentar y escanealo con el telefono."
        )

    # ---------------- preparar un mensaje ----------------
    def preparar(self, paso: Paso, progreso: Callable[[str], None] | None = None) -> None:
        """Deja el mensaje listo en pantalla. NO lo envia."""
        if self.driver is None:
            raise ErrorWhatsAppWeb("El navegador no esta abierto.")
        if paso.tipo == "documento":
            self._preparar_documento(paso, progreso)
        else:
            self._preparar_texto(paso, progreso)

    def _preparar_texto(self, paso: Paso, progreso) -> None:
        self._avisar(progreso, f"Abriendo el chat de {paso.vendedor}...")
        self.driver.get(url_chat(paso.telefono, paso.texto))
        self._revisar_numero()
        if self._buscar("caja_texto", espera=ESPERA_ELEMENTO) is None:
            raise ErrorWhatsAppWeb(
                f"No se pudo abrir el chat de {paso.vendedor} ({paso.telefono}).\n"
                "Puede que el numero no tenga WhatsApp, o que WhatsApp Web haya "
                "cambiado y el programa no reconozca la pantalla.\n\n"
                "El mensaje lo podes mandar a mano: esta copiado en la vista previa."
            )
        log(f"WhatsApp Web: texto listo para {paso.vendedor}")

    def _preparar_documento(self, paso: Paso, progreso) -> None:
        if not paso.ruta_pdf or not os.path.isfile(paso.ruta_pdf):
            raise ErrorWhatsAppWeb(f"No se encuentra el PDF: {paso.ruta_pdf}")

        self._avisar(progreso, f"Abriendo el chat de {paso.vendedor}...")
        self.driver.get(url_chat(paso.telefono))
        self._revisar_numero()
        if self._buscar("caja_texto", espera=ESPERA_ELEMENTO) is None:
            raise ErrorWhatsAppWeb(
                f"No se pudo abrir el chat de {paso.vendedor} ({paso.telefono})."
            )

        self._avisar(progreso, "Adjuntando el PDF...")
        entrada = self._input_de_archivo()
        if entrada is None:
            raise ErrorWhatsAppWeb(
                "No se encontro donde adjuntar el archivo. WhatsApp Web pudo "
                "haber cambiado.\n\n"
                f"Adjuntalo a mano desde:\n{paso.ruta_pdf}"
            )
        try:
            entrada.send_keys(os.path.abspath(paso.ruta_pdf))
        except Exception as e:
            raise ErrorWhatsAppWeb(
                f"No se pudo adjuntar el PDF: {e}\n\nAdjuntalo a mano desde:\n"
                f"{paso.ruta_pdf}"
            )

        # No alcanza con que send_keys no falle: hay que ver el adjunto en la
        # vista previa. Si no aparece, el texto se mandaria sin el PDF.
        if self._buscar("adjunto_listo", espera=ESPERA_ELEMENTO) is None:
            raise ErrorWhatsAppWeb(
                "El PDF no llego a la vista previa de WhatsApp.\n\n"
                f"Adjuntalo a mano desde:\n{paso.ruta_pdf}"
            )

        if paso.epigrafe:
            self._escribir_epigrafe(paso.epigrafe)
        log(f"WhatsApp Web: PDF adjuntado para {paso.vendedor}")

    def _input_de_archivo(self) -> Any:
        """El input[type=file] al que se le pasa la ruta del PDF.

        El camino, medido contra WhatsApp Web:

          1. La pagina arranca con UN solo input, y es el de fotos
             (accept="image/*"). Pasarle un PDF no da error: WhatsApp lo
             descarta en silencio y el adjunto nunca aparece.
          2. Hay que apretar "Adjuntar" con un click DE VERDAD. Los clicks
             hechos por JavaScript no abren el menu.
          3. Y despues apretar "Documento". Recien ahi WhatsApp agrega el
             segundo input, con accept="*", que es el que sirve.
        """
        entrada = self._input_documentos()
        if entrada is not None:
            return entrada

        boton = self._buscar("adjuntar", espera=8)
        if boton is None:
            log_error("WhatsApp Web: no se encontro el boton de adjuntar")
            return None
        # si el menu ya quedo abierto de antes, volver a apretar lo cerraria
        if (boton.get_attribute("aria-expanded") or "").lower() != "true":
            if not self._click_real(boton):
                log_error("WhatsApp Web: no se pudo abrir el menu de adjuntar")
                return None
            time.sleep(1.2)

        opcion = self._buscar("menu_documento", espera=8)
        if opcion is None:
            log_error("WhatsApp Web: no aparecio la opcion 'Documento' del menu")
            return None
        if not self._click_real(opcion):
            log_error("WhatsApp Web: no se pudo apretar 'Documento'")
            return None

        # el input tarda un momento en aparecer
        for _ in range(12):
            entrada = self._input_documentos()
            if entrada is not None:
                log("WhatsApp Web: input de documentos listo")
                return entrada
            time.sleep(0.5)
        log_error("WhatsApp Web: 'Documento' no creo el input de archivos")
        return None

    def _input_documentos(self) -> Any:
        """El input que acepta documentos. NUNCA devuelve el de fotos."""
        from selenium.webdriver.common.by import By

        entradas = self.driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
        acepta_de = [(e, (e.get_attribute("accept") or "").strip()) for e in entradas]
        if acepta_de:
            log("WhatsApp Web: inputs presentes -> " + str([a for _, a in acepta_de]))
        for entrada, acepta in acepta_de:
            if _acepta_documentos(acepta):
                return entrada
        return None

    def _click_real(self, elemento: Any) -> bool:
        """Click de verdad. WhatsApp Web ignora los clicks hechos por JavaScript."""
        from selenium.webdriver.common.action_chains import ActionChains

        try:
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", elemento
            )
        except Exception:
            pass
        for estrategia in ("directo", "acciones", "javascript"):
            try:
                if estrategia == "directo":
                    elemento.click()
                elif estrategia == "acciones":
                    ActionChains(self.driver).move_to_element(elemento).pause(
                        0.3
                    ).click().perform()
                else:
                    self.driver.execute_script("arguments[0].click();", elemento)
                return True
            except Exception as e:
                log(f"WhatsApp Web: click {estrategia} fallo: {str(e)[:110]}")
        return False

    def _escribir_epigrafe(self, epigrafe: str) -> None:
        """Escribe el epigrafe del documento. Si falla, no se corta el envio."""
        caja = self._buscar("caja_epigrafe", espera=ESPERA_ELEMENTO)
        if caja is None:
            log_error("WhatsApp Web: no se encontro la caja de epigrafe")
            return
        try:
            caja.click()
            # sin saltos de linea: en la vista previa, Enter tambien envia
            caja.send_keys(epigrafe.replace("\n", " ").replace("\r", " ")[:900])
        except Exception as e:
            log_error(f"WhatsApp Web: no se pudo escribir el epigrafe: {e}")

    def _revisar_numero(self) -> None:
        if self._buscar("numero_invalido", espera=3) is not None:
            raise ErrorWhatsAppWeb(
                "WhatsApp dice que ese numero no tiene cuenta.\n"
                "Revisa el telefono del vendedor en la pestana Vendedores."
            )

    # ---------------- cierre ----------------
    def cerrar(self) -> None:
        """Cierra Chrome dandole tiempo a guardar la sesion.

        WhatsApp guarda las claves de la sesion en IndexedDB. Matar el navegador
        de golpe puede dejar esa escritura a medias, y entonces la proxima vez
        vuelve a pedir el QR. Un respiro antes de cerrar lo evita.
        """
        if self.driver is None:
            return
        try:
            time.sleep(2.5)
        except Exception:
            pass
        try:
            self.driver.quit()
        except Exception:
            pass
        self.driver = None
        log("WhatsApp Web: navegador cerrado")

    # ---------------- auxiliares ----------------
    def _buscar(self, clave: str, espera: int) -> Any:
        """Prueba los selectores de `clave` hasta que uno aparezca. None si no."""
        from selenium.common.exceptions import TimeoutException
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        opciones = _selectores()[clave]
        por_selector = max(1, espera // max(1, len(opciones)))
        for localizador in opciones:
            try:
                return WebDriverWait(self.driver, por_selector).until(
                    EC.presence_of_element_located(localizador)
                )
            except TimeoutException:
                continue
            except Exception:
                continue
        return None

    @staticmethod
    def _avisar(progreso: Callable[[str], None] | None, texto: str) -> None:
        if progreso:
            try:
                progreso(texto)
            except Exception:
                pass
