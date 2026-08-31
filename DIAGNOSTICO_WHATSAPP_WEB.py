# -*- coding: utf-8 -*-
"""
Diagnostico de WhatsApp Web: averigua por que no se adjunta el PDF.

No manda ningun mensaje. Solo abre un chat, mira como esta armada la pagina y
escribe todo en el archivo DIAGNOSTICO_WHATSAPP_WEB.txt, al lado de este
archivo. Ese txt es el que hay que mandarle al programador.

Correr con:   python DIAGNOSTICO_WHATSAPP_WEB.py [telefono]
El telefono es solo para abrir un chat. Si no se pasa, usa el del
primer vendedor cargado.
"""

from __future__ import annotations

import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import whatsapp_web as ww  # noqa: E402

SALIDA = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "DIAGNOSTICO_WHATSAPP_WEB.txt"
)


def numero_de_prueba() -> str:
    """El numero con el que abrir un chat. NO se le manda nada.

    Se toma del primer argumento, o del primer vendedor cargado. No hay ninguno
    escrito en el codigo: seria el telefono de una persona real dando vueltas
    en el repositorio.
    """
    if len(sys.argv) > 1:
        solo = "".join(c for c in sys.argv[1] if c.isdigit())
        if solo:
            return solo
    try:
        import vendedores

        for v in vendedores.cargar_vendedores():
            if v.telefono:
                return v.telefono
    except Exception:
        pass
    print(
        "Falta el numero con el que abrir un chat de prueba.\n"
        "Pasalo como argumento:  python DIAGNOSTICO_WHATSAPP_WEB.py 1155551234\n"
        "(o carga un vendedor en el programa y volve a correr esto)"
    )
    sys.exit(1)

_lineas: list[str] = []


def anotar(texto: str = "") -> None:
    print(texto)
    _lineas.append(texto)


def main() -> None:
    numero = numero_de_prueba()
    anotar("=" * 70)
    anotar("DIAGNOSTICO DE WHATSAPP WEB")
    anotar("=" * 70)

    listo, motivo = ww.disponible()
    anotar(f"selenium y Chrome disponibles: {listo}")
    if not listo:
        anotar(motivo)
        return
    anotar(f"carpeta del perfil: {ww.CARPETA_PERFIL}")
    anotar(f"el perfil existe  : {os.path.isdir(ww.CARPETA_PERFIL)}")

    sesion = ww.Sesion()
    try:
        anotar("\nAbriendo Chrome...")
        sesion.abrir(progreso=lambda t: anotar(f"  {t}"))
        time.sleep(6)

        if not sesion.hay_sesion():
            anotar("\n>>> NO hay sesion iniciada: aparecio el codigo QR.")
            anotar(">>> Escanealo AHORA con el telefono.")
            anotar(">>> (WhatsApp > Dispositivos vinculados > Vincular un dispositivo)")
            input("\n    Cuando veas tus chats, apreta ENTER aca... ")
            time.sleep(4)

        anotar(f"\nsesion iniciada: {sesion.hay_sesion()}   |   QR a la vista: {sesion.hay_qr()}")
        if not sesion.hay_sesion():
            anotar("Sigue sin haber sesion. No se puede seguir.")
            return

        anotar("\nAbriendo un chat (NO se envia nada)...")
        sesion.driver.get(ww.url_chat(numero))
        time.sleep(9)
        anotar(f"url actual: {sesion.driver.current_url[:80]}")
        anotar(f"caja de texto encontrada: "
               f"{sesion._buscar('caja_texto', espera=15) is not None}")

        # ---------------- inputs de archivo ----------------
        def volcar_inputs(momento: str) -> None:
            datos = sesion.driver.execute_script(
                """
                return Array.from(document.querySelectorAll('input[type=file]'))
                  .map((e,i)=>({i:i, accept:e.getAttribute('accept'),
                                multiple:e.multiple,
                                oculto:e.offsetParent===null}));
                """
            )
            anotar(f"\n--- input[type=file] {momento}: {len(datos)} ---")
            for d in datos:
                anotar("   " + json.dumps(d, ensure_ascii=False))

        volcar_inputs("ANTES de tocar nada")

        # ---------------- botones del pie ----------------
        anotar("\n--- botones y iconos del pie del chat ---")
        info = sesion.driver.execute_script(
            """
            const pie = document.querySelector('footer') || document.body;
            const iconos = Array.from(pie.querySelectorAll('[data-icon]'))
                .map(e=>e.getAttribute('data-icon'));
            const botones = Array.from(pie.querySelectorAll('button,[role=button]'))
                .slice(0,20).map(e=>({
                    aria:e.getAttribute('aria-label'),
                    title:e.getAttribute('title'),
                    icono:(e.querySelector('[data-icon]')||{}).getAttribute
                        ? e.querySelector('[data-icon]').getAttribute('data-icon') : null
                }));
            return {iconos: iconos, botones: botones};
            """
        )
        anotar("   data-icon: " + json.dumps(info["iconos"], ensure_ascii=False))
        for b in info["botones"]:
            anotar("   boton: " + json.dumps(b, ensure_ascii=False))

        # ---------------- probar el boton de adjuntar ----------------
        anotar("\n--- probando el boton de adjuntar ---")
        boton = sesion._buscar("adjuntar", espera=8)
        anotar(f"   encontrado por los selectores del programa: {boton is not None}")
        if boton is not None:
            try:
                anotar(f"   aria-label: {boton.get_attribute('aria-label')}")
                anotar(f"   title     : {boton.get_attribute('title')}")
                sesion.driver.execute_script("arguments[0].click();", boton)
                time.sleep(2)
                anotar("   click dado")
            except Exception as e:
                anotar(f"   fallo el click: {e}")
            volcar_inputs("DESPUES de abrir el menu")

            anotar("\n--- opciones del menu que aparecio ---")
            opciones = sesion.driver.execute_script(
                """
                return Array.from(document.querySelectorAll('li,[role=button]'))
                  .map(e=>(e.innerText||'').trim())
                  .filter(t=>t && t.length<40).slice(0,20);
                """
            )
            anotar("   " + json.dumps(opciones, ensure_ascii=False))

        # ---------------- que dice el programa ----------------
        anotar("\n--- lo que resuelve el programa ---")
        entrada = sesion._input_de_archivo()
        anotar(f"   _input_de_archivo() devolvio: "
               f"{'un input' if entrada is not None else 'NADA (este es el problema)'}")
        if entrada is not None:
            anotar(f"   accept del input elegido: {entrada.get_attribute('accept')}")

        anotar(f"\nversion de Chrome: "
               f"{sesion.driver.capabilities.get('browserVersion','?')}")
    except Exception as e:
        anotar(f"\nERROR: {type(e).__name__}: {e}")
    finally:
        try:
            sesion.cerrar()
        except Exception:
            pass
        anotar("\nnavegador cerrado")
        try:
            with open(SALIDA, "w", encoding="utf-8") as fh:
                fh.write("\n".join(_lineas))
            print(f"\n\nListo. Mandale este archivo al programador:\n  {SALIDA}")
        except OSError as e:
            print(f"No se pudo escribir el informe: {e}")
    input("\nApreta ENTER para cerrar...")


if __name__ == "__main__":
    main()
