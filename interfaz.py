# -*- coding: utf-8 -*-
"""
Gestor de Cobranzas Wood Tools - interfaz de escritorio.

Flujo de uso:
  1. Cargar el reporte "A cobrar por cliente" (Excel .xls/.xlsx o PDF).
  2. Revisar los clientes vencidos y a vencer, con su saldo.
  3. Asignar los clientes a cada vendedor (queda guardado para la proxima vez).
  4. Enviar por WhatsApp el resumen: el PDF + el listado en un mensaje.

Todo lo lento (parseo, API de Claude, envios) corre en un hilo aparte y se
comunica con la ventana por una cola, para que la interfaz no se congele.
"""

from __future__ import annotations

import os
import queue
import threading
import traceback
import webbrowser
from datetime import date
from typing import Any, Callable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import claude_api
import mensajes
import parser_excel
import parser_pdf
import pdf_resumen
import vendedores as vend_mod
import whatsapp
from config import (
    APP_NOMBRE,
    APP_VERSION,
    ARCHIVO_LOG,
    CARPETA_HISTORIAL,
    CARPETA_SALIDA,
    CONFIG,
    asegurar_carpetas,
    hora_arg,
    log,
    log_error,
)
from modelos import (
    A_VENCER,
    Cliente,
    ETIQUETA_ESTADO,
    Reporte,
    VENCE_HOY,
    VENCIDO,
    formatear_fecha,
    formatear_pesos,
)

COLOR_VENCIDO = "#FDE3E1"
COLOR_A_VENCER = "#FFF9E0"
COLOR_RESALTADO = "#FFFACD"


# ==========================================================================
# VENTANA PRINCIPAL
# ==========================================================================
class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        asegurar_carpetas()

        self.title(f"{APP_NOMBRE} v{APP_VERSION}")
        # Se adapta a la pantalla: en notebooks de 1366x768 la ventana fija de
        # 780 px quedaba por debajo de la barra de tareas y cortaba los botones.
        ancho = min(1280, max(1024, self.winfo_screenwidth() - 80))
        alto = min(820, max(600, self.winfo_screenheight() - 120))
        self.geometry(f"{ancho}x{alto}+{max(0, (self.winfo_screenwidth() - ancho) // 2)}+20")
        self.minsize(1024, 600)

        self.reporte: Reporte | None = None
        self.ruta_archivo: str = ""
        self.mapa_filas: dict[str, Cliente] = {}
        self.mapa_vendedores: dict[str, vend_mod.Vendedor] = {}
        self.cola: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.trabajando = False
        self._orden_invertido: dict[str, bool] = {}

        self._crear_estilos()
        self._crear_widgets()
        self._cargar_config_en_ui()
        self.after(120, self._procesar_cola)

    # ------------------------------------------------------------------
    # ESTILOS
    # ------------------------------------------------------------------
    def _crear_estilos(self) -> None:
        estilo = ttk.Style(self)
        try:
            estilo.theme_use("vista")
        except tk.TclError:
            pass
        estilo.configure("Titulo.TLabel", font=("Segoe UI", 13, "bold"))
        estilo.configure("Sub.TLabel", foreground="#555555")
        estilo.configure("Total.TLabel", font=("Segoe UI", 11, "bold"))
        estilo.configure("Rojo.TLabel", foreground="#9C0006", font=("Segoe UI", 10, "bold"))
        estilo.configure("Verde.TLabel", foreground="#1F6F3F")
        estilo.configure("Treeview", rowheight=24)

    # ------------------------------------------------------------------
    # ARMADO DE LA VENTANA
    # ------------------------------------------------------------------
    def _crear_widgets(self) -> None:
        self.libreta = ttk.Notebook(self)
        self.libreta.pack(fill="both", expand=True, padx=8, pady=(8, 0))

        self.tab_archivo = ttk.Frame(self.libreta, padding=12)
        self.tab_clientes = ttk.Frame(self.libreta, padding=8)
        self.tab_vendedores = ttk.Frame(self.libreta, padding=12)
        self.tab_envio = ttk.Frame(self.libreta, padding=8)
        self.tab_config = ttk.Frame(self.libreta, padding=12)

        self.libreta.add(self.tab_archivo, text="  1. Cargar reporte  ")
        self.libreta.add(self.tab_clientes, text="  2. Clientes  ")
        self.libreta.add(self.tab_vendedores, text="  3. Vendedores  ")
        self.libreta.add(self.tab_envio, text="  4. Enviar  ")
        self.libreta.add(self.tab_config, text="  Configuracion  ")

        self._armar_tab_archivo()
        self._armar_tab_clientes()
        self._armar_tab_vendedores()
        self._armar_tab_envio()
        self._armar_tab_config()

        # barra de estado
        barra = ttk.Frame(self)
        barra.pack(fill="x", side="bottom", padx=8, pady=6)
        self.var_estado = tk.StringVar(value="Listo. Carga un reporte para empezar.")
        ttk.Label(barra, textvariable=self.var_estado, style="Sub.TLabel").pack(side="left")
        self.barra_progreso = ttk.Progressbar(barra, mode="indeterminate", length=180)
        self.barra_progreso.pack(side="right")

    # ---------------- TAB 1: CARGAR ----------------
    def _armar_tab_archivo(self) -> None:
        t = self.tab_archivo
        ttk.Label(t, text="Cargar el resumen de cuenta", style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(
            t,
            text=(
                "Acepta el reporte 'A cobrar por cliente' de Gestion Comercial en "
                "Excel (.xls / .xlsx) o en PDF.\n"
                "El Excel se lee directamente, incluyendo lo que marcaste en amarillo. "
                "Los PDF se interpretan con la API de Claude."
            ),
            style="Sub.TLabel",
            justify="left",
        ).pack(anchor="w", pady=(2, 12))

        fila = ttk.Frame(t)
        fila.pack(fill="x")
        self.var_ruta = tk.StringVar()
        ttk.Entry(fila, textvariable=self.var_ruta).pack(
            side="left", fill="x", expand=True, ipady=3
        )
        ttk.Button(fila, text="Buscar archivo...", command=self._elegir_archivo).pack(
            side="left", padx=(8, 0)
        )

        botones = ttk.Frame(t)
        botones.pack(fill="x", pady=12)
        self.btn_leer = ttk.Button(
            botones, text="Leer el archivo", command=self._leer_archivo
        )
        self.btn_leer.pack(side="left")
        self.btn_claude = ttk.Button(
            botones,
            text="Interpretar con Claude",
            command=self._interpretar_con_claude,
        )
        self.btn_claude.pack(side="left", padx=6)
        self.btn_verificar = ttk.Button(
            botones, text="Verificar con Claude", command=self._verificar_con_claude
        )
        self.btn_verificar.pack(side="left", padx=6)
        self.btn_telefonos = ttk.Button(
            botones,
            text="Ordenar telefonos con Claude",
            command=self._telefonos_con_claude,
        )
        self.btn_telefonos.pack(side="left", padx=6)

        # resumen
        marco = ttk.LabelFrame(t, text="Resumen de lo leido", padding=10)
        marco.pack(fill="x", pady=(4, 10))
        self.var_res_titulo = tk.StringVar(value="-")
        self.var_res_fecha = tk.StringVar(value="-")
        self.var_res_clientes = tk.StringVar(value="-")
        self.var_res_vencido = tk.StringVar(value="-")
        self.var_res_avencer = tk.StringVar(value="-")
        self.var_res_total = tk.StringVar(value="-")
        self.var_res_control = tk.StringVar(value="-")

        campos = [
            ("Reporte", self.var_res_titulo),
            ("Fecha del reporte", self.var_res_fecha),
            ("Clientes", self.var_res_clientes),
            ("Saldo vencido", self.var_res_vencido),
            ("Saldo a vencer", self.var_res_avencer),
            ("Total a cobrar", self.var_res_total),
            ("Control de totales", self.var_res_control),
        ]
        for i, (etiqueta, var) in enumerate(campos):
            ttk.Label(marco, text=f"{etiqueta}:").grid(
                row=i, column=0, sticky="w", padx=(0, 12), pady=1
            )
            estilo = "Total.TLabel" if etiqueta == "Total a cobrar" else "TLabel"
            ttk.Label(marco, textvariable=var, style=estilo).grid(
                row=i, column=1, sticky="w", pady=1
            )
        marco.columnconfigure(1, weight=1)

        # avisos
        marco_avisos = ttk.LabelFrame(t, text="Avisos y controles", padding=6)
        marco_avisos.pack(fill="both", expand=True)
        self.txt_avisos = tk.Text(
            marco_avisos, height=8, wrap="word", font=("Consolas", 9)
        )
        scroll = ttk.Scrollbar(marco_avisos, command=self.txt_avisos.yview)
        self.txt_avisos.configure(yscrollcommand=scroll.set, state="disabled")
        self.txt_avisos.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    # ---------------- TAB 2: CLIENTES ----------------
    def _armar_tab_clientes(self) -> None:
        t = self.tab_clientes

        filtros = ttk.LabelFrame(t, text="Filtros", padding=8)
        filtros.pack(fill="x")

        self.var_f_vencidos = tk.BooleanVar(value=True)
        self.var_f_avencer = tk.BooleanVar(value=True)
        self.var_f_resaltados = tk.BooleanVar(value=False)
        self.var_f_sin_vendedor = tk.BooleanVar(value=False)
        self.var_f_buscar = tk.StringVar()
        self.var_f_saldo_min = tk.StringVar(value="0")
        self.var_f_dias = tk.StringVar(value="7")

        ttk.Checkbutton(
            filtros, text="Vencidos", variable=self.var_f_vencidos,
            command=self._refrescar_clientes,
        ).grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Checkbutton(
            filtros, text="A vencer", variable=self.var_f_avencer,
            command=self._refrescar_clientes,
        ).grid(row=0, column=1, sticky="w", padx=(0, 10))
        ttk.Checkbutton(
            filtros, text="Solo los marcados en amarillo",
            variable=self.var_f_resaltados, command=self._refrescar_clientes,
        ).grid(row=0, column=2, sticky="w", padx=(0, 10))
        ttk.Checkbutton(
            filtros, text="Solo sin vendedor asignado",
            variable=self.var_f_sin_vendedor, command=self._refrescar_clientes,
        ).grid(row=0, column=3, sticky="w", padx=(0, 10))

        ttk.Label(filtros, text="A vencer dentro de").grid(row=1, column=0, sticky="e", pady=(6, 0))
        ttk.Entry(filtros, textvariable=self.var_f_dias, width=5).grid(
            row=1, column=1, sticky="w", pady=(6, 0)
        )
        ttk.Label(filtros, text="dias").grid(row=1, column=1, sticky="e", pady=(6, 0))

        ttk.Label(filtros, text="Saldo minimo $").grid(row=1, column=2, sticky="e", pady=(6, 0))
        ttk.Entry(filtros, textvariable=self.var_f_saldo_min, width=12).grid(
            row=1, column=3, sticky="w", pady=(6, 0)
        )

        ttk.Label(filtros, text="Buscar").grid(row=1, column=4, sticky="e", padx=(10, 4), pady=(6, 0))
        entrada = ttk.Entry(filtros, textvariable=self.var_f_buscar, width=26)
        entrada.grid(row=1, column=5, sticky="w", pady=(6, 0))
        entrada.bind("<KeyRelease>", lambda _e: self._refrescar_clientes())
        ttk.Button(filtros, text="Aplicar", command=self._refrescar_clientes).grid(
            row=1, column=6, padx=8, pady=(6, 0)
        )

        # La barra de acciones se reserva primero (abajo) para que nunca quede
        # tapada cuando la ventana es baja; la tabla usa el resto.
        acciones = ttk.Frame(t)
        acciones.pack(fill="x", side="bottom", pady=(6, 0))

        marco = ttk.Frame(t)
        marco.pack(fill="both", expand=True, pady=8)

        columnas = (
            "marca", "codigo", "cliente", "comp", "vencido", "avencer",
            "total", "situacion", "vendedor", "telefono",
        )
        self.tabla = ttk.Treeview(
            marco, columns=columnas, show="headings", selectmode="extended"
        )
        encabezados = {
            "marca": ("", 30), "codigo": ("Codigo", 70),
            "cliente": ("Cliente", 260), "comp": ("Cp.", 40),
            "vencido": ("Vencido", 105), "avencer": ("A vencer", 105),
            "total": ("Saldo total", 110), "situacion": ("Situacion", 165),
            "vendedor": ("Vendedor", 145), "telefono": ("Telefono", 120),
        }
        for col, (texto, ancho) in encabezados.items():
            self.tabla.heading(col, text=texto, command=lambda c=col: self._ordenar_por(c))
            alineacion = "e" if col in ("vencido", "avencer", "total") else "w"
            if col in ("marca", "comp"):
                alineacion = "center"
            self.tabla.column(col, width=ancho, anchor=alineacion, stretch=(col == "cliente"))

        vscroll = ttk.Scrollbar(marco, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=vscroll.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        self.tabla.tag_configure("vencido", background=COLOR_VENCIDO)
        self.tabla.tag_configure("avencer", background=COLOR_A_VENCER)
        self.tabla.bind("<Double-1>", self._ver_detalle_cliente)

        # contenido de la barra de acciones (el frame ya se empaqueto arriba)
        ttk.Button(acciones, text="Seleccionar todo", command=self._seleccionar_todo).pack(
            side="left"
        )
        ttk.Button(acciones, text="Ver comprobantes", command=self._ver_detalle_cliente).pack(
            side="left", padx=6
        )
        ttk.Label(acciones, text="   Asignar seleccionados a:").pack(side="left")
        self.var_asignar = tk.StringVar()
        self.combo_asignar = ttk.Combobox(
            acciones, textvariable=self.var_asignar, state="readonly", width=30
        )
        self.combo_asignar.pack(side="left", padx=6)
        ttk.Button(acciones, text="Asignar", command=self._asignar_seleccionados).pack(
            side="left"
        )
        ttk.Button(
            acciones, text="Quitar asignacion", command=self._desasignar_seleccionados
        ).pack(side="left", padx=6)

        self.var_totales_tabla = tk.StringVar(value="")
        ttk.Label(acciones, textvariable=self.var_totales_tabla, style="Total.TLabel").pack(
            side="right"
        )

    # ---------------- TAB 3: VENDEDORES ----------------
    def _armar_tab_vendedores(self) -> None:
        t = self.tab_vendedores
        ttk.Label(t, text="Vendedores", style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(
            t,
            text=(
                "Estos son los numeros a los que se manda el resumen. El telefono "
                "se guarda en formato internacional (549 + area + numero)."
            ),
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 10))

        cuerpo = ttk.Frame(t)
        cuerpo.pack(fill="both", expand=True)

        # lista
        izq = ttk.Frame(cuerpo)
        izq.pack(side="left", fill="both", expand=True)
        cols = ("nombre", "codigo", "telefono", "activo", "notas")
        self.tabla_vend = ttk.Treeview(izq, columns=cols, show="headings", height=14)
        for col, texto, ancho in (
            ("nombre", "Nombre", 190), ("codigo", "Codigo", 70),
            ("telefono", "Telefono WhatsApp", 150), ("activo", "Activo", 60),
            ("notas", "Notas", 200),
        ):
            self.tabla_vend.heading(col, text=texto)
            self.tabla_vend.column(col, width=ancho, anchor="w")
        vs = ttk.Scrollbar(izq, orient="vertical", command=self.tabla_vend.yview)
        self.tabla_vend.configure(yscrollcommand=vs.set)
        self.tabla_vend.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        self.tabla_vend.bind("<<TreeviewSelect>>", self._cargar_vendedor_en_form)

        # formulario
        der = ttk.LabelFrame(cuerpo, text="Datos del vendedor", padding=10)
        der.pack(side="left", fill="y", padx=(12, 0))

        self.var_v_id = tk.StringVar()
        self.var_v_nombre = tk.StringVar()
        self.var_v_codigo = tk.StringVar()
        self.var_v_telefono = tk.StringVar()
        self.var_v_activo = tk.BooleanVar(value=True)
        self.var_v_notas = tk.StringVar()

        campos = [
            ("Nombre", self.var_v_nombre, 26),
            ("Codigo del ERP", self.var_v_codigo, 10),
            ("Telefono", self.var_v_telefono, 26),
            ("Notas", self.var_v_notas, 26),
        ]
        for i, (etiqueta, var, ancho) in enumerate(campos):
            ttk.Label(der, text=etiqueta).grid(row=i, column=0, sticky="w", pady=3)
            ttk.Entry(der, textvariable=var, width=ancho).grid(
                row=i, column=1, sticky="w", pady=3, padx=(6, 0)
            )
        ttk.Checkbutton(der, text="Activo", variable=self.var_v_activo).grid(
            row=len(campos), column=1, sticky="w", pady=3
        )
        ttk.Label(
            der,
            text="Ej: 11 3900-9488  o  5491139009488",
            style="Sub.TLabel",
        ).grid(row=len(campos) + 1, column=0, columnspan=2, sticky="w", pady=(0, 8))

        bot = ttk.Frame(der)
        bot.grid(row=len(campos) + 2, column=0, columnspan=2, sticky="w")
        ttk.Button(bot, text="Agregar nuevo", command=self._agregar_vendedor).pack(
            side="left"
        )
        ttk.Button(bot, text="Guardar cambios", command=self._guardar_vendedor).pack(
            side="left", padx=6
        )
        ttk.Button(bot, text="Eliminar", command=self._eliminar_vendedor).pack(side="left")
        ttk.Button(bot, text="Limpiar", command=self._limpiar_form_vendedor).pack(
            side="left", padx=6
        )

    # ---------------- TAB 4: ENVIAR ----------------
    def _armar_tab_envio(self) -> None:
        t = self.tab_envio

        arriba = ttk.Frame(t)
        arriba.pack(fill="x")
        ttk.Label(arriba, text="Enviar el resumen por WhatsApp", style="Titulo.TLabel").pack(
            anchor="w"
        )
        self.var_aviso_envio = tk.StringVar(
            value=(
                "Modo libre: WhatsApp solo permite mensajes libres si el vendedor "
                "te escribio en las ultimas 24 horas."
            )
        )
        ttk.Label(arriba, textvariable=self.var_aviso_envio, style="Sub.TLabel").pack(
            anchor="w", pady=(2, 8)
        )

        cuerpo = ttk.Frame(t)
        cuerpo.pack(fill="both", expand=True)

        izq = ttk.LabelFrame(cuerpo, text="Destinatarios", padding=6)
        izq.pack(side="left", fill="both", expand=True)
        cols = ("vendedor", "telefono", "clientes", "vencido", "total")
        self.tabla_envio = ttk.Treeview(
            izq, columns=cols, show="headings", selectmode="extended", height=10
        )
        for col, texto, ancho, anchor in (
            ("vendedor", "Vendedor", 190, "w"),
            ("telefono", "Telefono", 130, "w"),
            ("clientes", "Clientes", 70, "center"),
            ("vencido", "Vencido", 110, "e"),
            ("total", "Total", 110, "e"),
        ):
            self.tabla_envio.heading(col, text=texto)
            self.tabla_envio.column(col, width=ancho, anchor=anchor)
        ve = ttk.Scrollbar(izq, orient="vertical", command=self.tabla_envio.yview)
        self.tabla_envio.configure(yscrollcommand=ve.set)
        self.tabla_envio.pack(side="left", fill="both", expand=True)
        ve.pack(side="right", fill="y")

        opciones = ttk.LabelFrame(cuerpo, text="Opciones del envio", padding=10)
        opciones.pack(side="left", fill="y", padx=(10, 0))

        self.var_e_pdf = tk.BooleanVar(value=True)
        self.var_e_comprobantes = tk.BooleanVar(value=False)
        self.var_e_telefonos_pdf = tk.BooleanVar(value=True)
        self.var_e_original = tk.BooleanVar(value=False)

        ttk.Checkbutton(
            opciones, text="Adjuntar PDF del resumen", variable=self.var_e_pdf
        ).pack(anchor="w", pady=2)
        ttk.Checkbutton(
            opciones,
            text="Detallar comprobantes en el mensaje",
            variable=self.var_e_comprobantes,
        ).pack(anchor="w", pady=2)
        ttk.Checkbutton(
            opciones, text="Incluir telefonos en el PDF",
            variable=self.var_e_telefonos_pdf,
        ).pack(anchor="w", pady=2)
        ttk.Checkbutton(
            opciones,
            text="Adjuntar el PDF original cargado",
            variable=self.var_e_original,
        ).pack(anchor="w", pady=2)

        ttk.Separator(opciones).pack(fill="x", pady=8)
        ttk.Button(
            opciones, text="Actualizar destinatarios", command=self._refrescar_envio
        ).pack(fill="x", pady=2)
        ttk.Button(opciones, text="Vista previa", command=self._vista_previa).pack(
            fill="x", pady=2
        )
        ttk.Button(opciones, text="Generar solo el PDF", command=self._solo_pdf).pack(
            fill="x", pady=2
        )
        ttk.Separator(opciones).pack(fill="x", pady=8)
        self.btn_enviar = ttk.Button(
            opciones, text="ENVIAR POR WHATSAPP", command=self._enviar
        )
        self.btn_enviar.pack(fill="x", pady=2)
        ttk.Button(
            opciones, text="Abrir en WhatsApp Web", command=self._abrir_wa_web
        ).pack(fill="x", pady=2)
        ttk.Button(
            opciones, text="Abrir carpeta de resumenes", command=self._abrir_carpeta
        ).pack(fill="x", pady=(8, 2))

        # log
        marco_log = ttk.LabelFrame(t, text="Registro del envio", padding=6)
        marco_log.pack(fill="both", expand=True, pady=(8, 0))
        self.txt_log = tk.Text(marco_log, height=9, wrap="word", font=("Consolas", 9))
        sl = ttk.Scrollbar(marco_log, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=sl.set, state="disabled")
        self.txt_log.pack(side="left", fill="both", expand=True)
        sl.pack(side="right", fill="y")

    # ---------------- TAB 5: CONFIGURACION ----------------
    def _armar_tab_config(self) -> None:
        t = self.tab_config
        contenedor = ttk.Frame(t)
        contenedor.pack(fill="both", expand=True)

        # --- Claude ---
        gc = ttk.LabelFrame(contenedor, text="API de Claude (Anthropic)", padding=10)
        gc.pack(fill="x", pady=(0, 10))
        self.var_c_apikey = tk.StringVar()
        self.var_c_modelo = tk.StringVar()
        self.var_c_esfuerzo = tk.StringVar()
        self.var_c_claude_pdf = tk.BooleanVar()
        self.var_c_claude_tel = tk.BooleanVar()

        ttk.Label(gc, text="API key").grid(row=0, column=0, sticky="w")
        self.entry_apikey = ttk.Entry(gc, textvariable=self.var_c_apikey, width=58, show="*")
        self.entry_apikey.grid(row=0, column=1, sticky="w", padx=6)
        self.var_ver_clave = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            gc, text="Ver", variable=self.var_ver_clave, command=self._alternar_clave
        ).grid(row=0, column=2)
        ttk.Label(gc, text="Modelo").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Combobox(
            gc,
            textvariable=self.var_c_modelo,
            values=["claude-opus-5", "claude-sonnet-5", "claude-opus-4-8"],
            width=26,
            state="readonly",
        ).grid(row=1, column=1, sticky="w", padx=6, pady=(6, 0))
        ttk.Label(gc, text="Esfuerzo").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Combobox(
            gc,
            textvariable=self.var_c_esfuerzo,
            values=["low", "medium", "high", "xhigh", "max"],
            width=12,
            state="readonly",
        ).grid(row=2, column=1, sticky="w", padx=6, pady=(6, 0))
        ttk.Checkbutton(
            gc, text="Interpretar los PDF con Claude", variable=self.var_c_claude_pdf
        ).grid(row=3, column=1, sticky="w", padx=6, pady=(6, 0))
        ttk.Checkbutton(
            gc,
            text="Ordenar los telefonos con Claude al cargar",
            variable=self.var_c_claude_tel,
        ).grid(row=4, column=1, sticky="w", padx=6)

        # --- WhatsApp ---
        gw = ttk.LabelFrame(contenedor, text="WhatsApp Cloud API (Meta)", padding=10)
        gw.pack(fill="x", pady=(0, 10))
        self.var_w_token = tk.StringVar()
        self.var_w_pid = tk.StringVar()
        self.var_w_version = tk.StringVar()
        self.var_w_modo = tk.StringVar()
        self.var_w_plantilla = tk.StringVar()

        ttk.Label(gw, text="Token permanente").grid(row=0, column=0, sticky="w")
        self.entry_token = ttk.Entry(gw, textvariable=self.var_w_token, width=58, show="*")
        self.entry_token.grid(row=0, column=1, sticky="w", padx=6)
        self.var_ver_token = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            gw, text="Ver", variable=self.var_ver_token, command=self._alternar_token
        ).grid(row=0, column=2)

        ttk.Label(gw, text="ID del numero").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(gw, textvariable=self.var_w_pid, width=28).grid(
            row=1, column=1, sticky="w", padx=6, pady=(6, 0)
        )
        ttk.Label(gw, text="Version API").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(gw, textvariable=self.var_w_version, width=12).grid(
            row=2, column=1, sticky="w", padx=6, pady=(6, 0)
        )
        ttk.Label(gw, text="Modo de envio").grid(row=3, column=0, sticky="w", pady=(6, 0))
        combo_modo = ttk.Combobox(
            gw,
            textvariable=self.var_w_modo,
            values=["libre", "plantilla"],
            width=12,
            state="readonly",
        )
        combo_modo.grid(row=3, column=1, sticky="w", padx=6, pady=(6, 0))
        combo_modo.bind("<<ComboboxSelected>>", lambda _e: self._actualizar_aviso_envio())
        ttk.Label(gw, text="Nombre de plantilla").grid(row=4, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(gw, textvariable=self.var_w_plantilla, width=28).grid(
            row=4, column=1, sticky="w", padx=6, pady=(6, 0)
        )
        ttk.Label(
            gw,
            text=(
                "Modo libre: mensajes de texto y PDF, requiere que el vendedor haya "
                "escrito en las ultimas 24 h.\n"
                "Modo plantilla: funciona siempre, pero Meta no permite saltos de "
                "linea, asi que el listado va en una sola linea."
            ),
            style="Sub.TLabel",
            justify="left",
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(8, 0))

        # --- criterios ---
        gcr = ttk.LabelFrame(contenedor, text="Criterios de negocio", padding=10)
        gcr.pack(fill="x", pady=(0, 10))
        self.var_k_fecha_reporte = tk.BooleanVar()
        self.var_k_firma = tk.StringVar()
        self.var_k_max_car = tk.StringVar()
        self.var_k_pausa = tk.StringVar()
        self.var_k_area = tk.StringVar()

        ttk.Checkbutton(
            gcr,
            text="Usar la fecha del reporte para calcular los dias "
            "(si se destilda, se calcula contra hoy)",
            variable=self.var_k_fecha_reporte,
        ).grid(row=0, column=0, columnspan=4, sticky="w")

        ttk.Label(gcr, text="Firma del mensaje").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(gcr, textvariable=self.var_k_firma, width=34).grid(
            row=1, column=1, sticky="w", padx=6, pady=(6, 0)
        )
        ttk.Label(gcr, text="Maximo de caracteres").grid(row=1, column=2, sticky="e", pady=(6, 0))
        ttk.Entry(gcr, textvariable=self.var_k_max_car, width=8).grid(
            row=1, column=3, sticky="w", padx=6, pady=(6, 0)
        )
        ttk.Label(gcr, text="Pausa entre envios (seg)").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(gcr, textvariable=self.var_k_pausa, width=8).grid(
            row=2, column=1, sticky="w", padx=6, pady=(6, 0)
        )
        ttk.Label(gcr, text="Area telefonica por defecto").grid(
            row=2, column=2, sticky="e", pady=(6, 0)
        )
        ttk.Entry(gcr, textvariable=self.var_k_area, width=8).grid(
            row=2, column=3, sticky="w", padx=6, pady=(6, 0)
        )

        # --- botones ---
        bot = ttk.Frame(contenedor)
        bot.pack(fill="x", pady=6)
        ttk.Button(bot, text="Guardar configuracion", command=self._guardar_config).pack(
            side="left"
        )
        ttk.Button(bot, text="Probar WhatsApp", command=self._probar_whatsapp).pack(
            side="left", padx=6
        )
        ttk.Button(bot, text="Probar Claude", command=self._probar_claude).pack(side="left")
        ttk.Button(bot, text="Ver archivo de errores", command=self._abrir_log).pack(
            side="right"
        )

    # ==================================================================
    # HILOS Y COLA
    # ==================================================================
    def _progreso(self, texto: str) -> None:
        """Callback que pasa a las funciones lentas (corre en el hilo worker)."""
        self.cola.put(("progreso", texto))

    def _en_hilo(
        self,
        trabajo: Callable[[Callable[[str], None]], Any],
        al_terminar: Callable[[Any], None],
        etiqueta: str = "Procesando...",
    ) -> None:
        """Corre `trabajo` en un hilo y llama a `al_terminar` en el hilo de la UI."""
        if self.trabajando:
            messagebox.showinfo(
                "Un momento", "Hay una operacion en curso. Espera a que termine."
            )
            return
        self.trabajando = True
        self._bloquear_botones(True)
        self.var_estado.set(etiqueta)
        self.barra_progreso.start(12)

        def correr() -> None:
            try:
                resultado = trabajo(self._progreso)
                self.cola.put(("listo", (al_terminar, resultado)))
            except Exception as e:  # noqa: BLE001
                log_error(f"{etiqueta} fallo: {e}\n{traceback.format_exc()}")
                self.cola.put(("fallo", e))

        threading.Thread(target=correr, daemon=True).start()

    def _procesar_cola(self) -> None:
        try:
            while True:
                tipo, dato = self.cola.get_nowait()
                if tipo == "progreso":
                    self.var_estado.set(str(dato))
                    self._log(str(dato))
                elif tipo == "listo":
                    al_terminar, resultado = dato
                    self._fin_trabajo()
                    try:
                        al_terminar(resultado)
                    except Exception as e:  # noqa: BLE001
                        log_error(f"Error mostrando el resultado: {e}\n{traceback.format_exc()}")
                        messagebox.showerror("Error", f"Error mostrando el resultado:\n{e}")
                elif tipo == "fallo":
                    self._fin_trabajo()
                    self.var_estado.set("La operacion no se pudo completar.")
                    messagebox.showerror("No se pudo completar", str(dato))
        except queue.Empty:
            pass
        self.after(120, self._procesar_cola)

    def _fin_trabajo(self) -> None:
        self.trabajando = False
        self._bloquear_botones(False)
        self.barra_progreso.stop()

    def _bloquear_botones(self, bloquear: bool) -> None:
        estado = "disabled" if bloquear else "normal"
        for btn in (
            self.btn_leer, self.btn_claude, self.btn_verificar,
            self.btn_telefonos, self.btn_enviar,
        ):
            try:
                btn.configure(state=estado)
            except tk.TclError:
                pass

    # ==================================================================
    # TAB 1: CARGA
    # ==================================================================
    def _elegir_archivo(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Elegi el reporte de cuentas a cobrar",
            filetypes=[
                ("Reportes (Excel o PDF)", "*.xls *.xlsx *.xlsm *.pdf"),
                ("Excel", "*.xls *.xlsx *.xlsm"),
                ("PDF", "*.pdf"),
                ("Todos", "*.*"),
            ],
        )
        if ruta:
            self.var_ruta.set(ruta)

    def _leer_archivo(self) -> None:
        ruta = self.var_ruta.get().strip('" ')
        if not ruta:
            messagebox.showwarning("Falta el archivo", "Elegi primero un archivo.")
            return
        if not os.path.exists(ruta):
            messagebox.showerror("No existe", f"No se encuentra el archivo:\n{ruta}")
            return

        ext = os.path.splitext(ruta)[1].lower()
        area = str(CONFIG.get("area_telefono_defecto", "11"))

        if ext == ".pdf":
            usar_claude = bool(
                CONFIG.get("usar_claude_en_pdf", True) and claude_api.esta_disponible()
            )

            def trabajo_pdf(prog: Callable[[str], None]) -> Reporte:
                # Con la API disponible se intenta primero Claude, que es el
                # camino confiable para PDF. Si falla (sin saldo, sin internet),
                # se cae al parser propio en vez de dejar al usuario sin nada.
                if usar_claude:
                    try:
                        return claude_api.interpretar_pdf(ruta, progreso=prog)
                    except claude_api.ErrorClaude as e:
                        prog("Claude no pudo: se intenta con el lector propio...")
                        rep = parser_pdf.parsear_pdf(ruta, area)
                        rep.avisos.insert(
                            0,
                            f"No se pudo usar Claude para este PDF: {e}\n"
                            "Se leyo con el lector propio, que es menos preciso "
                            "en PDF: revisa bien los importes.",
                        )
                        return rep

                prog("Leyendo el PDF...")
                rep = parser_pdf.parsear_pdf(ruta, area)
                if not rep.clientes:
                    rep.avisos.append(
                        "Configura la API key de Claude para interpretar PDF con "
                        "mucha mas precision."
                    )
                return rep

            self._en_hilo(trabajo_pdf, self._recibir_reporte, "Leyendo el PDF...")
            return

        if ext not in (".xls", ".xlsx", ".xlsm"):
            messagebox.showerror(
                "Formato no soportado",
                f"No se puede leer '{ext}'. Usa .xls, .xlsx o .pdf.",
            )
            return

        def trabajo_excel(prog: Callable[[str], None]) -> Reporte:
            prog("Leyendo el Excel...")
            rep = parser_excel.parsear_excel(ruta, area)
            if rep.clientes and CONFIG.get("usar_claude_para_telefonos", False) and (
                claude_api.esta_disponible()
            ):
                try:
                    prog("Ordenando los telefonos con Claude...")
                    claude_api.normalizar_telefonos(rep, prog)
                except claude_api.ErrorClaude as e:
                    rep.avisos.append(f"No se pudieron ordenar los telefonos: {e}")
            return rep

        self._en_hilo(trabajo_excel, self._recibir_reporte, "Leyendo el Excel...")

    def _interpretar_con_claude(self) -> None:
        ruta = self.var_ruta.get().strip('" ')
        if not ruta or not os.path.exists(ruta):
            messagebox.showwarning("Falta el archivo", "Elegi primero un archivo valido.")
            return
        if not claude_api.esta_disponible():
            messagebox.showwarning(
                "Falta la API key",
                "Para usar Claude tenes que pegar tu API key en la pestana "
                "Configuracion.",
            )
            self.libreta.select(self.tab_config)
            return

        ext = os.path.splitext(ruta)[1].lower()

        def trabajo(prog: Callable[[str], None]) -> Reporte:
            if ext == ".pdf":
                return claude_api.interpretar_pdf(ruta, progreso=prog)
            prog("Preparando el volcado del Excel...")
            volcado = parser_excel.volcar_para_claude(ruta)
            return claude_api.interpretar_volcado_excel(
                volcado, nombre_archivo=ruta, progreso=prog
            )

        self._en_hilo(trabajo, self._recibir_reporte, "Claude esta interpretando...")

    def _verificar_con_claude(self) -> None:
        if not self._hay_reporte():
            return
        if not claude_api.esta_disponible():
            messagebox.showwarning(
                "Falta la API key", "Configura la API key de Claude primero."
            )
            return
        ruta = self.ruta_archivo
        rep = self.reporte
        assert rep is not None

        def trabajo(prog: Callable[[str], None]) -> dict[str, Any]:
            prog("Preparando los datos para auditar...")
            ext = os.path.splitext(ruta)[1].lower()
            if ext == ".pdf":
                volcado = parser_pdf.volcar_para_claude(ruta)
            else:
                volcado = parser_excel.volcar_para_claude(ruta)
            return claude_api.verificar_reporte(rep, volcado, prog)

        self._en_hilo(trabajo, self._mostrar_verificacion, "Auditando con Claude...")

    def _mostrar_verificacion(self, resultado: dict[str, Any]) -> None:
        lineas: list[str] = []
        if resultado.get("coincide"):
            lineas.append("Claude confirma que los datos leidos coinciden con el original.")
        else:
            lineas.append("Claude encontro diferencias con el original:")
        for clave, titulo in (
            ("clientes_faltantes", "Clientes que faltan"),
            ("diferencias", "Diferencias de importes o comprobantes"),
            ("observaciones", "Observaciones"),
        ):
            items = resultado.get(clave) or []
            if items:
                lineas.append(f"\n{titulo}:")
                lineas.extend(f"  - {x}" for x in items)

        texto = "\n".join(lineas)
        self._escribir_avisos(texto, agregar=True)
        self.libreta.select(self.tab_archivo)
        if resultado.get("coincide"):
            messagebox.showinfo("Verificacion", "Los datos coinciden con el original.")
        else:
            messagebox.showwarning(
                "Revisar", "Claude encontro diferencias. Mira el detalle en Avisos."
            )

    def _telefonos_con_claude(self) -> None:
        if not self._hay_reporte():
            return
        if not claude_api.esta_disponible():
            messagebox.showwarning(
                "Falta la API key", "Configura la API key de Claude primero."
            )
            return
        rep = self.reporte
        assert rep is not None

        def trabajo(prog: Callable[[str], None]) -> int:
            return claude_api.normalizar_telefonos(rep, prog)

        def al_terminar(cantidad: int) -> None:
            self._refrescar_clientes()
            messagebox.showinfo(
                "Telefonos", f"Se ordenaron los telefonos de {cantidad} clientes."
            )

        self._en_hilo(trabajo, al_terminar, "Ordenando telefonos con Claude...")

    def _recibir_reporte(self, rep: Reporte) -> None:
        self.reporte = rep
        self.ruta_archivo = rep.origen_archivo or self.var_ruta.get().strip('" ')

        # Recalcular dias contra hoy, si asi esta configurado
        if not CONFIG.get("usar_fecha_reporte", True):
            hoy = hora_arg().date()
            cambios = parser_excel.recalcular_dias(rep, hoy)
            if cambios:
                rep.avisos.append(
                    f"Se recalcularon los dias de {cambios} comprobantes contra hoy "
                    f"({formatear_fecha(hoy)}) en lugar de la fecha del reporte."
                )

        self._mostrar_resumen(rep)
        self._refrescar_clientes()
        self._refrescar_envio()

        problemas = rep.validar()
        if not rep.clientes:
            self.var_estado.set("No se reconocieron clientes en el archivo.")
            self.libreta.select(self.tab_archivo)
            messagebox.showwarning(
                "No se reconocio el reporte",
                "No se pudo identificar ningun cliente en el archivo.\n\n"
                "Si es un PDF o el formato del reporte cambio, usa el boton "
                "'Interpretar con Claude'.",
            )
            return

        if problemas:
            self.libreta.select(self.tab_archivo)
            self.var_estado.set(
                f"Cargado CON AVISOS: {rep.resumen_corto()} "
                f"({len(problemas)} controles no cierran)"
            )
            muestra = "\n".join(f"  - {p}" for p in problemas[:6])
            extra = (
                f"\n  ...y {len(problemas) - 6} mas (mira la pestana Avisos)."
                if len(problemas) > 6
                else ""
            )
            sugerencia = (
                "\n\nEste archivo es un PDF: conviene interpretarlo con Claude, "
                "que lee el documento completo."
                if (rep.motor == "pdf")
                else "\n\nPodes usar 'Verificar con Claude' para que audite la lectura."
            )
            messagebox.showwarning(
                "Los totales no cierran",
                "Se leyo el archivo, pero hay controles que no coinciden:\n\n"
                f"{muestra}{extra}{sugerencia}\n\n"
                "Revisa los datos antes de mandar los mensajes.",
            )
        else:
            self.libreta.select(self.tab_clientes)
            self.var_estado.set(f"Cargado: {rep.resumen_corto()}")

    def _mostrar_resumen(self, rep: Reporte) -> None:
        self.var_res_titulo.set(
            f"{rep.titulo or 'Sin titulo'}  ({rep.motor})"
        )
        self.var_res_fecha.set(formatear_fecha(rep.fecha_reporte))
        vencidos = len(rep.clientes_vencidos)
        self.var_res_clientes.set(
            f"{len(rep.clientes)}  ({vencidos} con deuda vencida, "
            f"{len(rep.clientes_resaltados)} marcados en amarillo)"
        )
        self.var_res_vencido.set(formatear_pesos(rep.saldo_vencido))
        self.var_res_avencer.set(formatear_pesos(rep.saldo_a_vencer))
        self.var_res_total.set(formatear_pesos(rep.saldo_total))

        if rep.total_general_saldo is None:
            self.var_res_control.set("El reporte no trae total general para comparar.")
        else:
            dif = rep.descuadre_general()
            if abs(dif) <= 0.05:
                self.var_res_control.set(
                    f"OK: coincide con el total del reporte "
                    f"({formatear_pesos(rep.total_general_saldo)})."
                )
            else:
                self.var_res_control.set(
                    f"ATENCION: el reporte dice "
                    f"{formatear_pesos(rep.total_general_saldo)} y la suma da "
                    f"{formatear_pesos(rep.saldo_total)} "
                    f"(diferencia {formatear_pesos(dif)})."
                )

        problemas = rep.validar()
        partes: list[str] = []
        if rep.avisos:
            partes.append("AVISOS DE LA LECTURA")
            partes.extend(f"  - {a}" for a in rep.avisos)
        if problemas:
            partes.append("\nCONTROLES QUE NO CIERRAN")
            partes.extend(f"  - {p}" for p in problemas)
        if not partes:
            partes.append("Sin avisos: los totales de todos los clientes cierran.")
        self._escribir_avisos("\n".join(partes))

    def _escribir_avisos(self, texto: str, agregar: bool = False) -> None:
        self.txt_avisos.configure(state="normal")
        if not agregar:
            self.txt_avisos.delete("1.0", "end")
        else:
            self.txt_avisos.insert("end", "\n\n")
        self.txt_avisos.insert("end", texto)
        self.txt_avisos.configure(state="disabled")
        self.txt_avisos.see("end")

    # ==================================================================
    # TAB 2: CLIENTES
    # ==================================================================
    def _clientes_filtrados(self) -> list[Cliente]:
        if self.reporte is None:
            return []

        try:
            saldo_min = float(str(self.var_f_saldo_min.get()).replace(",", ".") or 0)
        except ValueError:
            saldo_min = 0.0
        try:
            dias_limite = int(self.var_f_dias.get() or 0)
        except ValueError:
            dias_limite = 0

        buscar = self.var_f_buscar.get().strip().lower()
        asignaciones = vend_mod.cargar_asignaciones()

        salida: list[Cliente] = []
        for cli in self.reporte.clientes:
            es_vencido = cli.estado == VENCIDO
            if es_vencido and not self.var_f_vencidos.get():
                continue
            if not es_vencido and not self.var_f_avencer.get():
                continue
            if not es_vencido and dias_limite > 0:
                d = cli.dias_hasta_proximo_vto
                if d is not None and d > dias_limite:
                    continue
            if self.var_f_resaltados.get() and not cli.resaltado:
                continue
            if cli.saldo_total < saldo_min:
                continue
            if self.var_f_sin_vendedor.get():
                clave = cli.codigo_corto
                if asignaciones.get(clave):
                    continue
            if buscar:
                blanco = f"{cli.nombre} {cli.codigo} {cli.codigo_corto} {cli.nombre_fantasia}"
                if buscar not in blanco.lower():
                    continue
            salida.append(cli)

        salida.sort(key=lambda c: (c.estado != VENCIDO, -c.dias_max_vencido, -c.saldo_total))
        return salida

    def _refrescar_clientes(self) -> None:
        for fila in self.tabla.get_children():
            self.tabla.delete(fila)
        self.mapa_filas.clear()

        clientes = self._clientes_filtrados()
        asignaciones = vend_mod.cargar_asignaciones()
        cache_vend: dict[str, str] = {}

        for cli in clientes:
            id_v = asignaciones.get(cli.codigo_corto, "")
            if id_v and id_v not in cache_vend:
                v = vend_mod.vendedor_por_id(id_v)
                cache_vend[id_v] = v.etiqueta() if v else "(vendedor borrado)"
            nombre_v = cache_vend.get(id_v, "")

            tel = cli.telefonos[0] if cli.telefonos else ""
            etiqueta = "vencido" if cli.estado == VENCIDO else "avencer"
            iid = self.tabla.insert(
                "",
                "end",
                values=(
                    "*" if cli.resaltado else "",
                    cli.codigo_corto,
                    cli.nombre,
                    cli.cantidad_comprobantes,
                    formatear_pesos(cli.saldo_vencido) if cli.saldo_vencido else "",
                    formatear_pesos(cli.saldo_a_vencer) if cli.saldo_a_vencer else "",
                    formatear_pesos(cli.saldo_total),
                    cli.descripcion_tiempo(),
                    nombre_v,
                    tel,
                ),
                tags=(etiqueta,),
            )
            self.mapa_filas[iid] = cli

        total = round(sum(c.saldo_total for c in clientes), 2)
        venc = round(sum(c.saldo_vencido for c in clientes), 2)
        self.var_totales_tabla.set(
            f"{len(clientes)} clientes | vencido {formatear_pesos(venc)} | "
            f"total {formatear_pesos(total)}"
        )
        self._refrescar_combo_vendedores()

    def _ordenar_por(self, columna: str) -> None:
        """Ordena la tabla por la columna clickeada."""
        filas = list(self.tabla.get_children())
        if not filas:
            return

        def clave(iid: str) -> Any:
            cli = self.mapa_filas.get(iid)
            if cli is None:
                return ""
            if columna == "codigo":
                return int(cli.codigo_corto or 0)
            if columna == "cliente":
                return cli.nombre.lower()
            if columna == "comp":
                return cli.cantidad_comprobantes
            if columna == "vencido":
                return cli.saldo_vencido
            if columna == "avencer":
                return cli.saldo_a_vencer
            if columna == "total":
                return cli.saldo_total
            if columna == "situacion":
                return -cli.dias_max_vencido
            if columna == "marca":
                return not cli.resaltado
            return self.tabla.set(iid, columna).lower()

        invertido = self._orden_invertido.get(columna, False)
        try:
            filas.sort(key=clave, reverse=not invertido)
        except Exception:
            return
        self._orden_invertido[columna] = not invertido
        for pos, iid in enumerate(filas):
            self.tabla.move(iid, "", pos)

    def _seleccionar_todo(self) -> None:
        self.tabla.selection_set(self.tabla.get_children())

    def _clientes_seleccionados(self) -> list[Cliente]:
        return [
            self.mapa_filas[iid]
            for iid in self.tabla.selection()
            if iid in self.mapa_filas
        ]

    def _refrescar_combo_vendedores(self) -> None:
        lista = vend_mod.cargar_vendedores()
        self.mapa_vendedores = {v.descripcion(): v for v in lista}
        valores = list(self.mapa_vendedores.keys())
        self.combo_asignar.configure(values=valores)
        if valores and self.var_asignar.get() not in valores:
            self.var_asignar.set(valores[0])

    def _asignar_seleccionados(self) -> None:
        seleccion = self._clientes_seleccionados()
        if not seleccion:
            messagebox.showinfo(
                "Sin seleccion",
                "Marca uno o mas clientes en la lista (podes usar Ctrl o Shift).",
            )
            return
        vendedor = self.mapa_vendedores.get(self.var_asignar.get())
        if vendedor is None:
            messagebox.showwarning(
                "Sin vendedor",
                "Primero carga un vendedor en la pestana Vendedores.",
            )
            return
        cantidad = vend_mod.asignar_varios([c.codigo for c in seleccion], vendedor.id)
        self._refrescar_clientes()
        self._refrescar_envio()
        self.var_estado.set(
            f"{cantidad} clientes asignados a {vendedor.etiqueta()}."
        )

    def _desasignar_seleccionados(self) -> None:
        seleccion = self._clientes_seleccionados()
        if not seleccion:
            return
        for cli in seleccion:
            vend_mod.desasignar_cliente(cli.codigo)
        self._refrescar_clientes()
        self._refrescar_envio()
        self.var_estado.set(f"Se quito la asignacion de {len(seleccion)} clientes.")

    def _ver_detalle_cliente(self, _evento: Any = None) -> None:
        seleccion = self._clientes_seleccionados()
        if not seleccion:
            return
        cli = seleccion[0]

        ventana = tk.Toplevel(self)
        ventana.title(f"{cli.nombre} - Cod. {cli.codigo_corto}")
        ventana.geometry("880x520")
        ventana.transient(self)

        cabecera = ttk.Frame(ventana, padding=10)
        cabecera.pack(fill="x")
        ttk.Label(cabecera, text=cli.nombre, style="Titulo.TLabel").pack(anchor="w")
        datos = [
            ("Codigo", f"{cli.codigo}  (corto: {cli.codigo_corto})"),
            ("Nombre de fantasia", cli.nombre_fantasia or "-"),
            ("Direccion", cli.direccion or "-"),
            ("Localidad", cli.localidad or "-"),
            ("Contacto", cli.contacto or "-"),
            ("Telefonos (texto original)", cli.telefonos_raw or "-"),
            ("Telefonos detectados", ", ".join(cli.telefonos) or "-"),
            ("Marcado en amarillo", "Si" if cli.resaltado else "No"),
            ("Saldo vencido", formatear_pesos(cli.saldo_vencido)),
            ("Saldo a vencer", formatear_pesos(cli.saldo_a_vencer)),
            ("Saldo total", formatear_pesos(cli.saldo_total)),
            (
                "Total segun el reporte",
                formatear_pesos(cli.total_saldo_reporte)
                if cli.total_saldo_reporte is not None
                else "no informado",
            ),
        ]
        grilla = ttk.Frame(cabecera)
        grilla.pack(fill="x", pady=(6, 0))
        for i, (k, v) in enumerate(datos):
            ttk.Label(grilla, text=f"{k}:").grid(row=i, column=0, sticky="w", padx=(0, 10))
            ttk.Label(grilla, text=str(v)).grid(row=i, column=1, sticky="w")

        marco = ttk.LabelFrame(ventana, text="Comprobantes pendientes", padding=6)
        marco.pack(fill="both", expand=True, padx=10, pady=10)
        cols = ("comp", "emision", "vto", "condicion", "original", "saldo", "dias", "estado")
        tabla = ttk.Treeview(marco, columns=cols, show="headings")
        for col, texto, ancho, anchor in (
            ("comp", "Comprobante", 140, "w"),
            ("emision", "Emision", 80, "w"),
            ("vto", "Vencimiento", 88, "w"),
            ("condicion", "Condicion de pago", 190, "w"),
            ("original", "Importe original", 110, "e"),
            ("saldo", "Saldo", 110, "e"),
            ("dias", "Dias", 50, "e"),
            ("estado", "Estado", 80, "w"),
        ):
            tabla.heading(col, text=texto)
            tabla.column(col, width=ancho, anchor=anchor)
        tabla.tag_configure("vencido", background=COLOR_VENCIDO)
        tabla.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(marco, orient="vertical", command=tabla.yview)
        tabla.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        for comp in sorted(
            cli.comprobantes,
            key=lambda k: -(k.dias_venc if k.dias_venc is not None else -99999),
        ):
            tabla.insert(
                "",
                "end",
                values=(
                    comp.comprobante_str,
                    formatear_fecha(comp.fecha_emision),
                    formatear_fecha(comp.fecha_vto),
                    comp.condicion_pago,
                    formatear_pesos(comp.importe_original),
                    formatear_pesos(comp.saldo),
                    "-" if comp.dias_venc is None else comp.dias_venc,
                    ETIQUETA_ESTADO.get(comp.estado, comp.estado),
                ),
                tags=("vencido",) if comp.esta_vencido else (),
            )

        ttk.Button(ventana, text="Cerrar", command=ventana.destroy).pack(pady=(0, 10))

    # ==================================================================
    # TAB 3: VENDEDORES
    # ==================================================================
    def _refrescar_vendedores(self) -> None:
        for fila in self.tabla_vend.get_children():
            self.tabla_vend.delete(fila)
        for v in vend_mod.cargar_vendedores():
            self.tabla_vend.insert(
                "",
                "end",
                iid=v.id,
                values=(v.nombre, v.codigo, v.telefono, "Si" if v.activo else "No", v.notas),
            )
        self._refrescar_combo_vendedores()

    def _cargar_vendedor_en_form(self, _evento: Any = None) -> None:
        seleccion = self.tabla_vend.selection()
        if not seleccion:
            return
        v = vend_mod.vendedor_por_id(seleccion[0])
        if v is None:
            return
        self.var_v_id.set(v.id)
        self.var_v_nombre.set(v.nombre)
        self.var_v_codigo.set(v.codigo)
        self.var_v_telefono.set(v.telefono)
        self.var_v_activo.set(v.activo)
        self.var_v_notas.set(v.notas)

    def _limpiar_form_vendedor(self) -> None:
        self.var_v_id.set("")
        self.var_v_nombre.set("")
        self.var_v_codigo.set("")
        self.var_v_telefono.set("")
        self.var_v_activo.set(True)
        self.var_v_notas.set("")
        self.tabla_vend.selection_remove(self.tabla_vend.selection())

    def _agregar_vendedor(self) -> None:
        v, motivo = vend_mod.agregar_vendedor(
            self.var_v_nombre.get(),
            self.var_v_telefono.get(),
            self.var_v_codigo.get(),
            self.var_v_notas.get(),
        )
        if v is None:
            messagebox.showwarning("No se pudo agregar", motivo)
            return
        self._limpiar_form_vendedor()
        self._refrescar_vendedores()
        self._refrescar_envio()
        self.var_estado.set(f"Vendedor agregado: {v.etiqueta()} ({v.telefono}).")

    def _guardar_vendedor(self) -> None:
        id_v = self.var_v_id.get()
        if not id_v:
            messagebox.showinfo(
                "Elegi un vendedor",
                "Seleccionalo de la lista, o usa 'Agregar nuevo'.",
            )
            return
        ok, motivo = vend_mod.actualizar_vendedor(
            id_v,
            nombre=self.var_v_nombre.get(),
            codigo=self.var_v_codigo.get(),
            telefono=self.var_v_telefono.get(),
            activo=self.var_v_activo.get(),
            notas=self.var_v_notas.get(),
        )
        if not ok:
            messagebox.showwarning("No se pudo guardar", motivo)
            return
        self._refrescar_vendedores()
        self._refrescar_clientes()
        self._refrescar_envio()
        self.var_estado.set("Vendedor actualizado.")

    def _eliminar_vendedor(self) -> None:
        id_v = self.var_v_id.get()
        if not id_v:
            return
        v = vend_mod.vendedor_por_id(id_v)
        nombre = v.etiqueta() if v else id_v
        if not messagebox.askyesno(
            "Confirmar",
            f"Eliminar a {nombre}?\n\nSus clientes van a quedar sin vendedor asignado.",
        ):
            return
        if vend_mod.eliminar_vendedor(id_v):
            self._limpiar_form_vendedor()
            self._refrescar_vendedores()
            self._refrescar_clientes()
            self._refrescar_envio()
            self.var_estado.set(f"Vendedor eliminado: {nombre}.")

    # ==================================================================
    # TAB 4: ENVIO
    # ==================================================================
    def _grupos_envio(self) -> dict[str, list[Cliente]]:
        """Agrupa los clientes filtrados por vendedor asignado."""
        return vend_mod.agrupar_por_vendedor(self._clientes_filtrados())

    def _refrescar_envio(self) -> None:
        for fila in self.tabla_envio.get_children():
            self.tabla_envio.delete(fila)

        grupos = self._grupos_envio()
        for id_v, clientes in sorted(
            grupos.items(),
            key=lambda kv: -round(sum(c.saldo_total for c in kv[1]), 2),
        ):
            if id_v == vend_mod.SIN_ASIGNAR:
                etiqueta, telefono = "(clientes sin vendedor asignado)", ""
            else:
                v = vend_mod.vendedor_por_id(id_v)
                if v is None:
                    etiqueta, telefono = "(vendedor borrado)", ""
                elif not v.activo:
                    etiqueta, telefono = f"{v.etiqueta()} [inactivo]", v.telefono
                else:
                    etiqueta, telefono = v.etiqueta(), v.telefono

            self.tabla_envio.insert(
                "",
                "end",
                iid=id_v or "__sin__",
                values=(
                    etiqueta,
                    telefono,
                    len(clientes),
                    formatear_pesos(round(sum(c.saldo_vencido for c in clientes), 2)),
                    formatear_pesos(round(sum(c.saldo_total for c in clientes), 2)),
                ),
            )
        self._actualizar_aviso_envio()

    def _actualizar_aviso_envio(self) -> None:
        modo = self.var_w_modo.get() or str(CONFIG.get("modo_envio", "libre"))
        if modo == "plantilla":
            self.var_aviso_envio.set(
                "Modo plantilla: funciona con la ventana de 24 h cerrada, pero Meta "
                "no acepta saltos de linea, asi que el listado va en una sola linea."
            )
        else:
            self.var_aviso_envio.set(
                "Modo libre: WhatsApp solo permite mensajes libres si el vendedor te "
                "escribio en las ultimas 24 horas. Si da error 131047, pedile que "
                "mande cualquier mensaje al numero de la empresa."
            )

    def _destinatarios_elegidos(self) -> list[tuple[vend_mod.Vendedor, list[Cliente]]]:
        seleccion = self.tabla_envio.selection()
        if not seleccion:
            messagebox.showinfo(
                "Sin destinatarios",
                "Elegi en la lista a que vendedores queres mandarle el resumen.",
            )
            return []

        grupos = self._grupos_envio()
        salida: list[tuple[vend_mod.Vendedor, list[Cliente]]] = []
        for iid in seleccion:
            if iid == "__sin__":
                messagebox.showwarning(
                    "Sin vendedor",
                    "Hay clientes sin vendedor asignado. Asignalos en la pestana "
                    "Clientes antes de enviar.",
                )
                continue
            v = vend_mod.vendedor_por_id(iid)
            clientes = grupos.get(iid, [])
            if v is None or not clientes:
                continue
            salida.append((v, clientes))
        return salida

    def _armar_partes(
        self,
        vendedor: vend_mod.Vendedor,
        clientes: list[Cliente],
        incluir_comprobantes: bool | None = None,
    ) -> list[str]:
        """Arma las partes del mensaje.

        `incluir_comprobantes` se pasa explicito cuando esto corre en un hilo
        secundario: leer una variable de Tkinter fuera del hilo principal no es
        seguro.
        """
        assert self.reporte is not None
        if incluir_comprobantes is None:
            incluir_comprobantes = bool(self.var_e_comprobantes.get())
        texto = mensajes.construir_mensaje(
            self.reporte,
            clientes,
            vendedor.nombre,
            vendedor.codigo,
            incluir_comprobantes=incluir_comprobantes,
        )
        return mensajes.partir_mensaje(texto)

    def _vista_previa(self) -> None:
        if not self._hay_reporte():
            return
        destinos = self._destinatarios_elegidos()
        if not destinos:
            return

        ventana = tk.Toplevel(self)
        ventana.title("Vista previa de los mensajes")
        ventana.geometry("760x640")
        ventana.transient(self)

        cuaderno = ttk.Notebook(ventana)
        cuaderno.pack(fill="both", expand=True, padx=8, pady=8)

        for vendedor, clientes in destinos:
            marco = ttk.Frame(cuaderno, padding=6)
            cuaderno.add(marco, text=f" {vendedor.nombre} ({len(clientes)}) ")
            partes = self._armar_partes(vendedor, clientes)
            info = (
                f"A: {vendedor.telefono}   |   {len(partes)} mensaje(s)   |   "
                f"{sum(len(p) for p in partes)} caracteres"
            )
            ttk.Label(marco, text=info, style="Sub.TLabel").pack(anchor="w", pady=(0, 4))
            caja = tk.Text(marco, wrap="word", font=("Segoe UI", 10))
            sb = ttk.Scrollbar(marco, command=caja.yview)
            caja.configure(yscrollcommand=sb.set)
            caja.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            for i, parte in enumerate(partes, 1):
                caja.insert("end", f"----- MENSAJE {i} de {len(partes)} -----\n")
                caja.insert("end", parte + "\n\n")
            caja.configure(state="disabled")

        ttk.Button(ventana, text="Cerrar", command=ventana.destroy).pack(pady=(0, 8))

    def _generar_pdf_para(
        self,
        vendedor: vend_mod.Vendedor,
        clientes: list[Cliente],
        incluir_telefonos: bool,
    ) -> tuple[bool, str]:
        assert self.reporte is not None
        return pdf_resumen.generar_pdf_en_carpeta_salida(
            self.reporte,
            clientes,
            nombre_vendedor=vendedor.nombre,
            codigo_vendedor=vendedor.codigo,
            incluir_comprobantes=True,
            incluir_telefonos=incluir_telefonos,
        )

    def _solo_pdf(self) -> None:
        if not self._hay_reporte():
            return
        destinos = self._destinatarios_elegidos()
        if not destinos:
            return
        # Se leen aca, en el hilo principal, y se pasan como valores al hilo.
        con_telefonos = bool(self.var_e_telefonos_pdf.get())

        def trabajo(prog: Callable[[str], None]) -> list[str]:
            hechos: list[str] = []
            for vendedor, clientes in destinos:
                prog(f"Generando el PDF de {vendedor.nombre}...")
                ok, ruta = self._generar_pdf_para(vendedor, clientes, con_telefonos)
                hechos.append(
                    f"{vendedor.etiqueta()}: {os.path.basename(ruta) if ok else ruta}"
                )
            return hechos

        def al_terminar(hechos: list[str]) -> None:
            for linea in hechos:
                self._log(linea)
            messagebox.showinfo(
                "PDF generados",
                "Se generaron en la carpeta Resumenes:\n\n" + "\n".join(hechos),
            )

        self._en_hilo(trabajo, al_terminar, "Generando PDF...")

    def _abrir_wa_web(self) -> None:
        """Abre WhatsApp Web con el texto cargado, para mandarlo a mano."""
        if not self._hay_reporte():
            return
        destinos = self._destinatarios_elegidos()
        if not destinos:
            return
        if len(destinos) > 1:
            messagebox.showinfo(
                "Uno a la vez",
                "Para abrir WhatsApp Web elegi un solo vendedor.",
            )
            return
        vendedor, clientes = destinos[0]
        partes = self._armar_partes(vendedor, clientes)
        enlace = whatsapp.link_wa_me(vendedor.telefono, partes[0])
        if len(partes) > 1:
            messagebox.showinfo(
                "Mensaje largo",
                f"El resumen ocupa {len(partes)} mensajes. Se va a abrir el "
                "primero; el resto lo podes copiar desde la vista previa.",
            )
        webbrowser.open(enlace)
        self._log(f"Abierto WhatsApp Web para {vendedor.etiqueta()}")

    def _enviar(self) -> None:
        if not self._hay_reporte():
            return
        if not whatsapp.esta_configurado():
            messagebox.showwarning(
                "Falta configurar WhatsApp",
                "Carga el token y el ID de numero de WhatsApp en Configuracion.",
            )
            self.libreta.select(self.tab_config)
            return

        destinos = self._destinatarios_elegidos()
        if not destinos:
            return

        # Barrera de seguridad: no mandar plata mal calculada sin avisar.
        assert self.reporte is not None
        problemas = self.reporte.validar()
        if problemas:
            muestra = "\n".join(f"  - {p}" for p in problemas[:5])
            extra = (
                f"\n  ...y {len(problemas) - 5} mas."
                if len(problemas) > 5
                else ""
            )
            if not messagebox.askyesno(
                "Los importes no cierran",
                "ATENCION: la lectura del archivo tiene controles que no "
                f"coinciden:\n\n{muestra}{extra}\n\n"
                "Si mandas los mensajes ahora, los vendedores pueden recibir "
                "importes equivocados.\n\n"
                "Recomendado: cancela, y usa 'Interpretar con Claude' o "
                "'Verificar con Claude' en la pestana 1.\n\n"
                "Enviar igual?",
                icon="warning",
                default="no",
            ):
                self.libreta.select(self.tab_archivo)
                return

        detalle = "\n".join(
            f"  - {v.etiqueta()} ({v.telefono}): {len(cs)} clientes, "
            f"{formatear_pesos(round(sum(c.saldo_total for c in cs), 2))}"
            for v, cs in destinos
        )
        if not messagebox.askyesno(
            "Confirmar envio",
            f"Se va a enviar el resumen por WhatsApp a:\n\n{detalle}\n\n"
            "Confirmas el envio?",
        ):
            return

        # Todos los valores de la interfaz se leen aca, en el hilo principal:
        # Tkinter no es seguro de leer desde un hilo secundario.
        adjuntar_pdf = bool(self.var_e_pdf.get())
        con_telefonos = bool(self.var_e_telefonos_pdf.get())
        con_comprobantes = bool(self.var_e_comprobantes.get())
        adjuntar_original = bool(self.var_e_original.get()) and self.ruta_archivo.lower().endswith(
            ".pdf"
        )
        ruta_original = self.ruta_archivo

        def trabajo(prog: Callable[[str], None]) -> list[str]:
            reporte_lineas: list[str] = []
            for vendedor, clientes in destinos:
                prog(f"--- {vendedor.etiqueta()} ---")
                ruta_pdf: str | None = None
                nombre_pdf = "resumen.pdf"

                if adjuntar_original:
                    ruta_pdf = ruta_original
                    nombre_pdf = os.path.basename(ruta_original)
                elif adjuntar_pdf:
                    prog(f"Generando el PDF de {vendedor.nombre}...")
                    ok, resultado = self._generar_pdf_para(
                        vendedor, clientes, con_telefonos
                    )
                    if ok:
                        ruta_pdf = resultado
                        nombre_pdf = os.path.basename(resultado)
                    else:
                        reporte_lineas.append(
                            f"{vendedor.etiqueta()}: no se pudo generar el PDF "
                            f"({resultado}). Se envia solo el texto."
                        )

                partes = self._armar_partes(vendedor, clientes, con_comprobantes)
                epigrafe = mensajes.epigrafe_pdf(
                    self.reporte, clientes, vendedor.nombre  # type: ignore[arg-type]
                )

                resultados = whatsapp.enviar_resumen(
                    vendedor.telefono,
                    partes,
                    ruta_pdf=ruta_pdf,
                    nombre_pdf=nombre_pdf,
                    epigrafe_pdf=epigrafe,
                    progreso=prog,
                )
                fallos = [r for r in resultados if not r.ok]
                if fallos:
                    for r in fallos:
                        reporte_lineas.append(f"ERROR - {vendedor.etiqueta()}: {r.mensaje}")
                else:
                    reporte_lineas.append(
                        f"OK - {vendedor.etiqueta()}: {len(resultados)} mensaje(s) "
                        f"enviados, {len(clientes)} clientes."
                    )
                self._registrar_historial(vendedor, clientes, not fallos)
            return reporte_lineas

        def al_terminar(lineas: list[str]) -> None:
            for linea in lineas:
                self._log(linea)
            hubo_error = any(l.startswith("ERROR") for l in lineas)
            if hubo_error:
                messagebox.showwarning(
                    "Envio con problemas", "\n\n".join(lineas)
                )
            else:
                messagebox.showinfo("Envio completo", "\n".join(lineas))

        self._en_hilo(trabajo, al_terminar, "Enviando por WhatsApp...")

    def _registrar_historial(
        self, vendedor: vend_mod.Vendedor, clientes: list[Cliente], ok: bool
    ) -> None:
        """Deja constancia del envio en Historial/envios.csv."""
        try:
            ruta = os.path.join(CARPETA_HISTORIAL, "envios.csv")
            nuevo = not os.path.exists(ruta)
            with open(ruta, "a", encoding="utf-8-sig", newline="") as fh:
                if nuevo:
                    fh.write(
                        "fecha_hora;vendedor;codigo;telefono;clientes;saldo_total;resultado\n"
                    )
                total = round(sum(c.saldo_total for c in clientes), 2)
                fh.write(
                    f"{hora_arg():%Y-%m-%d %H:%M:%S};{vendedor.nombre};"
                    f"{vendedor.codigo};{vendedor.telefono};{len(clientes)};"
                    f"{total:.2f};{'enviado' if ok else 'con errores'}\n"
                )
        except Exception as e:
            log_error(f"No se pudo escribir el historial: {e}")

    def _abrir_carpeta(self) -> None:
        try:
            os.startfile(CARPETA_SALIDA)  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showinfo("Carpeta", f"Los PDF quedan en:\n{CARPETA_SALIDA}\n\n{e}")

    def _log(self, texto: str) -> None:
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", f"[{hora_arg():%H:%M:%S}] {texto}\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    # ==================================================================
    # TAB 5: CONFIGURACION
    # ==================================================================
    def _cargar_config_en_ui(self) -> None:
        self.var_c_apikey.set(str(CONFIG.get("anthropic_api_key", "")))
        self.var_c_modelo.set(str(CONFIG.get("modelo_claude", "claude-opus-5")))
        self.var_c_esfuerzo.set(str(CONFIG.get("esfuerzo_claude", "high")))
        self.var_c_claude_pdf.set(bool(CONFIG.get("usar_claude_en_pdf", True)))
        self.var_c_claude_tel.set(bool(CONFIG.get("usar_claude_para_telefonos", True)))

        self.var_w_token.set(str(CONFIG.get("whatsapp_token", "")))
        self.var_w_pid.set(str(CONFIG.get("whatsapp_phone_number_id", "")))
        self.var_w_version.set(str(CONFIG.get("whatsapp_version", "v21.0")))
        self.var_w_modo.set(str(CONFIG.get("modo_envio", "libre")))
        self.var_w_plantilla.set(str(CONFIG.get("plantilla_nombre", "")))

        self.var_k_fecha_reporte.set(bool(CONFIG.get("usar_fecha_reporte", True)))
        self.var_k_firma.set(str(CONFIG.get("firma_mensaje", "")))
        self.var_k_max_car.set(str(CONFIG.get("max_caracteres_mensaje", 3800)))
        self.var_k_pausa.set(str(CONFIG.get("pausa_entre_envios_seg", 1.5)))
        self.var_k_area.set(str(CONFIG.get("area_telefono_defecto", "11")))

        self.var_f_dias.set(str(CONFIG.get("dias_aviso_previo", 7)))
        self.var_f_saldo_min.set(str(CONFIG.get("saldo_minimo", 0)))
        self.var_f_resaltados.set(bool(CONFIG.get("solo_resaltados", False)))
        self.var_e_pdf.set(bool(CONFIG.get("generar_pdf_por_vendedor", True)))
        self.var_e_original.set(bool(CONFIG.get("adjuntar_pdf_original", False)))

        self._refrescar_vendedores()
        self._actualizar_aviso_envio()

    def _guardar_config(self) -> None:
        CONFIG["anthropic_api_key"] = self.var_c_apikey.get().strip()
        CONFIG["modelo_claude"] = self.var_c_modelo.get().strip() or "claude-opus-5"
        CONFIG["esfuerzo_claude"] = self.var_c_esfuerzo.get().strip() or "high"
        CONFIG["usar_claude_en_pdf"] = bool(self.var_c_claude_pdf.get())
        CONFIG["usar_claude_para_telefonos"] = bool(self.var_c_claude_tel.get())

        CONFIG["whatsapp_token"] = self.var_w_token.get().strip()
        CONFIG["whatsapp_phone_number_id"] = self.var_w_pid.get().strip()
        CONFIG["whatsapp_version"] = self.var_w_version.get().strip() or "v21.0"
        CONFIG["modo_envio"] = self.var_w_modo.get().strip() or "libre"
        CONFIG["plantilla_nombre"] = self.var_w_plantilla.get().strip()

        CONFIG["usar_fecha_reporte"] = bool(self.var_k_fecha_reporte.get())
        CONFIG["firma_mensaje"] = self.var_k_firma.get().strip()
        CONFIG["area_telefono_defecto"] = self.var_k_area.get().strip() or "11"
        CONFIG["solo_resaltados"] = bool(self.var_f_resaltados.get())
        CONFIG["generar_pdf_por_vendedor"] = bool(self.var_e_pdf.get())
        CONFIG["adjuntar_pdf_original"] = bool(self.var_e_original.get())

        for clave, var, conversor, defecto in (
            ("max_caracteres_mensaje", self.var_k_max_car, int, 3800),
            ("pausa_entre_envios_seg", self.var_k_pausa, float, 1.5),
            ("dias_aviso_previo", self.var_f_dias, int, 7),
            ("saldo_minimo", self.var_f_saldo_min, float, 0.0),
        ):
            try:
                CONFIG[clave] = conversor(str(var.get()).replace(",", "."))
            except (ValueError, TypeError):
                CONFIG[clave] = defecto
                var.set(str(defecto))

        if CONFIG.guardar():
            self.var_estado.set("Configuracion guardada.")
            self._actualizar_aviso_envio()
            messagebox.showinfo("Guardado", "La configuracion quedo guardada.")
        else:
            messagebox.showerror(
                "Error",
                "No se pudo guardar config.json. Revisa los permisos de la carpeta.",
            )

    def _alternar_clave(self) -> None:
        self.entry_apikey.configure(show="" if self.var_ver_clave.get() else "*")

    def _alternar_token(self) -> None:
        self.entry_token.configure(show="" if self.var_ver_token.get() else "*")

    def _probar_whatsapp(self) -> None:
        self._guardar_config_silencioso()

        def trabajo(prog: Callable[[str], None]) -> Any:
            prog("Consultando a Meta...")
            return whatsapp.diagnosticar()

        def al_terminar(res: Any) -> None:
            if res.ok:
                messagebox.showinfo("WhatsApp", res.mensaje)
            else:
                messagebox.showerror("WhatsApp", res.mensaje)

        self._en_hilo(trabajo, al_terminar, "Probando WhatsApp...")

    def _probar_claude(self) -> None:
        self._guardar_config_silencioso()
        if not claude_api.esta_disponible():
            messagebox.showwarning("Falta la API key", "Pega tu API key de Claude primero.")
            return

        def trabajo(prog: Callable[[str], None]) -> str:
            prog("Consultando a la API de Claude...")
            return claude_api.estimar_costo(
                "Prueba de conexion con la API de Claude.", salida_estimada=100
            )

        def al_terminar(texto: str) -> None:
            # Si algo falla, estimar_costo levanta ErrorClaude y el motivo real
            # (saldo agotado, clave invalida, sin internet) lo muestra _en_hilo.
            messagebox.showinfo(
                "Claude",
                f"Conexion OK con {CONFIG.get('modelo_claude')}.\n\n"
                f"Prueba de conteo de tokens: {texto}",
            )

        self._en_hilo(trabajo, al_terminar, "Probando Claude...")

    def _guardar_config_silencioso(self) -> None:
        """Pasa los datos de la pantalla a CONFIG sin mostrar mensajes."""
        CONFIG["anthropic_api_key"] = self.var_c_apikey.get().strip()
        CONFIG["modelo_claude"] = self.var_c_modelo.get().strip() or "claude-opus-5"
        CONFIG["whatsapp_token"] = self.var_w_token.get().strip()
        CONFIG["whatsapp_phone_number_id"] = self.var_w_pid.get().strip()
        CONFIG["whatsapp_version"] = self.var_w_version.get().strip() or "v21.0"
        CONFIG["modo_envio"] = self.var_w_modo.get().strip() or "libre"
        CONFIG["plantilla_nombre"] = self.var_w_plantilla.get().strip()
        CONFIG.guardar()

    def _abrir_log(self) -> None:
        if not os.path.exists(ARCHIVO_LOG):
            messagebox.showinfo("Sin errores", "Todavia no hay errores registrados.")
            return
        try:
            os.startfile(ARCHIVO_LOG)  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showinfo("Archivo de errores", f"{ARCHIVO_LOG}\n\n{e}")

    # ==================================================================
    # AUXILIARES
    # ==================================================================
    def _hay_reporte(self) -> bool:
        if self.reporte is None or not self.reporte.clientes:
            messagebox.showinfo(
                "Falta el reporte",
                "Primero carga un reporte en la pestana 'Cargar reporte'.",
            )
            self.libreta.select(self.tab_archivo)
            return False
        return True


# ==========================================================================
# ARRANQUE
# ==========================================================================
def main() -> None:
    log(f"=== {APP_NOMBRE} v{APP_VERSION} iniciado ===")
    try:
        app = App()
        app.mainloop()
    except Exception as e:  # noqa: BLE001
        log_error(f"Fallo critico: {e}\n{traceback.format_exc()}")
        try:
            messagebox.showerror(
                "Error critico",
                f"El programa no pudo iniciar:\n\n{e}\n\nRevisa errores_log.txt",
            )
        except Exception:
            pass


if __name__ == "__main__":
    main()
