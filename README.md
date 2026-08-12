# Gestor de Cobranzas Wood Tools

Programa de escritorio para Windows que lee el reporte **"A COBRAR POR CLIENTE
EXPRESADO EN $"** de Gestión Comercial (en Excel o PDF), identifica los clientes
**vencidos** y **a vencer** con su saldo, y le manda a cada vendedor su resumen
de cuenta por WhatsApp: un PDF con el detalle más un mensaje con el listado.

---

## Instalación

1. Instalar Python desde [python.org](https://www.python.org/downloads/),
   tildando **"Add Python to PATH"**.
2. Doble clic en **`INSTALAR.bat`** (crea el entorno e instala las librerías).
3. Doble clic en **`ABRIR_PROGRAMA.bat`**.

Para generar el `.exe` distribuible: **`CONSTRUIR_EXE.bat`** (queda en `dist\`).

---

## Cómo se usa

### 1. Cargar reporte
Elegís el archivo y apretás **Leer el archivo**.

- **Excel (`.xls` / `.xlsx`)** → lo lee el parser propio: es exacto, instantáneo
  y gratis. **Detecta lo que marcaste en amarillo.**
- **PDF** → lo interpreta la **API de Claude**, que lee el documento completo
  (incluso escaneado). Si no hay API key configurada, intenta un parseo propio
  por patrones de texto, que es mucho menos confiable.

Siempre se controla que **la suma de los comprobantes de cada cliente coincida
con su fila "Total cliente"**, y que la suma de los clientes coincida con el
"Total general". Si algo no cierra, el programa avisa fuerte y **vuelve a
preguntar antes de enviar**.

Botones adicionales:

| Botón | Qué hace |
|---|---|
| **Interpretar con Claude** | Fuerza la lectura con la API, sirva o no el parser propio. Útil si el ERP cambió el formato del reporte. |
| **Verificar con Claude** | Audita lo que leyó el programa contra el archivo original y lista las diferencias. No modifica nada. |
| **Ordenar teléfonos con Claude** | La columna Teléfonos del ERP viene sucia (`"45246075 LAUTARO 1140647300 KAREN PAGO"`). Esto separa nombre y número y normaliza a formato WhatsApp. |

### 2. Clientes
Lista con **Vencido / A vencer / Saldo total** por cliente y el tiempo
(`vencido hace 10 días`, `vence hoy`, `faltan 5 días`).

- Los clientes marcados en amarillo en el Excel llevan un **`*`** en la primera
  columna, y podés filtrarlos con **"Solo los marcados en amarillo"**.
- Doble clic en un cliente abre el **detalle de comprobantes** con el saldo
  parcial de cada factura.
- Seleccionás clientes (Ctrl / Shift) y los **asignás a un vendedor**. La
  asignación queda guardada en `asignaciones.json`, así que la próxima vez que
  cargues un reporte ya viene resuelta.

### 3. Vendedores
Los números a los que se manda el resumen. El teléfono se normaliza solo:
podés escribir `11 3900-9488` y se guarda como `5491139009488`.

### 4. Enviar
Una fila por vendedor, con la cantidad de clientes y el total. Elegís a quiénes
mandar y:

- **Vista previa** — ves el mensaje exacto, ya partido en la cantidad de
  mensajes de WhatsApp que va a ocupar.
- **Generar solo el PDF** — deja los PDF en la carpeta `Resumenes\`.
- **ENVIAR POR WHATSAPP** — sube el PDF, lo manda y después manda el texto.
- **Abrir en WhatsApp Web** — plan B: abre el chat con el texto ya cargado para
  mandarlo a mano.

Cada vendedor recibe **un PDF con solo sus clientes**, no la cartera completa.

---

## Configuración

### API de Claude
La clave se saca de [console.anthropic.com](https://console.anthropic.com) y
empieza con `sk-ant-`. Se pega en la pestaña **Configuración**.

También se toma de la variable de entorno `ANTHROPIC_API_KEY` si existe.

Modelo por defecto: **`claude-opus-5`**. El costo de interpretar un reporte
típico es de centavos de dólar; el botón **Probar Claude** muestra la estimación.

> **Hace falta saldo cargado.** Una clave válida sin crédito devuelve un error
> HTTP 400 con el texto `credit balance is too low`, que es fácil confundir con
> "el pedido está mal armado". El programa lo detecta y te dice explícitamente
> que hay que cargar crédito en **Plans & Billing**, no te manda a revisar el
> archivo. Con USD 5 alcanza para cientos de reportes.
>
> **Sin saldo, el programa igual sirve:** el lector de Excel es propio y no toca
> la API. Lo único que queda sin funcionar son los PDF (que caen al lector
> propio, menos preciso) y los tres botones de Claude.

#### Límites de la API (importan con reportes grandes)

| Límite | Valor | Qué pasa si se supera |
|---|---|---|
| Tamaño del PDF | 30 MB | Se rechaza antes de enviar, con mensaje claro. |
| Salida de la interpretación | 64.000 tokens (~400-600 clientes) | La respuesta se corta. **Se detecta** y se pide exportar el reporte en partes; no se devuelven datos a medias. |
| Entrada del Excel | 1.500.000 caracteres | Se recorta y **se avisa arriba del reporte que faltan clientes**, con el conteo de filas enviadas. |
| Entrada de la auditoría | 800.000 caracteres | Se recorta y la auditoría **no puede devolver "coincide"**, para no dar falsa tranquilidad. |
| Teléfonos | se procesan en lotes de 60 clientes | Sin límite práctico. |

Los importes que devuelve Claude **nunca se toman como verdad final**: siempre se
controlan contra la fila `Total cliente` y el `Total general` del propio reporte,
y los días de vencimiento se recalculan desde las fechas. Si Claude informa un
número de días que contradice las fechas, **gana la fecha** y queda anotado en los
avisos.

### WhatsApp Cloud API (Meta)
Se necesitan dos datos de [Meta for Developers](https://developers.facebook.com):

- **Token permanente** de la app de WhatsApp Business.
- **ID del número** (`phone_number_id`) del número emisor.

Son los mismos datos que ya usa `WhatsAppMessage`. El botón **Probar WhatsApp**
confirma que andan y muestra el número y el nombre verificado.

> ### ⚠️ La ventana de 24 horas — leer esto
>
> WhatsApp **no permite** mandar mensajes de texto ni documentos libres a alguien
> que no te escribió en las últimas 24 horas. Es una regla de Meta, no del
> programa. Si aparece el error **131047**, es esto.
>
> **Dos soluciones:**
>
> 1. **Modo `libre`** (por defecto, y el que da mejores mensajes): que cada
>    vendedor mande **cualquier mensaje** al número de la empresa antes de que le
>    envíes. Con eso queda la ventana abierta 24 h. Es lo más práctico: son
>    empleados, se les pide una vez al día.
> 2. **Modo `plantilla`**: funciona siempre, incluso con la ventana cerrada, pero
>    hay que crear y hacer aprobar una plantilla en Meta. **Limitación de Meta:
>    los parámetros de plantilla no aceptan saltos de línea**, así que el listado
>    de clientes va todo en una sola línea separada por ` | ` y truncado a 1000
>    caracteres. Se lee bastante peor.
>
> El botón **Abrir en WhatsApp Web** no tiene esta limitación, porque ahí mandás
> vos el mensaje a mano.

---

## Cómo interpreta el archivo

Estructura detectada en el reporte del ERP (columnas del Excel):

| Bloque | Cómo se reconoce | Datos que saca |
|---|---|---|
| Cabecera | filas 1-14 | Empresa, **Fecha del reporte** (J1), filtros (`Cobrador:110-110 Vendedor:002-044`) |
| Cabecera de cliente | col. A = código de 11 dígitos + col. C = razón social | código, razón social, teléfonos (G), localidad (F), dirección y nombre de fantasía (C de las filas siguientes), contacto (I) |
| Comprobante | col. B = tipo (`FE A`) + col. H numérica | talonario (A), tipo (B), número (C), condición de pago (D), F. emisión (E), **F. Vto (F)**, cuota (G), **Importe original (H)**, **Saldo (I)** ← *saldo parcial*, **Días Venc. (J)** |
| Total cliente | col. B empieza con `Total cliente` | importe original y **saldo total** del cliente |
| Total general | `Total general:` | totales de todo el reporte |

**La regla de los días:** `Días Venc. = fecha del reporte − fecha de vencimiento`.

- Positivo → **VENCIDO** (hace esa cantidad de días)
- Cero → **vence hoy**
- Negativo → **A VENCER** (faltan esos días)

En Configuración se puede destildar *"Usar la fecha del reporte"* para que los
días se recalculen **contra hoy**, útil si el reporte es de días atrás.

**El resaltado amarillo** del Excel marca el código y la razón social del
cliente, y la columna Saldo. El programa lo lee (`xlrd` para `.xls`, `openpyxl`
para `.xlsx`) y lo expone como marca de atención y filtro. **No** se usa para
decidir si algo está vencido: eso sale siempre de las fechas, que son
verificables contra los totales del reporte.

---

## Archivos que genera

| Archivo | Contenido |
|---|---|
| `config.json` | Claves y preferencias. **No se sube a git.** |
| `vendedores.json` | Los vendedores cargados. |
| `asignaciones.json` | Qué cliente le corresponde a qué vendedor. |
| `errores_log.txt` | Log técnico. Se abre desde Configuración. |
| `Resumenes\` | Los PDF generados. |
| `Historial\envios.csv` | Constancia de cada envío (fecha, vendedor, clientes, total, resultado). |

---

## Si algo falla

| Síntoma | Causa y solución |
|---|---|
| **131047** al enviar | Ventana de 24 h cerrada. Ver el recuadro de arriba. |
| **131030** | La cuenta de Meta está en modo prueba: hay que agregar el número a la lista de destinatarios permitidos. |
| **190 / 401** | El token de WhatsApp venció. Generar uno nuevo en Meta. |
| **132001** | El nombre de la plantilla no coincide con el de Meta (respeta minúsculas y guiones bajos). |
| "Los totales no cierran" | El parser leyó mal. En Excel es raro; en PDF es esperable: usá **Interpretar con Claude**. |
| "Falta la API key de Claude" | Pegar la clave en Configuración. |
| "La cuenta de Anthropic no tiene crédito" | Cargar saldo en console.anthropic.com → Plans & Billing. El Excel sigue funcionando sin esto. |
| El programa no abre | Correr `interfaz.py` desde una consola para ver el error, o abrir `errores_log.txt`. |

---

## Estado y limitaciones conocidas

**Verificado contra el archivo real `V:\MARKETING\LISTADO VENCIDO.XLS`:**
5 clientes, total **$418.666,62**, descuadre **$0,00** contra el "Total general"
del reporte; resaltado amarillo detectado en los 4 clientes que lo tenían.

**API de Claude — conexión probada, falta cargar saldo.** Con la clave real se
llegó a la API y respondió: la autenticación funciona y el pedido se acepta. Lo
que devolvió es `credit balance is too low`, o sea que la cuenta de Anthropic
todavía no tiene crédito. **Falta entonces verificar la calidad de la
interpretación** (que Claude lea bien un PDF del reporte); todo lo demás del
camino está verificado: el payload no lleva parámetros prohibidos, `effort` va
dentro de `output_config`, el bloque del PDF va antes del texto, si el endpoint
beta rechaza el pedido se reintenta por el estándar, un rechazo por política se
convierte en un error legible, y la conversión de la respuesta a datos internos
es correcta.

**WhatsApp sin probar:** no hay token de Meta cargado, así que **ninguna llamada
real a Meta fue ejecutada**. La primera vez conviene probar con **un solo
vendedor**.

**El parseo determinístico de PDF es best-effort.** Los PDF no tienen columnas
de verdad, así que si el reporte se exportó partido en varias páginas los
importes pueden quedar separados de su comprobante. Cuando eso pasa el control
de totales lo detecta y avisa. **Para PDF, el camino recomendado es Claude.**

---

## Estructura del código

| Archivo | Responsabilidad |
|---|---|
| `modelos.py` | `Reporte` / `Cliente` / `Comprobante`, cálculo de saldos y estados, formato de pesos y normalización de teléfonos argentinos. Es el contrato común. |
| `config.py` | Configuración, rutas persistentes (compatibles con PyInstaller) y log. |
| `parser_excel.py` | Lector determinístico de `.xls` / `.xlsx`, **con colores de relleno**. |
| `parser_pdf.py` | Extracción de texto de PDF y parseo best-effort. |
| `claude_api.py` | API de Claude: interpretación, auditoría, teléfonos y redacción. |
| `whatsapp.py` | Cloud API de Meta, con los errores traducidos al castellano. |
| `mensajes.py` | Armado y partido del texto del WhatsApp. |
| `pdf_resumen.py` | PDF del resumen por vendedor (reportlab). |
| `vendedores.py` | Vendedores y asignación de clientes. |
| `interfaz.py` | Ventana Tkinter. Todo lo lento va en un hilo aparte. |
