# Panel Cepas · Olliari

Panel de seguimiento del proveedor estratégico **CEPAS ARGENTINA S.A.- GANCIA**
(código de proveedor **6** en BSGestión, distinto del proveedor 336 "5 CEPAS
SRL" — no confundir): ventas a nuestros clientes y compras que le hacemos,
con foco en sus 4 marcas principales (**Martini Rosso, Bacardi, Gancia,
Amargo Obrero**), clientes con compra por trimestre (Q1/Q2/avance Q3,
identificados), y una **comparación de la categoría Vermouth completa**
contra la competencia (Cinzano y el resto), independiente de Panel Campari.

Construido replicando la arquitectura de `Panel_Campari` (mismo Excel en
Drive + Streamlit Cloud, mismo Service Account) — ver
`../Panel_Campari/CLAUDE.md` para el detalle general del patrón si hace
falta.

> **⚠️ Pendiente de setup (2026-09-09)**: creado por Claude en esta sesión,
> con `extract_data.py` validado a mano contra la API (corrida de prueba,
> sin subir a Drive), pero **todavía no está desplegado**. Faltan 3 pasos
> manuales que solo puede hacer Sofia (ver "Setup — primera vez" abajo):
> 1. Crear la carpeta "Cepas" en Drive y compartirla con la Service Account.
> 2. Repo de GitHub + Streamlit Community Cloud.
> 3. Tarea programada de Windows para la extracción diaria.
> Hasta que no esté el paso 1, `DRIVE_FOLDER_ID` en `extract_data.py` y
> `dashboard.py` queda en `"PENDIENTE_CREAR_CARPETA_DRIVE"` (placeholder, no
> anda así).

## Hallazgo importante sobre "los demás vermouth" (2026-09-09)

El pedido original mencionaba comparar contra vermouth de **Evo/Ajenjo,
Grupo Peñaflor y Mosquita Muerta** — se revisó la data real de BSGestión
antes de programar nada y **dos de esas tres no son lo que parecían**:

- **Grupo Peñaflor no vende vermouth** en absoluto — su catálogo en
  BSGestión (proveedor 33) es 100% vinos (Alma Mora, Trapiche, Fond de
  Cave, Don David, etc.), gin, whisky, vodka. Ningún vermouth.
- **"Mosquita Muerta" no es una marca de vermouth propia** — es una
  bodega/proyecto cuya línea "Cordero con Piel de Lobo" (que sí incluye un
  `VERMOUTH CORDERO CPL ROSSO 750cc`) factura bajo la razón social
  **Cosmopolitan S.A.** (proveedor 361). Confirmado con Sofia: es así, el
  campo "proveedor" de BSGestión es la razón social que emite la factura,
  no necesariamente la marca comercial — mismo patrón que ya se había visto
  con Cinzano en Panel Campari (factura bajo Campari Argentina aunque
  Sofia indica que hoy lo mueve comercialmente Grupo Peñaflor).
- **Evo & Ajenjo sí es real**: proveedor 552, vende `VERMUT AJENJO
  ROSSO/ROJO 750cc`.

Por eso `clasificar_marca_vermouth()` en `extract_data.py` clasifica por
**texto del nombre de producto**, no por proveedor — así no importa bajo
qué razón social facture cada uno. Las marcas de vermouth que aparecen hoy
en los datos (2024-12 a 2026-09): Martini Rosso y Martini
Bianco/Extra Dry (Cepas), Cinzano y Carpano/Antica Formula (facturan como
Campari Argentina / Fratelli Branca), Vermut Ajenjo (Evo & Ajenjo), Vermu
Lunfa (factura como "5 Cepas SRL" o "Lunfa Vermu SRL" según el mes),
Vermouth Unión Federal (Cosmopolitan), Vermouth La Fuerza (La Agrícola),
Vermouth Vincenzo (Bodegas Esmeralda), Vermouth Cordero/Mosquita Muerta
(Cosmopolitan), y Vermu "Siete Cuatro Seis" (Bodegas y Viñedos López) —
estas últimas muy chicas en volumen.

**Si en el futuro aparece un vermouth de Peñaflor** (lanzamiento nuevo,
por ejemplo) hay que agregarlo a mano en `clasificar_marca_vermouth()` —
no hay forma de que el filtro por Rubro/Categoria lo capture "raro", ya
capta cualquier producto con `Categoria == 'VERMOUTH'` sin importar
proveedor, solo falta que el nombre matchee alguna regla (si no, cae en
"Otros vermouth", que sirve de red de contención — revisar esa categoría
de tanto en tanto por si aparece algo nuevo sin clasificar).

## Arquitectura

```
API BSGestión (empresa '6' = consolidado, red interna/VPN)
        │  solo alcanzable desde la PC de Sofia — NO desde internet
        ▼
extract_data.py  (corre LOCAL, 1 vez por día vía tarea programada de Windows)
        │  trae ventas + compras del proveedor Cepas (código 6), y ADEMÁS
        │  trae aparte TODA la categoría Vermouth (todos los proveedores)
        ▼
Google Drive — carpeta "Cepas" del proyecto
        │  panel_cepas.xlsx (Datos Ventas, Datos Compras, Datos Vermouth, Metadata)
        ▼
dashboard.py  (Streamlit Community Cloud — URL pública con contraseña)
        │  lee panel_cepas.xlsx vía Service Account, todo se recalcula
        │  en vivo con pandas según los filtros que elija el usuario
        ▼
Usuarios (link + contraseña compartida)
```

## Componentes

| Archivo | Dónde corre | Qué hace |
|---|---|---|
| `extract_data.py` | Local (PC de Sofia, tarea programada) | Trae ventas Cepas + compras Cepas + ventas de TODA la categoría Vermouth desde BSGestión, arma `panel_cepas.xlsx` y lo sube a Drive |
| `google_drive_auth.py` | Streamlit Cloud | Descarga el Excel de Drive usando `st.secrets["gcp_service_account"]` |
| `metrics.py` | Streamlit Cloud | Funciones pandas compartidas: filtros, YTD, evolución, mix por marca, clientes por marca/trimestre, comparación de vermouth |
| `dashboard.py` | Streamlit Cloud | La app en sí — password, filtros, tabs |

## De dónde sale cada dato

### Ventas — `gr_reporteFacturacionPorItem('6', desde, hasta, 'VEN')`

Mismo stored procedure y mismo patrón de tramos de 90 días que Panel
Campari. De **cada tramo** se sacan DOS recortes (una sola llamada a la
API, dos filtros locales — no se pide dos veces lo mismo):

- **`Datos Ventas`**: `codigoProveedor == 6` (CEPAS ARGENTINA S.A.- GANCIA;
  confirmado 2026-09-09, no confundir con proveedor 336 "5 CEPAS SRL", que
  es otra empresa que vende Vermu Lunfa). Clasificado por
  `clasificar_marca()`: Martini Rosso, Martini (otras variantes: Bianco,
  Extra Dry), Bacardi, Gancia, Amargo Obrero, Otras marcas Cepas (Gin
  Bombay, Jack Daniel's, Ron Jamaica, Fernet 1882, Vodka Nikov, Amarula,
  vinos Omnium/Chances, etc.).
- **`Datos Vermouth`**: `Rubro == 'BEBIDAS CON ALCOHOL'` y
  `Categoria == 'VERMOUTH'`, **sin filtrar por proveedor** — clasificado
  por `clasificar_marca_vermouth()` (ver hallazgo arriba). Nota: Gancia y
  Amargo Obrero están categorizados como "APERITIVOS" en BSGestión, no
  "VERMOUTH" — por eso no aparecen en esta hoja, solo Martini sí.

### Compras — mismo patrón SELECT directo que Panel Campari

`labelCliente_Proveedor LIKE '%CEPAS ARGENTINA%'` (confirmado: en compras
aparece como `(00006) -CEPAS ARGENTINA S.A.- GANCIA`). Solo se traen
compras al proveedor Cepas — no tiene sentido traer compras de la
competencia de vermouth (no les compramos a ellos, son otros proveedores
del grupo).

### Fecha de inicio

`FECHA_DESDE_API = '2024-12-29'` — misma fecha que Panel Campari, por
consistencia.

## Decisiones de diseño

**Por qué "Martini Rosso" separado de "Martini (otras variantes)":** el
pedido original nombra específicamente "Martini Rosso" como una de las 4
marcas principales de Cepas (junto a Bacardi/Gancia/Amargo Obrero), y la
comparación de vermouth pide explícitamente "Martini Rosso vs Cinzano" —
Cepas también vende Martini Bianco y Martini Extra Dry, pero esos quedan en
un bucket aparte para no inflar el número de Martini Rosso ni mezclar
peras con manzanas en la comparación cara a cara contra Cinzano (que
también es específicamente la variante Rosso/Bianco, no un "Cinzano
genérico").

**Sin pestaña "Objetivos" (a diferencia de Panel Campari):** Cepas
Argentina no pasa targets trimestrales tipo PPT como sí hace Campari (Sell
Out/Sell In/Marca x PDV/Cobertura) — no hay `OBJETIVOS_CEPAS` en
`metrics.py` ni pestaña Objetivos en el dashboard. Si en el futuro Cepas
empieza a pasar objetivos, replicar el patrón de
`Panel_Campari/metrics.py::objetivos_trimestre()`.

**Clientes por marca y trimestre, con detalle (pestaña Clientes):**
`clientes_por_marca_trimestre()` en `metrics.py` arma, para cada una de las
4 marcas principales, el listado de clientes distintos que compraron en
Q1, Q2 y lo que va de Q3 del año en curso (Q3 se corta a la fecha de hoy,
marcado como "parcial" — no es un trimestre cerrado). El dashboard muestra
primero una tabla resumen (cantidad por marca x trimestre) y después un
selector Marca + Trimestre para ver el listado completo de clientes
(razón social) de esa combinación puntual — a diferencia de Panel Campari,
que solo mostraba conteos agregados, acá el pedido explícito fue "cantidad
e identificarlos".

**Comparación de vermouth con overlap de clientes:** además del ranking de
unidades/clientes por marca, `clientes_overlap()` calcula cuántos clientes
compran SOLO Martini Rosso, SOLO Cinzano, o las dos — pensado como
oportunidad comercial (clientes que ya tienen Cinzano en el surtido son
candidatos naturales para ofrecerles Martini Rosso, y viceversa), no fue
pedido explícito pero se agregó porque la pregunta original ("a cuántos le
vendieron uno y el otro") lo pide implícitamente.

## Setup — primera vez

### 1. Google Drive

Crear una carpeta nueva en Drive (por ejemplo "Cepas", al lado de la
carpeta "Campari" que ya existe) y compartirla como **Editor** con la
Service Account:

```
cheques@horacio-pagos.iam.gserviceaccount.com
```

(la misma que ya usan Panel_Campari / Olliari-Panel-Locales /
Rentabilidad_impresos — el `service_account.json` de esta carpeta ya es
una copia del mismo archivo, no hace falta crear una Service Account
nueva.)

Copiar el ID de esa carpeta (de la URL
`drive.google.com/drive/folders/<ID>`) y reemplazar el placeholder
`"PENDIENTE_CREAR_CARPETA_DRIVE"` en **dos lugares**:
- `extract_data.py` → `DRIVE_FOLDER_ID`
- `dashboard.py` → `DRIVE_FOLDER_ID`

**Nota (2026-09-09):** el Service Account no tiene forma de crear esta
carpeta él solo de manera que Sofia la vea (no comparte automáticamente
con una carpeta raíz visible) — confirmado explorando qué carpetas puede
listar (`Campari`, `Panel_diario`, `Tablero_Vendedores`, etc., todas
compartidas individualmente por Sofia). Por eso este paso es manual.

### 2. `service_account.json`

Ya copiado a esta carpeta (mismo archivo que Panel_Campari). Está en
`.gitignore`, nunca se commitea.

### 3. Repo GitHub + Streamlit Community Cloud

1. Crear un repo nuevo (privado) y pushear esta carpeta (sin
   `service_account.json` ni `panel_cepas.xlsx`, ya excluidos por
   `.gitignore`).
2. En [share.streamlit.io](https://share.streamlit.io): **New app** →
   apuntar al repo → main file `dashboard.py`.
3. **Settings → Secrets**, pegar (mismo formato que Panel_Campari — copiar
   el `[gcp_service_account]` completo de los secrets de
   `Olliari-Panel-Locales` o `Panel_Campari`):

```toml
PASSWORD = "elegir-una-contraseña"

[gcp_service_account]
type = "service_account"
project_id = "horacio-pagos"
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "cheques@horacio-pagos.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
```

### 4. Tarea programada de Windows

Copiar `run_extract.bat` a una ruta **sin espacios** (mismo bug ya
documentado en Panel_Campari — `schtasks /tr` con espacios en la ruta tira
`ERROR_FILE_NOT_FOUND` aunque la tarea se cree sin error):

```
C:\Users\Sofia\run_extract_cepas.bat
```

Para crear la tarea (correr una sola vez, **después** de terminar los
pasos 1 y 2):

```powershell
schtasks /create /tn "Olliari - Panel Cepas" /tr "C:\Users\Sofia\run_extract_cepas.bat" /sc daily /st 08:15 /f
```

(08:15, no 08:00, para no pisarse con la tarea de Panel Campari que ya
corre a esa hora.)

Para probarla a mano:

```powershell
schtasks /run /tn "Olliari - Panel Cepas"
```

## Uso diario

- El panel se actualiza solo, 1 vez por día, mientras la tarea programada
  esté corriendo y la VPN Radmin esté activa en ese momento.
- Para forzar una actualización fuera de horario: `python extract_data.py`
  a mano desde esta carpeta (con Anaconda, con VPN activa).
- Revisar `extract_log.txt` si el panel deja de actualizarse.

## Troubleshooting

Ver `../Panel_Campari/CLAUDE.md`, sección Troubleshooting — los mismos
problemas (archivo no encontrado en Drive, error de conexión VPN,
BrokenPipeError transitorio, versiones de `requirements.txt`) aplican acá
con la misma causa y la misma solución, cambiando "Campari" por "Cepas".
