# -*- coding: utf-8 -*-
"""
extract_data.py — Corre LOCAL (tarea programada diaria, con VPN Radmin activa).

Trae ventas + compras del proveedor CEPAS ARGENTINA S.A.- GANCIA (Martini,
Bacardi, Gancia, Amargo Obrero, Gin Bombay, Jack Daniel's, Ron Jamaica,
Fernet 1882, etc.) desde BSGestión, y ADEMÁS trae, aparte, TODAS las ventas
de la categoría "VERMOUTH" (rubro "BEBIDAS CON ALCOHOL") sin importar el
proveedor — para poder comparar Martini Rosso (Cepas) contra Cinzano,
Carpano, Ajenjo, Lunfa, Unión Federal, La Fuerza, Vincenzo, Cordero (Mosquita
Muerta) y el resto de las marcas de vermouth que vendemos. Deja todo
consolidado en panel_cepas.xlsx, que sube/actualiza en Google Drive.

Mismo patrón que Panel_Campari (ver ese CLAUDE.md para el detalle de la
arquitectura general) — la diferencia es que acá hay un tercer dataset
("Datos Vermouth") que NO se filtra por proveedor, sino por categoría de
producto, porque el pedido es comparar Cepas contra la competencia dentro
de la misma categoría, no solo mirar el proveedor Cepas aislado.

Requiere en esta misma carpeta:
  - service_account.json   (misma Service Account que usan
    Panel_Campari / Olliari-Panel-Locales / Rentabilidad_impresos)

Ventas: CALL gr_reporteFacturacionPorItem('6', desde, hasta, 'VEN')
  — empresa '6' = consolidado de las 4 empresas del grupo. Se pide en tramos
  de 90 días. De cada tramo se sacan DOS recortes: uno por codigoProveedor=6
  (Cepas Argentina) y otro por Rubro/Categoria = Vermouth (todos los
  proveedores) — así no hace falta pedirle dos veces lo mismo a la API.

Compras: mismo patrón SELECT directo que Panel_Campari (el SP
  gr_FacturacionComprasItemGrupo no funciona, confirmado ahí). Solo se traen
  compras a Cepas Argentina (no a la competencia de vermouth — no tiene
  sentido, no le compramos a ellos).

Nota sobre "proveedor" vs "marca comercial real" (igual que con Cinzano en
Campari): el campo `proveedor`/`labelCliente_Proveedor` de BSGestión es la
razón social que emite la factura, no necesariamente quién es dueño de la
marca hoy. Ej.: "VERMOUTH CORDERO CPL ROSSO" factura bajo "COSMOPOLITAN
S.A." pero es la línea Cordero de la bodega Mosquita Muerta; Cinzano puede
facturar bajo Campari Argentina aunque comercialmente hoy lo mueva Grupo
Peñaflor. Por eso la clasificación de marca (`clasificar_marca_vermouth`)
se hace por texto del nombre de producto, no por el campo proveedor.
"""

import re
import sys
from datetime import datetime

import pandas as pd
import requests

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ══════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════════════

# TODO (setup inicial): reemplazar por el ID real de la carpeta de Drive
# "Cepas" creada para este proyecto y compartida como Editor con la Service
# Account cheques@horacio-pagos.iam.gserviceaccount.com (ver CLAUDE.md).
DRIVE_FOLDER_ID = "1al-GEctDtC-oIgXWPzlmxNqzKlp30wmq"
SERVICE_ACCOUNT_FILE = "service_account.json"
PANEL_FILENAME = "panel_cepas.xlsx"

URL_NOVARA = "http://26.169.81.157:8080/BSGestion/rest/consultar"
HEADERS_NOVARA = {"bduser": "consultas", "bdpass": "consultas1234"}
EMPRESA_CONSOLIDADO = "6"

CODIGO_PROVEEDOR_CEPAS = 6
PROVEEDOR_CEPAS_LABEL = "CEPAS ARGENTINA S.A.- GANCIA"

# Igual que Campari, arranca a fin de diciembre 2024 — ampliar/acortar acá
# si hace falta (afecta el tiempo que tarda la extracción).
FECHA_DESDE_API = "2024-12-29"

TIPOS_VENTAS_OK = ["FAC"]
TIPOS_VENTAS_NC = ["NC"]

TIPOS_COMPRA_IDS = (67, 43, 80, 98, 46, 88, 17, 68)

LINEAS_NO_PRODUCTO = [
    "IMPUESTOS INTERNOS", "IVA PERCEPCIONES", "PERCEPCIONES INGRESOS BRUTOS",
    "DESCUENTOS COMERCIAL COMPRA", "PERCEPCION IVA", "PERCEPCION IIBB",
]

RUBRO_ALCOHOL = "BEBIDAS CON ALCOHOL"
CATEGORIA_VERMOUTH = "VERMOUTH"


# ── Clasificación de marca (dentro del proveedor Cepas) ────────────────────
# El campo "Marca" que trae la API no siempre distingue submarca — la forma
# confiable de separar Martini/Bacardi/Gancia/Amargo Obrero del resto del
# portafolio de Cepas (Gin Bombay, Jack Daniel's, Ron Jamaica, Fernet 1882,
# Vodka Nikov, Amarula, vinos Omnium/Chances) es el nombre del producto.
def clasificar_marca(producto):
    p = (producto or "").upper()
    if any(np in p for np in LINEAS_NO_PRODUCTO):
        return "Impuestos/Otros (no producto)"
    if "MARTINI ROSSO" in p:
        return "Martini Rosso"
    if "MARTINI" in p:
        return "Martini (otras variantes)"
    if "BACARDI" in p:
        return "Bacardi"
    if "GANCIA" in p:
        return "Gancia"
    if "AMARGO OBRERO" in p:
        return "Amargo Obrero"
    return "Otras marcas Cepas"  # Gin Bombay, Jack Daniel's, Ron Jamaica, Fernet 1882, etc.


# ── Clasificación de marca DENTRO de la categoría Vermouth (todo el rubro,
# cualquier proveedor) — para la comparación Martini Rosso vs Cinzano vs el
# resto. Confirmado contra datos reales (2026-09): las marcas de vermouth que
# vendemos hoy son Martini (Cepas), Cinzano/Carpano (facturan bajo Campari /
# Fratelli Branca), Vermut Ajenjo (Evo & Ajenjo), Vermu Lunfa (5 Cepas SRL /
# Lunfa Vermu SRL), Vermouth Unión Federal (Cosmopolitan), Vermouth La
# Fuerza (La Agrícola), Vermouth Vincenzo (Bodegas Esmeralda), Vermouth
# Cordero (línea de Mosquita Muerta, factura como Cosmopolitan S.A.),
# Vermouth Antica Formula (Fratelli Branca) y Vermu Rojo "Siete Cuatro Seis"
# (Bodegas y Viñedos López). Grupo Peñaflor y "Mosquita Muerta" como tal NO
# tienen un vermouth propio con otro nombre en los datos — Mosquita Muerta
# es la línea Cordero de Cosmopolitan.
def clasificar_marca_vermouth(producto):
    p = (producto or "").upper()
    if "MARTINI ROSSO" in p:
        return "Martini Rosso"
    if "MARTINI" in p:
        return "Martini (otras variantes)"
    if "CINZANO" in p and "ROSSO" in p:
        return "Cinzano Rosso"
    if "CINZANO" in p:
        return "Cinzano (otras variantes)"
    if "CARPANO" in p:
        return "Carpano"
    if "ANTICA FORMULA" in p:
        return "Antica Formula"
    if "AJENJO" in p:
        return "Ajenjo (Evo & Ajenjo)"
    if "LUNFA" in p:
        return "Lunfa"
    if "UNION FEDERAL" in p or "UNIÓN FEDERAL" in p:
        return "Unión Federal"
    if "LA FUERZA" in p:
        return "La Fuerza"
    if "VINCENZO" in p:
        return "Vincenzo"
    if "CORDERO" in p:
        return "Cordero (Mosquita Muerta)"
    if "SIETE CUATRO SEIS" in p or "746" in p:
        return "Siete Cuatro Seis"
    return "Otros vermouth"


def nombre_fantasia(cliente_raw):
    """'(01179) - JUANOJUAN S.A. - GOMERIA JUAN & JUAN' -> 'GOMERIA JUAN & JUAN'."""
    partes = [p.strip() for p in str(cliente_raw or "").split(" - ")]
    return " - ".join(partes[2:]) if len(partes) > 2 else ""


def normalizar_producto(articulo):
    """'BV1899 | CAMPARI 1000cc' -> 'CAMPARI 1000cc'; 'BV1899 - CAMPARI 1000cc'
    -> 'CAMPARI 1000cc'."""
    p = (articulo or "").strip()
    if " | " in p:
        return p.split(" | ")[-1].strip()
    m = re.match(r"^[A-Z]{2}\d+\s*-\s*(.+)$", p)
    if m:
        return m.group(1).strip()
    return p


# ══════════════════════════════════════════════════════════════════════════
# DRIVE
# ══════════════════════════════════════════════════════════════════════════

def _drive_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)


def drive_find_file_id(service, folder_id, filename):
    q = f"'{folder_id}' in parents and name = '{filename}' and trashed = false"
    res = service.files().list(q=q, spaces="drive", fields="files(id)").execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None


def drive_upload_or_update(service, folder_id, filename, local_path):
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    media = MediaIoBaseUpload(open(local_path, "rb"), mimetype=mime, resumable=True)
    existing_id = drive_find_file_id(service, folder_id, filename)
    if existing_id:
        service.files().update(fileId=existing_id, media_body=media).execute()
        return existing_id
    try:
        metadata = {"name": filename, "parents": [folder_id]}
        created = service.files().create(body=metadata, media_body=media, fields="id").execute()
        return created["id"]
    except Exception as e:
        if "storageQuotaExceeded" in str(e):
            raise SystemExit(
                f"No pude crear '{filename}' en Drive porque la Service Account no tiene "
                "cuota propia. Solución (una sola vez): crear manualmente un archivo llamado "
                f"'{filename}' (aunque esté vacío) en esa carpeta de Drive, y volver a correr "
                "este script — a partir de ahí sí lo puede actualizar."
            )
        raise


# ══════════════════════════════════════════════════════════════════════════
# API BSGestión
# ══════════════════════════════════════════════════════════════════════════

def api_call(query, timeout=300):
    r = requests.post(URL_NOVARA, headers=HEADERS_NOVARA, json={"query": query}, timeout=timeout)
    r.raise_for_status()
    return r.json()


def _armar_df_ventas(df):
    df = df[df["Tipo comprobante"].isin(TIPOS_VENTAS_OK + TIPOS_VENTAS_NC)].copy()
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    df["Producto"] = df["Articulo"].apply(normalizar_producto)
    df["NombreFantasia"] = df["Cliente"].apply(nombre_fantasia)
    df["Cantidad"] = pd.to_numeric(df["Cantidad"], errors="coerce").fillna(0)
    df["PrecioUnitario"] = pd.to_numeric(df["PrecioUnitario"], errors="coerce").fillna(0)
    df["Total"] = pd.to_numeric(df["PrecioTotal"], errors="coerce").fillna(0)
    df["TotalNeto"] = pd.to_numeric(df["PrecioTotalNeto"], errors="coerce").fillna(0)
    df = df.rename(columns={"empresa": "Empresa"})
    return df


def fetch_ventas(fecha_desde_str, chunk_dias=90):
    """Trae ventas en tramos de 90 días y arma DOS datasets a partir del
    mismo pull (para no pedirle dos veces lo mismo a la API):
    - ventas_cepas: codigoProveedor == 6 (CEPAS ARGENTINA S.A.- GANCIA),
      clasificadas por clasificar_marca().
    - ventas_vermouth: Rubro=BEBIDAS CON ALCOHOL, Categoria=VERMOUTH, TODOS
      los proveedores, clasificadas por clasificar_marca_vermouth()."""
    fecha_desde = pd.Timestamp(fecha_desde_str)
    fecha_hasta = pd.Timestamp.now().normalize()
    partes_cepas, partes_verm = [], []
    cursor = fecha_desde
    while cursor <= fecha_hasta:
        cursor_fin = min(cursor + pd.Timedelta(days=chunk_dias), fecha_hasta)
        print(f"  Ventas {cursor.date()} -> {cursor_fin.date()}...")
        query = (f"CALL gr_reporteFacturacionPorItem('{EMPRESA_CONSOLIDADO}', "
                  f"'{cursor.strftime('%Y-%m-%d')}', '{cursor_fin.strftime('%Y-%m-%d')}', 'VEN')")
        filas = api_call(query)
        if filas:
            df_tramo = pd.DataFrame(filas)

            df_cepas = df_tramo[df_tramo["codigoProveedor"] == CODIGO_PROVEEDOR_CEPAS]
            if not df_cepas.empty:
                partes_cepas.append(df_cepas)

            df_verm = df_tramo[(df_tramo["Rubro"] == RUBRO_ALCOHOL)
                                & (df_tramo["Categoria"] == CATEGORIA_VERMOUTH)]
            if not df_verm.empty:
                partes_verm.append(df_verm)
        cursor = cursor_fin + pd.Timedelta(days=1)

    cols_cepas = ["Fecha", "Empresa", "Cliente", "NombreFantasia", "Vendedor", "Producto", "Marca",
                  "Categoria", "NroComprobante", "Tipo", "Cantidad", "PrecioUnitario",
                  "Total", "TotalNeto"]
    if partes_cepas:
        ventas_cepas = pd.concat(partes_cepas, ignore_index=True)
        ventas_cepas = _armar_df_ventas(ventas_cepas)
        ventas_cepas["Tipo"] = ventas_cepas["Tipo comprobante"].map({"FAC": "FC", "NC": "NC"}).fillna(ventas_cepas["Tipo comprobante"])
        ventas_cepas["Marca"] = ventas_cepas["Producto"].apply(clasificar_marca)
        for c in cols_cepas:
            if c not in ventas_cepas.columns:
                ventas_cepas[c] = pd.NA
        ventas_cepas = ventas_cepas[cols_cepas].sort_values("Fecha").reset_index(drop=True)
    else:
        ventas_cepas = pd.DataFrame(columns=cols_cepas)

    cols_verm = ["Fecha", "Empresa", "Cliente", "NombreFantasia", "Vendedor", "Producto",
                 "MarcaVermouth", "Proveedor", "NroComprobante", "Tipo", "Cantidad",
                 "PrecioUnitario", "Total", "TotalNeto"]
    if partes_verm:
        ventas_verm = pd.concat(partes_verm, ignore_index=True)
        ventas_verm = _armar_df_ventas(ventas_verm)
        ventas_verm["Tipo"] = ventas_verm["Tipo comprobante"].map({"FAC": "FC", "NC": "NC"}).fillna(ventas_verm["Tipo comprobante"])
        ventas_verm["MarcaVermouth"] = ventas_verm["Producto"].apply(clasificar_marca_vermouth)
        ventas_verm = ventas_verm.rename(columns={"proveedor": "Proveedor"})
        for c in cols_verm:
            if c not in ventas_verm.columns:
                ventas_verm[c] = pd.NA
        ventas_verm = ventas_verm[cols_verm].sort_values("Fecha").reset_index(drop=True)
    else:
        ventas_verm = pd.DataFrame(columns=cols_verm)

    return ventas_cepas, ventas_verm


def fetch_compras_cepas(fecha_desde_str, chunk_dias=180):
    """Compras a Cepas Argentina — mismo patrón SELECT directo que
    Panel_Campari (el SP gr_FacturacionComprasItemGrupo no funciona)."""
    fecha_desde = pd.Timestamp(fecha_desde_str)
    fecha_hasta = pd.Timestamp.now().normalize()
    partes = []
    cursor = fecha_desde
    tipos_str = ", ".join(str(t) for t in TIPOS_COMPRA_IDS)
    while cursor <= fecha_hasta:
        cursor_fin = min(cursor + pd.Timedelta(days=chunk_dias), fecha_hasta)
        print(f"  Compras {cursor.date()} -> {cursor_fin.date()}...")
        query = f"""
            SELECT
                c.fecha                                AS Fecha,
                c.labelCliente_Proveedor               AS Proveedor,
                c.idEmpresa                            AS IdEmpresa,
                c.numeroPuntoVenta                     AS PuntoVenta,
                c.nroComprobante                       AS NroComprobante,
                t.codigo                                AS Tipo,
                t.coeficiente                           AS Coeficiente,
                p.descripcion                           AS Producto,
                i.cantidad                              AS Cantidad,
                i.precioUnitario                        AS PrecioUnitario,
                i.importeTotalIva_Desc * t.coeficiente  AS Total
            FROM ac_co_din_comprobantes c
            JOIN ac_co_ref_tiposcomprobante t  ON t.ID  = c.idTipoComprobante
            JOIN ac_co_din_itemscomprobante i  ON i.IDCOMPROBANTE = c.ID
            JOIN pr_din_producto p             ON p.ID  = i.IDPRODUCTO
            WHERE c.fecha >= '{cursor.strftime('%Y-%m-%d')}'
            AND c.fecha <= '{cursor_fin.strftime('%Y-%m-%d')}'
            AND c.anulado = 0
            AND c.idTipoComprobante IN ({tipos_str})
            AND i.IDPRODUCTO > 0
            AND c.labelCliente_Proveedor LIKE '%CEPAS ARGENTINA%'
        """
        filas = api_call(query, timeout=180)
        if filas:
            partes.append(pd.DataFrame(filas))
        cursor = cursor_fin + pd.Timedelta(days=1)

    if not partes:
        return pd.DataFrame()

    df = pd.concat(partes, ignore_index=True)
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    df["Producto"] = df["Producto"].apply(normalizar_producto)
    df["Marca"] = df["Producto"].apply(clasificar_marca)
    df["Cantidad"] = pd.to_numeric(df["Cantidad"], errors="coerce").fillna(0)
    df["PrecioUnitario"] = pd.to_numeric(df["PrecioUnitario"], errors="coerce").fillna(0)
    df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0)
    df["NroComp"] = df.apply(
        lambda r: f"{int(r['PuntoVenta']):04d}-{int(r['NroComprobante']):08d}"
        if pd.notna(r["PuntoVenta"]) and pd.notna(r["NroComprobante"]) else "",
        axis=1,
    )
    cols = ["Fecha", "IdEmpresa", "NroComp", "Tipo", "Producto", "Marca",
            "Cantidad", "PrecioUnitario", "Total"]
    df = df[cols].sort_values("Fecha").reset_index(drop=True)
    return df


# ══════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════

def main():
    print(f"=== Panel Cepas · extract_data.py · {datetime.now().strftime('%d/%m/%Y %H:%M')} ===")

    print("\nDescargando ventas (Cepas Argentina + categoría Vermouth completa)...")
    ventas, ventas_vermouth = fetch_ventas(FECHA_DESDE_API)
    print(f"  Ventas Cepas: {len(ventas):,} filas")
    if not ventas.empty:
        print(f"  Rango: {ventas['Fecha'].min().date()} -> {ventas['Fecha'].max().date()}")
        print(f"  Facturación total (con IVA): $ {ventas['Total'].sum():,.0f}")
    print(f"  Ventas Vermouth (todos los proveedores): {len(ventas_vermouth):,} filas")
    if not ventas_vermouth.empty:
        print(f"  Marcas de vermouth detectadas: {sorted(ventas_vermouth['MarcaVermouth'].unique())}")

    print("\nDescargando compras a Cepas Argentina...")
    compras = fetch_compras_cepas(FECHA_DESDE_API)
    print(f"  Total compras: {len(compras):,} filas")
    if not compras.empty:
        print(f"  Rango: {compras['Fecha'].min().date()} -> {compras['Fecha'].max().date()}")
        print(f"  Compras totales: $ {compras['Total'].sum():,.0f}")

    if ventas.empty and compras.empty and ventas_vermouth.empty:
        print("\nSin datos — no se genera ni sube el Excel.")
        sys.exit(1)

    local_path = PANEL_FILENAME
    with pd.ExcelWriter(local_path, engine="openpyxl") as writer:
        ventas.to_excel(writer, sheet_name="Datos Ventas", index=False)
        compras.to_excel(writer, sheet_name="Datos Compras", index=False)
        ventas_vermouth.to_excel(writer, sheet_name="Datos Vermouth", index=False)
        meta = pd.DataFrame([{
            "Actualizado": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Proveedor": PROVEEDOR_CEPAS_LABEL,
            "Fecha desde": FECHA_DESDE_API,
            "Filas Ventas": len(ventas),
            "Filas Compras": len(compras),
            "Filas Vermouth": len(ventas_vermouth),
        }])
        meta.to_excel(writer, sheet_name="Metadata", index=False)
    print(f"\nArchivo local generado: {local_path}")

    print("Subiendo a Drive...")
    service = _drive_service()
    file_id = drive_upload_or_update(service, DRIVE_FOLDER_ID, PANEL_FILENAME, local_path)
    print(f"Listo. {PANEL_FILENAME} actualizado en Drive (id={file_id})")


if __name__ == "__main__":
    main()
