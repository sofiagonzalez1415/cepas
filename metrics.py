"""
metrics.py — Funciones de agregación (pandas) para el dashboard de Panel Cepas.
Todo se recalcula en vivo a partir de las hojas "Datos Ventas" / "Datos Compras"
/ "Datos Vermouth" de panel_cepas.xlsx, según los filtros que elija el usuario
en el dashboard. Varias funciones son las mismas (o casi) que en
Panel_Campari/metrics.py — ver ese archivo si hace falta comparar.
"""

import pandas as pd

ORDEN_MARCAS = ["Martini Rosso", "Bacardi", "Gancia", "Amargo Obrero",
                 "Martini (otras variantes)", "Otras marcas Cepas",
                 "Impuestos/Otros (no producto)"]

# Las 4 marcas principales de Cepas que Olliari sigue de cerca.
MARCAS_PRINCIPALES = ["Martini Rosso", "Bacardi", "Gancia", "Amargo Obrero"]

# Orden "natural" para mostrar la comparación de vermouth — Martini Rosso y
# Cinzano primero (los dos grandes que se quieren comparar cara a cara),
# después el resto de marcas chicas.
ORDEN_MARCAS_VERMOUTH = ["Martini Rosso", "Cinzano Rosso", "Martini (otras variantes)",
                          "Cinzano (otras variantes)", "Carpano", "Antica Formula",
                          "Ajenjo (Evo & Ajenjo)", "Lunfa", "Unión Federal", "La Fuerza",
                          "Vincenzo", "Cordero (Mosquita Muerta)", "Siete Cuatro Seis", "Otros vermouth"]

FECHA_DESDE_GRAFICOS = pd.Timestamp("2025-01-01")


def var(a, b):
    """Variación porcentual de b sobre a. None si a es 0/NaN."""
    if a in (0, None) or pd.isna(a):
        return None
    return (b - a) / a


def trimestre(fecha):
    return (fecha.month - 1) // 3 + 1


def filtrar(df, marcas=None, productos=None, fecha_desde=None, fecha_hasta=None,
            excluir_no_producto=True, col_marca="Marca"):
    out = df.copy()
    if excluir_no_producto and col_marca in out.columns:
        out = out[out[col_marca] != "Impuestos/Otros (no producto)"]
    if marcas:
        out = out[out[col_marca].isin(marcas)]
    if productos:
        out = out[out["Producto"].isin(productos)]
    if fecha_desde is not None:
        out = out[out["Fecha"] >= pd.Timestamp(fecha_desde)]
    if fecha_hasta is not None:
        out = out[out["Fecha"] <= pd.Timestamp(fecha_hasta)]
    return out


# ── YTD ──────────────────────────────────────────────────────────────────

def resumen_ytd(ventas, año, fecha_corte=None):
    """Compara YTD del año dado contra el mismo período (mismo día-del-año)
    del año anterior."""
    fecha_corte = fecha_corte or pd.Timestamp.now().normalize()
    doy = fecha_corte.dayofyear

    def _slice(año_x):
        ini = pd.Timestamp(year=año_x, month=1, day=1)
        fin = ini + pd.Timedelta(days=doy - 1)
        d = ventas[(ventas["Fecha"] >= ini) & (ventas["Fecha"] <= fin)]
        return d["Total"].sum(), d["Cantidad"].sum()

    total_actual, unid_actual = _slice(año)
    total_prev, unid_prev = _slice(año - 1)
    return {
        "total_actual": total_actual, "total_prev": total_prev,
        "unid_actual": unid_actual, "unid_prev": unid_prev,
        "var_total": var(total_prev, total_actual),
        "var_unid": var(unid_prev, unid_actual),
        "precio_prom_actual": (total_actual / unid_actual) if unid_actual else 0,
        "precio_prom_prev": (total_prev / unid_prev) if unid_prev else 0,
        "fecha_corte": fecha_corte,
    }


# ── Evolución mensual / trimestral ──────────────────────────────────────

def evolucion_mensual(ventas, n_meses=13):
    d = ventas[ventas["Fecha"] >= FECHA_DESDE_GRAFICOS].copy()
    d["Periodo"] = d["Fecha"].dt.to_period("M")
    g = d.groupby("Periodo").agg(Total=("Total", "sum"), Unidades=("Cantidad", "sum")).reset_index()
    g = g.sort_values("Periodo").tail(n_meses)
    g["Etiqueta"] = g["Periodo"].astype(str)
    return g


def evolucion_trimestral(ventas, aplicar_corte=True):
    d = ventas[ventas["Fecha"] >= FECHA_DESDE_GRAFICOS] if aplicar_corte else ventas
    d = d.copy()
    d["Año"] = d["Fecha"].dt.year
    d["Trim"] = d["Fecha"].apply(trimestre)
    g = d.groupby(["Año", "Trim"]).agg(Total=("Total", "sum"), Unidades=("Cantidad", "sum")).reset_index()
    g = g.sort_values(["Año", "Trim"]).reset_index(drop=True)
    g["Etiqueta"] = g["Año"].astype(str) + "-Q" + g["Trim"].astype(str)
    g["VarTotal"] = g["Total"].pct_change()
    g["VarUnidades"] = g["Unidades"].pct_change()
    g = g.merge(
        g[["Año", "Trim", "Total", "Unidades"]].assign(Año=lambda x: x["Año"] + 1)
         .rename(columns={"Total": "TotalAñoAnt", "Unidades": "UnidadesAñoAnt"}),
        on=["Año", "Trim"], how="left",
    )
    g["VarTotalYoY"] = g.apply(lambda r: var(r["TotalAñoAnt"], r["Total"]), axis=1)
    g["VarUnidadesYoY"] = g.apply(lambda r: var(r["UnidadesAñoAnt"], r["Unidades"]), axis=1)
    return g.drop(columns=["TotalAñoAnt", "UnidadesAñoAnt"])


def evolucion_mensual_por_marca(ventas, n_meses=13, col_marca="Marca"):
    d = ventas[ventas["Fecha"] >= FECHA_DESDE_GRAFICOS].copy()
    d["Periodo"] = d["Fecha"].dt.to_period("M")
    g = d.groupby(["Periodo", col_marca]).agg(Total=("Total", "sum"), Unidades=("Cantidad", "sum")).reset_index()
    periodos = sorted(d["Periodo"].unique())[-n_meses:]
    g = g[g["Periodo"].isin(periodos)]
    g["Etiqueta"] = g["Periodo"].astype(str)
    return g.sort_values("Periodo")


# ── Mix por marca ────────────────────────────────────────────────────────

def mix_por_marca(ventas, fecha_desde=None, fecha_hasta=None, col_marca="Marca", orden=None):
    d = filtrar(ventas, fecha_desde=fecha_desde, fecha_hasta=fecha_hasta, col_marca=col_marca)
    g = d.groupby(col_marca).agg(Total=("Total", "sum"), Unidades=("Cantidad", "sum"),
                                  Clientes=("Cliente", "nunique")).reset_index()
    total_gral = g["Unidades"].sum()
    g["Participacion"] = g["Unidades"] / total_gral if total_gral else 0
    orden = orden or ORDEN_MARCAS
    orden_map = {m: i for i, m in enumerate(orden)}
    g["_orden"] = g[col_marca].map(orden_map).fillna(99)
    return g.sort_values("_orden").drop(columns="_orden").reset_index(drop=True)


# ── Ranking de clientes ──────────────────────────────────────────────────

def ranking_clientes(ventas, fecha_desde=None, fecha_hasta=None, top_n=25):
    d = filtrar(ventas, fecha_desde=fecha_desde, fecha_hasta=fecha_hasta)
    g = d.groupby("Cliente").agg(
        NombreFantasia=("NombreFantasia", "first"),
        Total=("Total", "sum"), Unidades=("Cantidad", "sum"),
        Comprobantes=("NroComprobante", "nunique"),
    ).reset_index()
    total_gral = g["Unidades"].sum()
    g["Participacion"] = g["Unidades"] / total_gral if total_gral else 0
    return g.sort_values("Total", ascending=False).head(top_n).reset_index(drop=True)


def razon_social(cliente_raw):
    """'(01827) - FORTY FOUR AVENUE S.A. - CORTEZ' -> 'FORTY FOUR AVENUE S.A.'"""
    partes = [p.strip() for p in str(cliente_raw or "").split(" - ")]
    return partes[1] if len(partes) > 1 else partes[0]


def vendedor_nombre(vendedor_raw):
    if not vendedor_raw or (isinstance(vendedor_raw, float) and pd.isna(vendedor_raw)):
        return ""
    partes = str(vendedor_raw).split(" - ", 1)
    return partes[1].strip() if len(partes) > 1 else partes[0].strip()


def mapa_vendedor_por_cliente(ventas):
    d = ventas.sort_values("Fecha")
    return d.groupby("Cliente")["Vendedor"].last().apply(vendedor_nombre).to_dict()


# ── Resultados mensuales por SKU ─────────────────────────────────────────

def resultados_mensuales_por_producto(ventas, col_marca="Marca"):
    """Unidades y facturación por Producto (SKU) x mes calendario, para
    TODO el histórico disponible (sin el corte de FECHA_DESDE_GRAFICOS que
    usan los gráficos de tendencia) — pensado para exportar/tabular, no
    para graficar. Devuelve una fila por (Producto, Periodo)."""
    d = ventas.copy()
    d["Periodo"] = d["Fecha"].dt.to_period("M")
    g = d.groupby([col_marca, "Producto", "Periodo"]).agg(
        Unidades=("Cantidad", "sum"), Total=("Total", "sum"),
        Clientes=("Cliente", "nunique"),
    ).reset_index()
    g["Etiqueta"] = g["Periodo"].astype(str)
    return g.sort_values([col_marca, "Producto", "Periodo"])


def pivot_mensual_por_producto(ventas, valor="Unidades", col_marca="Marca"):
    """Tabla ancha: filas = (Marca, Producto), columnas = mes (YYYY-MM),
    valores = Unidades o Total. Pensada para pegar directo en una hoja de
    Excel tipo 'un SKU por fila, un mes por columna'."""
    g = resultados_mensuales_por_producto(ventas, col_marca=col_marca)
    piv = g.pivot_table(index=[col_marca, "Producto"], columns="Etiqueta", values=valor,
                         aggfunc="sum", fill_value=0)
    piv = piv.reindex(sorted(piv.columns), axis=1)
    return piv.reset_index()


# ── Ranking de producto individual ──────────────────────────────────────

def ranking_productos(ventas, fecha_desde=None, fecha_hasta=None, marcas=None):
    d = filtrar(ventas, marcas=marcas, fecha_desde=fecha_desde, fecha_hasta=fecha_hasta)
    g = d.groupby(["Marca", "Producto"]).agg(
        Total=("Total", "sum"), Unidades=("Cantidad", "sum"),
        PrecioProm=("PrecioUnitario", "mean"),
        Compradores=("Cliente", "nunique"),
    ).reset_index()
    return g.sort_values("Total", ascending=False).reset_index(drop=True)


def producto_comparativo(ventas, producto):
    d = ventas[(ventas["Producto"] == producto) & (ventas["Fecha"] >= FECHA_DESDE_GRAFICOS)].copy()
    d["Periodo"] = d["Fecha"].dt.to_period("M")
    g = d.groupby("Periodo").agg(Total=("Total", "sum"), Unidades=("Cantidad", "sum")).reset_index()
    g["Etiqueta"] = g["Periodo"].astype(str)
    return g.sort_values("Periodo")


def producto_trimestral(ventas, producto):
    d = ventas[ventas["Producto"] == producto]
    return evolucion_trimestral(d)


def producto_ytd(ventas, producto, año, fecha_corte=None):
    fecha_corte = fecha_corte or pd.Timestamp.now().normalize()
    doy = fecha_corte.dayofyear
    d = ventas[ventas["Producto"] == producto]

    def _slice(año_x):
        ini = pd.Timestamp(year=año_x, month=1, day=1)
        fin = ini + pd.Timedelta(days=doy - 1)
        dd = d[(d["Fecha"] >= ini) & (d["Fecha"] <= fin)]
        return dd["Total"].sum(), dd["Cantidad"].sum()

    total_actual, unid_actual = _slice(año)
    total_prev, unid_prev = _slice(año - 1)
    return {
        "total_actual": total_actual, "total_prev": total_prev,
        "unid_actual": unid_actual, "unid_prev": unid_prev,
        "var_total": var(total_prev, total_actual),
        "var_unid": var(unid_prev, unid_actual),
    }


def ranking_productos_ytd(ventas, año, fecha_corte=None, marcas=None, productos=None, top_n=20):
    fecha_corte = fecha_corte or pd.Timestamp.now().normalize()
    doy = fecha_corte.dayofyear
    d = filtrar(ventas, marcas=marcas, productos=productos)

    def _slice(año_x):
        ini = pd.Timestamp(year=año_x, month=1, day=1)
        fin = ini + pd.Timedelta(days=doy - 1)
        dd = d[(d["Fecha"] >= ini) & (d["Fecha"] <= fin)]
        return dd.groupby("Producto").agg(Total=("Total", "sum"), Unidades=("Cantidad", "sum"),
                                            Compradores=("Cliente", "nunique"))

    actual = _slice(año).rename(columns={"Total": "TotalActual", "Unidades": "UnidadesActual",
                                          "Compradores": "CompradoresActual"})
    prev = _slice(año - 1).rename(columns={"Total": "TotalPrev", "Unidades": "UnidadesPrev",
                                            "Compradores": "CompradoresPrev"})
    g = actual.join(prev, how="outer").fillna(0).reset_index()
    g["VarUnidades"] = g.apply(lambda r: var(r["UnidadesPrev"], r["UnidadesActual"]), axis=1)
    return g.sort_values("UnidadesActual", ascending=False).head(top_n).reset_index(drop=True)


# ── Tabla Marca x Trimestre (para la pestaña Resumen) ───────────────────

def _inicio_trimestre(año, q):
    return pd.Timestamp(year=año, month=3 * (q - 1) + 1, day=1)


def tabla_marca_trimestral(ventas, año, marcas=None, fecha_corte=None, col_marca="Marca"):
    """Fila por marca (+ fila Total), columnas Q1/Q2/Q3 (año y año-1, en
    unidades) + Total YTD, con Var. %. El trimestre EN CURSO se corta al
    mismo día en los dos años que se comparan."""
    marcas = marcas or MARCAS_PRINCIPALES
    fecha_corte = fecha_corte or pd.Timestamp.now().normalize()
    doy = fecha_corte.dayofyear
    trim_actual = trimestre(fecha_corte)
    dias_en_trim_actual = (fecha_corte - _inicio_trimestre(fecha_corte.year, trim_actual)).days
    d = ventas[ventas[col_marca].isin(marcas)].copy()

    filas = []
    for marca in marcas:
        dm = d[d[col_marca] == marca]
        fila = {"Marca": marca}
        for q in (1, 2, 3):
            for año_x in (año - 1, año):
                ini_q = _inicio_trimestre(año_x, q)
                if q == trim_actual:
                    fin_q = ini_q + pd.Timedelta(days=dias_en_trim_actual)
                else:
                    fin_q = _inicio_trimestre(año_x, q + 1) - pd.Timedelta(days=1)
                mask = (dm["Fecha"] >= ini_q) & (dm["Fecha"] <= fin_q)
                fila[f"Q{q}_{año_x}"] = dm.loc[mask, "Cantidad"].sum()
        for año_x in (año - 1, año):
            ini = pd.Timestamp(year=año_x, month=1, day=1)
            fin = ini + pd.Timedelta(days=doy - 1)
            mask = (dm["Fecha"] >= ini) & (dm["Fecha"] <= fin)
            fila[f"YTD_{año_x}"] = dm.loc[mask, "Cantidad"].sum()
        filas.append(fila)

    tabla = pd.DataFrame(filas)
    total = {"Marca": "Total"}
    for c in tabla.columns:
        if c != "Marca":
            total[c] = tabla[c].sum()
    tabla = pd.concat([tabla, pd.DataFrame([total])], ignore_index=True)

    for q in (1, 2, 3):
        tabla[f"Q{q}_Var"] = tabla.apply(lambda r, q=q: var(r[f"Q{q}_{año-1}"], r[f"Q{q}_{año}"]), axis=1)
    tabla["YTD_Var"] = tabla.apply(lambda r: var(r[f"YTD_{año-1}"], r[f"YTD_{año}"]), axis=1)
    return tabla


# ── Clientes por marca y trimestre (Q1/Q2/Q3), identificados ────────────
# A pedido explícito: para Martini Rosso/Bacardi/Gancia/Amargo Obrero, saber
# CUÁNTOS clientes compraron cada marca en Q1, Q2 y lo que va de Q3 — y
# poder ver el listado de quiénes son, no solo el número.

def _rango_trimestre_cortado(año, q, fecha_corte=None):
    """Rango de un trimestre calendario, cortado a fecha_corte si es el
    trimestre en curso (para no comparar un Q parcial contra uno cerrado sin
    avisar)."""
    fecha_corte = fecha_corte or pd.Timestamp.now().normalize()
    ini = _inicio_trimestre(año, q)
    fin_completo = _inicio_trimestre(año, q + 1) - pd.Timedelta(days=1)
    if año == fecha_corte.year and q == trimestre(fecha_corte):
        fin = min(fecha_corte, fin_completo)
    else:
        fin = fin_completo
    return ini, fin


def clientes_por_marca_trimestre(ventas, marcas=None, año=None, fecha_corte=None, col_marca="Marca"):
    """{marca: {Q1: {"clientes": [...], "cantidad": n, "ini":.., "fin":..},
                Q2: {...}, Q3: {...}}} — Q3 puede estar parcial si es el
    trimestre en curso."""
    marcas = marcas or MARCAS_PRINCIPALES
    fecha_corte = fecha_corte or pd.Timestamp.now().normalize()
    año = año or fecha_corte.year
    out = {}
    for marca in marcas:
        d = ventas[ventas[col_marca] == marca]
        out[marca] = {}
        for q in (1, 2, 3):
            ini, fin = _rango_trimestre_cortado(año, q, fecha_corte)
            mask = (d["Fecha"] >= ini) & (d["Fecha"] <= fin)
            clientes = sorted(d.loc[mask, "Cliente"].unique())
            out[marca][f"Q{q}"] = {
                "clientes": clientes, "cantidad": len(clientes),
                "ini": ini.date(), "fin": fin.date(),
                "parcial": fin < (_inicio_trimestre(año, q + 1) - pd.Timedelta(days=1)),
            }
    return out


def tabla_resumen_clientes_por_marca(clientes_trim):
    """A partir de clientes_por_marca_trimestre(), arma una tabla chica
    Marca x (Q1/Q2/Q3) con la CANTIDAD de clientes distintos en cada uno."""
    filas = []
    for marca, qs in clientes_trim.items():
        filas.append({
            "Marca": marca,
            "Clientes Q1": qs["Q1"]["cantidad"],
            "Clientes Q2": qs["Q2"]["cantidad"],
            "Clientes Q3": qs["Q3"]["cantidad"],
        })
    return pd.DataFrame(filas)


# ── Comparación de vermouth (todas las marcas, todos los proveedores) ───

def es_rosso(producto):
    """True si el nombre del producto indica variante 'Rosso' (Martini Rosso,
    Cinzano Rosso, Carpano Rosso, Vermu Lunfa Rosso, Vermouth Unión Federal
    Rosso, Vermut Ajenjo Rosso, Vermouth Cordero CPL Rosso, etc.) — busca
    literal 'ROSSO' en el nombre, NO 'ROJO' (La Fuerza Rojo, Ajenjo Rojo,
    Siete Cuatro Seis son otra variante/nomenclatura, no Rosso)."""
    return "ROSSO" in (producto or "").upper()


def filtrar_solo_rosso(ventas_verm):
    """Recorta el dataset de vermouth a SOLO los productos 'Rosso' (ver
    es_rosso) — para comparaciones cara a cara donde no tiene sentido mezclar
    con Bianco/Extra Dry/Segundo/1757/To Spritz/Bitter/Blanco/etc."""
    return ventas_verm[ventas_verm["Producto"].apply(es_rosso)]


def resumen_rosso_ytd_por_marca(ventas_verm, año, fecha_corte=None):
    """YTD (mismo día del año) de vermouth Rosso, año actual vs anterior:
    unidades totales de la categoría Rosso + unidades y participación por
    marca — para ver cómo cambió el reparto entre Martini Rosso, Cinzano
    Rosso y el resto, no solo el total."""
    fecha_corte = fecha_corte or pd.Timestamp.now().normalize()
    doy = fecha_corte.dayofyear
    d = filtrar_solo_rosso(ventas_verm)

    def _slice(año_x):
        ini = pd.Timestamp(year=año_x, month=1, day=1)
        fin = ini + pd.Timedelta(days=doy - 1)
        dd = d[(d["Fecha"] >= ini) & (d["Fecha"] <= fin)]
        por_marca = dd.groupby("MarcaVermouth")["Cantidad"].sum()
        return dd["Cantidad"].sum(), por_marca

    total_actual, marca_actual = _slice(año)
    total_prev, marca_prev = _slice(año - 1)

    marcas = sorted(set(marca_actual.index) | set(marca_prev.index))
    filas = []
    for m in marcas:
        u_prev = marca_prev.get(m, 0.0)
        u_act = marca_actual.get(m, 0.0)
        part_prev = (u_prev / total_prev) if total_prev else 0.0
        part_act = (u_act / total_actual) if total_actual else 0.0
        filas.append({
            "Marca": m,
            f"Unidades {año-1}": u_prev, f"Unidades {año}": u_act,
            "Var. % unidades": var(u_prev, u_act),
            f"Participación {año-1}": part_prev, f"Participación {año}": part_act,
            "Var. participación (pp)": (part_act - part_prev) * 100,
        })
    tabla = pd.DataFrame(filas).sort_values(f"Unidades {año}", ascending=False).reset_index(drop=True)
    total_fila = {
        "Marca": "Total vermouth Rosso",
        f"Unidades {año-1}": total_prev, f"Unidades {año}": total_actual,
        "Var. % unidades": var(total_prev, total_actual),
        f"Participación {año-1}": 1.0 if total_prev else 0.0,
        f"Participación {año}": 1.0 if total_actual else 0.0,
        "Var. participación (pp)": 0.0,
    }
    tabla = pd.concat([tabla, pd.DataFrame([total_fila])], ignore_index=True)
    return tabla, total_prev, total_actual

def resumen_anual_marca_vermouth(ventas_verm, marca, año, fecha_corte=None):
    """Facturación y unidades del año anterior (año completo cerrado) y del
    año en curso (YTD al mismo día), para una marca de vermouth puntual
    (ej. Martini Rosso) — más variación mensual (último mes cerrado vs. el
    anterior), interanual (ese mismo mes vs. un año atrás) y YTD (año en
    curso vs año anterior, mismo día). Pensada para las tarjetas resumen de
    la pestaña Vermouth."""
    fecha_corte = fecha_corte or pd.Timestamp.now().normalize()
    d = ventas_verm[ventas_verm["MarcaVermouth"] == marca]

    ini_prev, fin_prev = pd.Timestamp(año - 1, 1, 1), pd.Timestamp(año - 1, 12, 31)
    d_prev = d[(d["Fecha"] >= ini_prev) & (d["Fecha"] <= fin_prev)]
    total_año_prev, unid_año_prev = d_prev["Total"].sum(), d_prev["Cantidad"].sum()

    doy = fecha_corte.dayofyear

    def _ytd(año_x):
        ini = pd.Timestamp(año_x, 1, 1)
        fin = ini + pd.Timedelta(days=doy - 1)
        dd = d[(d["Fecha"] >= ini) & (d["Fecha"] <= fin)]
        return dd["Total"].sum(), dd["Cantidad"].sum()

    total_ytd_act, unid_ytd_act = _ytd(año)
    total_ytd_prev, unid_ytd_prev = _ytd(año - 1)

    dm = d.copy()
    dm["Periodo"] = dm["Fecha"].dt.to_period("M")
    mensual = dm.groupby("Periodo").agg(Total=("Total", "sum"), Unidades=("Cantidad", "sum"))

    def _mes(periodo):
        if periodo in mensual.index:
            return mensual.loc[periodo, "Total"], mensual.loc[periodo, "Unidades"]
        return 0.0, 0.0

    periodo_actual = fecha_corte.to_period("M")
    ult_cerrado = periodo_actual - 1
    penult_cerrado = periodo_actual - 2
    mismo_mes_año_ant = ult_cerrado - 12

    total_ult, unid_ult = _mes(ult_cerrado)
    total_penult, unid_penult = _mes(penult_cerrado)
    total_ia, unid_ia = _mes(mismo_mes_año_ant)

    return {
        "total_año_prev": total_año_prev, "unid_año_prev": unid_año_prev,
        "total_ytd_act": total_ytd_act, "unid_ytd_act": unid_ytd_act,
        "var_mensual_total": var(total_penult, total_ult),
        "var_mensual_unid": var(unid_penult, unid_ult),
        "var_interanual_total": var(total_ia, total_ult),
        "var_interanual_unid": var(unid_ia, unid_ult),
        "var_ytd_total": var(total_ytd_prev, total_ytd_act),
        "var_ytd_unid": var(unid_ytd_prev, unid_ytd_act),
    }


def ranking_marcas_vermouth(ventas_verm, fecha_desde=None, fecha_hasta=None):
    """Unidades, clientes distintos y facturación por marca de vermouth —
    para comparar Martini Rosso contra Cinzano y el resto del rubro."""
    d = filtrar(ventas_verm, fecha_desde=fecha_desde, fecha_hasta=fecha_hasta, col_marca="MarcaVermouth")
    g = d.groupby("MarcaVermouth").agg(
        Total=("Total", "sum"), Unidades=("Cantidad", "sum"),
        Clientes=("Cliente", "nunique"),
    ).reset_index()
    total_gral = g["Unidades"].sum()
    g["ParticipacionUnidades"] = g["Unidades"] / total_gral if total_gral else 0
    orden_map = {m: i for i, m in enumerate(ORDEN_MARCAS_VERMOUTH)}
    g["_orden"] = g["MarcaVermouth"].map(orden_map).fillna(99)
    return g.sort_values("_orden").drop(columns="_orden").reset_index(drop=True)


def clientes_overlap(ventas_verm, marca_a, marca_b, fecha_desde=None, fecha_hasta=None):
    """Clientes que compraron marca_a, marca_b, ambas, y solo una — para
    'Martini Rosso vs Cinzano': cuántos clientes son exclusivos de cada uno
    y cuántos compran las dos."""
    d = filtrar(ventas_verm, fecha_desde=fecha_desde, fecha_hasta=fecha_hasta, col_marca="MarcaVermouth")
    clientes_a = set(d[d["MarcaVermouth"] == marca_a]["Cliente"].unique())
    clientes_b = set(d[d["MarcaVermouth"] == marca_b]["Cliente"].unique())
    return {
        "solo_a": sorted(clientes_a - clientes_b),
        "solo_b": sorted(clientes_b - clientes_a),
        "ambas": sorted(clientes_a & clientes_b),
        "total_a": len(clientes_a), "total_b": len(clientes_b),
    }


def evolucion_mensual_por_marca_vermouth(ventas_verm, marcas=None, n_meses=13):
    """Serie MENSUAL de unidades por marca de vermouth — mismo espíritu que
    evolucion_trimestral_por_marca_vermouth, pero mes a mes (para ver el
    detalle fino de Martini Rosso vs. Cinzano Rosso, no solo el agregado
    trimestral)."""
    marcas = marcas or ["Martini Rosso", "Cinzano Rosso"]
    d = ventas_verm[ventas_verm["MarcaVermouth"].isin(marcas)].copy()
    d = d[d["Fecha"] >= FECHA_DESDE_GRAFICOS]
    d["Periodo"] = d["Fecha"].dt.to_period("M")
    g = d.groupby(["Periodo", "MarcaVermouth"]).agg(Unidades=("Cantidad", "sum")).reset_index()
    periodos = sorted(d["Periodo"].unique())[-n_meses:]
    g = g[g["Periodo"].isin(periodos)]
    g["Etiqueta"] = g["Periodo"].astype(str)
    return g.sort_values("Periodo")


def evolucion_trimestral_por_marca_vermouth(ventas_verm, marcas=None):
    """Serie trimestral de unidades por marca de vermouth — para el gráfico
    de evolución comparada."""
    marcas = marcas or ["Martini Rosso", "Cinzano"]
    d = ventas_verm[ventas_verm["MarcaVermouth"].isin(marcas)].copy()
    d = d[d["Fecha"] >= FECHA_DESDE_GRAFICOS]
    d["Año"] = d["Fecha"].dt.year
    d["Trim"] = d["Fecha"].apply(trimestre)
    g = d.groupby(["Año", "Trim", "MarcaVermouth"]).agg(Unidades=("Cantidad", "sum")).reset_index()
    g["Etiqueta"] = g["Año"].astype(str) + "-Q" + g["Trim"].astype(str)
    return g.sort_values(["Año", "Trim"])


# ── Ventas vs Compras ────────────────────────────────────────────────────

def comparativo_ventas_compras_mensual(ventas, compras, n_meses=13):
    v = ventas[ventas["Fecha"] >= FECHA_DESDE_GRAFICOS].copy()
    v["Periodo"] = v["Fecha"].dt.to_period("M")
    gv = v.groupby("Periodo")["Total"].sum().rename("Ventas")

    c = compras[compras["Fecha"] >= FECHA_DESDE_GRAFICOS].copy()
    c["Periodo"] = c["Fecha"].dt.to_period("M")
    gc = c.groupby("Periodo")["Total"].sum().rename("Compras")

    g = pd.concat([gv, gc], axis=1).fillna(0).reset_index()
    g = g.sort_values("Periodo").tail(n_meses)
    g["Etiqueta"] = g["Periodo"].astype(str)
    return g
