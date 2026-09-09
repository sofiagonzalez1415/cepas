"""
dashboard.py — Panel Cepas: seguimiento de ventas y compras al proveedor
CEPAS ARGENTINA S.A.- GANCIA (Martini, Bacardi, Gancia, Amargo Obrero y el
resto del portafolio), más una comparación de la categoría Vermouth completa
contra la competencia (Cinzano, Ajenjo, Lunfa, Unión Federal, La Fuerza,
Vincenzo, Cordero/Mosquita Muerta, etc.). Se despliega en Streamlit
Community Cloud. Lee panel_cepas.xlsx desde Drive (lo genera y sube
extract_data.py, corriendo local 1 vez por día). Mismo patrón general que
Panel_Campari — ver ese CLAUDE.md para el detalle de la arquitectura.
"""

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from google_drive_auth import download_file, find_file_id
import metrics as M

st.set_page_config(page_title="Panel Cepas", page_icon="🍸", layout="wide")

DRIVE_FOLDER_ID = "1al-GEctDtC-oIgXWPzlmxNqzKlp30wmq"
PANEL_FILENAME = "panel_cepas.xlsx"

# ── Tokens de color ──────────────────────────────────────────────────────
ACCENT = "#8a1538"       # bordó Martini/Gancia, para acentos de UI
SERIE_1 = "#2a78d6"
SERIE_2 = "#eb6834"
GOOD = "#006300"
BAD = "#d03b3b"
GRID = "#e1e0d9"
TEXT_SECONDARY = "#52514e"
MARCA_COLORS = {
    "Martini Rosso": "#8a1538", "Bacardi": "#c0242a", "Gancia": "#e8a33d",
    "Amargo Obrero": "#5b4636", "Martini (otras variantes)": "#c78fa0",
    "Otras marcas Cepas": "#898781",
}
VERMOUTH_COLORS = {
    "Martini Rosso": "#8a1538", "Cinzano Rosso": "#1a5fb4", "Martini (otras variantes)": "#c78fa0",
    "Cinzano (otras variantes)": "#7fa6d9",
    "Carpano": "#5b3a29", "Antica Formula": "#8f6a4d", "Ajenjo (Evo & Ajenjo)": "#2f7d5e",
    "Lunfa": "#c9962c", "Unión Federal": "#6a5acd", "La Fuerza": "#946b3f",
    "Vincenzo": "#4c8577", "Cordero (Mosquita Muerta)": "#a63b5a", "Siete Cuatro Seis": "#7a7a7a",
    "Otros vermouth": "#898781",
}


def fmt_money(n):
    try:
        if abs(n) >= 1_000_000:
            return f"$ {n/1_000_000:.1f}M"
        if abs(n) >= 1_000:
            return f"$ {n/1_000:.0f}k"
        return f"$ {n:,.0f}".replace(",", ".")
    except Exception:
        return "$ 0"


def fmt_int(n):
    try:
        return f"{n:,.0f}".replace(",", ".")
    except Exception:
        return "0"


def fmt_pct(n):
    if n is None or (isinstance(n, float) and pd.isna(n)):
        return "—"
    try:
        return f"{n*100:+.1f}%"
    except Exception:
        return "—"


def tabla_variacion_trimestral(trimestral, col_val="Unidades", fmt_val=None):
    fmt_val = fmt_val or (fmt_int if col_val == "Unidades" else fmt_money)
    var_qoq_col = "VarUnidades" if col_val == "Unidades" else "VarTotal"
    var_yoy_col = "VarUnidadesYoY" if col_val == "Unidades" else "VarTotalYoY"
    out = pd.DataFrame({
        "Trimestre": trimestral["Etiqueta"],
        col_val: trimestral[col_val].apply(fmt_val),
        "Var. trim. anterior": trimestral[var_qoq_col].apply(fmt_pct),
        "Var. interanual": trimestral[var_yoy_col].apply(fmt_pct),
    })
    return out


def render_tabla_marca_trimestral(tabla, año):
    cols_order, rename, var_cols = [], {}, []
    for q in (1, 2, 3):
        c_prev, c_act, c_var = f"Q{q}_{año-1}", f"Q{q}_{año}", f"Q{q}_Var"
        cols_order += [c_prev, c_act, c_var]
        rename[c_prev], rename[c_act], rename[c_var] = f"Q{q} {año-1}", f"Q{q} {año}", f"Var Q{q}"
        var_cols.append(rename[c_var])
    c_prev, c_act, c_var = f"YTD_{año-1}", f"YTD_{año}", "YTD_Var"
    cols_order += [c_prev, c_act, c_var]
    rename[c_prev], rename[c_act], rename[c_var] = f"YTD {año-1}", f"YTD {año}", "Var YTD"
    var_cols.append("Var YTD")

    disp = tabla.set_index("Marca")[cols_order].rename(columns=rename)

    def color_var(v):
        if pd.isna(v):
            return ""
        return ("background-color: #c6efce; color: #006100" if v >= 0
                else "background-color: #ffc7ce; color: #9c0006")

    fmt = {c: fmt_int for c in disp.columns if c not in var_cols}
    fmt.update({c: fmt_pct for c in var_cols})
    return disp.style.format(fmt).map(color_var, subset=var_cols)


# ── Password ─────────────────────────────────────────────────────────────

def check_password():
    if "PASSWORD" not in st.secrets:
        return True
    if st.session_state.get("auth_ok"):
        return True
    pwd = st.text_input("Contraseña", type="password")
    if pwd == st.secrets["PASSWORD"]:
        st.session_state["auth_ok"] = True
        st.rerun()
    elif pwd:
        st.error("Contraseña incorrecta")
    return False


if not check_password():
    st.stop()

# ── Carga de datos ───────────────────────────────────────────────────────

@st.cache_data(ttl=300)
def cargar_datos():
    file_id = find_file_id(DRIVE_FOLDER_ID, PANEL_FILENAME)
    if not file_id:
        return None, None, None, None
    stream = download_file(file_id)
    ventas = pd.read_excel(stream, sheet_name="Datos Ventas")
    stream.seek(0)
    compras = pd.read_excel(stream, sheet_name="Datos Compras")
    stream.seek(0)
    ventas_verm = pd.read_excel(stream, sheet_name="Datos Vermouth")
    stream.seek(0)
    meta = pd.read_excel(stream, sheet_name="Metadata")
    ventas["Fecha"] = pd.to_datetime(ventas["Fecha"])
    compras["Fecha"] = pd.to_datetime(compras["Fecha"])
    ventas_verm["Fecha"] = pd.to_datetime(ventas_verm["Fecha"])
    return ventas, compras, ventas_verm, meta


ventas, compras, ventas_verm, meta = cargar_datos()

if ventas is None:
    st.error(f"No se encontró '{PANEL_FILENAME}' en la carpeta de Drive. "
             "¿Ya corrió extract_data.py al menos una vez?")
    st.stop()

st.title("🍸 Panel Cepas")
st.caption(
    f"Proveedor estratégico CEPAS ARGENTINA S.A.- GANCIA (Martini, Bacardi, Gancia, "
    f"Amargo Obrero) · Actualizado: {meta['Actualizado'].iloc[0]} · "
    f"{int(meta['Filas Ventas'].iloc[0]):,} filas de ventas Cepas, "
    f"{int(meta['Filas Compras'].iloc[0]):,} de compras, "
    f"{int(meta['Filas Vermouth'].iloc[0]):,} filas de vermouth (todas las marcas)"
)

# ── Filtros (sidebar) ────────────────────────────────────────────────────

st.sidebar.header("Filtros")
marcas_disp = [m for m in M.ORDEN_MARCAS if m != "Impuestos/Otros (no producto)"
               and m in ventas["Marca"].unique()]
marcas_sel = st.sidebar.multiselect("Marcas", marcas_disp, default=marcas_disp)

productos_disp_sidebar = sorted(ventas[ventas["Marca"].isin(marcas_sel or marcas_disp)]["Producto"].unique())
productos_sel = st.sidebar.multiselect("Producto", productos_disp_sidebar, default=[],
                                         help="Vacío = todos los productos de las marcas elegidas arriba.")

fecha_min, fecha_max = ventas["Fecha"].min().date(), ventas["Fecha"].max().date()
hoy = pd.Timestamp.now()
inicio_trim_actual = pd.Timestamp(year=hoy.year, month=3 * ((hoy.month - 1) // 3) + 1, day=1).date()
default_desde = max(fecha_min, inicio_trim_actual)
rango = st.sidebar.date_input("Rango de fechas", value=(default_desde, fecha_max),
                                min_value=fecha_min, max_value=fecha_max)
fecha_desde, fecha_hasta = (rango if isinstance(rango, tuple) and len(rango) == 2
                             else (fecha_min, fecha_max))

ventas_f = M.filtrar(ventas, marcas=marcas_sel or None, productos=productos_sel or None,
                      fecha_desde=fecha_desde, fecha_hasta=fecha_hasta)
ventas_mp = M.filtrar(ventas, marcas=marcas_sel or None, productos=productos_sel or None)
compras_mp = M.filtrar(compras, marcas=marcas_sel or None, productos=productos_sel or None)

tab_resumen, tab_mix, tab_evolucion, tab_clientes, tab_vermouth, tab_productos, tab_compras = st.tabs(
    ["Resumen", "Mix de producto", "Evolución", "Clientes", "Vermouth vs. competencia",
     "Detalle por producto", "Compras vs Ventas"]
)

# ── Tab Resumen ──────────────────────────────────────────────────────────

with tab_resumen:
    año_actual = pd.Timestamp.now().year
    r = M.resumen_ytd(ventas_mp, año_actual)

    c1, c2, c3 = st.columns(3)
    c1.metric(f"Facturación YTD {año_actual}", fmt_money(r["total_actual"]), fmt_pct(r["var_total"]))
    c2.metric(f"Unidades YTD {año_actual}", fmt_int(r["unid_actual"]), fmt_pct(r["var_unid"]))
    c3.metric("Precio promedio / unidad", fmt_money(r["precio_prom_actual"]),
              fmt_pct(M.var(r["precio_prom_prev"], r["precio_prom_actual"])))
    st.caption("YTD comparado contra el mismo día del año del período anterior.")

    st.subheader("Marca x Trimestre (unidades) — Martini Rosso / Bacardi / Gancia / Amargo Obrero")
    tabla_marca = M.tabla_marca_trimestral(ventas_mp, año_actual)
    st.dataframe(render_tabla_marca_trimestral(tabla_marca, año_actual), use_container_width=True)
    st.caption(
        "Q1 y Q2 son trimestres completos. El trimestre en curso (Q3) se corta al mismo "
        "día en el año actual y el anterior, para que la comparación sea pareja."
    )

    st.subheader("Evolución trimestral (unidades)")
    trimestral = M.evolucion_trimestral(ventas_mp)
    fig = go.Figure()
    fig.add_bar(x=trimestral["Etiqueta"], y=trimestral["Unidades"], marker_color=ACCENT,
                text=[fmt_int(v) for v in trimestral["Unidades"]], textposition="outside")
    fig.update_layout(showlegend=False, height=360, margin=dict(t=10, b=10),
                       yaxis_title="Unidades", plot_bgcolor="white", yaxis=dict(gridcolor=GRID))
    st.plotly_chart(fig, use_container_width=True)

# ── Tab Mix de producto ──────────────────────────────────────────────────

with tab_mix:
    st.subheader("Mix de producto (Martini Rosso / Bacardi / Gancia / Amargo Obrero / resto)")
    mix = M.mix_por_marca(ventas_f)

    col1, col2 = st.columns([1, 1])
    with col1:
        fig = go.Figure(go.Pie(
            labels=mix["Marca"], values=mix["Total"], hole=0.55,
            marker_colors=[MARCA_COLORS.get(m, "#898781") for m in mix["Marca"]],
            textinfo="label+percent",
        ))
        fig.update_layout(height=420, margin=dict(t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        tabla = mix.copy()
        tabla["Total"] = tabla["Total"].apply(fmt_money)
        tabla["Unidades"] = tabla["Unidades"].apply(fmt_int)
        tabla["Clientes"] = tabla["Clientes"].apply(fmt_int)
        tabla["Participacion"] = tabla["Participacion"].apply(lambda x: f"{x*100:.1f}%")
        st.dataframe(tabla, use_container_width=True, hide_index=True)

    st.subheader("Mix mensual por marca (unidades, últimos 13 meses)")
    mix_mes = M.evolucion_mensual_por_marca(ventas_mp)
    fig2 = go.Figure()
    for marca in marcas_sel or marcas_disp:
        d = mix_mes[mix_mes["Marca"] == marca]
        fig2.add_bar(x=d["Etiqueta"], y=d["Unidades"], name=marca, marker_color=MARCA_COLORS.get(marca, "#898781"))
    totales_mes = mix_mes.groupby("Etiqueta")["Unidades"].sum()
    fig2.update_layout(
        barmode="stack", height=420, margin=dict(t=40, b=10),
        plot_bgcolor="white", yaxis=dict(gridcolor=GRID, title="Unidades"),
        annotations=[
            dict(x=etq, y=total, text=fmt_int(total), showarrow=False,
                 yanchor="bottom", font=dict(size=11, color=TEXT_SECONDARY))
            for etq, total in totales_mes.items()
        ],
    )
    st.plotly_chart(fig2, use_container_width=True)

# ── Tab Evolución ────────────────────────────────────────────────────────

with tab_evolucion:
    unidad_vista = st.radio("Ver en", ["Unidades", "Pesos"], horizontal=True, index=0, key="evol_unidad")
    col_val = "Unidades" if unidad_vista == "Unidades" else "Total"
    fmt_val = fmt_int if unidad_vista == "Unidades" else fmt_money

    st.subheader(f"Evolución mensual ({unidad_vista.lower()})")
    mensual = M.evolucion_mensual(ventas_mp, n_meses=18)
    fig = go.Figure()
    fig.add_scatter(x=mensual["Etiqueta"], y=mensual[col_val], mode="lines+markers",
                     line=dict(color=ACCENT, width=2), marker=dict(size=6))
    fig.update_layout(height=380, margin=dict(t=10, b=10), plot_bgcolor="white",
                       yaxis=dict(gridcolor=GRID, title=unidad_vista))
    st.plotly_chart(fig, use_container_width=True)

    st.subheader(f"Evolución trimestral ({unidad_vista.lower()}) — desde 2025")
    trimestral = M.evolucion_trimestral(ventas_mp)
    fig2 = go.Figure()
    fig2.add_bar(x=trimestral["Etiqueta"], y=trimestral[col_val], marker_color=SERIE_1, name=unidad_vista)
    fig2.update_layout(height=380, margin=dict(t=10, b=10), plot_bgcolor="white",
                        yaxis=dict(gridcolor=GRID, title=unidad_vista))
    st.plotly_chart(fig2, use_container_width=True)

    st.dataframe(tabla_variacion_trimestral(trimestral, col_val, fmt_val), use_container_width=True, hide_index=True)
    st.caption(
        "El último trimestre del año en curso suele estar parcial — compararlo contra "
        "un trimestre cerrado de otro año sobreestima la caída/suba."
    )

# ── Tab Clientes ─────────────────────────────────────────────────────────

with tab_clientes:
    año_cli = pd.Timestamp.now().year
    st.subheader(f"Clientes con compra por trimestre — {año_cli} (Q1, Q2 y avance de Q3)")

    marcas_para_clientes = [m for m in (marcas_sel or marcas_disp) if m in M.MARCAS_PRINCIPALES] or M.MARCAS_PRINCIPALES
    clientes_trim = M.clientes_por_marca_trimestre(ventas, marcas=marcas_para_clientes, año=año_cli)

    tabla_resumen = M.tabla_resumen_clientes_por_marca(clientes_trim)
    q3_info = next(iter(clientes_trim.values()))["Q3"]
    tabla_resumen = tabla_resumen.rename(columns={
        "Clientes Q3": f"Clientes Q3 (al {q3_info['fin'].strftime('%d/%m')})"
    })
    st.dataframe(tabla_resumen, use_container_width=True, hide_index=True)
    st.caption(
        f"Cantidad de clientes DISTINTOS que compraron cada marca en cada trimestre "
        f"{año_cli}. Q3 está en curso (datos hasta el {q3_info['fin'].strftime('%d/%m/%Y')}), "
        "no es un trimestre cerrado — va a seguir subiendo."
    )

    st.subheader("Detalle: qué clientes compraron cada marca, por trimestre")
    col_marca_det, col_q_det = st.columns(2)
    with col_marca_det:
        marca_detalle = st.selectbox("Marca", marcas_para_clientes, key="cli_marca_detalle")
    with col_q_det:
        q_detalle = st.selectbox("Trimestre", ["Q1", "Q2", "Q3"], index=2, key="cli_q_detalle")

    info_q = clientes_trim[marca_detalle][q_detalle]
    aviso_parcial = " (trimestre en curso, parcial)" if info_q["parcial"] else ""
    st.markdown(
        f"**{marca_detalle} — {q_detalle} {año_cli}** "
        f"({info_q['ini'].strftime('%d/%m')} al {info_q['fin'].strftime('%d/%m')}{aviso_parcial}): "
        f"**{info_q['cantidad']} clientes**"
    )
    if info_q["cantidad"] > 0:
        tabla_det = pd.DataFrame({
            "Cliente": info_q["clientes"],
            "Razón Social": [M.razon_social(c) for c in info_q["clientes"]],
        })
        st.dataframe(tabla_det, use_container_width=True, hide_index=True)
    else:
        st.info("Sin compras de esta marca en el trimestre seleccionado.")

    st.subheader("Clientes a los que vendimos productos Cepas (ranking por facturación)")
    top_n = st.slider("Top N clientes", 10, 100, 25, step=5)
    ranking = M.ranking_clientes(ventas_f, top_n=top_n)
    tabla = ranking.copy()
    tabla["Total"] = tabla["Total"].apply(fmt_money)
    tabla["Unidades"] = tabla["Unidades"].apply(fmt_int)
    tabla["Participacion"] = tabla["Participacion"].apply(lambda x: f"{x*100:.1f}%")
    tabla = tabla.rename(columns={"NombreFantasia": "Nombre"})
    st.dataframe(tabla[["Cliente", "Nombre", "Total", "Unidades", "Comprobantes", "Participacion"]],
                 use_container_width=True, hide_index=True)

# ── Tab Vermouth vs. competencia ─────────────────────────────────────────

with tab_vermouth:
    st.subheader("Categoría Vermouth completa — todas las marcas que vendemos, no solo Cepas")
    st.caption(
        "Rubro 'Bebidas con alcohol', categoría 'Vermouth' en BSGestión, sin filtrar por "
        "proveedor — así se compara Martini Rosso (Cepas) contra Cinzano y el resto de la "
        "competencia dentro de la misma categoría de producto."
    )

    fecha_min_v, fecha_max_v = ventas_verm["Fecha"].min().date(), ventas_verm["Fecha"].max().date()
    rango_v = st.date_input("Rango de fechas (vermouth)", value=(fecha_min_v, fecha_max_v),
                              min_value=fecha_min_v, max_value=fecha_max_v, key="verm_rango")
    fv_desde, fv_hasta = (rango_v if isinstance(rango_v, tuple) and len(rango_v) == 2
                           else (fecha_min_v, fecha_max_v))

    ranking_verm = M.ranking_marcas_vermouth(ventas_verm, fecha_desde=fv_desde, fecha_hasta=fv_hasta)

    col1, col2 = st.columns([1, 1])
    with col1:
        fig = go.Figure(go.Pie(
            labels=ranking_verm["MarcaVermouth"], values=ranking_verm["Unidades"], hole=0.55,
            marker_colors=[VERMOUTH_COLORS.get(m, "#898781") for m in ranking_verm["MarcaVermouth"]],
            textinfo="label+percent",
        ))
        fig.update_layout(height=440, margin=dict(t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        tabla_v = ranking_verm.copy()
        tabla_v["Total"] = tabla_v["Total"].apply(fmt_money)
        tabla_v["Unidades"] = tabla_v["Unidades"].apply(fmt_int)
        tabla_v["Clientes"] = tabla_v["Clientes"].apply(fmt_int)
        tabla_v["ParticipacionUnidades"] = tabla_v["ParticipacionUnidades"].apply(lambda x: f"{x*100:.1f}%")
        tabla_v = tabla_v.rename(columns={"MarcaVermouth": "Marca", "ParticipacionUnidades": "% Unidades"})
        st.dataframe(tabla_v, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("Cara a cara: Martini Rosso vs. Cinzano Rosso")
    st.caption(
        "Comparación específica variante Rosso contra Rosso (no incluye Cinzano Bianco, "
        "Segundo, 1757 ni To Spritz — esos quedan aparte, en 'Cinzano (otras variantes)', "
        "para no mezclar peras con manzanas)."
    )

    fila_mr = ranking_verm[ranking_verm["MarcaVermouth"] == "Martini Rosso"]
    fila_ci = ranking_verm[ranking_verm["MarcaVermouth"] == "Cinzano Rosso"]
    unid_mr = fila_mr["Unidades"].iloc[0] if not fila_mr.empty else 0
    unid_ci = fila_ci["Unidades"].iloc[0] if not fila_ci.empty else 0
    cli_mr = fila_mr["Clientes"].iloc[0] if not fila_mr.empty else 0
    cli_ci = fila_ci["Clientes"].iloc[0] if not fila_ci.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Unidades Martini Rosso", fmt_int(unid_mr))
    c2.metric("Unidades Cinzano Rosso", fmt_int(unid_ci))
    c3.metric("Clientes Martini Rosso", fmt_int(cli_mr))
    c4.metric("Clientes Cinzano Rosso", fmt_int(cli_ci))
    if unid_ci:
        st.caption(f"Martini Rosso vendió {unid_mr/unid_ci*100:.0f}% de las unidades de Cinzano Rosso en el período elegido.")

    overlap = M.clientes_overlap(ventas_verm, "Martini Rosso", "Cinzano Rosso", fecha_desde=fv_desde, fecha_hasta=fv_hasta)
    oc1, oc2, oc3 = st.columns(3)
    oc1.metric("Solo compran Martini Rosso", fmt_int(len(overlap["solo_a"])))
    oc2.metric("Compran las dos marcas", fmt_int(len(overlap["ambas"])))
    oc3.metric("Solo compran Cinzano Rosso", fmt_int(len(overlap["solo_b"])))
    st.caption(
        "Clientes distintos en el rango de fechas elegido. 'Compran las dos' son clientes "
        "que ya tienen las dos marcas en su surtido — potencial para empujar Martini Rosso "
        "sin necesidad de abrir cliente nuevo."
    )

    with st.expander("Ver el listado de clientes por grupo"):
        cc1, cc2, cc3 = st.columns(3)
        with cc1:
            st.markdown(f"**Solo Martini Rosso ({len(overlap['solo_a'])})**")
            st.dataframe(pd.DataFrame({"Razón Social": [M.razon_social(c) for c in overlap["solo_a"]]}),
                         use_container_width=True, hide_index=True)
        with cc2:
            st.markdown(f"**Compran las dos ({len(overlap['ambas'])})**")
            st.dataframe(pd.DataFrame({"Razón Social": [M.razon_social(c) for c in overlap["ambas"]]}),
                         use_container_width=True, hide_index=True)
        with cc3:
            st.markdown(f"**Solo Cinzano Rosso ({len(overlap['solo_b'])})**")
            st.dataframe(pd.DataFrame({"Razón Social": [M.razon_social(c) for c in overlap["solo_b"]]}),
                         use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("Evolución trimestral: Martini Rosso vs. Cinzano Rosso (unidades)")
    marcas_evol = st.multiselect("Marcas a comparar", ranking_verm["MarcaVermouth"].tolist(),
                                   default=[m for m in ["Martini Rosso", "Cinzano Rosso"] if m in ranking_verm["MarcaVermouth"].tolist()])
    if marcas_evol:
        evol_verm = M.evolucion_trimestral_por_marca_vermouth(ventas_verm, marcas=marcas_evol)
        fig3 = go.Figure()
        for marca in marcas_evol:
            d = evol_verm[evol_verm["MarcaVermouth"] == marca]
            fig3.add_bar(x=d["Etiqueta"], y=d["Unidades"], name=marca, marker_color=VERMOUTH_COLORS.get(marca, "#898781"))
        fig3.update_layout(barmode="group", height=400, margin=dict(t=10, b=10), plot_bgcolor="white",
                            yaxis=dict(gridcolor=GRID, title="Unidades"),
                            legend=dict(orientation="h", yanchor="bottom", y=1.02))
        st.plotly_chart(fig3, use_container_width=True)

# ── Tab Detalle por producto ─────────────────────────────────────────────

with tab_productos:
    año_actual = pd.Timestamp.now().year

    st.subheader("Resultados mensuales, por SKU (todos los productos Cepas)")
    valor_mensual = st.radio("Ver en", ["Unidades", "Pesos"], horizontal=True, index=0, key="mensual_sku_valor")
    col_valor = "Unidades" if valor_mensual == "Unidades" else "Total"
    piv_mensual = M.pivot_mensual_por_producto(ventas, valor=col_valor)
    piv_mensual_disp = piv_mensual.copy()
    cols_mes = [c for c in piv_mensual_disp.columns if c not in ("Marca", "Producto")]
    fmt_col = fmt_int if col_valor == "Unidades" else fmt_money
    for c in cols_mes:
        piv_mensual_disp[c] = piv_mensual_disp[c].apply(fmt_col)
    st.dataframe(piv_mensual_disp, use_container_width=True, hide_index=True)
    st.caption(
        "Una fila por SKU, una columna por mes calendario — todo el histórico disponible "
        "(no respeta el filtro de fecha de la izquierda, para poder ver la serie completa)."
    )

    st.subheader("Ranking de productos YTD, en unidades")
    ranking_ytd = M.ranking_productos_ytd(ventas, año_actual, marcas=marcas_sel or None, productos=productos_sel or None)
    tabla_ytd = ranking_ytd.copy()
    tabla_ytd["Var. % unidades"] = tabla_ytd["VarUnidades"].apply(fmt_pct)
    tabla_ytd = tabla_ytd.rename(columns={
        "UnidadesActual": f"Unidades YTD {año_actual}",
        "UnidadesPrev": f"Unidades YTD {año_actual - 1}",
        "CompradoresActual": "Compradores únicos",
    })[["Producto", f"Unidades YTD {año_actual - 1}", f"Unidades YTD {año_actual}", "Var. % unidades", "Compradores únicos"]]
    for c in [f"Unidades YTD {año_actual}", f"Unidades YTD {año_actual - 1}", "Compradores únicos"]:
        tabla_ytd[c] = tabla_ytd[c].apply(fmt_int)
    st.dataframe(tabla_ytd, use_container_width=True, hide_index=True)

    st.subheader("Ranking completo (todo el histórico filtrado), en unidades")
    ranking_prod = M.ranking_productos(ventas_f).sort_values("Unidades", ascending=False)
    tabla = ranking_prod.copy()
    tabla["Unidades"] = tabla["Unidades"].apply(fmt_int)
    tabla["Total"] = tabla["Total"].apply(fmt_money)
    tabla["PrecioProm"] = tabla["PrecioProm"].apply(fmt_money)
    tabla["Compradores"] = tabla["Compradores"].apply(fmt_int)
    tabla = tabla.rename(columns={"Compradores": "Compradores únicos"})
    st.dataframe(tabla[["Marca", "Producto", "Unidades", "Compradores únicos", "Total", "PrecioProm"]],
                 use_container_width=True, hide_index=True)

    productos_disp = sorted(ventas_f["Producto"].unique())
    if productos_disp:
        prod_sel = st.selectbox("Producto puntual", productos_disp)

        pytd = M.producto_ytd(ventas, prod_sel, año_actual)
        c1, c2 = st.columns(2)
        c1.metric(f"Unidades YTD {año_actual}", fmt_int(pytd["unid_actual"]), fmt_pct(pytd["var_unid"]))
        c2.metric(f"Unidades YTD {año_actual - 1}", fmt_int(pytd["unid_prev"]))

        st.markdown("**Evolución mensual (unidades)**")
        serie = M.producto_comparativo(ventas, prod_sel)
        fig = go.Figure()
        fig.add_scatter(x=serie["Etiqueta"], y=serie["Unidades"], mode="lines+markers",
                         line=dict(color=ACCENT, width=2), name="Unidades")
        fig.update_layout(height=300, margin=dict(t=10, b=10), plot_bgcolor="white",
                           yaxis=dict(gridcolor=GRID, title="Unidades"))
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("**Evolución trimestral (unidades)**")
        serie_t = M.producto_trimestral(ventas, prod_sel)
        fig2 = go.Figure()
        fig2.add_bar(
            x=serie_t["Etiqueta"], y=serie_t["Unidades"], marker_color=SERIE_1,
            text=[fmt_pct(v) if pd.notna(v) else "" for v in serie_t["VarUnidades"]],
            textposition="outside",
        )
        fig2.update_layout(height=300, margin=dict(t=10, b=10), plot_bgcolor="white",
                            yaxis=dict(gridcolor=GRID, title="Unidades"))
        st.plotly_chart(fig2, use_container_width=True)
        st.dataframe(tabla_variacion_trimestral(serie_t, "Unidades"), use_container_width=True, hide_index=True)

# ── Tab Compras vs Ventas ────────────────────────────────────────────────

with tab_compras:
    st.subheader("Lo que compramos a Cepas Argentina vs. lo que facturamos a nuestros clientes")
    comp = M.comparativo_ventas_compras_mensual(ventas_mp, compras_mp, n_meses=18)
    fig = go.Figure()
    fig.add_bar(x=comp["Etiqueta"], y=comp["Compras"], name="Compras a Cepas", marker_color=SERIE_1)
    fig.add_bar(x=comp["Etiqueta"], y=comp["Ventas"], name="Ventas a clientes", marker_color=SERIE_2)
    fig.update_layout(barmode="group", height=420, margin=dict(t=10, b=10),
                       plot_bgcolor="white", yaxis=dict(gridcolor=GRID))
    st.plotly_chart(fig, use_container_width=True)
    st.caption(
        "Compras: Novara comprándole a Cepas Argentina (costo). Ventas: lo que Novara le "
        "factura a sus propios clientes con esos productos."
    )
