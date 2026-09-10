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


def fila_variaciones(vals):
    """HTML de una fila compacta con 2-3 variaciones (label, var_pct) lado a
    lado — pensada para ir debajo de un st.metric grande, sin el achique de
    texto que sufre st.metric cuando se lo mete en columnas angostas. Se
    arma en una sola línea (sin indentación por renglón) para que Markdown
    no la confunda con un bloque de código, como pasó con la versión
    anterior de las tarjetas de Vermouth."""
    partes = []
    for label, v in vals:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            txt, color = "—", TEXT_SECONDARY
        else:
            positivo = v >= 0
            color = GOOD if positivo else BAD
            flecha = "▲" if positivo else "▼"
            txt = f"{flecha} {fmt_pct(v)}"
        partes.append(
            f'<div style="flex:1;white-space:nowrap;"><div style="font-size:11px;color:{TEXT_SECONDARY};'
            f'white-space:nowrap;">{label}</div>'
            f'<div style="font-weight:700;font-size:13px;color:{color};white-space:nowrap;">{txt}</div></div>'
        )
    return f'<div style="display:flex;gap:14px;margin-top:10px;">{"".join(partes)}</div>'


def fmt_share(n):
    """Como fmt_pct pero sin el signo '+' — para columnas de participación
    (% del total), no de variación."""
    if n is None or (isinstance(n, float) and pd.isna(n)):
        return "—"
    try:
        return f"{n*100:.1f}%"
    except Exception:
        return "—"


def color_var(v):
    """Verde si la variación es >= 0, rojo si es negativa — para columnas de
    Var. % en cualquier tabla del tablero."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        return ("background-color: #c6efce; color: #006100" if v >= 0
                else "background-color: #ffc7ce; color: #9c0006")
    except Exception:
        return ""


def estilizar_tabla(df, int_cols=None, money_cols=None, share_cols=None, var_cols=None):
    """Arma un pandas Styler para mostrar una tabla con números formateados
    (miles con punto, $ con miles, % con un decimal) SIN convertirlos a texto
    en el DataFrame — Streamlit ordena/filtra sobre el valor crudo aunque la
    celda se vea formateada, así que esto es lo que mantiene el sort nativo
    funcionando bien (el bug de 'no ordena bien' era por pre-convertir los
    números a string antes de pasarlos a st.dataframe). var_cols además se
    colorea rojo/verde."""
    fmt = {}
    for c in (int_cols or []):
        fmt[c] = fmt_int
    for c in (money_cols or []):
        fmt[c] = fmt_money
    for c in (share_cols or []):
        fmt[c] = fmt_share
    for c in (var_cols or []):
        fmt[c] = fmt_pct
    styler = df.style.format(fmt)
    if var_cols:
        styler = styler.map(color_var, subset=var_cols)
    return styler


def tabla_variacion_trimestral(trimestral, col_val="Unidades"):
    var_qoq_col = "VarUnidades" if col_val == "Unidades" else "VarTotal"
    var_yoy_col = "VarUnidadesYoY" if col_val == "Unidades" else "VarTotalYoY"
    out = pd.DataFrame({
        "Trimestre": trimestral["Etiqueta"],
        col_val: trimestral[col_val],
        "Var. trim. anterior": trimestral[var_qoq_col],
        "Var. interanual": trimestral[var_yoy_col],
    })
    cols_num = [col_val] if col_val == "Unidades" else []
    cols_money = [col_val] if col_val == "Total" else []
    return estilizar_tabla(out, int_cols=cols_num, money_cols=cols_money,
                            var_cols=["Var. trim. anterior", "Var. interanual"])


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
    ["Resumen", "Mix de producto", "Evolución", "Clientes", "Vermouth",
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

    st.subheader("Marca x Trimestre (unidades)")
    tabla_marca = M.tabla_marca_trimestral(ventas_mp, año_actual)
    st.dataframe(render_tabla_marca_trimestral(tabla_marca, año_actual), use_container_width=True)

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
    st.subheader("Mix de producto — Q3 en curso")
    mix = M.mix_por_marca(ventas_f)

    col1, col2 = st.columns([1, 1])
    with col1:
        fig = go.Figure(go.Pie(
            labels=mix["Marca"], values=mix["Unidades"], hole=0.55,
            marker_colors=[MARCA_COLORS.get(m, "#898781") for m in mix["Marca"]],
            textinfo="label+percent",
        ))
        fig.update_layout(height=420, margin=dict(t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        tabla_mix = mix.rename(columns={"Participacion": "% Unidades"})
        tabla_estilo = estilizar_tabla(tabla_mix, money_cols=["Total"], int_cols=["Unidades", "Clientes"],
                                        share_cols=["% Unidades"])
        st.dataframe(tabla_estilo, use_container_width=True, hide_index=True)

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

    st.dataframe(tabla_variacion_trimestral(trimestral, col_val), use_container_width=True, hide_index=True)
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
        f"Cantidad de clientes DISTINTOS que compraron cada marca en cada trimestre {año_cli}."
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
    tabla = ranking.rename(columns={"NombreFantasia": "Nombre", "Participacion": "% Unidades"})
    tabla = tabla[["Cliente", "Nombre", "Total", "Unidades", "Comprobantes", "% Unidades"]]
    tabla_estilo = estilizar_tabla(tabla, money_cols=["Total"], int_cols=["Unidades", "Comprobantes"],
                                    share_cols=["% Unidades"])
    st.dataframe(tabla_estilo, use_container_width=True, hide_index=True)

# ── Tab Vermouth vs. competencia ─────────────────────────────────────────

with tab_vermouth:
    st.subheader("Vermouth")

    año_cards = pd.Timestamp.now().year
    resumen_cat = M.resumen_anual_categoria_vermouth(ventas_verm, año_cards)
    prev, act = resumen_cat["prev"], resumen_cat["act"]

    cc1, cc2, cc3, cc4 = st.columns(4)
    with cc1.container(border=True):
        st.metric(f"Facturación {año_cards - 1}", fmt_money(prev["total"]))
    with cc2.container(border=True):
        st.metric(f"Unidades {año_cards - 1}", fmt_int(prev["unidades"]))
    for col, titulo, valor, campo in (
        (cc3, f"Facturación YTD {año_cards}", fmt_money(act["total"]), "total"),
        (cc4, f"Unidades YTD {año_cards}", fmt_int(act["unidades"]), "unid"),
    ):
        with col.container(border=True):
            st.metric(titulo, valor)
            st.markdown(fila_variaciones([
                ("Mensual", act[f"var_mensual_{campo}"]),
                ("Interanual", act[f"var_interanual_{campo}"]),
                ("Acumulado", act[f"var_acumulado_{campo}"]),
            ]), unsafe_allow_html=True)
    st.caption("Categoría Vermouth completa (todas las variantes) — Martini Rosso, Cinzano y el resto de las marcas juntas.")

    fecha_min_v, fecha_max_v = ventas_verm["Fecha"].min().date(), ventas_verm["Fecha"].max().date()
    hoy_date = pd.Timestamp.now().date()
    default_desde_v = max(fecha_min_v, pd.Timestamp(year=hoy_date.year, month=1, day=1).date())
    default_hasta_v = min(fecha_max_v, hoy_date)
    rango_v = st.date_input("Rango de fechas (vermouth)", value=(default_desde_v, default_hasta_v),
                              min_value=fecha_min_v, max_value=fecha_max_v, key="verm_rango")
    fv_desde, fv_hasta = (rango_v if isinstance(rango_v, tuple) and len(rango_v) == 2
                           else (default_desde_v, default_hasta_v))

    ventas_verm_mix = M.filtrar_solo_rosso(ventas_verm)
    ranking_verm = M.ranking_marcas_vermouth(ventas_verm_mix, fecha_desde=fv_desde, fecha_hasta=fv_hasta)

    st.subheader("Vermouth Rosso")
    col1, col2 = st.columns([1, 1])
    with col1:
        fig = go.Figure(go.Pie(
            labels=ranking_verm["MarcaVermouth"], values=ranking_verm["Unidades"], hole=0.55,
            marker_colors=[VERMOUTH_COLORS.get(m, "#898781") for m in ranking_verm["MarcaVermouth"]],
            textinfo="percent", textposition="inside",
        ))
        fig.update_layout(height=460, margin=dict(t=10, b=80),
                           legend=dict(orientation="h", yanchor="top", y=-0.1, font=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        tabla_v = ranking_verm.rename(columns={"MarcaVermouth": "Marca", "ParticipacionUnidades": "% Unidades"})
        tabla_v_estilo = estilizar_tabla(tabla_v, money_cols=["Total"], int_cols=["Unidades", "Clientes"],
                                          share_cols=["% Unidades"])
        st.dataframe(tabla_v_estilo, use_container_width=True, hide_index=True)

    # NOTA: el ranking de arriba (ranking_verm) ya está filtrado a solo Rosso.
    # Todo lo que sigue en la pestaña (cara a cara, evolución, overlap de
    # clientes) sigue usando ventas_verm completo a propósito, salvo donde
    # se indique lo contrario — son comparaciones puntuales Martini Rosso vs
    # Cinzano Rosso, no dependen del recorte de arriba.

    st.markdown("---")
    st.subheader(f"Evolución YTD: vermouth Rosso, {pd.Timestamp.now().year} vs. {pd.Timestamp.now().year - 1}")
    año_rosso = pd.Timestamp.now().year
    tabla_rosso, total_rosso_prev, total_rosso_act = M.resumen_rosso_ytd_por_marca(ventas_verm, año_rosso)
    tabla_rosso = tabla_rosso[tabla_rosso["Marca"] != "Total vermouth Rosso"]

    cr1, cr2 = st.columns(2)
    cr1.metric(f"Unidades vermouth Rosso YTD {año_rosso - 1}", fmt_int(total_rosso_prev))
    cr2.metric(f"Unidades vermouth Rosso YTD {año_rosso}", fmt_int(total_rosso_act),
               fmt_pct(M.var(total_rosso_prev, total_rosso_act)))

    tabla_rosso_estilo = estilizar_tabla(
        tabla_rosso,
        int_cols=[f"Unidades {año_rosso - 1}", f"Unidades {año_rosso}"],
        share_cols=[f"Participación {año_rosso - 1}", f"Participación {año_rosso}"],
        var_cols=["Var. % unidades"],
    )
    st.dataframe(tabla_rosso_estilo, use_container_width=True, hide_index=True)

    st.subheader("Martini Rosso vs Cinzano Rosso en 2026")

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
    n_solo_mr, n_ambas, n_solo_ci = len(overlap["solo_a"]), len(overlap["ambas"]), len(overlap["solo_b"])
    st.markdown(
        f"**Es el desglose de los {fmt_int(cli_mr)} / {fmt_int(cli_ci)} clientes de arriba:** "
        f"Martini Rosso = {n_solo_mr} exclusivos + {n_ambas} que compran las dos "
        f"({n_solo_mr}+{n_ambas}={n_solo_mr + n_ambas}) · Cinzano Rosso = {n_solo_ci} exclusivos + "
        f"{n_ambas} que compran las dos ({n_solo_ci}+{n_ambas}={n_solo_ci + n_ambas})"
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
    st.subheader("Evolución: Martini Rosso vs. Otro Rosso (unidades)")
    marcas_evol = st.multiselect("Marcas a comparar", ranking_verm["MarcaVermouth"].tolist(),
                                   default=[m for m in ["Martini Rosso", "Cinzano Rosso"] if m in ranking_verm["MarcaVermouth"].tolist()])
    if marcas_evol:
        st.markdown("**Mes a mes**")
        evol_verm_mes = M.evolucion_mensual_por_marca_vermouth(ventas_verm, marcas=marcas_evol, n_meses=18)
        fig_mes = go.Figure()
        for marca in marcas_evol:
            d = evol_verm_mes[evol_verm_mes["MarcaVermouth"] == marca]
            fig_mes.add_scatter(x=d["Etiqueta"], y=d["Unidades"], mode="lines+markers", name=marca,
                                 line=dict(color=VERMOUTH_COLORS.get(marca, "#898781"), width=2), marker=dict(size=6))
        fig_mes.update_layout(height=380, margin=dict(t=10, b=10), plot_bgcolor="white",
                               yaxis=dict(gridcolor=GRID, title="Unidades"),
                               legend=dict(orientation="h", yanchor="bottom", y=1.02))
        st.plotly_chart(fig_mes, use_container_width=True)

        st.markdown("**Por trimestre**")
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

    st.subheader(f"Resultados mensuales, por SKU — {año_actual} (todos los productos Cepas)")
    valor_mensual = st.radio("Ver en", ["Unidades", "Pesos"], horizontal=True, index=0, key="mensual_sku_valor")
    col_valor = "Unidades" if valor_mensual == "Unidades" else "Total"
    ventas_año_actual = ventas[ventas["Fecha"].dt.year == año_actual]
    piv_mensual = M.pivot_mensual_por_producto(ventas_año_actual, valor=col_valor)
    cols_mes = [c for c in piv_mensual.columns if c not in ("Marca", "Producto")]
    piv_mensual_estilo = estilizar_tabla(
        piv_mensual,
        int_cols=cols_mes if col_valor == "Unidades" else [],
        money_cols=cols_mes if col_valor == "Total" else [],
    )
    st.dataframe(piv_mensual_estilo, use_container_width=True, hide_index=True)

    st.subheader("Ranking de productos YTD, en unidades")
    ranking_ytd = M.ranking_productos_ytd(ventas, año_actual, marcas=marcas_sel or None, productos=productos_sel or None)
    tabla_ytd = ranking_ytd.rename(columns={
        "UnidadesActual": f"Unidades YTD {año_actual}",
        "UnidadesPrev": f"Unidades YTD {año_actual - 1}",
        "CompradoresActual": "Compradores únicos",
        "VarUnidades": "Var. % unidades",
    })[["Producto", f"Unidades YTD {año_actual - 1}", f"Unidades YTD {año_actual}", "Var. % unidades", "Compradores únicos"]]
    tabla_ytd_estilo = estilizar_tabla(
        tabla_ytd,
        int_cols=[f"Unidades YTD {año_actual}", f"Unidades YTD {año_actual - 1}", "Compradores únicos"],
        var_cols=["Var. % unidades"],
    )
    st.dataframe(tabla_ytd_estilo, use_container_width=True, hide_index=True)

    productos_disp = sorted(ventas_f["Producto"].unique())
    if productos_disp:
        prod_sel = st.selectbox("Producto puntual", productos_disp)

        pytd = M.producto_ytd(ventas, prod_sel, año_actual)
        c1, c2 = st.columns(2)
        c1.metric(f"Unidades YTD {año_actual - 1}", fmt_int(pytd["unid_prev"]))
        c2.metric(f"Unidades YTD {año_actual}", fmt_int(pytd["unid_actual"]), fmt_pct(pytd["var_unid"]))

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
