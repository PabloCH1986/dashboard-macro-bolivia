import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(
    page_title="Dashboard Macroeconómico Ejecutivo - Bolivia",
    layout="wide",
    page_icon="📊"
)

EXCEL_FILE = "Info.xlsx"
SHEET_NAME = "data"

# =========================
# ESTILO PREMIUM
# =========================
st.markdown("""
<style>
.main {
    background-color: #0B1020;
}
.block-container {
    padding-top: 1.5rem;
}
h1, h2, h3, h4, h5, h6, p, label {
    color: #F8FAFC !important;
}
[data-testid="stMetric"] {
    background: linear-gradient(135deg, #111827, #1E293B);
    border: 1px solid #334155;
    padding: 18px;
    border-radius: 18px;
    box-shadow: 0 6px 20px rgba(0,0,0,0.25);
}
[data-testid="stMetricLabel"] {
    color: #CBD5E1 !important;
}
[data-testid="stMetricValue"] {
    color: #F8FAFC !important;
    font-size: 28px;
}
[data-testid="stMetricDelta"] {
    font-size: 15px;
}
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
}
.stTabs [data-baseweb="tab"] {
    background-color: #111827;
    color: #E5E7EB;
    border-radius: 12px;
    padding: 10px 18px;
}
.stTabs [aria-selected="true"] {
    background-color: #2563EB;
    color: white;
}
</style>
""", unsafe_allow_html=True)

# =========================
# CARGA DE DATOS
# =========================
@st.cache_data
def cargar_datos():
    raw = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=None)

    nombres = raw.iloc[1].copy()
    nombres.iloc[0] = "fecha"
    nombres = nombres.astype(str).str.strip()

    nombres_unicos = []
    contador = {}
    for n in nombres:
        if n in contador:
            contador[n] += 1
            nombres_unicos.append(f"{n}_{contador[n]}")
        else:
            contador[n] = 0
            nombres_unicos.append(n)

    data = raw.iloc[2:].copy()
    data.columns = nombres_unicos
    data = data.rename(columns={data.columns[0]: "fecha"})

    data["fecha"] = pd.to_datetime(data["fecha"], errors="coerce")
    data = data.dropna(subset=["fecha"])
    data = data.dropna(axis=1, how="all")

    for col in data.columns:
        if col != "fecha":
            data[col] = pd.to_numeric(data[col], errors="coerce")

    return data

df_original = cargar_datos()

# =========================
# FUNCIONES
# =========================
def buscar_columna(texto):
    texto = texto.lower()
    for col in df_original.columns:
        if col != "fecha" and texto in str(col).lower():
            return col
    return None

def ultimo_valor(df, col):
    if col is None:
        return None, None
    s = df[["fecha", col]].dropna()
    if s.empty:
        return None, None
    u = s.iloc[-1]
    return u[col], u["fecha"]

def variacion_interanual(df, col):
    if col is None:
        return None
    s = df[["fecha", col]].dropna().sort_values("fecha")
    if len(s) < 13:
        return None
    actual = s.iloc[-1]
    base_fecha = actual["fecha"] - pd.DateOffset(years=1)
    ant = s[s["fecha"] <= base_fecha]
    if ant.empty:
        return None
    base = ant.iloc[-1][col]
    if base == 0 or pd.isna(base):
        return None
    return ((actual[col] / base) - 1) * 100

def formato_numero(x):
    if x is None or pd.isna(x):
        return "Sin dato"
    return f"{x:,.2f}"

def kpi(df, titulo, col, unidad=""):
    valor, fecha = ultimo_valor(df, col)
    yoy = variacion_interanual(df, col)

    if valor is None:
        st.metric(titulo, "Sin dato")
        return

    delta = f"{yoy:,.1f}% interanual" if yoy is not None else None
    st.metric(titulo, f"{formato_numero(valor)} {unidad}", delta)
    st.caption(f"Último dato: {fecha.strftime('%d/%m/%Y')}")

def grafico_linea(df, col, titulo, unidad=""):
    if col is None:
        st.warning(f"No se encontró: {titulo}")
        return

    s = df[["fecha", col]].dropna()
    if s.empty:
        st.warning(f"Sin datos para: {titulo}")
        return

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=s["fecha"],
        y=s[col],
        mode="lines",
        line=dict(width=3),
        hovertemplate="%{x|%d/%m/%Y}<br>Valor: %{y:,.2f}<extra></extra>"
    ))

    fig.update_layout(
        title=titulo,
        height=430,
        template="plotly_dark",
        paper_bgcolor="#111827",
        plot_bgcolor="#111827",
        font=dict(color="#F8FAFC"),
        margin=dict(l=20, r=20, t=60, b=30),
        xaxis_title="",
        yaxis_title=unidad,
    )
    st.plotly_chart(fig, use_container_width=True)

def grafico_barras(df, cols, titulo):
    cols = [c for c in cols if c is not None]
    if not cols:
        st.warning("No hay variables disponibles.")
        return

    ultimos = []
    for c in cols:
        v, f = ultimo_valor(df, c)
        if v is not None:
            ultimos.append({"Variable": c[:45], "Valor": v})

    if not ultimos:
        st.warning("Sin datos.")
        return

    data = pd.DataFrame(ultimos)
    fig = px.bar(data, x="Variable", y="Valor", title=titulo, template="plotly_dark")
    fig.update_layout(
        height=430,
        paper_bgcolor="#111827",
        plot_bgcolor="#111827",
        font=dict(color="#F8FAFC"),
        margin=dict(l=20, r=20, t=60, b=80)
    )
    st.plotly_chart(fig, use_container_width=True)

def semaforo(nombre, valor, bajo, medio, invertido=False):
    if valor is None:
        estado, color = "Sin dato", "#64748B"
    else:
        if not invertido:
            if valor < bajo:
                estado, color = "Bajo", "#22C55E"
            elif valor < medio:
                estado, color = "Moderado", "#F59E0B"
            else:
                estado, color = "Alto", "#EF4444"
        else:
            if valor > medio:
                estado, color = "Adecuado", "#22C55E"
            elif valor > bajo:
                estado, color = "Moderado", "#F59E0B"
            else:
                estado, color = "Crítico", "#EF4444"

    st.markdown(f"""
    <div style="background:#111827;border:1px solid #334155;border-radius:18px;padding:18px;margin-bottom:10px">
        <h4 style="margin:0;color:#F8FAFC">{nombre}</h4>
        <p style="font-size:28px;margin:8px 0;color:{color};font-weight:700">{estado}</p>
    </div>
    """, unsafe_allow_html=True)

# =========================
# VARIABLES
# =========================
igae = buscar_columna("IGAE")
inflacion_12m = buscar_columna("Variación a doce meses")
rin = buscar_columna("Reservas internacionales netas")
tc_venta = buscar_columna("Valor referencial de venta")
base_monetaria = buscar_columna("Base monetaria")
credito_privado = buscar_columna("Crédito del sistema financiero al sector privado")
depositos = buscar_columna("Depósitos en entidades")
exportaciones = buscar_columna("Exportaciones")
importaciones = buscar_columna("Importaciones")
saldo_comercial = buscar_columna("Saldo Comercial")
divisas = buscar_columna("Divisas")
oro = buscar_columna("Oro")
m1 = buscar_columna("M’1")
m2 = buscar_columna("M’2")
m3 = buscar_columna("M’3")

# =========================
# SIDEBAR
# =========================
st.sidebar.title("⚙️ Panel de control")
fecha_min = df_original["fecha"].min()
fecha_max = df_original["fecha"].max()

rango = st.sidebar.date_input(
    "Rango de fechas",
    value=(fecha_min, fecha_max),
    min_value=fecha_min,
    max_value=fecha_max
)

df = df_original.copy()
if len(rango) == 2:
    inicio = pd.to_datetime(rango[0])
    fin = pd.to_datetime(rango[1])
    df = df[(df["fecha"] >= inicio) & (df["fecha"] <= fin)]

st.sidebar.markdown("---")
st.sidebar.metric("Variables disponibles", len(df.columns) - 1)
st.sidebar.metric("Última fecha", df["fecha"].max().strftime("%d/%m/%Y"))

# =========================
# HEADER
# =========================
st.markdown("""
# 📊 Dashboard Macroeconómico Ejecutivo - Bolivia
### Monitor de coyuntura económica, monetaria, externa y financiera
""")

st.markdown("---")

# =========================
# KPIs
# =========================
c1, c2, c3, c4 = st.columns(4)
with c1:
    kpi(df, "Actividad económica - IGAE", igae, "")
with c2:
    kpi(df, "Inflación interanual", inflacion_12m, "%")
with c3:
    kpi(df, "RIN", rin, "millones $us")
with c4:
    kpi(df, "Tipo de cambio venta", tc_venta, "Bs/$us")

c5, c6, c7, c8 = st.columns(4)
with c5:
    kpi(df, "Base monetaria", base_monetaria, "millones Bs")
with c6:
    kpi(df, "Crédito privado", credito_privado, "millones Bs")
with c7:
    kpi(df, "Depósitos", depositos, "millones Bs")
with c8:
    kpi(df, "Saldo comercial", saldo_comercial, "millones $us")

st.markdown("---")

# =========================
# TABS
# =========================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📈 Resumen",
    "🔥 Inflación",
    "🌎 Sector externo",
    "💵 Monetario",
    "🏦 Financiero",
    "⚠️ Riesgos"
])

with tab1:
    a, b = st.columns(2)
    with a:
        grafico_linea(df, igae, "IGAE - Actividad económica")
    with b:
        grafico_linea(df, inflacion_12m, "Inflación interanual")

    c, d = st.columns(2)
    with c:
        grafico_linea(df, rin, "Reservas internacionales netas")
    with d:
        grafico_linea(df, credito_privado, "Crédito al sector privado")

with tab2:
    grafico_linea(df, inflacion_12m, "Inflación a doce meses", "%")

with tab3:
    a, b = st.columns(2)
    with a:
        grafico_linea(df, rin, "Reservas internacionales netas")
    with b:
        grafico_linea(df, tc_venta, "Tipo de cambio referencial de venta")

    c, d = st.columns(2)
    with c:
        grafico_linea(df, exportaciones, "Exportaciones")
    with d:
        grafico_linea(df, importaciones, "Importaciones")

    grafico_barras(df, [rin, divisas, oro], "Composición reciente de reservas")

with tab4:
    a, b = st.columns(2)
    with a:
        grafico_linea(df, base_monetaria, "Base monetaria")
    with b:
        grafico_barras(df, [m1, m2, m3], "Agregados monetarios")

with tab5:
    a, b = st.columns(2)
    with a:
        grafico_linea(df, credito_privado, "Crédito al sector privado")
    with b:
        grafico_linea(df, depositos, "Depósitos del sistema financiero")

with tab6:
    st.subheader("Semáforo macroeconómico")
    infl_val, _ = ultimo_valor(df, inflacion_12m)
    rin_val, _ = ultimo_valor(df, rin)
    tc_yoy = variacion_interanual(df, tc_venta)
    cred_yoy = variacion_interanual(df, credito_privado)

    r1, r2, r3, r4 = st.columns(4)
    with r1:
        semaforo("Riesgo inflacionario", infl_val, 3, 6)
    with r2:
        semaforo("Posición externa - RIN", rin_val, 2000, 5000, invertido=True)
    with r3:
        semaforo("Presión cambiaria", tc_yoy, 2, 8)
    with r4:
        semaforo("Expansión crediticia", cred_yoy, 5, 15)

    st.info("Los umbrales del semáforo son referenciales y pueden ajustarse según criterio técnico.")

st.markdown("---")

# =========================
# EXPLORADOR
# =========================
st.subheader("🔎 Explorador de variables")
variables = [c for c in df.columns if c != "fecha"]
seleccion = st.selectbox("Selecciona cualquier variable del Excel", variables)
grafico_linea(df, seleccion, seleccion)

st.download_button(
    "⬇️ Descargar base filtrada",
    data=df.to_csv(index=False).encode("utf-8"),
    file_name="base_macro_filtrada.csv",
    mime="text/csv"
)
