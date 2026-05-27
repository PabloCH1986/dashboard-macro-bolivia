import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="Dashboard Macroeconómico Ejecutivo - Bolivia",
    layout="wide"
)

EXCEL_FILE = "Info.xlsx"
SHEET_NAME = "data"

st.title("Dashboard Macroeconómico Ejecutivo - Bolivia")
st.caption("Fuente: base macroeconómica en Excel")

@st.cache_data
def cargar_datos():
    raw = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=None)

    nombres = raw.iloc[1].copy()
    nombres.iloc[0] = "fecha"

    # Evitar columnas repetidas
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
    data = data.dropna(subset=["fecha"])
    data["fecha"] = pd.to_datetime(data["fecha"], errors="coerce")
    data = data.dropna(subset=["fecha"])

    data = data.dropna(axis=1, how="all")

    for col in data.columns:
        if col != "fecha":
            data[col] = pd.to_numeric(data[col], errors="coerce")

    return data

df = cargar_datos()

# -----------------------------
# Funciones auxiliares
# -----------------------------

def buscar_columna(texto):
    texto = texto.lower()
    for col in df.columns:
        if col != "fecha" and texto in str(col).lower():
            return col
    return None

def ultimo_valor(col):
    if col is None:
        return None, None

    serie = df[["fecha", col]].dropna()
    if serie.empty:
        return None, None

    ultimo = serie.iloc[-1]
    return ultimo[col], ultimo["fecha"]

def variacion_interanual(col):
    if col is None:
        return None

    serie = df[["fecha", col]].dropna().sort_values("fecha")
    if len(serie) < 13:
        return None

    actual = serie.iloc[-1]
    fecha_base = actual["fecha"] - pd.DateOffset(years=1)

    anterior = serie[serie["fecha"] <= fecha_base]
    if anterior.empty:
        return None

    base = anterior.iloc[-1][col]

    if base == 0:
        return None

    return ((actual[col] / base) - 1) * 100

def kpi(titulo, col, unidad=""):
    valor, fecha = ultimo_valor(col)
    var_yoy = variacion_interanual(col)

    if valor is None:
        st.metric(titulo, "Sin dato")
        return

    valor_txt = f"{valor:,.2f}"
    if unidad:
        valor_txt += f" {unidad}"

    delta = None
    if var_yoy is not None:
        delta = f"{var_yoy:,.1f}% interanual"

    st.metric(titulo, valor_txt, delta)
    st.caption(f"Último dato: {fecha.strftime('%d/%m/%Y')}")

def grafico_linea(col, titulo):
    if col is None:
        st.warning(f"No se encontró la variable: {titulo}")
        return

    serie = df[["fecha", col]].dropna()

    if serie.empty:
        st.warning(f"La variable no tiene datos: {titulo}")
        return

    fig = px.line(
        serie,
        x="fecha",
        y=col,
        title=titulo
    )

    fig.update_layout(
        template="plotly_white",
        height=420,
        margin=dict(l=20, r=20, t=60, b=20),
        xaxis_title="Fecha",
        yaxis_title=""
    )

    st.plotly_chart(fig, use_container_width=True)

# -----------------------------
# Variables recomendadas
# -----------------------------

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
m1 = buscar_columna("M’1")
m2 = buscar_columna("M’2")
m3 = buscar_columna("M’3")
divisas = buscar_columna("Divisas")
oro = buscar_columna("Oro")

# -----------------------------
# Sidebar
# -----------------------------

st.sidebar.header("Filtros")

fecha_min = df["fecha"].min()
fecha_max = df["fecha"].max()

rango = st.sidebar.date_input(
    "Rango de fechas",
    value=(fecha_min, fecha_max),
    min_value=fecha_min,
    max_value=fecha_max
)

if len(rango) == 2:
    inicio = pd.to_datetime(rango[0])
    fin = pd.to_datetime(rango[1])
    df = df[(df["fecha"] >= inicio) & (df["fecha"] <= fin)]

st.sidebar.markdown("---")
st.sidebar.write(f"Variables disponibles: {len(df.columns)-1}")

# -----------------------------
# KPIs principales
# -----------------------------

st.subheader("Indicadores principales")

c1, c2, c3, c4 = st.columns(4)

with c1:
    kpi("Actividad económica - IGAE", igae, "índice")

with c2:
    kpi("Inflación interanual", inflacion_12m, "%")

with c3:
    kpi("Reservas internacionales netas", rin, "millones $us")

with c4:
    kpi("Tipo de cambio venta", tc_venta, "Bs/$us")

c5, c6, c7, c8 = st.columns(4)

with c5:
    kpi("Base monetaria", base_monetaria, "millones Bs")

with c6:
    kpi("Crédito al sector privado", credito_privado, "millones Bs")

with c7:
    kpi("Depósitos", depositos, "millones Bs")

with c8:
    kpi("Saldo comercial", saldo_comercial, "millones $us")

st.markdown("---")

# -----------------------------
# Gráficos principales
# -----------------------------

st.subheader("Actividad económica e inflación")

g1, g2 = st.columns(2)

with g1:
    grafico_linea(igae, "IGAE - Actividad económica")

with g2:
    grafico_linea(inflacion_12m, "Inflación a doce meses")

st.subheader("Sector externo y cambiario")

g3, g4 = st.columns(2)

with g3:
    grafico_linea(rin, "Reservas internacionales netas")

with g4:
    grafico_linea(tc_venta, "Tipo de cambio referencial de venta")

st.subheader("Sector monetario y financiero")

g5, g6 = st.columns(2)

with g5:
    grafico_linea(base_monetaria, "Base monetaria")

with g6:
    grafico_linea(credito_privado, "Crédito al sector privado")

st.subheader("Comercio exterior")

g7, g8 = st.columns(2)

with g7:
    grafico_linea(exportaciones, "Exportaciones")

with g8:
    grafico_linea(importaciones, "Importaciones")

st.subheader("Reservas: composición")

g9, g10 = st.columns(2)

with g9:
    grafico_linea(divisas, "Divisas")

with g10:
    grafico_linea(oro, "Oro")

st.markdown("---")

# -----------------------------
# Explorador
# -----------------------------

st.subheader("Explorador de variables")

variables = [c for c in df.columns if c != "fecha"]

seleccion = st.selectbox(
    "Selecciona cualquier variable del Excel",
    variables
)

grafico_linea(seleccion, seleccion)

st.info("Dashboard construido con base transpuesta: fechas en filas y variables en columnas.")