import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import uuid
import base64
import os
import unicodedata

# =========================
# CONFIGURACIÓN GENERAL
# =========================

st.set_page_config(
    page_title="Dashboard Macroeconómico CENGOB - Bolivia",
    layout="wide",
    page_icon="📊"
)

EXCEL_FILE = "Info.xlsx"
SHEET_NAME = "data"

# =========================
# TEMA AUTOMÁTICO
# =========================

st.markdown("""
<style>

.stApp{
    background-color:#EEF2F5;
}

/* HEADER */
[data-testid="stHeader"]{
    background:#0B3B36;
}

/* SIDEBAR */
[data-testid="stSidebar"]{
    background:#F8FAFC;
}

/* TITULOS */
h1,h2,h3,h4,h5,h6{
    color:#0B3B36 !important;
    font-weight:800 !important;
}

/* TEXTO */
p,label{
    color:#1E293B;
}

/* TABS */
.stTabs [data-baseweb="tab"]{
    background:#E7E5E4;
    color:#1E293B !important;
    border-radius:12px;
    padding:10px 18px;
    font-weight:700;
}

.stTabs [aria-selected="true"] *{
    color:white !important;
}

.stTabs [data-baseweb="tab"] *{
    color:#0F172A !important;
}

/* METRICAS */
[data-testid="stMetric"]{
    background:#F8FAFC;
    border-left:5px solid #C9A227;
    border-radius:18px;
    padding:18px;
    box-shadow:0 4px 12px rgba(0,0,0,0.05);
    border-top:1px solid #E5E7EB;
    border-right:1px solid #E5E7EB;
    border-bottom:1px solid #E5E7EB;
    transition:0.3s;
}

/* GRAFICOS */
.js-plotly-plot{
    border-radius:18px;
    overflow:hidden;
}

/* ALERTAS */
.stAlert{
    border-left:5px solid #C8A951;
    border-radius:14px;
}

</style>
""", unsafe_allow_html=True)

# =========================
# CARGA DE DATOS
# =========================

@st.cache_data(ttl=60)
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


try:
    df_original = cargar_datos()
except Exception as e:
    st.error(
        "No se pudo cargar la base de datos. Verifica que el archivo "
        f"'{EXCEL_FILE}' exista y que la hoja se llame '{SHEET_NAME}'."
    )
    st.exception(e)
    st.stop()

# =========================
# FUNCIONES BASE
# =========================

def normalizar_texto(texto):
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto


def buscar_columna(texto):
    texto_norm = normalizar_texto(texto)

    for col in df_original.columns:
        if col != "fecha" and texto_norm in normalizar_texto(col):
            return col

    return None


def buscar_columna_multiple(opciones):
    for opcion in opciones:
        col = buscar_columna(opcion)
        if col is not None:
            return col
    return None


def ultimo_valor(df, col):
    if col is None:
        return None, None

    if col not in df.columns:
        return None, None

    s = df[["fecha", col]].dropna()

    if s.empty:
        return None, None

    s = s.sort_values("fecha")
    u = s.iloc[-1]

    return u[col], u["fecha"]


def variacion_interanual(df, col):
    if col is None:
        return None

    if col not in df.columns:
        return None

    s = df[["fecha", col]].dropna().sort_values("fecha")

    if len(s) < 2:
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

    st.markdown(
        f"""
        <p style="
            color:#0B3B36;
            font-size:15px;
            margin-top:6px;
            font-weight:500;
        ">
            Último dato: {fecha.strftime('%d/%m/%Y')}
        </p>
        """,
        unsafe_allow_html=True
    )


def grafico_linea(df, col, titulo, unidad=""):
    if col is None:
        st.warning(f"No se encontró: {titulo}")
        return

    if col not in df.columns:
        st.warning(f"No se encontró la columna para: {titulo}")
        return

    s = df[["fecha", col]].dropna().sort_values("fecha")

    if s.empty:
        st.warning(f"Sin datos para: {titulo}")
        return

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=s["fecha"],
            y=s[col],
            mode="lines",
            line=dict(width=3.5, color="#0B3B36"),
            hovertemplate="%{x|%d/%m/%Y}<br>Valor: %{y:,.2f}<extra></extra>"
        )
    )

    fig.update_layout(
        title=titulo,
        height=430,
        template="plotly_white",
        paper_bgcolor="#DCEAF7",
        plot_bgcolor="#DCEAF7",
        font=dict(color="#000000", size=14),
        title_font=dict(color="#000000", size=22),
        margin=dict(l=20, r=20, t=60, b=30),
        xaxis_title="",
        yaxis_title=unidad,
        hovermode="x unified",
        legend=dict(font=dict(color="#000000", size=13)),
        xaxis=dict(
            rangeselector=dict(
                bgcolor="#FFFFFF",
                activecolor="#0B3B36",
                font=dict(color="#000000", size=13),
                buttons=list([
                    dict(count=1, label="1A", step="year", stepmode="backward"),
                    dict(count=5, label="5A", step="year", stepmode="backward"),
                    dict(count=10, label="10A", step="year", stepmode="backward"),
                    dict(step="all", label="Todo")
                ])
            ),
            rangeslider=dict(
                visible=True,
                bgcolor="#CFE3F5",
                bordercolor="#94A3B8"
            ),
            type="date",
            tickfont=dict(color="#000000"),
            gridcolor="rgba(0,0,0,0.08)"
        ),
        yaxis=dict(
            tickfont=dict(color="#000000"),
            gridcolor="rgba(0,0,0,0.12)",
            zerolinecolor="rgba(0,0,0,0.25)"
        )
    )

    fig.update_traces(line=dict(width=3.5))
    fig.update_xaxes(showgrid=True, gridcolor="rgba(0,0,0,0.08)")
    fig.update_yaxes(gridcolor="#D1D5DB")

    st.plotly_chart(
        fig,
        use_container_width=True,
        key=f"linea_{titulo}_{uuid.uuid4()}"
    )


def grafico_lineas_multiples(df, cols, titulo, unidad=""):
    cols = [c for c in cols if c is not None and c in df.columns]

    if not cols:
        st.warning(f"No hay variables disponibles para: {titulo}")
        return

    fig = go.Figure()

    colores = [
        "#0B3B36",
        "#C9A227",
        "#556B2F",
        "#C2410C",
        "#475569",
        "#2563EB",
        "#7C3AED"
    ]

    nombres = {
        credito_privado: "Crédito privado",
        depositos: "Depósitos",
        m1: "M1",
        m2: "M2",
        m3: "M3",
        tc_venta: "TC Referencial",
        tc_oficial: "TC Oficial",
        bol_dep: "Boliv. depósitos",
        bol_cred: "Boliv. créditos",
        consumo_hogares: "Consumo hogares",
        consumo_publico: "Consumo público",
        formacion_capital: "Formación bruta de capital",
        expo_bienes_servicios: "Exportaciones",
        impo_bienes_servicios: "Importaciones",
        ingresos_totales_spnf: "Ingresos totales",
        egresos_totales_spnf: "Egresos totales",
        resultado_corriente_spnf: "Resultado corriente",
        resultado_global_spnf: "Resultado global",
        ingresos_corrientes_spnf: "Ingresos corrientes",
        ingresos_capital_spnf: "Ingresos de capital",
        egresos_corrientes_spnf: "Egresos corrientes",
        egresos_capital_spnf: "Egresos de capital",
        pobreza_bolivia: "Pobreza",
        desocupacion_nacional: "Desocupación"
    }

    for i, c in enumerate(cols):
        s = df[["fecha", c]].dropna().sort_values("fecha")

        if not s.empty:
            fig.add_trace(
                go.Scatter(
                    x=s["fecha"],
                    y=s[c],
                    mode="lines",
                    name=nombres.get(c, c),
                    line=dict(
                        width=3.5,
                        color=colores[i % len(colores)]
                    ),
                    hovertemplate="%{x|%d/%m/%Y}<br>%{y:,.2f}<extra></extra>"
                )
            )

    fig.update_layout(
        title=titulo,
        height=430,
        template="plotly_white",
        paper_bgcolor="#DCEAF7",
        plot_bgcolor="#DCEAF7",
        font=dict(color="#000000", size=14),
        title_font=dict(color="#000000", size=20),
        margin=dict(l=20, r=20, t=60, b=30),
        xaxis_title="",
        yaxis_title=unidad,
        hovermode="x unified",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(color="#000000", size=13)
        ),
        xaxis=dict(
            rangeselector=dict(
                bgcolor="#FFFFFF",
                activecolor="#0B3B36",
                font=dict(color="#000000", size=13),
                buttons=list([
                    dict(count=1, label="1A", step="year", stepmode="backward"),
                    dict(count=5, label="5A", step="year", stepmode="backward"),
                    dict(count=10, label="10A", step="year", stepmode="backward"),
                    dict(step="all", label="Todo")
                ])
            ),
            rangeslider=dict(
                visible=True,
                bgcolor="#CFE3F5",
                bordercolor="#94A3B8"
            ),
            type="date",
            tickfont=dict(color="#000000"),
            gridcolor="rgba(0,0,0,0.08)"
        ),
        yaxis=dict(
            tickfont=dict(color="#000000"),
            gridcolor="rgba(0,0,0,0.12)",
            zerolinecolor="rgba(0,0,0,0.25)"
        )
    )

    fig.update_yaxes(
        gridcolor="rgba(0,0,0,0.12)",
        tickfont=dict(color="#000000")
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        key=f"multi_{titulo}_{uuid.uuid4()}"
    )


def grafico_barras(df, cols, titulo):
    cols = [c for c in cols if c is not None and c in df.columns]

    if not cols:
        st.warning("No hay variables disponibles.")
        return

    ultimos = []

    for c in cols:
        v, f = ultimo_valor(df, c)

        if v is not None:
            ultimos.append({
                "Variable": str(c)[:45],
                "Valor": v
            })

    if not ultimos:
        st.warning("Sin datos.")
        return

    data = pd.DataFrame(ultimos)

    fig = px.bar(
        data,
        x="Variable",
        y="Valor",
        title=titulo,
        template="plotly_white"
    )

    fig.update_layout(
        height=430,
        paper_bgcolor="#DCEAF7",
        plot_bgcolor="#DCEAF7",
        font=dict(color="#000000"),
        margin=dict(l=20, r=20, t=60, b=80),
        xaxis_title="",
        yaxis_title="Valor"
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        key=f"barras_{titulo}_{uuid.uuid4()}"
    )


def alerta_sector(titulo, mensaje, nivel="info"):
    iconos = {
        "info": "ℹ️",
        "ok": "✅",
        "warning": "⚠️",
        "danger": "🚨"
    }

    colores = {
        "info": "#DCEAF7",
        "ok": "#DCFCE7",
        "warning": "#FEF3C7",
        "danger": "#FEE2E2"
    }

    bordes = {
        "info": "#2563EB",
        "ok": "#16A34A",
        "warning": "#D97706",
        "danger": "#DC2626"
    }

    textos = {
        "info": "#0F172A",
        "ok": "#14532D",
        "warning": "#78350F",
        "danger": "#7F1D1D"
    }

    st.markdown(f"""
    <div style="
        background:{colores.get(nivel, '#DCEAF7')};
        border-left:7px solid {bordes.get(nivel, '#2563EB')};
        border-radius:16px;
        padding:16px 20px;
        margin:10px 0 22px 0;
        box-shadow:0 3px 10px rgba(0,0,0,0.05);
    ">
        <div style="
            font-size:18px;
            font-weight:800;
            color:{textos.get(nivel, '#0F172A')};
            margin-bottom:6px;
        ">
            {iconos.get(nivel, 'ℹ️')} {titulo}
        </div>
        <div style="
            font-size:15px;
            color:{textos.get(nivel, '#0F172A')};
            line-height:1.5;
        ">
            {mensaje}
        </div>
    </div>
    """, unsafe_allow_html=True)


def tarjeta_riesgo(titulo, nivel):

    colores = {
        "Alto": "#8B1A1A",
        "Moderado": "#8A3A0A",
        "Bajo": "#14532D",
        "Sin dato": "#475569"
    }

    borde = {
        "Alto": "#EF4444",
        "Moderado": "#F59E0B",
        "Bajo": "#22C55E",
        "Sin dato": "#94A3B8"
    }

    color = colores.get(nivel, "#475569")
    line = borde.get(nivel, "#94A3B8")

    html = f"""
    <div style="
        background:{color};
        padding:24px;
        border-radius:20px;
        border:2px solid {line};
        min-height:150px;
        box-shadow:0 8px 22px rgba(0,0,0,0.18);
        font-family:Arial, sans-serif;
    ">
        <div style="
            color:#FFFFFF;
            font-size:21px;
            font-weight:800;
            margin-bottom:32px;
        ">
            {titulo}
        </div>

        <div style="
            color:#FFFFFF;
            font-size:42px;
            font-weight:900;
        ">
            {nivel}
        </div>
    </div>
    """

    components.html(html, height=210)

def clasificar_normal(valor, bajo, medio):
    if valor is None:
        return "Sin dato"

    if valor < bajo:
        return "Bajo"

    if valor < medio:
        return "Moderado"

    return "Alto"


def clasificar_invertido(valor, bajo, medio):
    if valor is None:
        return "Sin dato"

    if valor > medio:
        return "Bajo"

    if valor > bajo:
        return "Moderado"

    return "Alto"

# =========================
# VARIABLES ECONÓMICAS
# =========================

# Sector real
igae = buscar_columna("IGAE")
pib_pm = buscar_columna("PIB a precios de mercado")
consumo_publico = buscar_columna("Gasto de consumo final de la administración pública")
consumo_hogares = buscar_columna("Gasto de consumo final de los hogares")
formacion_capital = buscar_columna("Formación bruta de capital")
expo_bienes_servicios = buscar_columna("Exportaciones de bienes y servicios")
impo_bienes_servicios = buscar_columna_multiple([
    "importaciones bienes y servicios",
    "Importaciones de bienes y servicios",
    "importaciones de bienes y servicios"
])

# Sector precios
inflacion_12m = buscar_columna("Variación a doce meses")
inflacion_mensual = buscar_columna_multiple([
    "Variación mensual inflacion total",
    "Variación mensual inflación total"
])
inflacion_acumulada = buscar_columna("Variación acumulada en el año")

# Sector externo
rin = buscar_columna("Reservas internacionales netas")
tc_venta = buscar_columna("Valor referencial de venta")
tc_oficial = buscar_columna("Tipo de cambio oficial")

if tc_oficial is None:
    tc_oficial = buscar_columna("Tipo de cambio de venta")

exportaciones = buscar_columna("Exportaciones")
importaciones = buscar_columna("Importaciones")
saldo_comercial = buscar_columna("Saldo Comercial")
divisas = buscar_columna("Divisas")
oro = buscar_columna("Oro")
oro_ley_tn = buscar_columna("d/c Oro según Ley N°1503 en Tn")
oro_ley = buscar_columna("d/c Oro según Ley N°1503")
recursos_alta_liquidez = buscar_columna("Recursos de Alta Liquidez")
oro_convertible = buscar_columna("Oro convertible en divisas")
posicion_fmi = buscar_columna("Posición con el FMI")

# Sector financiero
credito_privado = buscar_columna("Crédito del sistema financiero al sector privado")
depositos = buscar_columna("Depósitos en entidades")
bol_dep = buscar_columna("Bolivianización Depósitos")
bol_cred = buscar_columna("Bolivianización Créditos")

# Sector monetario
base_monetaria = buscar_columna("Base monetaria")
m1 = buscar_columna("M’1")
m2 = buscar_columna("M’2")
m3 = buscar_columna("M’3")

# Sector fiscal
ingresos_totales_spnf = buscar_columna("Ingresos Totales SPNF")
ingresos_corrientes_spnf = buscar_columna("Ingresos Corrientes del SPNF")
ingresos_capital_spnf = buscar_columna("Ingresos de Capital del SPNF")

egresos_totales_spnf = buscar_columna("Egresos Totales del SPNF")
egresos_corrientes_spnf = buscar_columna("Egresos Corrientes del SPNF")
egresos_capital_spnf = buscar_columna("Egresos de Capital del SPNF")

resultado_corriente_spnf = buscar_columna("Resultado Fiscal Corriente del SPNF")
resultado_global_spnf = buscar_columna("Resultado Fiscal Global del SPNF")

# Sector social
pobreza_bolivia = buscar_columna_multiple([
    "Bolivia: Indidencia de pobreza",
    "Bolivia: Incidencia de pobreza",
    "Incidencia de pobreza"
])
gini_bolivia = buscar_columna_multiple([
    "Bolivia: Índice de GINI",
    "Bolivia: Indice de GINI",
    "Índice de GINI",
    "Indice de GINI"
])
desocupacion_nacional = buscar_columna("Tasa de Desocupación Nacional")

# =========================
# ALERTAS AUTOMÁTICAS POR SECTOR
# =========================

def alerta_precios(df):
    infl_val, _ = ultimo_valor(df, inflacion_12m)

    if infl_val is None:
        return "Alerta de precios", "No se cuenta con dato suficiente de inflación interanual.", "info"

    if infl_val >= 6:
        return "Alerta de precios", "La inflación interanual se encuentra en zona de presión. Conviene monitorear alimentos, transporte y expectativas.", "danger"

    if infl_val >= 3:
        return "Alerta de precios", "La inflación se mantiene en rango de vigilancia. Se recomienda seguimiento preventivo.", "warning"

    return "Alerta de precios", "La inflación se encuentra en un rango relativamente contenido.", "ok"


def alerta_externo(df):
    rin_val, _ = ultimo_valor(df, rin)

    if rin_val is None:
        return "Alerta externa", "No se cuenta con dato suficiente de reservas internacionales.", "info"

    if rin_val < 2000:
        return "Alerta externa", "Las reservas internacionales se encuentran en zona crítica. Existe riesgo de presión cambiaria y restricción externa.", "danger"

    if rin_val < 5000:
        return "Alerta externa", "Las reservas internacionales están en zona de vigilancia. Se recomienda monitorear divisas, oro y balanza comercial.", "warning"

    return "Alerta externa", "La posición externa muestra un nivel relativamente adecuado de reservas.", "ok"


def alerta_monetario(df):
    bm_yoy = variacion_interanual(df, base_monetaria)

    if bm_yoy is None:
        return "Alerta monetaria", "No se cuenta con información suficiente para calcular la variación interanual de la base monetaria.", "info"

    if bm_yoy >= 15:
        return "Alerta monetaria", "La base monetaria muestra una expansión elevada. Puede generar presión sobre precios, liquidez y expectativas.", "warning"

    if bm_yoy < 0:
        return "Alerta monetaria", "La base monetaria presenta contracción, lo que puede reflejar menor liquidez en la economía.", "warning"

    return "Alerta monetaria", "La dinámica monetaria se mantiene en un rango de seguimiento regular.", "ok"


def alerta_financiero(df):
    cred_yoy = variacion_interanual(df, credito_privado)
    dep_yoy = variacion_interanual(df, depositos)

    if cred_yoy is None and dep_yoy is None:
        return "Alerta financiera", "No se cuenta con información suficiente de crédito y depósitos.", "info"

    if dep_yoy is not None and dep_yoy < 0:
        return "Alerta financiera", "Los depósitos muestran contracción interanual. Puede existir presión de liquidez en el sistema financiero.", "danger"

    if cred_yoy is not None and dep_yoy is not None and cred_yoy > dep_yoy + 5:
        return "Alerta financiera", "El crédito crece por encima de los depósitos. Conviene monitorear liquidez, fondeo y calidad de cartera.", "warning"

    return "Alerta financiera", "El sistema financiero muestra una dinámica relativamente estable entre crédito y depósitos.", "ok"


def alerta_real(df):
    pib_yoy = variacion_interanual(df, pib_pm)
    consumo_yoy = variacion_interanual(df, consumo_hogares)
    inversion_yoy = variacion_interanual(df, formacion_capital)

    if pib_yoy is None:
        return "Alerta sector real", "No se cuenta con información suficiente para calcular el crecimiento interanual del PIB.", "info"

    if pib_yoy < 0:
        return "Alerta sector real", "El PIB muestra contracción interanual. Se recomienda revisar consumo, inversión y sector externo real.", "danger"

    if inversion_yoy is not None and inversion_yoy < 0:
        return "Alerta sector real", "El PIB crece, pero la formación bruta de capital muestra debilidad. Existe riesgo sobre el crecimiento futuro.", "warning"

    if consumo_yoy is not None and consumo_yoy < 0:
        return "Alerta sector real", "El consumo de hogares muestra deterioro. Puede reflejar menor dinamismo de la demanda interna.", "warning"

    return "Alerta sector real", "La actividad real mantiene una trayectoria positiva según los últimos datos disponibles.", "ok"


def alerta_fiscal(df):
    resultado_global, _ = ultimo_valor(df, resultado_global_spnf)
    ingresos_yoy = variacion_interanual(df, ingresos_totales_spnf)
    egresos_yoy = variacion_interanual(df, egresos_totales_spnf)

    if resultado_global is None:
        return "Alerta fiscal", "No se cuenta con dato suficiente del resultado fiscal global del SPNF.", "info"

    if resultado_global < 0 and egresos_yoy is not None and ingresos_yoy is not None and egresos_yoy > ingresos_yoy:
        return "Alerta fiscal", "El resultado fiscal global es deficitario y los egresos crecen por encima de los ingresos. Riesgo fiscal elevado.", "danger"

    if resultado_global < 0:
        return "Alerta fiscal", "El SPNF registra déficit global. Se recomienda monitorear ingresos, gasto corriente y gasto de capital.", "warning"

    return "Alerta fiscal", "El resultado fiscal global se mantiene en terreno positivo o sin señales críticas inmediatas.", "ok"


def alerta_social(df):
    pobreza_val, _ = ultimo_valor(df, pobreza_bolivia)
    desocupacion_val, _ = ultimo_valor(df, desocupacion_nacional)
    gini_val, _ = ultimo_valor(df, gini_bolivia)

    if pobreza_val is None and desocupacion_val is None and gini_val is None:
        return "Alerta social", "No se cuenta con información suficiente de pobreza, desigualdad o desocupación.", "info"

    if pobreza_val is not None and pobreza_val >= 35:
        return "Alerta social", "La incidencia de pobreza se mantiene elevada. Riesgo social alto sobre ingresos, empleo y bienestar.", "danger"

    if desocupacion_val is not None and desocupacion_val >= 8:
        return "Alerta social", "La tasa de desocupación nacional se encuentra en zona de alerta. Conviene monitorear empleo e ingresos laborales.", "warning"

    if gini_val is not None and gini_val >= 0.45:
        return "Alerta social", "El índice de GINI muestra una desigualdad relevante. Se recomienda seguimiento distributivo.", "warning"

    return "Alerta social", "Los indicadores sociales no muestran una señal crítica inmediata según los últimos datos disponibles.", "ok"

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

if not df.empty:
    st.sidebar.metric("Última fecha", df["fecha"].max().strftime("%d/%m/%Y"))
else:
    st.sidebar.metric("Última fecha", "Sin dato")

# =========================
# HEADER
# =========================

col1, col2 = st.columns([1.2, 5])

with col1:
    if os.path.exists("logo_cengob.png"):
        with open("logo_cengob.png", "rb") as img_file:
            logo_base64 = base64.b64encode(img_file.read()).decode()

        st.markdown(f"""
        <div style="
            background:#FFFFFF;
            border-radius:22px;
            box-shadow:0 4px 14px rgba(0,0,0,0.08);
            height:200px;
            width:100%;
            display:flex;
            justify-content:center;
            align-items:center;
            padding:20px;
            margin-top:10px;
        ">
            <img src="data:image/png;base64,{logo_base64}"
                 style="
                    width:170px;
                    max-width:90%;
                    height:auto;
                    object-fit:contain;
                 ">
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div style="
            background:#FFFFFF;
            border-radius:22px;
            box-shadow:0 4px 14px rgba(0,0,0,0.08);
            height:200px;
            width:100%;
            display:flex;
            justify-content:center;
            align-items:center;
            padding:20px;
            margin-top:10px;
            color:#0B3B36;
            font-size:26px;
            font-weight:800;
        ">
            CENGOB
        </div>
        """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <h1 style="
        color:#0B3B36;
        margin-bottom:0;
        font-size: clamp(28px, 5vw, 48px);
        font-weight:800;
    ">
        Dashboard Macroeconómico Ejecutivo
    </h1>

    <h2 style="
        color:#0B3B36;
        margin-top:10px;
        font-size: clamp(18px, 3vw, 28px);
        font-weight:600;
    ">
        Centro de Gobierno - CENGOB
    </h2>

    <p style="
        color:#0B3B36;
        margin-top:14px;
        font-size: clamp(14px, 2vw, 20px);
    ">
        Monitor de coyuntura económica, monetaria, externa, fiscal, real y social
    </p>
    """, unsafe_allow_html=True)

st.markdown("---")

# =========================
# KPIs PRINCIPALES
# =========================

c1, c2, c3, c4 = st.columns(4)

with c1:
    kpi(df, "Actividad económica - IGAE", igae, "")

with c2:
    kpi(df, "Inflación interanual", inflacion_12m, "%")

with c3:
    kpi(df, "RIN", rin, "MM $us")

with c4:
    kpi(df, "Tipo de cambio venta", tc_venta, "Bs/$us")

c5, c6, c7, c8 = st.columns(4)

with c5:
    kpi(df, "Base monetaria", base_monetaria, "MM Bs")

with c6:
    kpi(df, "Crédito privado", credito_privado, "MM Bs")

with c7:
    kpi(df, "Depósitos", depositos, "MM Bs")

with c8:
    kpi(df, "Saldo comercial", saldo_comercial, "MM $us")

st.markdown("---")

titulo_precios, mensaje_precios, nivel_precios = alerta_precios(df)
titulo_externo, mensaje_externo, nivel_externo = alerta_externo(df)

nivel_general = "ok"

if nivel_precios in ["warning", "danger"] or nivel_externo in ["warning", "danger"]:
    nivel_general = "warning"

if nivel_precios == "danger" or nivel_externo == "danger":
    nivel_general = "danger"

alerta_sector(
    "Lectura ejecutiva general",
    f"{mensaje_precios} {mensaje_externo}",
    nivel_general
)

# =========================
# TABS
# =========================

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
    "📈 Resumen",
    "🔥 Inflación",
    "🌎 Sector externo",
    "💵 Monetario",
    "🏦 Financiero",
    "🏭 Sector real",
    "🏛️ Fiscal",
    "👥 Social",
    "⚠️ Riesgos"
])

# =========================
# TAB 1: RESUMEN
# =========================

with tab1:
    st.subheader("📈 Resumen ejecutivo de indicadores clave")

    titulo, mensaje, nivel = alerta_real(df)
    alerta_sector("Lectura ejecutiva general", mensaje, nivel)

    # =========================
    # FILA 1: INDICADORES MACRO PRINCIPALES
    # =========================

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi(df, "Actividad económica - IGAE", igae, "")

    with c2:
        kpi(df, "Inflación interanual", inflacion_12m, "%")

    with c3:
        kpi(df, "RIN", rin, "MM $us")

    with c4:
        kpi(df, "Tipo de cambio venta", tc_venta, "Bs/$us")

    # =========================
    # FILA 2: MONETARIO Y FINANCIERO
    # =========================

    c5, c6, c7, c8 = st.columns(4)

    with c5:
        kpi(df, "Base monetaria", base_monetaria, "MM Bs")

    with c6:
        kpi(df, "Crédito privado", credito_privado, "MM Bs")

    with c7:
        kpi(df, "Depósitos", depositos, "MM Bs")

    with c8:
        kpi(df, "Saldo comercial", saldo_comercial, "MM $us")

    # =========================
    # FILA 3: REAL, FISCAL Y SOCIAL
    # =========================

    c9, c10, c11, c12 = st.columns(4)

    with c9:
        kpi(df, "PIB a precios de mercado", pib_pm, "MM Bs")

    with c10:
        kpi(df, "Resultado Global SPNF", resultado_global_spnf, "MM Bs")

    with c11:
        kpi(df, "Incidencia de pobreza", pobreza_bolivia, "%")

    with c12:
        kpi(df, "Tasa de desocupación", desocupacion_nacional, "%")

# =========================
# TAB 2: INFLACIÓN
# =========================

with tab2:
    st.subheader("🔥 Sector precios - Inflación")

    titulo, mensaje, nivel = alerta_precios(df)
    alerta_sector(titulo, mensaje, nivel)

    c1, c2, c3 = st.columns(3)

    with c1:
        kpi(df, "Inflación interanual", inflacion_12m, "%")

    with c2:
        kpi(df, "Inflación mensual", inflacion_mensual, "%")

    with c3:
        kpi(df, "Inflación acumulada", inflacion_acumulada, "%")

    st.markdown("---")

    grafico_linea(df, inflacion_12m, "Inflación a doce meses", "%")

    st.markdown("### Indicadores complementarios de inflación")

    i1, i2 = st.columns(2)

    with i1:
        grafico_linea(df, inflacion_mensual, "Variación mensual inflación total", "%")

    with i2:
        grafico_linea(df, inflacion_acumulada, "Variación acumulada en el año", "%")
        
# =========================
# TAB 3: SECTOR EXTERNO
# =========================

with tab3:
    st.subheader("🌎 Sector externo")

    titulo, mensaje, nivel = alerta_externo(df)
    alerta_sector(titulo, mensaje, nivel)

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi(df, "RIN", rin, "MM $us")

    with c2:
        kpi(df, "Tipo de cambio referencial", tc_venta, "Bs/$us")

    with c3:
        kpi(df, "Exportaciones", exportaciones, "MM $us")

    with c4:
        kpi(df, "Importaciones", importaciones, "MM $us")

    c5, c6, c7, c8 = st.columns(4)

    with c5:
        kpi(df, "Saldo comercial", saldo_comercial, "MM $us")

    with c6:
        kpi(df, "Divisas", divisas, "MM $us")

    with c7:
        kpi(df, "Oro", oro, "MM $us")

    with c8:
        kpi(df, "Recursos alta liquidez", recursos_alta_liquidez, "MM $us")

    st.markdown("---")

    a, b = st.columns(2)

    with a:
        grafico_linea(df, rin, "Reservas internacionales netas", "MM $us")

    with b:
        grafico_lineas_multiples(
            df,
            [tc_venta, tc_oficial],
            "Tipo de cambio referencial vs oficial",
            "Bs/$us"
        )

    c, d = st.columns(2)

    with c:
        grafico_linea(df, exportaciones, "Exportaciones", "MM $us")

    with d:
        grafico_linea(df, importaciones, "Importaciones", "MM $us")

# =========================
# TAB 4: MONETARIO
# =========================

with tab4:
    st.subheader("💵 Sector monetario")

    titulo, mensaje, nivel = alerta_monetario(df)
    alerta_sector(titulo, mensaje, nivel)

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi(df, "Base monetaria", base_monetaria, "MM Bs")

    with c2:
        kpi(df, "M'1", m1, "MM Bs")

    with c3:
        kpi(df, "M'2", m2, "MM Bs")

    with c4:
        kpi(df, "M'3", m3, "MM Bs")

    st.markdown("---")

    a, b = st.columns(2)

    with a:
        grafico_linea(df, base_monetaria, "Base monetaria", "Millones de Bs")

    with b:
        grafico_lineas_multiples(
            df,
            [m1, m2, m3],
            "Agregados monetarios: M'1, M'2 y M'3",
            "Millones de Bs"
        )

# =========================
# TAB 5: FINANCIERO
# =========================

with tab5:
    st.subheader("🏦 Sector financiero")

    titulo, mensaje, nivel = alerta_financiero(df)
    alerta_sector(titulo, mensaje, nivel)

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi(df, "Crédito privado", credito_privado, "MM Bs")

    with c2:
        kpi(df, "Depósitos", depositos, "MM Bs")

    with c3:
        kpi(df, "Bolivianización depósitos", bol_dep, "%")

    with c4:
        kpi(df, "Bolivianización créditos", bol_cred, "%")

    st.markdown("---")

    a, b = st.columns(2)

    with a:
        grafico_lineas_multiples(
            df,
            [credito_privado, depositos],
            "Crédito y depósitos del sistema financiero",
            "Millones de Bs"
        )

    with b:
        grafico_lineas_multiples(
            df,
            [bol_dep, bol_cred],
            "Bolivianización de depósitos y créditos",
            "%"
        )

# =========================
# TAB 6: SECTOR REAL
# =========================

with tab6:
    st.subheader("🏭 Sector real - PIB por enfoque del gasto")

    titulo, mensaje, nivel = alerta_real(df)
    alerta_sector(titulo, mensaje, nivel)

    c1, c2, c3 = st.columns(3)

    with c1:
        kpi(df, "PIB a precios de mercado", pib_pm, "MM Bs")

    with c2:
        kpi(df, "Consumo de hogares", consumo_hogares, "MM Bs")

    with c3:
        kpi(df, "Formación bruta de capital", formacion_capital, "MM Bs")

    c4, c5, c6 = st.columns(3)

    with c4:
        kpi(df, "Consumo público", consumo_publico, "MM Bs")

    with c5:
        kpi(df, "Exportaciones", expo_bienes_servicios, "MM Bs")

    with c6:
        kpi(df, "Importaciones", impo_bienes_servicios, "MM Bs")

    st.markdown("---")

    a, b = st.columns(2)

    with a:
        grafico_linea(
            df,
            pib_pm,
            "PIB a precios de mercado",
            "Millones de Bs"
        )

    with b:
        grafico_lineas_multiples(
            df,
            [consumo_hogares, consumo_publico, formacion_capital],
            "Demanda interna: consumo e inversión",
            "Millones de Bs"
        )

    c, d = st.columns(2)

    with c:
        grafico_lineas_multiples(
            df,
            [expo_bienes_servicios, impo_bienes_servicios],
            "Sector externo real: exportaciones e importaciones",
            "Millones de Bs"
        )

    with d:
        grafico_lineas_multiples(
            df,
            [
                consumo_hogares,
                consumo_publico,
                formacion_capital,
                expo_bienes_servicios,
                impo_bienes_servicios
            ],
            "Componentes del PIB por gasto",
            "Millones de Bs"
        )

# =========================
# TAB 7: FISCAL
# =========================

with tab7:
    st.subheader("🏛️ Sector fiscal - Sector Público No Financiero")

    titulo, mensaje, nivel = alerta_fiscal(df)
    alerta_sector(titulo, mensaje, nivel)

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi(df, "Ingresos Totales SPNF", ingresos_totales_spnf, "MM Bs")

    with c2:
        kpi(df, "Egresos Totales SPNF", egresos_totales_spnf, "MM Bs")

    with c3:
        kpi(df, "Resultado Corriente SPNF", resultado_corriente_spnf, "MM Bs")

    with c4:
        kpi(df, "Resultado Global SPNF", resultado_global_spnf, "MM Bs")

    st.markdown("---")

    a, b = st.columns(2)

    with a:
        grafico_lineas_multiples(
            df,
            [ingresos_totales_spnf, egresos_totales_spnf],
            "Ingresos y egresos totales del SPNF",
            "Millones de Bs"
        )

    with b:
        grafico_lineas_multiples(
            df,
            [resultado_corriente_spnf, resultado_global_spnf],
            "Resultado fiscal corriente y global del SPNF",
            "Millones de Bs"
        )

    c, d = st.columns(2)

    with c:
        grafico_lineas_multiples(
            df,
            [
                ingresos_corrientes_spnf,
                ingresos_capital_spnf
            ],
            "Composición de ingresos del SPNF",
            "Millones de Bs"
        )

    with d:
        grafico_lineas_multiples(
            df,
            [
                egresos_corrientes_spnf,
                egresos_capital_spnf
            ],
            "Composición de egresos del SPNF",
            "Millones de Bs"
        )

    e, f = st.columns(2)

    with e:
        grafico_linea(
            df,
            resultado_corriente_spnf,
            "Resultado Fiscal Corriente del SPNF",
            "Millones de Bs"
        )

    with f:
        grafico_linea(
            df,
            resultado_global_spnf,
            "Resultado Fiscal Global del SPNF",
            "Millones de Bs"
        )

# =========================
# TAB 8: SOCIAL
# =========================

with tab8:
    st.subheader("👥 Sector social")

    titulo, mensaje, nivel = alerta_social(df)
    alerta_sector(titulo, mensaje, nivel)

    c1, c2, c3 = st.columns(3)

    with c1:
        kpi(df, "Incidencia de pobreza", pobreza_bolivia, "%")

    with c2:
        kpi(df, "Índice de GINI", gini_bolivia, "")

    with c3:
        kpi(df, "Tasa de desocupación nacional", desocupacion_nacional, "%")

    st.markdown("---")

    a, b = st.columns(2)

    with a:
        grafico_linea(
            df,
            pobreza_bolivia,
            "Bolivia: Incidencia de pobreza",
            "%"
        )

    with b:
        grafico_linea(
            df,
            gini_bolivia,
            "Bolivia: Índice de GINI",
            "Índice"
        )

    c, d = st.columns(2)

    with c:
        grafico_linea(
            df,
            desocupacion_nacional,
            "Tasa de Desocupación Nacional",
            "%"
        )

    with d:
        grafico_lineas_multiples(
            df,
            [
                pobreza_bolivia,
                desocupacion_nacional
            ],
            "Pobreza y desocupación nacional",
            "%"
        )

# =========================
# TAB 9: RIESGOS
# =========================

with tab9:
    st.subheader("🚦 Semáforo macroeconómico")

    infl_val, _ = ultimo_valor(df, inflacion_12m)
    rin_val, _ = ultimo_valor(df, rin)

    tc_ref_val, _ = ultimo_valor(df, tc_venta)
    tc_of_val, _ = ultimo_valor(df, tc_oficial)

    if tc_ref_val is not None and tc_of_val is not None and tc_of_val != 0:
        brecha_tc = abs(tc_ref_val - tc_of_val)
    else:
        brecha_tc = None

    cred_yoy = variacion_interanual(df, credito_privado)
    pib_yoy = variacion_interanual(df, pib_pm)
    resultado_global, _ = ultimo_valor(df, resultado_global_spnf)
    pobreza_val, _ = ultimo_valor(df, pobreza_bolivia)

    riesgo_inflacion = clasificar_normal(infl_val, 3, 6)
    riesgo_rin = clasificar_invertido(rin_val, 2000, 5000)
    riesgo_tc = clasificar_normal(brecha_tc, 0.20, 1.00)
    riesgo_credito = clasificar_normal(cred_yoy, 5, 15)

    if pib_yoy is None:
        riesgo_real = "Sin dato"
    elif pib_yoy < 0:
        riesgo_real = "Alto"
    elif pib_yoy < 2:
        riesgo_real = "Moderado"
    else:
        riesgo_real = "Bajo"

    if resultado_global is None:
        riesgo_fiscal = "Sin dato"
    elif resultado_global < 0:
        riesgo_fiscal = "Alto"
    else:
        riesgo_fiscal = "Bajo"

    if pobreza_val is None:
        riesgo_social = "Sin dato"
    elif pobreza_val >= 35:
        riesgo_social = "Alto"
    elif pobreza_val >= 25:
        riesgo_social = "Moderado"
    else:
        riesgo_social = "Bajo"

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        tarjeta_riesgo("Riesgo inflacionario", riesgo_inflacion)

    with c2:
        tarjeta_riesgo("Posición externa - RIN", riesgo_rin)

    with c3:
        tarjeta_riesgo("Presión cambiaria", riesgo_tc)

    with c4:
        tarjeta_riesgo("Expansión crediticia", riesgo_credito)

    st.markdown("<br>", unsafe_allow_html=True)

    c5, c6, c7 = st.columns(3)

    with c5:
        tarjeta_riesgo("Actividad real", riesgo_real)

    with c6:
        tarjeta_riesgo("Resultado fiscal", riesgo_fiscal)

    with c7:
        tarjeta_riesgo("Riesgo social", riesgo_social)

    st.info(
        "Los umbrales del semáforo son referenciales y pueden ajustarse según criterio técnico."
    )

st.markdown("---")

# =========================
# EXPLORADOR
# =========================

st.subheader("🔎 Explorador de variables")

variables_excluir = [
    "Bolivianización (%)_1",
    "Bolivianización (%)_2",
    "Bolivianización (%)_3",
    "Bolivianización (%)_4",
    "A la vista",
    "Caja de ahorro",
    "Plazo",
    "Otros"
]

variables = [
    c for c in df.columns
    if c != "fecha" and c not in variables_excluir
]

if variables:
    seleccion = st.selectbox("Selecciona cualquier variable del Excel", variables)
    grafico_linea(df, seleccion, seleccion)
else:
    st.warning("No existen variables disponibles para explorar.")

st.download_button(
    "⬇️ Descargar base filtrada",
    data=df.to_csv(index=False).encode("utf-8"),
    file_name="base_macro_filtrada.csv",
    mime="text/csv"
)
