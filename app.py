import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import uuid

st.set_page_config(
    page_title="Dashboard Macroeconómico CENGOB - Bolivia",
    layout="wide",
    page_icon="📊"
)

EXCEL_FILE = "Info.xlsx"
SHEET_NAME = "data"


# =========================
#    TEMA AUTOMÁTICO
# ========================= 

st.markdown("""
<style>

.stApp{
    background-color:#F3F4EF;
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
    background:white;
    border-left:6px solid #C9A227;
    border-radius:18px;
    padding:18px;
    box-shadow:0 4px 14px rgba(0,0,0,0.08);
}

/* GRAFICOS */
.js-plotly-plot{
    border-radius:18px;
    overflow:hidden;
}

/* ALERTAS */
.stAlert{
    border-left:6px solid #C9A227;
    border-radius:14px;
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
            template="plotly_dark",
        
            paper_bgcolor="#111827",
            plot_bgcolor="#111827",
        
            font=dict(
                color="#E5E7EB",
                size=14
            ),
        
            title_font=dict(
                color="#F9FAFB",
                size=22
            ),
        
            margin=dict(l=20, r=20, t=60, b=30),
        
            xaxis_title="",
            yaxis_title=unidad,
        
            xaxis=dict(
                rangeselector=dict(
                    bgcolor="#1F2937",
                    activecolor="#0B3B36",
        
                    font=dict(
                        color="#F9FAFB",
                        size=13
                    ),
        
                    buttons=list([
                        dict(count=1, label="1A", step="year", stepmode="backward"),
                        dict(count=5, label="5A", step="year", stepmode="backward"),
                        dict(count=10, label="10A", step="year", stepmode="backward"),
                        dict(step="all")
                    ])
                ),
        
                rangeslider=dict(
                    visible=True,
                    bgcolor="#111827"
                ),
        
                type="date",
        
                tickfont=dict(
                    color="#D1D5DB"
                )
            ),
        
            yaxis=dict(
                gridcolor="#374151",
                tickfont=dict(
                    color="#D1D5DB"
                )
            )
        )
    fig.update_traces(line=dict(width=3.5))
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#D1D5DB")

    st.plotly_chart(
    fig,
    use_container_width=True,
    key=f"linea_{titulo}_{uuid.uuid4()}"
    )



def grafico_lineas_multiples(df, cols, titulo, unidad=""):

    cols = [c for c in cols if c is not None]

    if not cols:
        st.warning(f"No hay variables disponibles para: {titulo}")
        return

    fig = go.Figure()

    # =====================
    # COLORES CENGOB
    # =====================

    colores = [
        "#0B3B36",  # verde petróleo
        "#C9A227",  # dorado
        "#556B2F",  # oliva
        "#C2410C",  # naranja
        "#475569"   # gris
    ]

    # =====================
    # NOMBRES CORTOS
    # =====================

    nombres = {
        credito_privado: "Crédito privado",
        depositos: "Depósitos",
        m1: "M1",
        m2: "M2",
        m3: "M3",
        tc_venta: "TC Referencial",
        tc_oficial: "TC Oficial",
        bol_dep: "Boliv. depósitos",
        bol_cred: "Boliv. créditos"
    }

    # =====================
    # GRAFICOS
    # =====================

    for i, c in enumerate(cols):

        s = df[["fecha", c]].dropna()

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

    # =====================
    # LAYOUT
    # =====================

    fig.update_layout(
        title=titulo,
        height=430,
        template="plotly_white",
        paper_bgcolor="#111827",
        plot_bgcolor="#111827",
        font=dict(color="#E5E7EB"),
        margin=dict(l=20, r=20, t=60, b=30),
        xaxis_title="",
        yaxis_title=unidad,
        hovermode="x unified",

        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),

        xaxis=dict(
            rangeselector=dict(
                bgcolor="#1F2937",
                activecolor="#0B3B36",

                font=dict(
                    color="#FFFFFF",
                    size=13
                ),

                buttons=list([
                    dict(count=1, label="1A", step="year", stepmode="backward"),
                    dict(count=5, label="5A", step="year", stepmode="backward"),
                    dict(count=10, label="10A", step="year", stepmode="backward"),
                    dict(step="all")
                ])
            ),

            rangeslider=dict(visible=True),
            type="date"
        )
    )

    fig.update_xaxes(
        showgrid=False,
        tickfont=dict(color="#D1D5DB")
    )
    
    fig.update_yaxes(
        gridcolor="#374151",
        tickfont=dict(color="#D1D5DB")
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        key=f"multi_{titulo}_{uuid.uuid4()}"
    )


def grafico_barras(df, cols, titulo):
    cols = [c for c in cols if c is not None]

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
        paper_bgcolor="#111827",
        plot_bgcolor="#111827",
        font=dict(color="#E5E7EB"),
        margin=dict(l=20, r=20, t=60, b=80),
        xaxis_title="",
        yaxis_title="Valor"
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        key=f"barras_{titulo}_{uuid.uuid4()}"
    )

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
inflacion_mensual = buscar_columna("Variación mensual inflacion total")
inflacion_acumulada = buscar_columna("Variación acumulada en el año")
rin = buscar_columna("Reservas internacionales netas")
tc_venta = buscar_columna("Valor referencial de venta")
tc_oficial = buscar_columna("Tipo de cambio oficial")
if tc_oficial is None:
    tc_oficial = buscar_columna("Tipo de cambio de venta")
bol_dep = buscar_columna("Bolivianización Depósitos")
bol_cred = buscar_columna("Bolivianización Créditos")
base_monetaria = buscar_columna("Base monetaria")
credito_privado = buscar_columna("Crédito del sistema financiero al sector privado")
depositos = buscar_columna("Depósitos en entidades")
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

col1, col2 = st.columns([1, 5])

with col1:
    st.image("logo_cengob.png", width=160)

with col2:
    st.markdown("""
    <h1 style="
        color:#F8FAFC;
        margin-bottom:0;
        font-size: clamp(28px, 5vw, 48px);
        font-weight:800;
    ">
        Dashboard Macroeconómico Ejecutivo
    </h1>

    <h2 style="
        color:#E5E7EB;
        margin-top:10px;
        font-size: clamp(18px, 3vw, 28px);
        font-weight:600;
    ">
        Centro de Gobierno - CENGOB
    </h2>

    <p style="
        color:#CBD5E1;
        margin-top:14px;
        font-size: clamp(14px, 2vw, 20px);
    ">
        Monitor de coyuntura económica, monetaria, externa y financiera
    </p>
    """, unsafe_allow_html=True)

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

st.info(
    "Lectura ejecutiva: la inflación se mantiene en zona de alerta, "
    "mientras las reservas internacionales continúan siendo el principal factor de riesgo externo."
)

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
        grafico_lineas_multiples(
            df,
            [credito_privado, depositos],
            "Crédito y depósitos del sistema financiero",
            "Millones de Bs"
        )

with tab2:
    grafico_linea(df, inflacion_12m, "Inflación a doce meses", "%")
    st.markdown("### Indicadores complementarios de inflación")

    i1, i2 = st.columns(2)

    with i1:
        grafico_linea(df, inflacion_mensual, "Variación mensual inflación total", "%")

    with i2:
        grafico_linea(df, inflacion_acumulada, "Variación acumulada en el año", "%")


with tab3:
    a, b = st.columns(2)

    with a:
        grafico_linea(df, rin, "Reservas internacionales netas")

    with b:
        grafico_lineas_multiples(
            df,
            [tc_venta, tc_oficial],
            "Tipo de cambio referencial vs oficial",
            "Bs/$us"
        )

    c, d = st.columns(2)

    with c:
        grafico_linea(df, exportaciones, "Exportaciones")

    with d:
        grafico_linea(df, importaciones, "Importaciones")


with tab4:
    a, b = st.columns(2)

    with a:
        grafico_linea(df, base_monetaria, "Base monetaria")

    with b:
        grafico_lineas_multiples(
            df,
            [m1, m2, m3],
            "Agregados monetarios: M'1, M'2 y M'3",
            "Millones de Bs"
        )


with tab5:
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
# FUNCION TARJETA RIESGO
# =========================

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
        padding:28px;
        border-radius:20px;
        border:2px solid {line};
        min-height:180px;
        box-shadow:0 8px 22px rgba(0,0,0,0.18);
    ">
        <div style="
            color:white;
            font-size:24px;
            font-weight:800;
            margin-bottom:45px;
        ">
            {titulo}
        </div>

        <div style="
            color:white;
            font-size:44px;
            font-weight:900;
        ">
            {nivel}
        </div>
    </div>
    """

    st.html(html)
    
# =========================
# RIESGOS
# =========================

with tab6:

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

    

    # =====================
    # FUNCIONES DE CLASIFICACION
    # =====================

    def clasificar_normal(valor, bajo, medio):

        if valor is None:
            return "Sin dato"

        if valor < bajo:
            return "Bajo"

        elif valor < medio:
            return "Moderado"

        else:
            return "Alto"

    def clasificar_invertido(valor, bajo, medio):

        if valor is None:
            return "Sin dato"

        if valor > medio:
            return "Bajo"

        elif valor > bajo:
            return "Moderado"

        else:
            return "Alto"

    # =====================
    # SEMAFORO
    # =====================

    riesgo_inflacion = clasificar_normal(
        infl_val,
        3,
        6
    )

    riesgo_rin = clasificar_invertido(
        rin_val,
        2000,
        5000
    )

    riesgo_tc = clasificar_normal(
        brecha_tc,
        0.20,
        1.00
    )
    
    riesgo_credito = clasificar_normal(
        cred_yoy,
        5,
        15
    )

    # =====================
    # TARJETAS
    # =====================

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        tarjeta_riesgo(
            "Riesgo inflacionario",
            riesgo_inflacion
        )

    with c2:
        tarjeta_riesgo(
            "Posición externa - RIN",
            riesgo_rin
        )

    with c3:
        tarjeta_riesgo(
            "Presión cambiaria",
            riesgo_tc
        )

    with c4:
        tarjeta_riesgo(
            "Expansión crediticia",
            riesgo_credito
        )

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
seleccion = st.selectbox("Selecciona cualquier variable del Excel", variables)
grafico_linea(df, seleccion, seleccion)

st.download_button(
    "⬇️ Descargar base filtrada",
    data=df.to_csv(index=False).encode("utf-8"),
    file_name="base_macro_filtrada.csv",
    mime="text/csv"
)
