import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import uuid
import base64
import os
import unicodedata
import io
import json
import requests
import time

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload


# =========================
# CONFIGURACIÓN GENERAL
# =========================

st.set_page_config(
    page_title="Sistema de Información Económica de Bolivia - CENGOB",
    layout="wide",
    page_icon="📊"
)

SHEET_NAME = "data"
DRIVE_CACHE_TTL = 60

# El ID se guarda en .streamlit/secrets.toml o en los Secrets de Streamlit Cloud.
# La aplicación admite dos modalidades:
# 1) Privada y recomendada: cuenta de servicio de Google.
# 2) Pública: archivo compartido como "Cualquier persona con el enlace".

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
# CARGA DE DATOS DESDE GOOGLE DRIVE
# =========================

GOOGLE_SHEETS_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.readonly"


def _leer_configuracion_drive():
    """Obtiene el ID y, si existen, las credenciales privadas."""
    try:
        drive_cfg = dict(st.secrets["drive"])
    except (FileNotFoundError, KeyError):
        drive_cfg = {}

    file_id = str(
        drive_cfg.get("file_id")
        or os.getenv("GOOGLE_DRIVE_FILE_ID", "")
    ).strip()

    if not file_id:
        raise RuntimeError(
            "No se configuró el ID del archivo de Google Drive. "
            "Agrega [drive] file_id = \"...\" en los Secrets de Streamlit."
        )

    credenciales = None

    # Opción A: credenciales guardadas como una sección TOML.
    try:
        credenciales = dict(st.secrets["gcp_service_account"])
    except (FileNotFoundError, KeyError):
        pass

    # Opción B: JSON completo guardado dentro de [drive].
    if not credenciales and drive_cfg.get("service_account_json"):
        try:
            credenciales = json.loads(drive_cfg["service_account_json"])
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                "El campo drive.service_account_json no contiene un JSON válido."
            ) from exc

    return file_id, credenciales


def _parece_archivo_excel(contenido):
    """Reconoce archivos XLSX/XLS por su firma binaria."""
    return (
        contenido.startswith(b"PK")
        or contenido.startswith(bytes.fromhex("D0CF11E0"))
    )


def _descargar_drive_privado(file_id, credenciales_info):
    """Descarga un Excel privado o exporta una hoja de Google como XLSX."""
    credenciales = service_account.Credentials.from_service_account_info(
        credenciales_info,
        scopes=[DRIVE_SCOPE],
    )

    servicio = build(
        "drive",
        "v3",
        credentials=credenciales,
        cache_discovery=False,
    )

    metadatos = (
        servicio.files()
        .get(
            fileId=file_id,
            fields="id,name,mimeType",
            supportsAllDrives=True,
        )
        .execute()
    )

    if metadatos.get("mimeType") == GOOGLE_SHEETS_MIME:
        solicitud = servicio.files().export_media(
            fileId=file_id,
            mimeType=XLSX_MIME,
        )
    else:
        solicitud = servicio.files().get_media(
            fileId=file_id,
            supportsAllDrives=True,
        )

    buffer = io.BytesIO()
    descargador = MediaIoBaseDownload(buffer, solicitud)
    terminado = False

    while not terminado:
        _, terminado = descargador.next_chunk()

    contenido = buffer.getvalue()

    if not _parece_archivo_excel(contenido):
        raise RuntimeError(
            "Google Drive respondió, pero el contenido descargado no parece "
            "ser un archivo Excel válido."
        )

    return contenido, metadatos.get("name", "archivo de Drive")


def _descargar_drive_publico(file_id):
    """Descarga un archivo público de Drive o exporta un Google Sheet público."""
    # Evita que Google/ navegador entregue una versión anterior del XLSX.
    cache_buster = int(time.time())
    urls = [
        f"https://drive.google.com/uc?export=download&id={file_id}&_={cache_buster}",
        f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx&_={cache_buster}",
    ]

    ultimo_error = None

    for url in urls:
        try:
            respuesta = requests.get(
                url,
                timeout=90,
                allow_redirects=True,
                headers={
                    "User-Agent": "Mozilla/5.0",
                    "Cache-Control": "no-cache",
                    "Pragma": "no-cache",
                },
            )
            respuesta.raise_for_status()

            if _parece_archivo_excel(respuesta.content):
                return respuesta.content, "archivo público de Google Drive"

            ultimo_error = RuntimeError(
                "La descarga devolvió una página HTML en lugar del Excel."
            )
        except requests.RequestException as exc:
            ultimo_error = exc

    raise RuntimeError(
        "No se pudo descargar el archivo públicamente. Compártelo como "
        "'Cualquier persona con el enlace' o configura una cuenta de servicio."
    ) from ultimo_error


@st.cache_data(ttl=DRIVE_CACHE_TTL, show_spinner=False)
def descargar_excel_drive():
    """Devuelve los bytes del Excel y el nombre de la fuente."""
    file_id, credenciales_info = _leer_configuracion_drive()

    if credenciales_info:
        try:
            return _descargar_drive_privado(file_id, credenciales_info)
        except HttpError as exc:
            raise RuntimeError(
                "Google Drive rechazó el acceso. Verifica que el archivo esté "
                "compartido con el correo de la cuenta de servicio."
            ) from exc

    return _descargar_drive_publico(file_id)


@st.cache_data(ttl=DRIVE_CACHE_TTL, show_spinner="Actualizando datos desde Google Drive...")
def cargar_datos():
    contenido_excel, nombre_fuente = descargar_excel_drive()

    raw = pd.read_excel(
        io.BytesIO(contenido_excel),
        sheet_name=SHEET_NAME,
        header=None,
    )

    if raw.shape[0] < 3:
        raise ValueError(
            "La hoja no tiene la estructura esperada: se requieren al menos "
            "dos filas de encabezado y una fila de datos."
        )

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

    # =========================
    # CONVERSIÓN DE FECHA
    # =========================
    data["fecha"] = pd.to_datetime(data["fecha"], errors="coerce")
    data = data.dropna(subset=["fecha"])

    # Quita horas ocultas si existieran
    data["fecha"] = data["fecha"].dt.normalize()

    # =========================
    # AJUSTE DE FECHAS
    # =========================
    # Regla:
    # 1. Meses cerrados: último día del mes.
    # 2. Último mes disponible incompleto: mantiene fecha real de corte.

    if data.empty:
        raise ValueError(
            "No se encontraron fechas válidas en la primera columna de la hoja."
        )

    fecha_max = data["fecha"].max()
    ultimo_mes_disponible = fecha_max.to_period("M")
    fecha_max_fin_mes = fecha_max + pd.offsets.MonthEnd(0)

    ultimo_mes_esta_cerrado = fecha_max == fecha_max_fin_mes

    def ajustar_fecha_mensual(fecha):
        if pd.isna(fecha):
            return fecha

        if (
            not ultimo_mes_esta_cerrado
            and fecha.to_period("M") == ultimo_mes_disponible
        ):
            return fecha

        return fecha + pd.offsets.MonthEnd(0)

    data["fecha"] = data["fecha"].apply(ajustar_fecha_mensual)

    # =========================
    # LIMPIEZA DE COLUMNAS
    # =========================
    data = data.dropna(axis=1, how="all")

    for col in data.columns:
        if col != "fecha":
            data[col] = pd.to_numeric(data[col], errors="coerce")

    data.attrs["fuente_drive"] = nombre_fuente
    return data

# =========================
# CARGAR BASE ORIGINAL
# =========================

try:
    df_original = cargar_datos()
    st.sidebar.caption(
        f"☁️ Fuente: Google Drive · actualización automática cada "
        f"{DRIVE_CACHE_TTL} segundos"
    )
except Exception as e:
    st.error(
        "No se pudo cargar la base desde Google Drive. Verifica el ID del "
        f"archivo, los permisos de acceso y que la hoja se llame '{SHEET_NAME}'."
    )
    st.info(
        "Para acceso privado, comparte el archivo con el correo de la cuenta "
        "de servicio. Para acceso público, habilita 'Cualquier persona con el enlace'."
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

    # Unifica espacios, tabulaciones y saltos de línea de los encabezados de Excel.
    # Esto permite reconocer títulos visualmente partidos en varias líneas.
    texto = " ".join(texto.split())

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

    texto = f"{x:,.2f}"

    # Formato español/boliviano:
    # 36,114.29 -> 36.114,29
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")

    return texto


def formato_numero_1d(x):
    if x is None or pd.isna(x):
        return "Sin dato"

    texto = f"{x:,.1f}"

    # Formato español/boliviano:
    # 12.5 -> 12,5
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")

    return texto


def kpi(df, titulo, col, unidad="", tipo="ultimo", delta_tipo="interanual"):
    """
    tipo:
    - "ultimo": muestra último dato disponible.
    - "acumulado": muestra acumulado anual al último dato disponible.

    delta_tipo:
    - "interanual": variación porcentual interanual.
    - "pp": variación interanual en puntos porcentuales.
    - "acumulado": variación porcentual del acumulado anual.
    - "ninguno": no muestra delta.
    """

    if tipo == "acumulado":
        valor, fecha = valor_acumulado_anual(df, col)
    else:
        valor, fecha = ultimo_valor(df, col)

    if valor is None:
        st.metric(titulo, "Sin dato")
        return

    delta = None

    if delta_tipo == "interanual":
        yoy = variacion_interanual(df, col)
        delta = f"{formato_numero_1d(yoy)}% interanual" if yoy is not None else None

    elif delta_tipo == "pp":
        pp = variacion_interanual_pp(df, col)
        delta = f"{formato_numero_1d(pp)} p.p. interanual" if pp is not None else None

    elif delta_tipo == "acumulado_pct":
        yoy_acum = variacion_acumulada_interanual(df, col)
        delta = f"{formato_numero_1d(yoy_acum)}% acum. interanual" if yoy_acum is not None else None

    elif delta_tipo == "ninguno":
        delta = None

    st.metric(titulo, f"{formato_numero(valor)} {unidad}", delta)

    if tipo == "acumulado":
        texto_fecha = f"Acumulado a: {fecha.strftime('%d/%m/%Y')}"
    else:
        texto_fecha = f"Último dato: {fecha.strftime('%d/%m/%Y')}"

    st.markdown(
        f"""
        <p style="
            color:#0B3B36;
            font-size:15px;
            margin-top:6px;
            font-weight:500;
        ">
            {texto_fecha}
        </p>
        """,
        unsafe_allow_html=True
    )


def valor_acumulado_anual(df, col):
    if col is None:
        return None, None

    if col not in df.columns:
        return None, None

    s = df[["fecha", col]].dropna().sort_values("fecha")

    if s.empty:
        return None, None

    ultima_fecha = s["fecha"].max()
    gestion = ultima_fecha.year

    s_gestion = s[
        (s["fecha"].dt.year == gestion) &
        (s["fecha"] <= ultima_fecha)
    ]

    if s_gestion.empty:
        return None, None

    return s_gestion[col].sum(), ultima_fecha


def variacion_acumulada_interanual(df, col):
    if col is None:
        return None

    if col not in df.columns:
        return None

    s = df[["fecha", col]].dropna().sort_values("fecha")

    if s.empty:
        return None

    ultima_fecha = s["fecha"].max()
    gestion_actual = ultima_fecha.year
    gestion_anterior = gestion_actual - 1
    mes_corte = ultima_fecha.month

    actual = s[
        (s["fecha"].dt.year == gestion_actual) &
        (s["fecha"].dt.month <= mes_corte)
    ][col].sum()

    anterior = s[
        (s["fecha"].dt.year == gestion_anterior) &
        (s["fecha"].dt.month <= mes_corte)
    ][col].sum()

    if anterior == 0 or pd.isna(anterior):
        return None

    return ((actual / anterior) - 1) * 100


def variacion_interanual_pp(df, col):
    if col is None:
        return None

    if col not in df.columns:
        return None

    s = df[["fecha", col]].dropna().sort_values("fecha")

    if s.empty:
        return None

    actual = s.iloc[-1]
    base_fecha = actual["fecha"] - pd.DateOffset(years=1)

    ant = s[s["fecha"] <= base_fecha]

    if ant.empty:
        return None

    base = ant.iloc[-1][col]

    if pd.isna(base):
        return None

    return actual[col] - base






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

    # =====================
    # LÍNEA PRINCIPAL CON SOMBRA
    # =====================
    fig.add_trace(
        go.Scatter(
            x=s["fecha"],
            y=s[col],
            mode="lines",
            line=dict(
                width=3.5,
                color="#0B3B36"
            ),
            marker=dict(
                size=7,
                color="#0B3B36",
                line=dict(
                    width=1.5,
                    color="#FFFFFF"
                )
            ),
            fill="tozeroy",
            fillcolor="rgba(11,59,54,0.16)",
            hovertemplate="%{x|%d/%m/%Y}<br>Valor: %{y:,.2f}<extra></extra>"
        )
    )

    # =====================
    # DISEÑO DEL GRÁFICO
    # =====================
    fig.update_layout(
        title=titulo,
        height=460,
        template="plotly_white",

        paper_bgcolor="#EEF2F5",
        plot_bgcolor="#EEF2F5",

        font=dict(
            color="#000000",
            size=14
        ),

        title_font=dict(
            color="#0B3B36",
            size=22
        ),

        margin=dict(
            l=25,
            r=25,
            t=65,
            b=25
        ),

        xaxis_title="",
        yaxis_title=unidad,

        hovermode="x unified",

        xaxis=dict(
            type="date",

            rangeselector=dict(
                bgcolor="#FFFFFF",
                activecolor="#C9A227",
                bordercolor="#0B3B36",
                borderwidth=1,
                font=dict(
                    color="#0B3B36",
                    size=12
                ),
                buttons=list([
                    dict(count=1, label="1A", step="year", stepmode="backward"),
                    dict(count=3, label="3A", step="year", stepmode="backward"),
                    dict(count=5, label="5A", step="year", stepmode="backward"),
                    dict(count=10, label="10A", step="year", stepmode="backward"),
                    dict(step="all", label="Todo")
                ])
            ),

            rangeslider=dict(
                visible=True,
                bgcolor="#E6D7A2",
                bordercolor="#0B3B36",
                borderwidth=2,
                thickness=0.12
            ),

            tickfont=dict(
                color="#000000",
                size=12
            ),

            gridcolor="rgba(0,0,0,0.10)",
            showgrid=True
        ),

        yaxis=dict(
            tickfont=dict(
                color="#000000",
                size=12
            ),
            gridcolor="rgba(0,0,0,0.12)",
            zeroline=True,
            zerolinecolor="rgba(11,59,54,0.45)",
            zerolinewidth=1.5
        )
    )

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
        "#0B3B36",  # verde petróleo CENGOB
        "#C9A227",  # dorado CENGOB
        "#556B2F",  # oliva
        "#C2410C",  # naranja institucional
        "#475569",  # gris
        "#2563EB",  # azul
        "#7C3AED"   # violeta
    ]

    rellenos = [
        "rgba(11,59,54,0.13)",
        "rgba(201,162,39,0.13)",
        "rgba(85,107,47,0.12)",
        "rgba(194,65,12,0.11)",
        "rgba(71,85,105,0.11)",
        "rgba(37,99,235,0.10)",
        "rgba(124,58,237,0.10)"
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
        encaje_constituido: "Encaje constituido",
        excedente_encaje_efectivo: "Liquidez EIF total",
        excedente_encaje_me: "Liquidez EIF en ME",
        consumo_hogares: "Consumo hogares",
        consumo_publico: "Consumo público",
        formacion_capital: "FBKF",
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
                        width=3.2,
                        color=colores[i % len(colores)]
                    ),
                    marker=dict(
                        size=6,
                        color=colores[i % len(colores)],
                        line=dict(
                            width=1,
                            color="#FFFFFF"
                        )
                    ),
                    fill="tozeroy" if i == 0 else None,
                    fillcolor=rellenos[i % len(rellenos)],
                    hovertemplate="%{x|%d/%m/%Y}<br>%{y:,.2f}<extra></extra>"
                )
            )

    fig.update_layout(
        title=titulo,
        height=460,
        template="plotly_white",

        paper_bgcolor="#EEF2F5",
        plot_bgcolor="#EEF2F5",

        font=dict(
            color="#000000",
            size=14
        ),

        title_font=dict(
            color="#0B3B36",
            size=20
        ),

        margin=dict(
            l=25,
            r=25,
            t=70,
            b=25
        ),

        xaxis_title="",
        yaxis_title=unidad,

        hovermode="x unified",

        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(
                color="#000000",
                size=13
            )
        ),

        xaxis=dict(
            type="date",

            rangeselector=dict(
                bgcolor="#FFFFFF",
                activecolor="#C9A227",
                bordercolor="#0B3B36",
                borderwidth=1,
                font=dict(
                    color="#0B3B36",
                    size=12
                ),
                buttons=list([
                    dict(count=1, label="1A", step="year", stepmode="backward"),
                    dict(count=3, label="3A", step="year", stepmode="backward"),
                    dict(count=5, label="5A", step="year", stepmode="backward"),
                    dict(count=10, label="10A", step="year", stepmode="backward"),
                    dict(step="all", label="Todo")
                ])
            ),

            rangeslider=dict(
                visible=True,
                bgcolor="#E6D7A2",
                bordercolor="#0B3B36",
                borderwidth=2,
                thickness=0.12
            ),

            tickfont=dict(
                color="#000000",
                size=12
            ),

            gridcolor="rgba(0,0,0,0.10)",
            showgrid=True
        ),

        yaxis=dict(
            tickfont=dict(
                color="#000000",
                size=12
            ),
            gridcolor="rgba(0,0,0,0.12)",
            zeroline=True,
            zerolinecolor="rgba(11,59,54,0.45)",
            zerolinewidth=1.5
        )
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        key=f"multi_{titulo}_{uuid.uuid4()}"
    )

def grafico_doble_eje(
    df,
    col_izq,
    col_der,
    titulo,
    nombre_izq,
    nombre_der,
    titulo_eje_izq="Eje izquierdo",
    titulo_eje_der="Eje derecho",
    unidad_izq="",
    unidad_der="",
    color_izq="#0B3B36",
    color_der="#C9A227",
    sombra_izq=True,
    sombra_der=False
):
    """
    Gráfico general de doble eje.
    
    col_izq: columna que irá en el eje izquierdo.
    col_der: columna que irá en el eje derecho.
    nombre_izq: nombre visible de la serie izquierda.
    nombre_der: nombre visible de la serie derecha.
    titulo_eje_izq: título del eje Y izquierdo.
    titulo_eje_der: título del eje Y derecho.
    unidad_izq: unidad para hover del eje izquierdo.
    unidad_der: unidad para hover del eje derecho.
    """

    if col_izq is None and col_der is None:
        st.warning(f"No se encontraron variables para: {titulo}")
        return

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # =====================
    # SERIE EJE IZQUIERDO
    # =====================
    if col_izq is not None and col_izq in df.columns:

        s_izq = df[["fecha", col_izq]].dropna().sort_values("fecha")

        if not s_izq.empty:

            fig.add_trace(
                go.Scatter(
                    x=s_izq["fecha"],
                    y=s_izq[col_izq],
                    mode="lines",
                    name=nombre_izq,
                    line=dict(
                        width=3.5,
                        color=color_izq
                    ),
                    marker=dict(
                        size=6,
                        color=color_izq,
                        line=dict(
                            width=1,
                            color="#FFFFFF"
                        )
                    ),
                    fill="tozeroy" if sombra_izq else None,
                    fillcolor="rgba(11,59,54,0.13)" if sombra_izq else None,
                    hovertemplate=(
                        "%{x|%d/%m/%Y}<br>"
                        + nombre_izq
                        + ": %{y:,.2f} "
                        + unidad_izq
                        + "<extra></extra>"
                    )
                ),
                secondary_y=False
            )

    # =====================
    # SERIE EJE DERECHO
    # =====================
    if col_der is not None and col_der in df.columns:

        s_der = df[["fecha", col_der]].dropna().sort_values("fecha")

        if not s_der.empty:

            fig.add_trace(
                go.Scatter(
                    x=s_der["fecha"],
                    y=s_der[col_der],
                    mode="lines",
                    name=nombre_der,
                    line=dict(
                        width=3.5,
                        color=color_der
                    ),
                    marker=dict(
                        size=6,
                        color=color_der,
                        line=dict(
                            width=1,
                            color="#FFFFFF"
                        )
                    ),
                    fill="tozeroy" if sombra_der else None,
                    fillcolor="rgba(201,162,39,0.13)" if sombra_der else None,
                    hovertemplate=(
                        "%{x|%d/%m/%Y}<br>"
                        + nombre_der
                        + ": %{y:,.2f} "
                        + unidad_der
                        + "<extra></extra>"
                    )
                ),
                secondary_y=True
            )

    # =====================
    # FORMATO GENERAL
    # =====================
    fig.update_layout(
        title=titulo,
        height=480,
        template="plotly_white",

        paper_bgcolor="#EEF2F5",
        plot_bgcolor="#EEF2F5",

        font=dict(
            color="#000000",
            size=14
        ),

        title_font=dict(
            color="#0B3B36",
            size=22
        ),

        margin=dict(
            l=25,
            r=25,
            t=70,
            b=25
        ),

        hovermode="x unified",

        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.03,
            xanchor="right",
            x=1,
            font=dict(
                color="#000000",
                size=13
            )
        ),

        xaxis=dict(
            type="date",

            rangeselector=dict(
                bgcolor="#FFFFFF",
                activecolor="#C9A227",
                bordercolor="#0B3B36",
                borderwidth=1,
                font=dict(
                    color="#0B3B36",
                    size=12
                ),
                buttons=list([
                    dict(count=1, label="1A", step="year", stepmode="backward"),
                    dict(count=3, label="3A", step="year", stepmode="backward"),
                    dict(count=5, label="5A", step="year", stepmode="backward"),
                    dict(count=10, label="10A", step="year", stepmode="backward"),
                    dict(step="all", label="Todo")
                ])
            ),

            rangeslider=dict(
                visible=True,
                bgcolor="#E6D7A2",
                bordercolor="#0B3B36",
                borderwidth=2,
                thickness=0.12
            ),

            tickfont=dict(
                color="#000000",
                size=12
            ),

            gridcolor="rgba(0,0,0,0.10)",
            showgrid=True
        )
    )

    # =====================
    # EJE IZQUIERDO
    # =====================
    fig.update_yaxes(
        title_text=titulo_eje_izq,
        secondary_y=False,
        tickfont=dict(
            color=color_izq
        ),
        title_font=dict(
            color=color_izq
        ),
        gridcolor="rgba(0,0,0,0.12)",
        zeroline=True,
        zerolinecolor="rgba(11,59,54,0.45)"
    )

    # =====================
    # EJE DERECHO
    # =====================
    fig.update_yaxes(
        title_text=titulo_eje_der,
        secondary_y=True,
        tickfont=dict(
            color=color_der
        ),
        title_font=dict(
            color=color_der
        ),
        gridcolor="rgba(0,0,0,0)"
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        key=f"doble_eje_{titulo}_{uuid.uuid4()}"
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


def crear_serie_interanual(df, col, nombre_nueva_columna):
    if col is None:
        return df

    if col not in df.columns:
        return df

    df = df.copy()
    df = df.sort_values("fecha")

    df[nombre_nueva_columna] = (
        df[col].pct_change(periods=12) * 100
    )

    return df


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

tc_venta = buscar_columna_multiple([
    "Valor referencial de venta del dólar estadounidense",
    "Valor referencial de venta",
    "Tipo de cambio referencial"
])

# El tipo de cambio oficial cambió de fuente dentro de la base:
# - Hasta mayo de 2026: tipo de cambio de venta en el Bolsín.
# - Desde junio de 2026: columna específica de Tipo de Cambio Oficial.
# Se construye una sola serie consolidada, priorizando el dato nuevo.
tc_oficial_nuevo = buscar_columna_multiple([
    "Tipo de Cambio Oficial (Bs/USD)",
    "Tipo de cambio oficial Bs/USD",
    "Tipo de cambio oficial"
])

tc_oficial_historico = buscar_columna_multiple([
    "Tipo de cambio de venta en el Bolsín",
    "Tipo de cambio de venta en el Bolsin"
])

tc_oficial = None

if tc_oficial_nuevo is not None or tc_oficial_historico is not None:
    tc_oficial = "Tipo de cambio oficial consolidado"

    if tc_oficial_nuevo is not None:
        serie_oficial_nueva = df_original[tc_oficial_nuevo]
    else:
        serie_oficial_nueva = pd.Series(index=df_original.index, dtype="float64")

    if tc_oficial_historico is not None:
        serie_oficial_historica = df_original[tc_oficial_historico]
    else:
        serie_oficial_historica = pd.Series(index=df_original.index, dtype="float64")

    df_original[tc_oficial] = serie_oficial_nueva.combine_first(
        serie_oficial_historica
    )

exportaciones_valor = buscar_columna_multiple([
    "Exportaciones (En millones de dólares)",
    "Exportaciones en millones de dólares",
    "Exportaciones"
])

exportaciones_peso = buscar_columna_multiple([
    "Exportaciones (Peso neto en toneladas)",
    "Exportaciones Peso neto en toneladas",
    "Peso neto en toneladas"
])

importaciones_valor = buscar_columna_multiple([
    "Importaciones (Valor CIF en millones dólares)",
    "Importaciones Valor CIF en millones dólares",
    "Importaciones en millones dólares",
    "Importaciones"
])

importaciones_peso = buscar_columna_multiple([
    "Importaciones (Peso Bruto en Toneladas)",
    "Importaciones Peso Bruto en Toneladas",
    "Peso Bruto en Toneladas"
])

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

encaje_constituido = buscar_columna_multiple([
    "Encaje constituido por el sistema financiero",
    "Encaje del sistema financiero"
])

excedente_encaje_efectivo = buscar_columna_multiple([
    "Excedente de Encaje en el BCB del sistema financiero ( en efectivo; liquidez del sistema financiero )",
    "Excedente de Encaje en el BCB del sistema financiero en efectivo",
])

excedente_encaje_me = buscar_columna_multiple([
    "Excedente de Encaje en el BCB del sistema financiero en ME",
    "Excedente de Encaje en ME",
])


# Sector monetario
base_monetaria = buscar_columna("Base monetaria")
m1 = buscar_columna("M’1")
m2 = buscar_columna("M’2")
m3 = buscar_columna("M’3")
crec_base_monetaria = "Crecimiento interanual Base monetaria"

titulos_bcb_usd = buscar_columna_multiple([
    "Saldo de Títulos del Banco Central de Bolivia (millones de $us)",
    "Saldo de Titulos del Banco Central de Bolivia (millones de $us)",
    "Saldo de Títulos del Banco Central de Bolivia",
    "Saldo de Titulos del Banco Central de Bolivia"
])

tasa_reporto_mn = buscar_columna_multiple([
    "Tasas premio de reporto del BCB en Moneda nacional",
    "Tasa premio de reporto del BCB en Moneda nacional",
    "Tasas premio de reporto",
    "Premio de reporto"
])








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


# ============================================================
# CENGOB MACRO MONITOR 2.0 - INTERFAZ EJECUTIVA
# ============================================================
# Esta capa reemplaza las pestañas del prototipo original por una navegación
# ejecutiva, incorpora fechas por indicador, cruces sincronizados y métricas
# derivadas. Mantiene la paleta institucional CENGOB.

from html import escape

CENGOB_GREEN = "#0B3B36"
CENGOB_GREEN_2 = "#145A50"
CENGOB_GOLD = "#C9A227"
CENGOB_GOLD_SOFT = "#E6D7A2"
CENGOB_BG = "#EEF2F5"
CENGOB_CARD = "#FFFFFF"
CENGOB_TEXT = "#17212B"
CENGOB_MUTED = "#64748B"
CENGOB_BORDER = "#D9E1E7"

# CSS reforzado: evita pérdida visual de títulos/números y mejora responsive.
st.markdown(f"""
<style>
.block-container {{
    padding-top: 1.05rem;
    padding-bottom: 2.5rem;
    max-width: 1680px;
}}
.stApp {{ background:{CENGOB_BG}; }}
[data-testid="stHeader"] {{ background:{CENGOB_GREEN}; }}
[data-testid="stSidebar"] {{
    background:#F7FAF9;
    border-right:1px solid {CENGOB_BORDER};
}}
[data-testid="stSidebar"] * {{ color:{CENGOB_TEXT}; }}

h1,h2,h3,h4,h5,h6 {{
    color:{CENGOB_GREEN} !important;
    letter-spacing:-0.02em;
}}

/* Refuerzo para métricas nativas que aún se usen */
[data-testid="stMetric"] {{
    background:#FFFFFF;
    border:1px solid {CENGOB_BORDER};
    border-left:5px solid {CENGOB_GOLD};
    border-radius:16px;
    padding:15px 16px;
    min-height:128px;
    overflow:visible !important;
}}
[data-testid="stMetricLabel"] p {{
    color:{CENGOB_TEXT} !important;
    font-weight:700 !important;
    white-space:normal !important;
    line-height:1.25 !important;
}}
[data-testid="stMetricValue"] {{
    color:{CENGOB_GREEN} !important;
    font-size:clamp(1.35rem,2vw,2.1rem) !important;
    font-weight:850 !important;
    white-space:nowrap !important;
}}
[data-testid="stMetricDelta"] {{
    font-size:.84rem !important;
    white-space:normal !important;
}}

/* Radio de navegación como menú lateral */
[data-testid="stSidebar"] [role="radiogroup"] label {{
    padding:.44rem .55rem;
    border-radius:10px;
    margin-bottom:.1rem;
}}
[data-testid="stSidebar"] [role="radiogroup"] label:hover {{
    background:#E8F0EE;
}}

/* Plotly: marco consistente */
[data-testid="stPlotlyChart"] {{
    background:#FFFFFF;
    border:1px solid {CENGOB_BORDER};
    border-radius:16px;
    padding:4px 6px 2px 6px;
    box-shadow:0 3px 12px rgba(15,23,42,.035);
}}

div[data-testid="stDataFrame"] {{
    border:1px solid {CENGOB_BORDER};
    border-radius:14px;
    overflow:hidden;
}}

.cengob-hero {{
    background:linear-gradient(115deg,{CENGOB_GREEN} 0%, #104A42 70%, #1A6156 100%);
    border-radius:20px;
    padding:22px 26px;
    color:white;
    box-shadow:0 8px 24px rgba(11,59,54,.14);
    margin-bottom:14px;
}}
.cengob-hero h1 {{
    color:white !important;
    margin:0 0 4px 0;
    font-size:clamp(1.55rem,3vw,2.55rem);
}}
.cengob-hero p {{
    margin:0;
    color:#E6F0EE;
    font-size:.98rem;
}}
.cengob-chip {{
    display:inline-block;
    padding:5px 9px;
    border-radius:999px;
    background:rgba(255,255,255,.12);
    border:1px solid rgba(255,255,255,.18);
    color:white;
    font-size:.78rem;
    margin-right:5px;
    margin-top:10px;
}}

.kpi-card {{
    background:{CENGOB_CARD};
    border:1px solid {CENGOB_BORDER};
    border-top:4px solid {CENGOB_GOLD};
    border-radius:16px;
    padding:14px 15px 13px 15px;
    min-height:156px;
    box-shadow:0 4px 14px rgba(15,23,42,.04);
    overflow:hidden;
}}
.kpi-title {{
    color:{CENGOB_TEXT};
    font-size:.86rem;
    font-weight:750;
    line-height:1.25;
    min-height:2.15em;
}}
.kpi-value {{
    color:{CENGOB_GREEN};
    font-size:clamp(1.35rem,2.15vw,2.15rem);
    font-weight:900;
    letter-spacing:-.035em;
    line-height:1.1;
    margin-top:7px;
    white-space:nowrap;
    overflow:hidden;
    text-overflow:ellipsis;
}}
.kpi-delta {{
    color:{CENGOB_TEXT};
    font-size:.80rem;
    font-weight:650;
    margin-top:7px;
    min-height:1.15em;
    white-space:normal;
}}
.kpi-meta {{
    color:{CENGOB_MUTED};
    font-size:.73rem;
    margin-top:7px;
    display:flex;
    gap:5px;
    flex-wrap:wrap;
    align-items:center;
}}
.badge-ok,.badge-warn,.badge-old,.badge-info {{
    border-radius:999px;
    padding:2px 6px;
    font-weight:750;
    font-size:.69rem;
}}
.badge-ok {{ background:#DCFCE7;color:#166534; }}
.badge-warn {{ background:#FEF3C7;color:#92400E; }}
.badge-old {{ background:#FEE2E2;color:#991B1B; }}
.badge-info {{ background:#E2E8F0;color:#334155; }}

.section-card {{
    background:#FFFFFF;
    border:1px solid {CENGOB_BORDER};
    border-radius:16px;
    padding:16px 18px;
    margin:7px 0 12px 0;
}}
.exec-title {{
    color:{CENGOB_GREEN};
    font-size:1rem;
    font-weight:850;
    margin-bottom:7px;
}}
.exec-line {{
    color:{CENGOB_TEXT};
    font-size:.91rem;
    line-height:1.48;
    margin:5px 0;
}}
.small-note {{ color:{CENGOB_MUTED}; font-size:.78rem; line-height:1.35; }}
</style>
""", unsafe_allow_html=True)

# -------------------------
# Variables adicionales de la base
# -------------------------
reservas_brutas = buscar_columna("Reservas internacionales brutas del BCB")
recursos_alta_liquidez = buscar_columna("Recursos de Alta Liquidez")

financiamiento_externo_spnf = buscar_columna("Financiamiento Externo del SPNF")
financiamiento_interno_spnf = buscar_columna("Financiamiento Interno del SPNF")

pobreza_urbana = buscar_columna_multiple(["Urbano: Incidencia de pobreza", "Urbano: Incidencia de pobreza "])
pobreza_rural = buscar_columna_multiple(["Rural: Incidencia de pobreza", "Rural: Incidencia de pobreza "])
pobreza_extrema_bolivia = buscar_columna_multiple(["Bolivia: Indidencia de pobreza extrema", "Bolivia: Incidencia de pobreza extrema"])
pobreza_extrema_urbana = buscar_columna_multiple(["Urbano: Incidencia de pobreza extrema"])
pobreza_extrema_rural = buscar_columna_multiple(["Rural: Incidencia de pobreza extrema"])
gini_urbano = buscar_columna_multiple(["Urbano: Índice de GINI", "Urbano: Indice de GINI"])
gini_rural = buscar_columna_multiple(["Rural: Índice de GINI", "Rural: Indice de GINI"])
pea = buscar_columna("Población Económicamente Activa")

billetes_publico = buscar_columna("Billetes y monedas en poder del público")
emision_monetaria = buscar_columna("Emisión monetaria")
reservas_bancarias_totales = buscar_columna("Reservas bancarias totales")
credito_neto_spnf_bcb = buscar_columna("Crédito neto del BCB al Sector Público No financiero")

saldo_tgn_usd = buscar_columna("Saldo de Títulos del Tesoro General de la Nación (millones de $us)")

# -------------------------
# Utilidades analíticas
# -------------------------
def _serie(df_, col):
    if col is None or col not in df_.columns:
        return pd.DataFrame(columns=["fecha", "valor"])
    s = df_[["fecha", col]].dropna().sort_values("fecha").copy()
    s = s.rename(columns={col: "valor"})
    return s


def ultimo_y_anterior(df_, col):
    s = _serie(df_, col)
    if s.empty:
        return None, None, None, None
    last = s.iloc[-1]
    prev = s.iloc[-2] if len(s) >= 2 else None
    return (
        last["valor"], last["fecha"],
        None if prev is None else prev["valor"],
        None if prev is None else prev["fecha"]
    )


def valor_en_o_antes(df_, col, fecha_objetivo):
    s = _serie(df_, col)
    if s.empty:
        return None, None
    tmp = s[s["fecha"] <= pd.Timestamp(fecha_objetivo)]
    if tmp.empty:
        return None, None
    r = tmp.iloc[-1]
    return r["valor"], r["fecha"]


def variacion_pct(valor, base):
    if valor is None or base is None or pd.isna(valor) or pd.isna(base) or base == 0:
        return None
    return (valor / base - 1) * 100


def variacion_pp(valor, base):
    if valor is None or base is None or pd.isna(valor) or pd.isna(base):
        return None
    return valor - base


def ultimo_par_sincronizado(df_, col1, col2, tolerancia_dias=0):
    """Retorna el par más reciente con fechas comparables; evita falsas brechas."""
    s1 = _serie(df_, col1)
    s2 = _serie(df_, col2)
    if s1.empty or s2.empty:
        return None
    a = s1.rename(columns={"valor":"v1"})
    b = s2.rename(columns={"valor":"v2"})
    if tolerancia_dias == 0:
        m = a.merge(b, on="fecha", how="inner")
        if m.empty:
            return None
        r = m.iloc[-1]
        return r["fecha"], r["v1"], r["v2"]
    # merge_asof para tolerancia controlada
    m = pd.merge_asof(
        a.sort_values("fecha"), b.sort_values("fecha"),
        on="fecha", direction="nearest",
        tolerance=pd.Timedelta(days=tolerancia_dias)
    ).dropna()
    if m.empty:
        return None
    r = m.iloc[-1]
    return r["fecha"], r["v1"], r["v2"]


def frescura(fecha, frecuencia="mensual"):
    if fecha is None or pd.isna(fecha):
        return "Sin dato", "old"
    hoy = pd.Timestamp.now(tz="America/La_Paz").tz_localize(None).normalize()
    dias = max(0, (hoy - pd.Timestamp(fecha).normalize()).days)
    limites = {
        "diaria": (3, 10),
        "mensual": (45, 75),
        "trimestral": (120, 180),
        "anual": (450, 650),
    }
    ok, old = limites.get(frecuencia, (45, 90))
    if dias <= ok:
        return "Actual", "ok"
    if dias <= old:
        return f"Rezago {dias}d", "warn"
    return f"Desactualizado {dias}d", "old"


def fmt(x, dec=1, abs_value=False):
    if x is None or pd.isna(x):
        return "s/d"
    if abs_value:
        x = abs(x)
    text = f"{x:,.{dec}f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def kpi_card(titulo, valor, unidad="", fecha=None, delta=None, frecuencia="mensual", nota=None):
    estado, cls = frescura(fecha, frecuencia)
    fecha_txt = fecha.strftime("%d/%m/%Y") if fecha is not None and not pd.isna(fecha) else "sin fecha"
    delta_html = escape(delta) if delta else "&nbsp;"
    nota_html = f"<span>· {escape(str(nota))}</span>" if nota else ""
    st.markdown(f"""
    <div class="kpi-card">
      <div class="kpi-title">{escape(str(titulo))}</div>
      <div class="kpi-value">{escape(str(valor))}{(' ' + escape(unidad)) if unidad else ''}</div>
      <div class="kpi-delta">{delta_html}</div>
      <div class="kpi-meta">
        <span>{fecha_txt}</span>
        <span class="badge-{cls}">{estado}</span>
        {nota_html}
      </div>
    </div>
    """, unsafe_allow_html=True)


def estilo_fig(fig, titulo=None, altura=390, unidad_y=""):
    fig.update_layout(
        height=altura,
        template="plotly_white",
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(color=CENGOB_TEXT, size=12),
        margin=dict(l=35, r=25, t=68 if titulo else 35, b=42),
        hovermode="x unified",
        title=(dict(text=titulo, x=0.02, xanchor="left", font=dict(size=17, color=CENGOB_GREEN)) if titulo else None),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0, font=dict(size=11)),
    )
    fig.update_xaxes(showgrid=False, automargin=True, tickfont=dict(size=10))
    fig.update_yaxes(gridcolor="rgba(100,116,139,.16)", automargin=True, title_text=unidad_y, tickfont=dict(size=10))
    return fig


def linea_v2(df_, col, titulo, unidad="", color=CENGOB_GREEN, transform=None, altura=380):
    s = _serie(df_, col)
    if s.empty:
        st.info(f"Sin datos para {titulo}.")
        return
    y = s["valor"].copy()
    if transform == "abs":
        y = y.abs()
    fig = go.Figure(go.Scatter(
        x=s["fecha"], y=y,
        mode="lines",
        line=dict(color=color, width=3),
        fill="tozeroy",
        fillcolor="rgba(11,59,54,.07)" if color == CENGOB_GREEN else "rgba(201,162,39,.08)",
        hovertemplate="%{x|%d/%m/%Y}<br>%{y:,.2f}<extra></extra>"
    ))
    estilo_fig(fig, titulo, altura, unidad)
    st.plotly_chart(fig, use_container_width=True, key=f"v2_line_{titulo}_{uuid.uuid4()}")


def multi_v2(df_, specs, titulo, unidad="", altura=390):
    palette=[CENGOB_GREEN,CENGOB_GOLD,"#3F6B63","#8A6F1A","#475569","#2563EB"]
    fig=go.Figure()
    n=0
    for spec in specs:
        col=spec.get("col")
        s=_serie(df_,col)
        if s.empty:
            continue
        y=s["valor"].copy()
        if spec.get("abs"):
            y=y.abs()
        fig.add_trace(go.Scatter(
            x=s["fecha"], y=y, mode="lines",
            name=spec.get("name",str(col)),
            line=dict(width=2.8,color=spec.get("color",palette[n%len(palette)])),
            hovertemplate="%{x|%d/%m/%Y}<br>%{y:,.2f}<extra></extra>"
        ))
        n+=1
    if n==0:
        st.info(f"Sin datos para {titulo}.")
        return
    estilo_fig(fig,titulo,altura,unidad)
    st.plotly_chart(fig,use_container_width=True,key=f"v2_multi_{titulo}_{uuid.uuid4()}")


def doble_v2(df_, col1, col2, titulo, name1, name2, unidad1="", unidad2="", altura=410):
    s1=_serie(df_,col1)
    s2=_serie(df_,col2)
    if s1.empty and s2.empty:
        st.info(f"Sin datos para {titulo}.")
        return
    fig=make_subplots(specs=[[{"secondary_y":True}]])
    if not s1.empty:
        fig.add_trace(go.Scatter(x=s1["fecha"],y=s1["valor"],name=name1,
                                 line=dict(color=CENGOB_GREEN,width=3)),secondary_y=False)
    if not s2.empty:
        fig.add_trace(go.Scatter(x=s2["fecha"],y=s2["valor"],name=name2,
                                 line=dict(color=CENGOB_GOLD,width=3)),secondary_y=True)
    estilo_fig(fig,titulo,altura,"")
    fig.update_yaxes(title_text=unidad1, secondary_y=False, color=CENGOB_GREEN)
    fig.update_yaxes(title_text=unidad2, secondary_y=True, color="#8A6F1A", showgrid=False)
    st.plotly_chart(fig,use_container_width=True,key=f"v2_double_{titulo}_{uuid.uuid4()}")


def ytd_suma(df_, col, year, month_cut):
    if col is None or col not in df_.columns:
        return None
    s=df_[["fecha",col]].dropna()
    s=s[(s["fecha"].dt.year==year)&(s["fecha"].dt.month<=month_cut)]
    if s.empty:
        return None
    return s[col].sum()


def tabla_cobertura(df_):
    items = [
        ("Tipo de cambio oficial", tc_oficial, "Diaria"),
        ("Base monetaria", base_monetaria, "Diaria / mensual"),
        ("Inflación 12 meses", inflacion_12m, "Mensual"),
        ("RIN", rin, "Mensual"),
        ("Crédito privado", credito_privado, "Mensual"),
        ("Depósitos", depositos, "Mensual"),
        ("Saldo comercial", saldo_comercial, "Mensual"),
        ("Resultado fiscal global", resultado_global_spnf, "Mensual"),
        ("IGAE", igae, "Mensual"),
        ("PIB / crecimiento", pib_pm, "Trimestral"),
        ("Desocupación", desocupacion_nacional, "Mensual"),
        ("Pobreza", pobreza_bolivia, "Anual"),
    ]
    rows=[]
    for nombre,col,frec in items:
        v,f=ultimo_valor(df_,col)
        rows.append({"Indicador":nombre,"Último dato":None if f is None else f.date(),"Frecuencia":frec,"Valor":v})
    return pd.DataFrame(rows)


# -------------------------
# Sidebar / filtros
# -------------------------
st.sidebar.markdown("## CENGOB Macro Monitor")
st.sidebar.caption("Sistema ejecutivo de seguimiento económico")

if st.sidebar.button("🔄 Actualizar desde Drive", use_container_width=True):
    descargar_excel_drive.clear()
    cargar_datos.clear()
    st.rerun()

auto_refresh = st.sidebar.toggle("Actualización automática", value=True, help="Recarga la aplicación cada 5 minutos para consultar cambios en Drive.")
if auto_refresh:
    components.html("""
    <script>
      setTimeout(function(){ window.parent.location.reload(); }, 300000);
    </script>
    """, height=0)

fecha_min=df_original["fecha"].min().date()
fecha_max=df_original["fecha"].max().date()
rango=st.sidebar.date_input("Ventana de gráficos", value=(max(fecha_min, (df_original["fecha"].max()-pd.DateOffset(years=5)).date()), fecha_max), min_value=fecha_min, max_value=fecha_max)

df_view=df_original.copy()
if isinstance(rango,(tuple,list)) and len(rango)==2:
    ini=pd.Timestamp(rango[0]); fin=pd.Timestamp(rango[1])
    df_view=df_view[(df_view["fecha"]>=ini)&(df_view["fecha"]<=fin)]

seccion=st.sidebar.radio(
    "Navegación",
    [
        "Resumen ejecutivo",
        "Cruces y señales",
        "Precios y cambio",
        "Monetario y financiero",
        "Sector externo",
        "Fiscal",
        "Actividad y social",
        "Cobertura de datos",
        "Explorador"
    ],
    label_visibility="collapsed"
)

st.sidebar.markdown("---")
st.sidebar.caption(f"Fuente: Google Drive · {len(df_original.columns)-1} variables procesadas")
st.sidebar.caption(f"Última fecha en la base: {df_original['fecha'].max().strftime('%d/%m/%Y')}")

# -------------------------
# Header ejecutivo
# -------------------------
st.markdown(f"""
<div class="cengob-hero">
  <h1>Sistema de Información Económica de Bolivia</h1>
  <p>Centro de Gobierno · CENGOB | monitoreo macroeconómico para toma de decisiones</p>
  <span class="cengob-chip">Última fila: {df_original['fecha'].max().strftime('%d/%m/%Y')}</span>
  <span class="cengob-chip">{len(df_original.columns)-1} variables</span>
  <span class="cengob-chip">Actualización automática desde Drive</span>
</div>
""", unsafe_allow_html=True)

# ============================================================
# 1. RESUMEN EJECUTIVO
# ============================================================
if seccion == "Resumen ejecutivo":
    st.subheader("Resumen ejecutivo")
    st.caption("Cada tarjeta conserva su propia fecha de corte. No se fuerza una fecha común cuando la base mezcla frecuencias.")

    # Datos clave
    tc_v,tc_f,tc_prev,tc_prev_f=ultimo_y_anterior(df_original,tc_oficial)
    infl_v,infl_f,_,_=ultimo_y_anterior(df_original,inflacion_12m)
    rin_v,rin_f,_,_=ultimo_y_anterior(df_original,rin)
    bm_v,bm_f,_,_=ultimo_y_anterior(df_original,base_monetaria)
    cred_v,cred_f,_,_=ultimo_y_anterior(df_original,credito_privado)
    dep_v,dep_f,_,_=ultimo_y_anterior(df_original,depositos)

    # YoY base monetaria usando fecha comparable anterior
    bm_yoy=None
    if bm_f is not None:
        bm_base,_=valor_en_o_antes(df_original,base_monetaria,bm_f-pd.DateOffset(years=1))
        bm_yoy=variacion_pct(bm_v,bm_base)

    # TC: cambio contra observación previa, explícitamente con fecha
    tc_delta=variacion_pct(tc_v,tc_prev)
    tc_delta_txt=(f"{fmt(tc_delta,1)}% vs {tc_prev_f.strftime('%d/%m')}" if tc_delta is not None and tc_prev_f is not None else None)

    # RIN: comparación mensual previa
    rin_s=_serie(df_original,rin)
    rin_prev=rin_s.iloc[-2]["valor"] if len(rin_s)>=2 else None
    rin_mom=variacion_pct(rin_v,rin_prev)

    # Crédito/depósitos: razón, que no cambia por unidad USD/Bs
    cred_dep=(cred_v/dep_v*100) if cred_v is not None and dep_v not in (None,0) else None

    c1,c2,c3,c4,c5,c6=st.columns(6)
    with c1:
        kpi_card("Tipo de cambio oficial",fmt(tc_v,2),"Bs/USD",tc_f,tc_delta_txt,"diaria")
    with c2:
        kpi_card("Inflación interanual",fmt(infl_v,2),"%",infl_f,"Variación a 12 meses","mensual")
    with c3:
        kpi_card("Reservas internacionales netas",fmt(rin_v,1),"MM USD",rin_f,(f"{fmt(rin_mom,1)}% vs dato previo" if rin_mom is not None else None),"mensual")
    with c4:
        kpi_card("Base monetaria",fmt(bm_v,0),"MM Bs",bm_f,(f"{fmt(bm_yoy,1)}% interanual" if bm_yoy is not None else None),"diaria")
    with c5:
        kpi_card("Crédito / depósitos",fmt(cred_dep,1),"%",min(cred_f,dep_f) if cred_f is not None and dep_f is not None else None,"Indicador de fondeo","mensual")
    with c6:
        liq_v,liq_f=ultimo_valor(df_original,excedente_encaje_efectivo)
        enc_v,enc_f=ultimo_valor(df_original,encaje_constituido)
        liq_ratio=(liq_v/enc_v*100) if liq_v is not None and enc_v not in (None,0) else None
        kpi_card("Excedente / encaje",fmt(liq_ratio,1),"%",liq_f,"Holgura relativa de liquidez","mensual")

    # Lectura automática sustentada en la base
    infl_m_v,infl_m_f=ultimo_valor(df_original,inflacion_mensual)
    oro_v,oro_f=ultimo_valor(df_original,oro)
    ral_v,ral_f=ultimo_valor(df_original,recursos_alta_liquidez)
    oro_share=(oro_v/rin_v*100) if oro_v is not None and rin_v not in (None,0) else None
    ral_share=(ral_v/rin_v*100) if ral_v is not None and rin_v not in (None,0) else None

    # Fiscal H1 (si existe junio del último año con fiscal)
    fisc_s=_serie(df_original,resultado_global_spnf)
    fisc_year=int(fisc_s.iloc[-1]["fecha"].year) if not fisc_s.empty else None
    fisc_month=int(fisc_s.iloc[-1]["fecha"].month) if not fisc_s.empty else None
    fisc_ytd=ytd_suma(df_original,resultado_global_spnf,fisc_year,fisc_month) if fisc_year else None
    fisc_prev_ytd=ytd_suma(df_original,resultado_global_spnf,fisc_year-1,fisc_month) if fisc_year else None
    mejora_def=None
    if fisc_ytd is not None and fisc_prev_ytd is not None:
        mejora_def=fisc_ytd-fisc_prev_ytd

    # Comercio YTD según última observación disponible
    trade_s=_serie(df_original,exportaciones_valor)
    trade_year=int(trade_s.iloc[-1]["fecha"].year) if not trade_s.empty else None
    trade_month=int(trade_s.iloc[-1]["fecha"].month) if not trade_s.empty else None
    exp_ytd=ytd_suma(df_original,exportaciones_valor,trade_year,trade_month) if trade_year else None
    exp_prev_ytd=ytd_suma(df_original,exportaciones_valor,trade_year-1,trade_month) if trade_year else None
    imp_ytd=ytd_suma(df_original,importaciones_valor,trade_year,trade_month) if trade_year else None
    saldo_ytd=ytd_suma(df_original,saldo_comercial,trade_year,trade_month) if trade_year else None
    exp_yoy=variacion_pct(exp_ytd,exp_prev_ytd)

    st.markdown(f"""
    <div class="section-card">
      <div class="exec-title">Lectura ejecutiva automática</div>
      <div class="exec-line"><b>Precios y cambio.</b> La inflación interanual disponible es <b>{fmt(infl_v,2)}%</b> ({infl_f.strftime('%d/%m/%Y') if infl_f is not None else 's/f'}) y la inflación mensual es <b>{fmt(infl_m_v,2)}%</b>. El TC oficial llega a <b>{fmt(tc_v,2)} Bs/USD</b> al {tc_f.strftime('%d/%m/%Y') if tc_f is not None else 's/f'}.</div>
      <div class="exec-line"><b>Reservas.</b> Las RIN ascienden a <b>{fmt(rin_v,1)} MM USD</b>; el oro representa aproximadamente <b>{fmt(oro_share,1)}%</b> y los recursos de alta liquidez <b>{fmt(ral_share,1)}%</b>. La composición importa tanto como el nivel total.</div>
      <div class="exec-line"><b>Fiscal.</b> El resultado global acumulado hasta {fisc_month:02d}/{fisc_year} es <b>{fmt(fisc_ytd,1)} MM Bs</b>; frente al mismo corte de {fisc_year-1}, la diferencia es <b>{fmt(mejora_def,1)} MM Bs</b>.</div>
      <div class="exec-line"><b>Comercio exterior.</b> Al corte {trade_month:02d}/{trade_year}, el saldo acumulado es <b>{fmt(saldo_ytd,1)} MM USD</b>; las exportaciones acumulan <b>{fmt(exp_ytd,1)} MM USD</b> ({fmt(exp_yoy,1)}% interanual) y las importaciones <b>{fmt(imp_ytd,1,abs_value=True)} MM USD</b>.</div>
    </div>
    """,unsafe_allow_html=True)

    # Tres gráficos que responden preguntas ejecutivas
    g1,g2=st.columns([1.15,1])
    with g1:
        doble_v2(df_view,tc_oficial,inflacion_12m,"Tipo de cambio e inflación","TC oficial","Inflación 12m","Bs/USD","%")
    with g2:
        doble_v2(df_view,rin,tc_oficial,"Reservas y tipo de cambio","RIN","TC oficial","MM USD","Bs/USD")

    g3,g4=st.columns(2)
    with g3:
        multi_v2(df_view,[{"col":base_monetaria,"name":"Base monetaria"},{"col":m3,"name":"M3"}],"Liquidez monetaria: base y M3","MM Bs")
    with g4:
        multi_v2(df_view,[{"col":credito_privado,"name":"Crédito"},{"col":depositos,"name":"Depósitos"}],"Crédito y depósitos reportados en USD","MM USD")
        st.caption("Desde el cambio del régimen cambiario, la expresión en USD puede incorporar un efecto de conversión. Ver 'Cruces y señales'.")

# ============================================================
# 2. CRUCES Y SEÑALES
# ============================================================
elif seccion == "Cruces y señales":
    st.subheader("Cruces de variables y señales tempranas")
    st.caption("Los cruces se construyen solo cuando las fechas son comparables o se documenta explícitamente la transformación.")

    # 1) Evitar falsa brecha TC
    par_tc=ultimo_par_sincronizado(df_original,tc_venta,tc_oficial,tolerancia_dias=0)
    if par_tc is None:
        st.warning("No existe una observación reciente en la que el TC referencial y el TC oficial estén disponibles el mismo día. Por ello, la V2 no calcula una 'brecha cambiaria actual' mezclando fechas distintas.")
    else:
        f,vref,vof=par_tc
        st.info(f"Última comparación sincronizada: {f:%d/%m/%Y} · Referencial {fmt(vref,2)} · Oficial {fmt(vof,2)}")

    # Reconversión aproximada de crédito y depósitos a Bs para aislar efecto denominación
    base=df_original[["fecha"]].copy()
    if credito_privado in df_original.columns and depositos in df_original.columns and tc_oficial in df_original.columns:
        base["credito_bs_aprox"]=df_original[credito_privado]*df_original[tc_oficial]
        base["depositos_bs_aprox"]=df_original[depositos]*df_original[tc_oficial]
        base["ratio_cd"]=df_original[credito_privado]/df_original[depositos]*100
    else:
        base["credito_bs_aprox"]=None; base["depositos_bs_aprox"]=None; base["ratio_cd"]=None

    c1,c2=st.columns(2)
    with c1:
        fig=go.Figure()
        s=base[["fecha","credito_bs_aprox","depositos_bs_aprox"]].dropna()
        if not s.empty:
            fig.add_trace(go.Scatter(x=s["fecha"],y=s["credito_bs_aprox"],name="Crédito aprox. Bs",line=dict(color=CENGOB_GREEN,width=3)))
            fig.add_trace(go.Scatter(x=s["fecha"],y=s["depositos_bs_aprox"],name="Depósitos aprox. Bs",line=dict(color=CENGOB_GOLD,width=3)))
            estilo_fig(fig,"Crédito y depósitos: aproximación en Bs",390,"MM Bs")
            st.plotly_chart(fig,use_container_width=True,key="cred_dep_bs")
            st.caption("Aproximación analítica = saldo reportado en MM USD × TC oficial de la base. Sirve para separar el efecto de conversión; no reemplaza una serie oficial en Bs.")
        else:
            st.info("No hay observaciones sincronizadas suficientes para la reconversión.")
    with c2:
        fig=go.Figure()
        s=base[["fecha","ratio_cd"]].dropna()
        if not s.empty:
            fig.add_trace(go.Scatter(x=s["fecha"],y=s["ratio_cd"],line=dict(color=CENGOB_GREEN,width=3)))
            fig.add_hline(y=100,line_dash="dash",line_color=CENGOB_GOLD,annotation_text="100%")
            estilo_fig(fig,"Crédito / depósitos",390,"%")
            st.plotly_chart(fig,use_container_width=True,key="ratio_cd")

    # RIN: composición y liquidez
    c3,c4=st.columns(2)
    with c3:
        comps=[]
        for nombre,col in [("Oro",oro),("Divisas",divisas),("DEG",buscar_columna("DEG")),("Posición FMI",posicion_fmi)]:
            v,f=ultimo_valor(df_original,col)
            if v is not None:
                comps.append((nombre,v))
        if comps:
            fig=go.Figure(go.Bar(x=[x[0] for x in comps],y=[x[1] for x in comps],marker_color=[CENGOB_GOLD,CENGOB_GREEN,"#64748B","#94A3B8"]))
            estilo_fig(fig,"Composición de reservas internacionales",390,"MM USD")
            st.plotly_chart(fig,use_container_width=True,key="rin_comp")
    with c4:
        doble_v2(df_view,recursos_alta_liquidez,rin,"Liquidez inmediata dentro de las RIN","Recursos alta liquidez","RIN","MM USD","MM USD")

    # Comercio: precio implícito por tonelada
    trade=df_original[["fecha",exportaciones_valor,exportaciones_peso,importaciones_valor,importaciones_peso]].copy()
    trade["px_exp"]=trade[exportaciones_valor]*1_000_000/trade[exportaciones_peso]
    trade["px_imp"]=trade[importaciones_valor].abs()*1_000_000/trade[importaciones_peso].abs()
    c5,c6=st.columns(2)
    with c5:
        fig=go.Figure()
        s=trade[["fecha","px_exp","px_imp"]].replace([float("inf"),-float("inf")],pd.NA).dropna()
        fig.add_trace(go.Scatter(x=s["fecha"],y=s["px_exp"],name="Exportación",line=dict(color=CENGOB_GREEN,width=3)))
        fig.add_trace(go.Scatter(x=s["fecha"],y=s["px_imp"],name="Importación",line=dict(color=CENGOB_GOLD,width=3)))
        estilo_fig(fig,"Valor unitario implícito del comercio",390,"USD/ton")
        st.plotly_chart(fig,use_container_width=True,key="trade_unit")
    with c6:
        multi_v2(df_view,[{"col":exportaciones_valor,"name":"Exportaciones"},{"col":importaciones_valor,"name":"Importaciones","abs":True}],"Exportaciones e importaciones","MM USD")

    # Heatmap de cambios estandarizados - variables seleccionadas
    st.markdown("### Radar de presiones")
    st.caption("Estandariza cambios históricos de variables seleccionadas para comparar señales con escalas distintas; no sustituye una regla de política.")
    radar_specs=[
        ("Inflación",inflacion_12m),
        ("TC oficial",tc_oficial),
        ("Base monetaria",base_monetaria),
        ("RIN",rin),
        ("Liquidez EIF",excedente_encaje_efectivo),
        ("Resultado fiscal",resultado_global_spnf),
        ("Desocupación",desocupacion_nacional),
    ]
    rows=[]
    for nombre,col in radar_specs:
        s=_serie(df_original,col)
        if len(s)>=8:
            # z-score robusto simple sobre últimas 60 observaciones
            tail=s.tail(60)["valor"].astype(float)
            mean=tail.mean(); std=tail.std()
            z=(tail.iloc[-1]-mean)/std if std not in (0,None) and not pd.isna(std) else None
            rows.append((nombre,z,s.iloc[-1]["fecha"]))
    if rows:
        heat=pd.DataFrame(rows,columns=["Indicador","z","fecha"])
        fig=go.Figure(go.Heatmap(z=[heat["z"].tolist()],x=heat["Indicador"].tolist(),y=["Último dato"],colorscale=[[0,"#B91C1C"],[.5,"#F8FAFC"],[1,"#0B3B36"]],zmid=0,colorbar=dict(title="z")))
        estilo_fig(fig,"Desvío respecto a la historia reciente",300,"")
        st.plotly_chart(fig,use_container_width=True,key="heat_risk")

# ============================================================
# 3. PRECIOS Y CAMBIO
# ============================================================
elif seccion == "Precios y cambio":
    st.subheader("Precios y mercado cambiario")
    v12,f12=ultimo_valor(df_original,inflacion_12m)
    vm,fm=ultimo_valor(df_original,inflacion_mensual)
    va,fa=ultimo_valor(df_original,inflacion_acumulada)
    tv,tf=ultimo_valor(df_original,tc_oficial)
    c1,c2,c3,c4=st.columns(4)
    with c1:kpi_card("Inflación 12 meses",fmt(v12,2),"%",f12,"Nivel interanual","mensual")
    with c2:kpi_card("Inflación mensual",fmt(vm,2),"%",fm,"Momentum de precios","mensual")
    with c3:kpi_card("Inflación acumulada",fmt(va,2),"%",fa,"Gestión en curso","mensual")
    with c4:kpi_card("TC oficial",fmt(tv,2),"Bs/USD",tf,"Último valor","diaria")
    a,b=st.columns(2)
    with a:multi_v2(df_view,[{"col":inflacion_12m,"name":"12 meses"},{"col":inflacion_acumulada,"name":"Acumulada"},{"col":inflacion_mensual,"name":"Mensual"}],"Inflación: tres velocidades","%")
    with b:doble_v2(df_view,tc_oficial,inflacion_12m,"TC oficial e inflación","TC oficial","Inflación 12m","Bs/USD","%")
    linea_v2(df_view,indice_tc_real if 'indice_tc_real' in globals() else buscar_columna("Índice de tipo de cambio real"),"Índice de tipo de cambio real","Índice")

# ============================================================
# 4. MONETARIO Y FINANCIERO
# ============================================================
elif seccion == "Monetario y financiero":
    st.subheader("Monetario y financiero")
    bm,bmf=ultimo_valor(df_original,base_monetaria)
    m3v,m3f=ultimo_valor(df_original,m3)
    lv,lf=ultimo_valor(df_original,excedente_encaje_efectivo)
    cv,cf=ultimo_valor(df_original,credito_privado)
    dv,df_=ultimo_valor(df_original,depositos)
    ratio=(cv/dv*100) if cv is not None and dv not in (None,0) else None
    c1,c2,c3,c4=st.columns(4)
    with c1:kpi_card("Base monetaria",fmt(bm,0),"MM Bs",bmf,None,"diaria")
    with c2:kpi_card("M3",fmt(m3v,0),"MM Bs",m3f,None,"mensual")
    with c3:kpi_card("Liquidez EIF",fmt(lv,1),"MM USD",lf,"Excedente de encaje","mensual")
    with c4:kpi_card("Crédito / depósitos",fmt(ratio,1),"%",min(cf,df_) if cf is not None and df_ is not None else None,None,"mensual")
    a,b=st.columns(2)
    with a:multi_v2(df_view,[{"col":base_monetaria,"name":"Base monetaria"},{"col":m1,"name":"M1"},{"col":m2,"name":"M2"},{"col":m3,"name":"M3"}],"Agregados monetarios","MM Bs")
    with b:multi_v2(df_view,[{"col":encaje_constituido,"name":"Encaje constituido"},{"col":excedente_encaje_efectivo,"name":"Excedente efectivo"},{"col":excedente_encaje_me,"name":"Excedente ME"}],"Encaje y liquidez","MM USD")
    c,d=st.columns(2)
    with c:multi_v2(df_view,[{"col":credito_privado,"name":"Crédito"},{"col":depositos,"name":"Depósitos"}],"Crédito y depósitos reportados","MM USD")
    with d:multi_v2(df_view,[{"col":bol_cred,"name":"Bolivianización crédito"},{"col":bol_dep,"name":"Bolivianización depósitos"}],"Bolivianización","%")
    doble_v2(df_view,tasa_reporto_mn,excedente_encaje_efectivo,"Tasa de reporto y liquidez EIF","Tasa reporto MN","Liquidez EIF","%","MM USD")

# ============================================================
# 5. SECTOR EXTERNO
# ============================================================
elif seccion == "Sector externo":
    st.subheader("Sector externo")
    rv,rf=ultimo_valor(df_original,rin)
    dv,dfx=ultimo_valor(df_original,divisas)
    ov,of=ultimo_valor(df_original,oro)
    ralv,ralf=ultimo_valor(df_original,recursos_alta_liquidez)
    c1,c2,c3,c4=st.columns(4)
    with c1:kpi_card("RIN",fmt(rv,1),"MM USD",rf,None,"mensual")
    with c2:kpi_card("Divisas",fmt(dv,1),"MM USD",dfx,(f"{fmt(dv/rv*100,1)}% de RIN" if rv else None),"mensual")
    with c3:kpi_card("Oro",fmt(ov,1),"MM USD",of,(f"{fmt(ov/rv*100,1)}% de RIN" if rv else None),"mensual")
    with c4:kpi_card("Recursos alta liquidez",fmt(ralv,1),"MM USD",ralf,(f"{fmt(ralv/rv*100,1)}% de RIN" if rv else None),"mensual")
    a,b=st.columns(2)
    with a:multi_v2(df_view,[{"col":rin,"name":"RIN"},{"col":divisas,"name":"Divisas"},{"col":oro,"name":"Oro"}],"Reservas y composición","MM USD")
    with b:doble_v2(df_view,rin,tc_oficial,"RIN y tipo de cambio","RIN","TC oficial","MM USD","Bs/USD")
    c,d=st.columns(2)
    with c:multi_v2(df_view,[{"col":exportaciones_valor,"name":"Exportaciones"},{"col":importaciones_valor,"name":"Importaciones","abs":True}],"Comercio exterior mensual","MM USD")
    with d:linea_v2(df_view,saldo_comercial,"Saldo comercial","MM USD")
    linea_v2(df_view,buscar_columna("Índice de Términos del Intercambio"),"Índice de términos del intercambio","Índice")

# ============================================================
# 6. FISCAL
# ============================================================
elif seccion == "Fiscal":
    st.subheader("Sector fiscal · SPNF")
    fs=_serie(df_original,resultado_global_spnf)
    if not fs.empty:
        y=int(fs.iloc[-1]["fecha"].year); m=int(fs.iloc[-1]["fecha"].month)
        ing=ytd_suma(df_original,ingresos_totales_spnf,y,m)
        egr=ytd_suma(df_original,egresos_totales_spnf,y,m)
        res=ytd_suma(df_original,resultado_global_spnf,y,m)
        res_prev=ytd_suma(df_original,resultado_global_spnf,y-1,m)
        fin_ext=ytd_suma(df_original,financiamiento_externo_spnf,y,m)
        fin_int=ytd_suma(df_original,financiamiento_interno_spnf,y,m)
        corte=fs.iloc[-1]["fecha"]
    else:
        y=m=None; ing=egr=res=res_prev=fin_ext=fin_int=None; corte=None
    c1,c2,c3,c4=st.columns(4)
    with c1:kpi_card("Ingresos acumulados",fmt(ing,1),"MM Bs",corte,None,"mensual")
    with c2:kpi_card("Egresos acumulados",fmt(egr,1),"MM Bs",corte,None,"mensual")
    with c3:kpi_card("Resultado global acumulado",fmt(res,1),"MM Bs",corte,(f"Año previo: {fmt(res_prev,1)}" if res_prev is not None else None),"mensual")
    with c4:kpi_card("Financiamiento interno",fmt(fin_int,1),"MM Bs",corte,(f"Externo: {fmt(fin_ext,1)}" if fin_ext is not None else None),"mensual")
    a,b=st.columns(2)
    with a:multi_v2(df_view,[{"col":ingresos_totales_spnf,"name":"Ingresos"},{"col":egresos_totales_spnf,"name":"Egresos"}],"Ingresos y egresos mensuales","MM Bs")
    with b:multi_v2(df_view,[{"col":resultado_corriente_spnf,"name":"Resultado corriente"},{"col":resultado_global_spnf,"name":"Resultado global"}],"Resultados fiscales","MM Bs")
    c,d=st.columns(2)
    with c:multi_v2(df_view,[{"col":ingresos_corrientes_spnf,"name":"Corrientes"},{"col":ingresos_capital_spnf,"name":"Capital"}],"Composición de ingresos","MM Bs")
    with d:multi_v2(df_view,[{"col":egresos_corrientes_spnf,"name":"Corrientes"},{"col":egresos_capital_spnf,"name":"Capital"}],"Composición de egresos","MM Bs")
    multi_v2(df_view,[{"col":financiamiento_externo_spnf,"name":"Financiamiento externo"},{"col":financiamiento_interno_spnf,"name":"Financiamiento interno"}],"Fuentes de financiamiento del resultado fiscal","MM Bs")

# ============================================================
# 7. ACTIVIDAD Y SOCIAL
# ============================================================
elif seccion == "Actividad y social":
    st.subheader("Actividad real y condiciones sociales")
    pibv,pibf=ultimo_valor(df_original,pib_pm)
    igaev,igaef=ultimo_valor(df_original,igae)
    uv,uf=ultimo_valor(df_original,desocupacion_nacional)
    pv,pf=ultimo_valor(df_original,pobreza_bolivia)
    c1,c2,c3,c4=st.columns(4)
    with c1:kpi_card("PIB · tasa reportada",fmt(pibv,2),"%",pibf,"No se vuelve a calcular YoY sobre esta tasa","trimestral")
    with c2:kpi_card("IGAE · tasa reportada",fmt(igaev,2),"%",igaef,"Mostrar rezago explícito","mensual")
    with c3:kpi_card("Desocupación",fmt(uv,2),"%",uf,None,"mensual")
    with c4:kpi_card("Pobreza",fmt(pv,2),"%",pf,"Dato estructural, no coyuntural","anual")
    a,b=st.columns(2)
    with a:multi_v2(df_view,[{"col":pib_pm,"name":"PIB"},{"col":consumo_hogares,"name":"Consumo hogares"},{"col":formacion_capital,"name":"FBKF"}],"Crecimiento y demanda interna","%")
    with b:doble_v2(df_view,igae,desocupacion_nacional,"Actividad y desocupación","IGAE","Desocupación","%","%")
    c,d=st.columns(2)
    with c:multi_v2(df_view,[{"col":pobreza_bolivia,"name":"Bolivia"},{"col":pobreza_urbana,"name":"Urbana"},{"col":pobreza_rural,"name":"Rural"}],"Incidencia de pobreza","%")
    with d:multi_v2(df_view,[{"col":pobreza_extrema_bolivia,"name":"Bolivia"},{"col":pobreza_extrema_urbana,"name":"Urbana"},{"col":pobreza_extrema_rural,"name":"Rural"}],"Pobreza extrema","%")
    multi_v2(df_view,[{"col":gini_bolivia,"name":"Bolivia"},{"col":gini_urbano,"name":"Urbano"},{"col":gini_rural,"name":"Rural"}],"Índice de Gini","")

# ============================================================
# 8. COBERTURA / CALIDAD DE DATOS
# ============================================================
elif seccion == "Cobertura de datos":
    st.subheader("Cobertura, frecuencia y calidad de datos")
    st.caption("Esta vista evita que la última fecha de la hoja se interprete como fecha común para todas las variables.")
    cov=tabla_cobertura(df_original)
    st.dataframe(cov,use_container_width=True,hide_index=True)

    st.markdown("### Controles automáticos recomendados")
    st.markdown("""
    - **Sincronización:** no calcular brechas entre series si sus fechas no coinciden o exceden la tolerancia definida.
    - **Semántica:** distinguir `nivel`, `tasa`, `stock`, `flujo` e `índice` para impedir transformaciones incorrectas.
    - **Frecuencia:** mostrar el rezago esperado por indicador y marcar automáticamente datos vencidos.
    - **Signos contables:** las importaciones pueden venir negativas por convención; se muestran en valor absoluto en KPIs, conservando el signo para el saldo comercial.
    - **Quiebres:** marcar cambios de fuente, metodología o régimen (por ejemplo, series cambiarias) antes de calcular tasas de variación.
    """)

# ============================================================
# 9. EXPLORADOR
# ============================================================
elif seccion == "Explorador":
    st.subheader("Explorador de variables")
    variables=[c for c in df_original.columns if c!="fecha"]
    sel=st.selectbox("Variable",variables)
    linea_v2(df_view,sel,sel,"")
    vv,vf=ultimo_valor(df_original,sel)
    st.caption(f"Último valor: {fmt(vv,4)} · {vf.strftime('%d/%m/%Y') if vf is not None else 'sin fecha'}")

    cols=st.multiselect("Comparar con",[c for c in variables if c!=sel],max_selections=4)
    if cols:
        specs=[{"col":sel,"name":sel}]+[{"col":c,"name":c} for c in cols]
        multi_v2(df_view,specs,"Comparación seleccionada","")

# Descarga consistente con la ventana visual
st.markdown("---")
st.download_button(
    "⬇️ Descargar ventana de datos",
    data=df_view.to_csv(index=False).encode("utf-8"),
    file_name="cengob_macro_monitor_datos.csv",
    mime="text/csv",
    use_container_width=False
)
st.caption("CENGOB Macro Monitor 2.0 · La interpretación automática se basa exclusivamente en las series disponibles en la hoja cargada.")
