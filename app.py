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

# =========================
# SIDEBAR
# =========================

st.sidebar.title("⚙️ Panel de control")

# Streamlit no se vuelve a ejecutar solo cuando cambia el Excel de Drive.
# Este botón fuerza una descarga nueva y elimina las dos capas de caché.
if st.sidebar.button(
    "🔄 Actualizar datos desde Drive",
    use_container_width=True,
    help="Descarga nuevamente el Excel y actualiza los indicadores."
):
    descargar_excel_drive.clear()
    cargar_datos.clear()
    st.rerun()

fecha_min = df_original["fecha"].min()
fecha_max = df_original["fecha"].max()
fecha_min_date = fecha_min.date()
fecha_max_date = fecha_max.date()

# Mantiene el rango hasta la última fecha disponible cuando la fuente incorpora
# un dato nuevo. Sin este control, la sesión puede quedarse cerrada en 24/07/2026
# aunque el Excel ya contenga una fila del 05/08/2026.
max_previa = st.session_state.get("_fecha_max_fuente_previa")
rango_guardado = st.session_state.get("rango_fechas")

if rango_guardado is None:
    st.session_state["rango_fechas"] = (fecha_min_date, fecha_max_date)
elif isinstance(rango_guardado, (tuple, list)) and len(rango_guardado) == 2:
    inicio_guardado, fin_guardado = rango_guardado

    # Si el usuario estaba viendo hasta el último dato anterior, ampliar
    # automáticamente el rango hasta el nuevo máximo de la fuente.
    if max_previa is not None and fin_guardado == max_previa and fecha_max_date > max_previa:
        st.session_state["rango_fechas"] = (inicio_guardado, fecha_max_date)

    # Evita rangos fuera de los límites actuales de la base.
    inicio_ajustado = max(inicio_guardado, fecha_min_date)
    fin_ajustado = min(
        st.session_state["rango_fechas"][1],
        fecha_max_date
    )
    st.session_state["rango_fechas"] = (inicio_ajustado, fin_ajustado)

st.session_state["_fecha_max_fuente_previa"] = fecha_max_date

rango = st.sidebar.date_input(
    "Rango de fechas",
    min_value=fecha_min_date,
    max_value=fecha_max_date,
    key="rango_fechas"
)

st.sidebar.caption(
    f"Última fecha detectada en el Excel: {fecha_max.strftime('%d/%m/%Y')}"
)

df = df_original.copy()

df = crear_serie_interanual(
    df,
    base_monetaria,
    crec_base_monetaria
)

if len(rango) == 2:
    inicio = pd.to_datetime(rango[0])
    fin = pd.to_datetime(rango[1])
    df = df[(df["fecha"] >= inicio) & (df["fecha"] <= fin)]

st.sidebar.markdown("---")
st.sidebar.metric("Variables disponibles", len(df.columns) - 1)

if not df.empty:
    st.sidebar.metric("Última fecha visible", df["fecha"].max().strftime("%d/%m/%Y"))
else:
    st.sidebar.metric("Última fecha visible", "Sin dato")

# Control de verificación específico para el tipo de cambio oficial.
tc_fuente_valor, tc_fuente_fecha = ultimo_valor(df_original, tc_oficial)
if tc_fuente_valor is not None and tc_fuente_fecha is not None:
    st.sidebar.success(
        "TC oficial leído del Excel: "
        f"{formato_numero(tc_fuente_valor)} Bs/$us · "
        f"{tc_fuente_fecha.strftime('%d/%m/%Y')}"
    )
else:
    st.sidebar.warning("No se detectó un dato de tipo de cambio oficial en la fuente.")

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
        Sistema de Información Económica de Bolivia
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
        Indicadores macroeconómicos, financieros, fiscales, externos y sociales para el seguimiento de la coyuntura nacional.
    </p>
    """, unsafe_allow_html=True)

st.markdown("---")


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
    # FILA 1: INDICADORES MACRO PRINCIPALES
    # =========================

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi(df, "PIB a precios de mercado", pib_pm, "%", tipo="ultimo", delta_tipo="pp")

    with c2:
        kpi(df, "Inflación interanual", inflacion_12m, "%", tipo="ultimo", delta_tipo="pp")

    with c3:
        kpi(df, "RIN", rin, "MM $us")

    with c4:
        kpi(
            df,
            "Tipo de cambio oficial",
            tc_oficial,
            "Bs/$us",
            delta_tipo="ninguno"
        )

    # =========================
    # FILA 2: MONETARIO Y FINANCIERO
    # =========================

    c5, c6, c7, c8 = st.columns(4)

    with c5:
        kpi(df, "Base monetaria", base_monetaria, "MM Bs")

    with c6:
        kpi(df, "Crédito privado", credito_privado, "MM $us")

    with c7:
        kpi(df, "Depósitos", depositos, "MM $us")

    with c8:
        kpi(df, "Saldo comercial", saldo_comercial, "MM $us")

    # =========================
    # FILA 3: REAL, FISCAL Y SOCIAL
    # =========================

    c9, c10, c11, c12 = st.columns(4)

    with c9:
        kpi(df, "Tasa premio reporto MN", tasa_reporto_mn, "%", tipo="ultimo", delta_tipo="pp")

    with c10:
        kpi(df, "Resultado Global SPNF", resultado_global_spnf, "MM Bs", tipo="acumulado", delta_tipo="acumulado")

    with c11:
        kpi(df, "Incidencia de pobreza", pobreza_bolivia, "%", tipo="ultimo", delta_tipo="pp")

    with c12:
        kpi(df, "Tasa de desocupación", desocupacion_nacional, "%", tipo="ultimo", delta_tipo="pp")

# =========================
# TAB 2: INFLACIÓN
# =========================

with tab2:
    st.subheader("🔥 Sector precios - Inflación")

    titulo, mensaje, nivel = alerta_precios(df)
    alerta_sector(titulo, mensaje, nivel)

    c1, c2, c3 = st.columns(3)

    with c1:
        kpi(df, "Inflación interanual", inflacion_12m, "%", tipo="ultimo", delta_tipo="pp")

    with c2:
        kpi(df, "Inflación mensual", inflacion_mensual, "%", tipo="ultimo", delta_tipo="pp")

    with c3:
        kpi(df, "Inflación acumulada", inflacion_acumulada, "%", tipo="ultimo", delta_tipo="pp")

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

    # Primera fila: posición externa y mercado cambiario
    c1, c2, c3 = st.columns(3)

    with c1:
        kpi(df, "RIN", rin, "MM $us")

    with c2:
        kpi(
            df,
            "Tipo de cambio referencial",
            tc_venta,
            "Bs/$us",
            delta_tipo="ninguno"
        )

    with c3:
        kpi(
            df,
            "Tipo de cambio oficial",
            tc_oficial,
            "Bs/$us",
            delta_tipo="ninguno"
        )

    # Segunda fila: comercio exterior
    c4, c5, c6 = st.columns(3)

    with c4:
        kpi(
            df,
            "Exportaciones",
            exportaciones_valor,
            "MM $us",
            tipo="acumulado",
            delta_tipo="acumulado"
        )

    with c5:
        kpi(
            df,
            "Importaciones",
            importaciones_valor,
            "MM $us",
            tipo="acumulado",
            delta_tipo="acumulado"
        )

    with c6:
        kpi(
            df,
            "Saldo comercial",
            saldo_comercial,
            "MM $us",
            tipo="acumulado",
            delta_tipo="acumulado"
        )

    # Tercera fila: composición de las reservas
    c7, c8, c9 = st.columns(3)

    with c7:
        kpi(df, "Divisas", divisas, "MM $us")

    with c8:
        kpi(df, "Oro", oro, "MM $us")

    with c9:
        kpi(df, "Recursos alta liquidez", recursos_alta_liquidez, "MM $us")

    st.markdown("---")

    a, b = st.columns(2)

    with a:
        grafico_linea(df, rin, "Reservas internacionales netas", "MM $us")

    with b:
        grafico_lineas_multiples(
            df,
            [tc_venta, tc_oficial],
            "Tipo de cambio referencial y oficial",
            "Bs/$us"
        )

    c, d = st.columns(2)
    
    with c:
        grafico_doble_eje(
            df=df,
            col_izq=exportaciones_peso,
            col_der=exportaciones_valor,
            titulo="Exportaciones: valor y peso neto",
            nombre_izq="Peso neto",
            nombre_der="Valor exportado",
            titulo_eje_izq="Peso neto en toneladas",
            titulo_eje_der="Millones de dólares",
            unidad_izq="toneladas",
            unidad_der="MM $us"
        )
    
    with d:
        grafico_doble_eje(
            df=df,
            col_izq=importaciones_peso,
            col_der=importaciones_valor,
            titulo="Importaciones: valor CIF y peso bruto",
            nombre_izq="Peso bruto",
            nombre_der="Valor CIF",
            titulo_eje_izq="Peso bruto en toneladas",
            titulo_eje_der="Millones de dólares",
            unidad_izq="toneladas",
            unidad_der="MM $us"
        )

# =========================
# TAB 4: MONETARIO
# =========================

with tab4:
    st.subheader("💵 Sector monetario")

    titulo, mensaje, nivel = alerta_monetario(df)
    alerta_sector(titulo, mensaje, nivel)

    # =========================
    # KPIs MONETARIOS
    # =========================

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi(df, "Base monetaria", base_monetaria, "MM Bs")

    with c2:
        kpi(df, "M'1", m1, "MM Bs")

    with c3:
        kpi(df, "M'2", m2, "MM Bs")

    with c4:
        kpi(df, "M'3", m3, "MM Bs")

    c5, c6 = st.columns(2)

    with c5:
        kpi(df, "Tasa premio reporto MN", tasa_reporto_mn, "%", tipo="ultimo", delta_tipo="pp")

    with c6:
        kpi(df, "Títulos BCB", titulos_bcb_usd, "MM $us")

    st.markdown("---")

    # =========================
    # GRÁFICOS MONETARIOS PRINCIPALES
    # =========================

    a, b = st.columns(2)

    with a:
        grafico_doble_eje(
            df=df,
            col_izq=crec_base_monetaria,
            col_der=base_monetaria,
            titulo="Base monetaria y crecimiento interanual",
            nombre_izq="Crecimiento interanual Base monetaria",
            nombre_der="Base monetaria",
            titulo_eje_izq="Crecimiento interanual (%)",
            titulo_eje_der="Base monetaria (millones de Bs)",
            unidad_izq="%",
            unidad_der="MM Bs",
            sombra_izq=False,
            sombra_der=False
        )

    with b:
        grafico_lineas_multiples(
            df,
            [m1, m2, m3],
            "Agregados monetarios: M'1, M'2 y M'3",
            "Millones de Bs"
        )

    # =========================
    # NUEVOS GRÁFICOS MONETARIOS
    # =========================

    c, d = st.columns(2)

    with c:
        grafico_linea(
            df,
            tasa_reporto_mn,
            "Tasas premio de reporto del BCB en moneda nacional",
            "%"
        )

    with d:
        grafico_linea(
            df,
            titulos_bcb_usd,
            "Saldo de títulos del Banco Central de Bolivia",
            "Millones de $us"
        )

# =========================
# TAB 5: FINANCIERO
# =========================

with tab5:
    st.subheader("🏦 Sector financiero")

    titulo, mensaje, nivel = alerta_financiero(df)
    alerta_sector(titulo, mensaje, nivel)

    # =========================
    # KPIs FINANCIEROS
    # =========================

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi(df, "Crédito privado", credito_privado, "MM $us")

    with c2:
        kpi(df, "Depósitos", depositos, "MM $us")

    with c3:
        kpi(df, "Bolivianización depósitos", bol_dep, "%", tipo="ultimo", delta_tipo="pp")

    with c4:
        kpi(df, "Bolivianización créditos", bol_cred, "%", tipo="ultimo", delta_tipo="pp")

    c5, c6, c7 = st.columns(3)

    with c5:
        kpi(df, "Encaje constituido", encaje_constituido, "MM $us")

    with c6:
        kpi(df, "Excedente encaje efectivo", excedente_encaje_efectivo, "MM $us")

    with c7:
        kpi(df, "Excedente encaje ME", excedente_encaje_me, "MM $us")

    st.markdown("---")

    # =========================
    # LIQUIDEZ Y ENCAJE
    # =========================

    a, b = st.columns(2)

    with a:
        grafico_linea(
            df,
            encaje_constituido,
            "Encaje constituido por el sistema financiero",
            "Millones de $us"
        )

    with b:
        grafico_lineas_multiples(
            df,
            [
                excedente_encaje_efectivo,
                excedente_encaje_me
            ],
            "Liquidez del sistema financiero: excedente de encaje en el BCB",
            "Millones de $us"
        )

    st.markdown("---")

    # =========================
    # CRÉDITOS, DEPÓSITOS Y BOLIVIANIZACIÓN
    # =========================

    c, d = st.columns(2)

    with c:
        grafico_lineas_multiples(
            df,
            [credito_privado, depositos],
            "Crédito y depósitos del sistema financiero",
            "Millones de $us"
        )

    with d:
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
        kpi(df, "PIB a precios de mercado", pib_pm, "%")

    with c2:
        kpi(df, "Consumo de hogares", consumo_hogares, "%")

    with c3:
        kpi(df, "Formación bruta de capital", formacion_capital, "%")

    c4, c5, c6 = st.columns(3)

    with c4:
        kpi(df, "Consumo público", consumo_publico, "%")

    with c5:
        kpi(df, "Exportaciones", expo_bienes_servicios, "%")

    with c6:
        kpi(df, "Importaciones", impo_bienes_servicios, "%")

    st.markdown("---")

    a, b = st.columns(2)

    with a:
        grafico_linea(
            df,
            pib_pm,
            "PIB a precios de mercado",
            "En porcentaje"
        )

    with b:
        grafico_lineas_multiples(
            df,
            [consumo_hogares, consumo_publico, formacion_capital],
            "Demanda interna: consumo e inversión",
            "En porcentaje"
        )

    c, d = st.columns(2)

    with c:
        grafico_lineas_multiples(
            df,
            [expo_bienes_servicios, impo_bienes_servicios],
            "Sector externo real: exportaciones e importaciones",
            "En porcentaje"
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
            "En porcentaje"
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
        kpi(df, "Ingresos Totales SPNF", ingresos_totales_spnf, "MM Bs", tipo="acumulado", delta_tipo="acumulado")

    with c2:
        kpi(df, "Egresos Totales SPNF", egresos_totales_spnf, "MM Bs", tipo="acumulado", delta_tipo="acumulado")

    with c3:
        kpi(df, "Resultado Corriente SPNF", resultado_corriente_spnf, "MM Bs", tipo="acumulado", delta_tipo="acumulado")

    with c4:
        kpi(df, "Resultado Global SPNF", resultado_global_spnf, "MM Bs", tipo="acumulado", delta_tipo="acumulado")

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
        kpi(df, "Incidencia de pobreza", pobreza_bolivia, "%", tipo="ultimo", delta_tipo="pp")

    with c2:
        kpi(df, "Índice de GINI", gini_bolivia, "", tipo="ultimo", delta_tipo="pp")

    with c3:
        kpi(df, "Tasa de desocupación nacional", desocupacion_nacional, "%", tipo="ultimo", delta_tipo="pp")

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
