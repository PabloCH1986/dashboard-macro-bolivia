# ============================================================
# CENGOB MACRO MONITOR V3
# Sistema de Información Económica de Bolivia
# Dashboard ejecutivo, dinámico, extensible y automático
# ============================================================

import os
import io
import json
import time
import uuid
import base64
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

# Auto-refresh opcional. Si no está instalado, el dashboard sigue funcionando.
try:
    from streamlit_autorefresh import st_autorefresh
except Exception:
    st_autorefresh = None


# ============================================================
# 1. CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Sistema de Información Económica de Bolivia - CENGOB",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded",
)

SHEET_NAME = "data"
DRIVE_CACHE_TTL = 60
AUTO_REFRESH_MS = 60_000
LOCAL_FALLBACK = "Info(4).xlsx"

CENGOB_GREEN = "#0B3B36"
CENGOB_GREEN_2 = "#0F5149"
CENGOB_GOLD = "#C9A227"
CENGOB_GOLD_2 = "#E3B341"
BG = "#EEF2F5"
CARD = "#FFFFFF"
TEXT = "#0F172A"
TEXT_2 = "#334155"
MUTED = "#64748B"
BORDER = "#D8E0E7"
GRID = "rgba(15,23,42,0.10)"
RED = "#B91C1C"
AMBER = "#B45309"
GREEN = "#15803D"
BLUE = "#1D4ED8"
GRAY = "#64748B"

PALETTE = [
    CENGOB_GREEN,
    CENGOB_GOLD,
    "#2563EB",
    "#C2410C",
    "#7C3AED",
    "#15803D",
    "#475569",
    "#0E7490",
]


# ============================================================
# 2. CSS - CONTRASTE, LEGIBILIDAD Y RESPONSIVE
# ============================================================

st.markdown(
    f"""
<style>
:root {{
    --cengob-green: {CENGOB_GREEN};
    --cengob-gold: {CENGOB_GOLD};
    --bg: {BG};
    --text: {TEXT};
    --text2: {TEXT_2};
    --border: {BORDER};
}}

html, body, [class*="css"] {{
    font-family: "Inter", "Segoe UI", Arial, sans-serif;
}}

.stApp {{
    background: {BG};
    color: {TEXT};
}}

[data-testid="stHeader"] {{
    background: {CENGOB_GREEN};
}}

[data-testid="stToolbar"] {{
    right: 1rem;
}}

[data-testid="stSidebar"] {{
    background: #F8FAFC;
    border-right: 1px solid {BORDER};
}}

[data-testid="stSidebar"] * {{
    color: {TEXT} !important;
}}

[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {{
    color: {CENGOB_GREEN} !important;
}}

h1, h2, h3, h4, h5, h6 {{
    color: {CENGOB_GREEN} !important;
    font-weight: 800 !important;
    letter-spacing: -0.015em;
}}

p, label, span {{
    color: inherit;
}}

/* Radio de navegación */
[data-testid="stSidebar"] div[role="radiogroup"] label {{
    background: #FFFFFF;
    border: 1px solid {BORDER};
    border-radius: 11px;
    padding: 9px 10px;
    margin-bottom: 5px;
    transition: 0.15s ease;
}}

[data-testid="stSidebar"] div[role="radiogroup"] label:hover {{
    border-color: {CENGOB_GOLD};
    background: #FFFCF1;
}}

/* Métricas nativas: por si se usan en sidebar */
[data-testid="stMetric"] {{
    background: #FFFFFF !important;
    border: 1px solid {BORDER} !important;
    border-left: 5px solid {CENGOB_GOLD} !important;
    border-radius: 14px !important;
    padding: 13px 14px !important;
    box-shadow: 0 3px 10px rgba(15, 23, 42, 0.05);
}}

[data-testid="stMetricLabel"] p,
[data-testid="stMetricLabel"] {{
    color: {TEXT_2} !important;
    font-size: 13px !important;
    font-weight: 700 !important;
    line-height: 1.25 !important;
    white-space: normal !important;
}}

[data-testid="stMetricValue"],
[data-testid="stMetricValue"] div {{
    color: {CENGOB_GREEN} !important;
    font-size: clamp(21px, 1.7vw, 31px) !important;
    font-weight: 850 !important;
    line-height: 1.05 !important;
}}

[data-testid="stMetricDelta"] {{
    font-weight: 700 !important;
}}

/* Inputs */
[data-baseweb="select"] > div,
[data-baseweb="input"] > div,
.stDateInput > div > div {{
    background: #FFFFFF !important;
    color: {TEXT} !important;
}}

/* Botones */
.stButton > button,
.stDownloadButton > button {{
    border-radius: 10px !important;
    font-weight: 750 !important;
}}

.stButton > button[kind="primary"] {{
    background: {CENGOB_GREEN} !important;
    color: #FFFFFF !important;
}}

/* Dataframes */
[data-testid="stDataFrame"] {{
    background: #FFFFFF;
    border: 1px solid {BORDER};
    border-radius: 14px;
    overflow: hidden;
}}

/* Plotly */
.js-plotly-plot {{
    border-radius: 14px;
    overflow: hidden;
}}

/* Alertas */
.stAlert {{
    border-radius: 12px;
}}

/* Expander */
[data-testid="stExpander"] {{
    background: #FFFFFF;
    border: 1px solid {BORDER};
    border-radius: 12px;
}}

/* Tarjetas ejecutivas */
.cengob-card {{
    background: #FFFFFF;
    border: 1px solid {BORDER};
    border-radius: 16px;
    padding: 16px 17px 14px 17px;
    box-shadow: 0 4px 14px rgba(15,23,42,0.055);
    min-height: 142px;
    height: 100%;
    overflow: hidden;
}}

.cengob-card:hover {{
    border-color: #C7D2DA;
    box-shadow: 0 7px 18px rgba(15,23,42,0.08);
}}

.kpi-title {{
    color: {TEXT_2} !important;
    font-size: 13px;
    font-weight: 800;
    line-height: 1.25;
    min-height: 33px;
    margin-bottom: 6px;
    overflow-wrap: anywhere;
}}

.kpi-value {{
    color: {CENGOB_GREEN} !important;
    font-size: clamp(23px, 2vw, 34px);
    font-weight: 900;
    line-height: 1.05;
    letter-spacing: -0.025em;
    overflow-wrap: anywhere;
}}

.kpi-unit {{
    color: {MUTED} !important;
    font-size: 12px;
    font-weight: 700;
}}

.kpi-delta {{
    font-size: 12.5px;
    font-weight: 800;
    margin-top: 8px;
    line-height: 1.25;
}}

.kpi-date {{
    color: {MUTED} !important;
    font-size: 11.5px;
    margin-top: 8px;
    font-weight: 650;
}}

.section-card {{
    background: #FFFFFF;
    border: 1px solid {BORDER};
    border-radius: 16px;
    padding: 16px 18px;
    box-shadow: 0 4px 14px rgba(15,23,42,0.045);
}}

.hero {{
    background: linear-gradient(120deg, {CENGOB_GREEN} 0%, {CENGOB_GREEN_2} 100%);
    border-radius: 18px;
    padding: 20px 24px;
    color: #FFFFFF !important;
    box-shadow: 0 7px 22px rgba(11,59,54,0.17);
    margin-bottom: 14px;
}}

.hero h1, .hero h2, .hero p, .hero div, .hero span {{
    color: #FFFFFF !important;
}}

.hero .gold {{
    color: #F2CC5C !important;
}}

.badge {{
    display: inline-block;
    border-radius: 999px;
    padding: 3px 8px;
    font-size: 11px;
    font-weight: 800;
}}

.signal-box {{
    border-radius: 14px;
    padding: 14px 15px;
    min-height: 140px;
    border: 1px solid {BORDER};
    background: #FFFFFF;
}}

.signal-title {{
    font-weight: 900;
    font-size: 14px;
    margin-bottom: 8px;
}}

.small-note {{
    color: {MUTED} !important;
    font-size: 12px;
    line-height: 1.4;
}}

@media (max-width: 1100px) {{
    .cengob-card {{ min-height: 132px; }}
    .kpi-value {{ font-size: 25px; }}
}}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# 3. CARGA AUTOMÁTICA DESDE GOOGLE DRIVE
# ============================================================

GOOGLE_SHEETS_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.readonly"


def _leer_configuracion_drive():
    try:
        drive_cfg = dict(st.secrets["drive"])
    except (FileNotFoundError, KeyError):
        drive_cfg = {}

    file_id = str(
        drive_cfg.get("file_id")
        or os.getenv("GOOGLE_DRIVE_FILE_ID", "")
    ).strip()

    credenciales = None

    try:
        credenciales = dict(st.secrets["gcp_service_account"])
    except (FileNotFoundError, KeyError):
        pass

    if not credenciales and drive_cfg.get("service_account_json"):
        try:
            credenciales = json.loads(drive_cfg["service_account_json"])
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                "drive.service_account_json no contiene un JSON válido."
            ) from exc

    return file_id, credenciales


def _parece_excel(contenido):
    return (
        contenido.startswith(b"PK")
        or contenido.startswith(bytes.fromhex("D0CF11E0"))
    )


def _descargar_drive_privado(file_id, credenciales_info):
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
            fields="id,name,mimeType,modifiedTime",
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

    if not _parece_excel(contenido):
        raise RuntimeError("Google Drive no devolvió un Excel válido.")

    return contenido, metadatos.get("name", "Google Drive"), metadatos.get("modifiedTime")


def _descargar_drive_publico(file_id):
    cache_buster = int(time.time())
    urls = [
        f"https://drive.google.com/uc?export=download&id={file_id}&_={cache_buster}",
        f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx&_={cache_buster}",
    ]

    ultimo_error = None

    for url in urls:
        try:
            r = requests.get(
                url,
                timeout=90,
                allow_redirects=True,
                headers={
                    "User-Agent": "Mozilla/5.0",
                    "Cache-Control": "no-cache",
                    "Pragma": "no-cache",
                },
            )
            r.raise_for_status()
            if _parece_excel(r.content):
                return r.content, "Google Drive público", None
            ultimo_error = RuntimeError("Drive devolvió HTML en lugar de Excel.")
        except requests.RequestException as exc:
            ultimo_error = exc

    raise RuntimeError(
        "No se pudo descargar el Excel de Drive. Revise permisos o use cuenta de servicio."
    ) from ultimo_error


def _leer_excel_bytes(contenido_excel, nombre_fuente="fuente"):
    raw = pd.read_excel(
        io.BytesIO(contenido_excel),
        sheet_name=SHEET_NAME,
        header=None,
    )

    if raw.shape[0] < 3:
        raise ValueError("La hoja requiere dos filas de encabezado y al menos una fila de datos.")

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
    data["fecha"] = data["fecha"].dt.normalize()

    # Mantener fecha real del último mes si está incompleto; cerrar meses anteriores.
    if data.empty:
        raise ValueError("No se detectaron fechas válidas.")

    fecha_max = data["fecha"].max()
    ultimo_mes = fecha_max.to_period("M")
    ultimo_mes_cerrado = fecha_max == fecha_max + pd.offsets.MonthEnd(0)

    def ajustar_fecha(fecha):
        if pd.isna(fecha):
            return fecha
        if not ultimo_mes_cerrado and fecha.to_period("M") == ultimo_mes:
            return fecha
        return fecha + pd.offsets.MonthEnd(0)

    data["fecha"] = data["fecha"].apply(ajustar_fecha)
    data = data.dropna(axis=1, how="all")

    for col in data.columns:
        if col != "fecha":
            data[col] = pd.to_numeric(data[col], errors="coerce")

    data = data.sort_values("fecha").reset_index(drop=True)
    data.attrs["fuente_drive"] = nombre_fuente
    return data


@st.cache_data(ttl=DRIVE_CACHE_TTL, show_spinner=False)
def descargar_excel_drive():
    file_id, credenciales = _leer_configuracion_drive()

    if file_id:
        if credenciales:
            try:
                return _descargar_drive_privado(file_id, credenciales)
            except HttpError as exc:
                raise RuntimeError(
                    "Drive rechazó el acceso. Comparta el archivo con la cuenta de servicio."
                ) from exc
        return _descargar_drive_publico(file_id)

    # Fallback local útil para desarrollo.
    if os.path.exists(LOCAL_FALLBACK):
        with open(LOCAL_FALLBACK, "rb") as f:
            return f.read(), LOCAL_FALLBACK, None

    raise RuntimeError(
        "No se configuró GOOGLE_DRIVE_FILE_ID / [drive] file_id y no existe Info(4).xlsx local."
    )


@st.cache_data(ttl=DRIVE_CACHE_TTL, show_spinner="Actualizando información económica...")
def cargar_datos():
    contenido, nombre, modified_time = descargar_excel_drive()
    data = _leer_excel_bytes(contenido, nombre)
    data.attrs["modified_time"] = modified_time
    return data


try:
    df_original = cargar_datos()
except Exception as exc:
    st.error("No se pudo cargar la base económica.")
    st.exception(exc)
    st.stop()


# ============================================================
# 4. FUNCIONES GENERALES DE DATOS
# ============================================================


def normalizar_texto(texto):
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.split())


def buscar_columna(texto):
    objetivo = normalizar_texto(texto)
    columnas = [c for c in df_original.columns if c != "fecha"]

    # 1) coincidencia exacta
    for col in columnas:
        if normalizar_texto(col) == objetivo:
            return col

    # 2) empieza por
    for col in columnas:
        if normalizar_texto(col).startswith(objetivo):
            return col

    # 3) contiene
    for col in columnas:
        if objetivo in normalizar_texto(col):
            return col

    return None


def buscar_columna_multiple(opciones):
    for op in opciones:
        col = buscar_columna(op)
        if col is not None:
            return col
    return None


def ultimo_valor(df, col):
    if col is None or col not in df.columns:
        return None, None
    s = df[["fecha", col]].dropna().sort_values("fecha")
    if s.empty:
        return None, None
    r = s.iloc[-1]
    return float(r[col]), r["fecha"]


def ultima_fecha_serie(df, col):
    _, f = ultimo_valor(df, col)
    return f


def valor_previo_12m(df, col, fecha_actual=None):
    if col is None or col not in df.columns:
        return None, None
    s = df[["fecha", col]].dropna().sort_values("fecha")
    if s.empty:
        return None, None
    if fecha_actual is None:
        fecha_actual = s["fecha"].max()
    objetivo = pd.Timestamp(fecha_actual) - pd.DateOffset(years=1)
    prev = s[s["fecha"] <= objetivo]
    if prev.empty:
        return None, None
    r = prev.iloc[-1]
    return float(r[col]), r["fecha"]


def variacion_interanual(df, col):
    actual, fecha = ultimo_valor(df, col)
    if actual is None:
        return None
    base, _ = valor_previo_12m(df, col, fecha)
    if base is None or base == 0:
        return None
    return (actual / base - 1) * 100


def variacion_interanual_pp(df, col):
    actual, fecha = ultimo_valor(df, col)
    if actual is None:
        return None
    base, _ = valor_previo_12m(df, col, fecha)
    if base is None:
        return None
    return actual - base


def valor_acumulado_anual(df, col):
    if col is None or col not in df.columns:
        return None, None
    s = df[["fecha", col]].dropna().sort_values("fecha")
    if s.empty:
        return None, None
    f = s["fecha"].max()
    actual = s[(s["fecha"].dt.year == f.year) & (s["fecha"] <= f)][col].sum()
    return float(actual), f


def variacion_acumulada_interanual(df, col):
    if col is None or col not in df.columns:
        return None
    s = df[["fecha", col]].dropna().sort_values("fecha")
    if s.empty:
        return None
    f = s["fecha"].max()
    y, m = f.year, f.month
    a = s[(s["fecha"].dt.year == y) & (s["fecha"].dt.month <= m)][col].sum()
    b = s[(s["fecha"].dt.year == y - 1) & (s["fecha"].dt.month <= m)][col].sum()
    if pd.isna(b) or b == 0:
        return None
    return (a / b - 1) * 100


def acumulado_a_mismo_mes(df, col, fecha_corte):
    if col is None or col not in df.columns or fecha_corte is None:
        return None
    s = df[["fecha", col]].dropna()
    y, m = pd.Timestamp(fecha_corte).year, pd.Timestamp(fecha_corte).month
    q = s[(s["fecha"].dt.year == y) & (s["fecha"].dt.month <= m)][col]
    return float(q.sum()) if not q.empty else None


def formato_numero(x, dec=2):
    if x is None or pd.isna(x):
        return "Sin dato"
    t = f"{x:,.{dec}f}"
    return t.replace(",", "X").replace(".", ",").replace("X", ".")


def formato_fecha(f):
    if f is None or pd.isna(f):
        return "Sin fecha"
    return pd.Timestamp(f).strftime("%d/%m/%Y")


def frecuencia_serie(df, col):
    if col is None or col not in df.columns:
        return "Sin dato"
    s = df[["fecha", col]].dropna().drop_duplicates("fecha").sort_values("fecha")
    if len(s) < 3:
        return "Irregular"
    dif = s["fecha"].diff().dt.days.dropna()
    med = dif.median()
    if med <= 3:
        return "Diaria"
    if med <= 40:
        return "Mensual"
    if med <= 110:
        return "Trimestral"
    if med <= 220:
        return "Semestral"
    return "Anual"


def estado_frescura(fecha, fecha_base=None):
    if fecha is None:
        return "Sin dato", GRAY
    if fecha_base is None:
        fecha_base = df_original["fecha"].max()
    dias = max(0, (pd.Timestamp(fecha_base) - pd.Timestamp(fecha)).days)
    if dias <= 45:
        return "Actual", GREEN
    if dias <= 120:
        return "Rezago normal", AMBER
    return "Rezago alto", RED


def serie_mensual(df, col, how="last"):
    if col is None or col not in df.columns:
        return pd.Series(dtype="float64")
    s = df[["fecha", col]].dropna().drop_duplicates("fecha", keep="last")
    if s.empty:
        return pd.Series(dtype="float64")
    s = s.set_index("fecha")[col].sort_index()
    if how == "sum":
        return s.resample("ME").sum(min_count=1)
    if how == "mean":
        return s.resample("ME").mean()
    return s.resample("ME").last()


def transformar_serie(s, tipo):
    if s is None or len(s) == 0:
        return s
    s = s.astype(float)
    if tipo == "Nivel":
        return s
    if tipo == "Variación interanual (%)":
        return s.pct_change(12) * 100
    if tipo == "Variación mensual (%)":
        return s.pct_change() * 100
    if tipo == "Cambio 12 meses (p.p./unidades)":
        return s.diff(12)
    if tipo == "Índice base 100":
        no_na = s.dropna()
        if no_na.empty or no_na.iloc[0] == 0:
            return s * np.nan
        return s / no_na.iloc[0] * 100
    if tipo == "Z-score":
        std = s.std()
        return (s - s.mean()) / std if std not in [0, np.nan] and not pd.isna(std) else s * np.nan
    return s


def alinear_mensual(df, col1, col2, t1="Nivel", t2="Nivel", how1="last", how2="last"):
    s1 = transformar_serie(serie_mensual(df, col1, how1), t1).rename("x")
    s2 = transformar_serie(serie_mensual(df, col2, how2), t2).rename("y")
    z = pd.concat([s1, s2], axis=1).dropna()
    return z


def ultimo_par_comun(df, col1, col2):
    if col1 is None or col2 is None or col1 not in df.columns or col2 not in df.columns:
        return None, None, None
    s = df[["fecha", col1, col2]].dropna().sort_values("fecha")
    if s.empty:
        return None, None, None
    r = s.iloc[-1]
    return float(r[col1]), float(r[col2]), r["fecha"]


def correlacion_segura(z):
    if z is None or len(z) < 6:
        return None
    c = z["x"].corr(z["y"])
    return None if pd.isna(c) else float(c)


def color_delta(delta, mejora_si_sube=True):
    if delta is None or pd.isna(delta):
        return MUTED
    bueno = delta >= 0 if mejora_si_sube else delta <= 0
    return GREEN if bueno else RED


# ============================================================
# 5. COMPONENTES VISUALES
# ============================================================


def tarjeta_kpi(
    titulo,
    valor,
    unidad="",
    delta=None,
    delta_texto=None,
    fecha=None,
    mejora_si_sube=True,
    dec=2,
    nota=None,
):
    if valor is None or pd.isna(valor):
        valor_txt = "Sin dato"
    else:
        valor_txt = formato_numero(valor, dec)

    if delta_texto is None and delta is not None and not pd.isna(delta):
        delta_texto = f"{formato_numero(delta, 1)}%"

    delta_html = ""
    if delta_texto:
        col = color_delta(delta, mejora_si_sube)
        delta_html = f'<div class="kpi-delta" style="color:{col} !important;">{delta_texto}</div>'

    nota_html = f'<div class="small-note">{nota}</div>' if nota else ""

    st.markdown(
        f"""
        <div class="cengob-card">
            <div class="kpi-title">{titulo}</div>
            <div class="kpi-value">{valor_txt} <span class="kpi-unit">{unidad}</span></div>
            {delta_html}
            <div class="kpi-date">Último dato: {formato_fecha(fecha)}</div>
            {nota_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi_desde_serie(
    df,
    titulo,
    col,
    unidad="",
    tipo="ultimo",
    delta_tipo="interanual",
    mejora_si_sube=True,
    dec=2,
    nota=None,
):
    if tipo == "acumulado":
        valor, fecha = valor_acumulado_anual(df, col)
    else:
        valor, fecha = ultimo_valor(df, col)

    delta = None
    texto = None

    if delta_tipo == "interanual":
        delta = variacion_interanual(df, col)
        if delta is not None:
            texto = f"{formato_numero(delta, 1)}% interanual"
    elif delta_tipo == "pp":
        delta = variacion_interanual_pp(df, col)
        if delta is not None:
            texto = f"{formato_numero(delta, 1)} p.p. interanual"
    elif delta_tipo in ["acumulado", "acumulado_pct"]:
        delta = variacion_acumulada_interanual(df, col)
        if delta is not None:
            texto = f"{formato_numero(delta, 1)}% acum. interanual"
    elif delta_tipo == "ninguno":
        delta = None

    tarjeta_kpi(
        titulo=titulo,
        valor=valor,
        unidad=unidad,
        delta=delta,
        delta_texto=texto,
        fecha=fecha,
        mejora_si_sube=mejora_si_sube,
        dec=dec,
        nota=nota,
    )


def base_layout(fig, titulo, height=390, legend=True):
    fig.update_layout(
        title=dict(text=titulo, font=dict(size=18, color=CENGOB_GREEN, family="Arial Black"), x=0.01),
        height=height,
        paper_bgcolor=CARD,
        plot_bgcolor=CARD,
        font=dict(color=TEXT, size=12),
        margin=dict(l=45, r=35, t=64, b=45),
        hovermode="x unified",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
            font=dict(size=11, color=TEXT),
            bgcolor="rgba(255,255,255,0.82)",
        ) if legend else dict(visible=False),
    )
    fig.update_xaxes(
        showgrid=False,
        linecolor=BORDER,
        tickfont=dict(color=TEXT_2, size=11),
        title_font=dict(color=TEXT_2),
    )
    fig.update_yaxes(
        gridcolor=GRID,
        zerolinecolor="rgba(11,59,54,0.22)",
        tickfont=dict(color=TEXT_2, size=11),
        title_font=dict(color=TEXT_2),
    )
    return fig


def mostrar_fig(fig, key=None):
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displaylogo": False, "responsive": True},
        key=key or f"plot_{uuid.uuid4()}",
    )


def grafico_linea(df, col, titulo, unidad="", color=CENGOB_GREEN, height=390, rango=False):
    if col is None or col not in df.columns:
        st.info(f"No se encontró la serie: {titulo}")
        return
    s = df[["fecha", col]].dropna().sort_values("fecha")
    if s.empty:
        st.info(f"Sin datos para: {titulo}")
        return

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=s["fecha"], y=s[col], name=titulo, mode="lines",
        line=dict(color=color, width=2.8),
        hovertemplate="%{x|%d/%m/%Y}<br>%{y:,.2f}<extra></extra>",
    ))
    base_layout(fig, titulo, height=height, legend=False)
    fig.update_yaxes(title_text=unidad)

    if rango:
        fig.update_xaxes(rangeselector=dict(
            buttons=[
                dict(count=1, label="1A", step="year", stepmode="backward"),
                dict(count=3, label="3A", step="year", stepmode="backward"),
                dict(count=5, label="5A", step="year", stepmode="backward"),
                dict(step="all", label="Todo"),
            ],
            bgcolor="#FFFFFF", activecolor=CENGOB_GOLD,
        ))
    mostrar_fig(fig)


def grafico_lineas_multiples(df, cols, titulo, unidad="", nombres=None, height=390):
    cols = [c for c in cols if c is not None and c in df.columns]
    if not cols:
        st.info(f"No hay series disponibles para: {titulo}")
        return

    nombres = nombres or {}
    fig = go.Figure()
    for i, col in enumerate(cols):
        s = df[["fecha", col]].dropna().sort_values("fecha")
        if s.empty:
            continue
        fig.add_trace(go.Scatter(
            x=s["fecha"], y=s[col], mode="lines",
            name=nombres.get(col, col),
            line=dict(width=2.6, color=PALETTE[i % len(PALETTE)]),
        ))

    base_layout(fig, titulo, height=height, legend=True)
    fig.update_yaxes(title_text=unidad)
    mostrar_fig(fig)


def grafico_doble_eje(
    df, col_izq, col_der, titulo, nombre_izq, nombre_der,
    titulo_eje_izq="", titulo_eje_der="", unidad_izq="", unidad_der="", height=405,
):
    if (col_izq is None or col_izq not in df.columns) and (col_der is None or col_der not in df.columns):
        st.info(f"No hay series disponibles para: {titulo}")
        return

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    if col_izq is not None and col_izq in df.columns:
        s = df[["fecha", col_izq]].dropna().sort_values("fecha")
        fig.add_trace(go.Scatter(
            x=s["fecha"], y=s[col_izq], name=nombre_izq, mode="lines",
            line=dict(color=CENGOB_GREEN, width=2.7),
            hovertemplate=f"%{{x|%d/%m/%Y}}<br>{nombre_izq}: %{{y:,.2f}} {unidad_izq}<extra></extra>",
        ), secondary_y=False)

    if col_der is not None and col_der in df.columns:
        s = df[["fecha", col_der]].dropna().sort_values("fecha")
        fig.add_trace(go.Scatter(
            x=s["fecha"], y=s[col_der], name=nombre_der, mode="lines",
            line=dict(color=CENGOB_GOLD, width=2.7),
            hovertemplate=f"%{{x|%d/%m/%Y}}<br>{nombre_der}: %{{y:,.2f}} {unidad_der}<extra></extra>",
        ), secondary_y=True)

    base_layout(fig, titulo, height=height, legend=True)
    fig.update_yaxes(title_text=titulo_eje_izq, secondary_y=False, title_font=dict(color=CENGOB_GREEN))
    fig.update_yaxes(title_text=titulo_eje_der, secondary_y=True, title_font=dict(color=CENGOB_GOLD), showgrid=False)
    mostrar_fig(fig)


def grafico_cruce(
    df,
    col1,
    col2,
    titulo,
    nombre1,
    nombre2,
    transform1="Nivel",
    transform2="Nivel",
    how1="last",
    how2="last",
    eje1="",
    eje2="",
    height=390,
):
    z = alinear_mensual(df, col1, col2, transform1, transform2, how1, how2)
    if z.empty:
        st.info(f"No existen observaciones sincronizadas suficientes para: {titulo}")
        return

    corr = correlacion_segura(z)
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Scatter(
        x=z.index, y=z["x"], name=nombre1,
        mode="lines", line=dict(color=CENGOB_GREEN, width=2.6),
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=z.index, y=z["y"], name=nombre2,
        mode="lines", line=dict(color=CENGOB_GOLD, width=2.6),
    ), secondary_y=True)

    titulo2 = titulo + (f" · ρ={corr:.2f}" if corr is not None else "")
    base_layout(fig, titulo2, height=height, legend=True)
    fig.update_yaxes(title_text=eje1 or transform1, secondary_y=False, title_font=dict(color=CENGOB_GREEN))
    fig.update_yaxes(title_text=eje2 or transform2, secondary_y=True, title_font=dict(color=CENGOB_GOLD), showgrid=False)
    mostrar_fig(fig)


def grafico_scatter_cruce(df, col1, col2, titulo, nombre1, nombre2, t1="Nivel", t2="Nivel"):
    z = alinear_mensual(df, col1, col2, t1, t2)
    if len(z) < 6:
        st.info(f"Observaciones insuficientes para dispersión: {titulo}")
        return

    corr = correlacion_segura(z)
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=z["x"], y=z["y"], mode="markers",
        marker=dict(size=8, color=CENGOB_GREEN, opacity=0.75),
        text=z.index.strftime("%m/%Y"),
        hovertemplate="%{text}<br>X: %{x:,.2f}<br>Y: %{y:,.2f}<extra></extra>",
        name="Observaciones",
    ))

    try:
        coef = np.polyfit(z["x"], z["y"], 1)
        xs = np.linspace(z["x"].min(), z["x"].max(), 100)
        ys = coef[0] * xs + coef[1]
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="lines", name="Tendencia lineal",
            line=dict(color=CENGOB_GOLD, width=2),
        ))
    except Exception:
        pass

    base_layout(fig, f"{titulo} · ρ={corr:.2f}" if corr is not None else titulo, height=390, legend=True)
    fig.update_xaxes(title_text=f"{nombre1} · {t1}")
    fig.update_yaxes(title_text=f"{nombre2} · {t2}")
    mostrar_fig(fig)


def grafico_correlacion_movil(df, col1, col2, titulo, t1="Nivel", t2="Nivel", ventana=12):
    z = alinear_mensual(df, col1, col2, t1, t2)
    if len(z) < ventana + 2:
        st.info(f"Observaciones insuficientes para correlación móvil: {titulo}")
        return
    roll = z["x"].rolling(ventana).corr(z["y"])
    fig = go.Figure(go.Scatter(
        x=roll.index, y=roll.values, mode="lines",
        line=dict(color=CENGOB_GREEN, width=2.7), name="Correlación móvil",
    ))
    base_layout(fig, f"{titulo} · ventana {ventana} meses", height=350, legend=False)
    fig.update_yaxes(range=[-1.05, 1.05], title_text="Correlación")
    mostrar_fig(fig)


def tarjeta_riesgo(titulo, nivel, detalle=""):
    colores = {
        "Alto": ("#FEE2E2", RED),
        "Moderado": ("#FEF3C7", AMBER),
        "Bajo": ("#DCFCE7", GREEN),
        "Sin dato": ("#F1F5F9", GRAY),
    }
    fondo, color = colores.get(nivel, colores["Sin dato"])
    st.markdown(
        f"""
        <div style="background:{fondo}; border:1px solid {color}55; border-left:6px solid {color};
                    border-radius:15px; padding:16px; min-height:130px;">
            <div style="color:{TEXT_2}; font-size:13px; font-weight:800; min-height:32px;">{titulo}</div>
            <div style="color:{color}; font-size:28px; font-weight:900; margin-top:8px;">{nivel}</div>
            <div style="color:{TEXT_2}; font-size:11.5px; margin-top:7px; line-height:1.35;">{detalle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def caja_senal(titulo, items, tipo="ok"):
    estilos = {
        "ok": ("#F0FDF4", GREEN, "↗"),
        "danger": ("#FEF2F2", RED, "↘"),
        "watch": ("#FFFBEB", AMBER, "!"),
        "info": ("#EFF6FF", BLUE, "i"),
    }
    fondo, color, icono = estilos.get(tipo, estilos["info"])
    lis = "".join([f"<li style='margin-bottom:6px;'>{x}</li>" for x in items]) or "<li>Sin señales suficientes.</li>"
    st.markdown(
        f"""
        <div class="signal-box" style="background:{fondo}; border-color:{color}55;">
            <div class="signal-title" style="color:{color} !important;">{icono} {titulo}</div>
            <ul style="margin:0; padding-left:18px; color:{TEXT}; font-size:12.5px; line-height:1.35;">{lis}</ul>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# 6. VARIABLES DEL EXCEL
# ============================================================

# Sector real
igae = buscar_columna("IGAE (indice)")
pib_pm = buscar_columna("PIB a precios de mercado")
consumo_publico = buscar_columna("Gasto de consumo final de la administración pública")
consumo_hogares = buscar_columna("Gasto de consumo final de los hogares")
formacion_capital = buscar_columna("Formación bruta de capital")
expo_bienes_servicios = buscar_columna("Exportaciones de bienes y servicios")
impo_bienes_servicios = buscar_columna("importaciones bienes y servicios")

# Monetario / BCB
balance_bcb = buscar_columna("Balance del Banco Central de Bolivia")
base_monetaria = buscar_columna("Base monetaria")
billetes_publico = buscar_columna("Billetes y monedas en poder del público")
emision_monetaria = buscar_columna("Emisión monetaria")
reservas_bancarias_totales = buscar_columna("Reservas bancarias totales")
reservas_bancarias_mn = buscar_columna("Reservas bancarias moneda nacional")
reservas_bancarias_me = buscar_columna("Reservas bancarias moneda extranjera")
credito_interno_neto_bcb = buscar_columna("Crédito interno neto del BCB")
credito_neto_bcb_spnf = buscar_columna("Crédito neto del BCB al Sector Público No financiero")
creditos_bcb_spnf = buscar_columna("Créditos del BCB al SPNF")
depositos_spnf_bcb = buscar_columna("Depósitos del SPNF en el BCB")
credito_otros_sectores = buscar_columna("Crédito a otros sectores")
m1 = buscar_columna("M’1")
m2 = buscar_columna("M’2")
m3 = buscar_columna("M’3")
participacion_mn_agregados = buscar_columna("Participación de MN y UFV en Agregados Monetarios")
m1_ratio = buscar_columna("M1 / M’1")
m2_ratio = buscar_columna("M2 / M’2")
m3_ratio = buscar_columna("M3 / M’3")

titulos_tgn_usd = buscar_columna("Saldo de Títulos del Tesoro General de la Nación (millones de $us)")
titulos_bcb_usd = buscar_columna("Saldo de Títulos del Banco Central de Bolivia (millones de $us)")
financiamiento_corto_bcb = buscar_columna("Financiamiento de corto plazo al sistema financiero por parte del BCB")
creditos_liquidez_bcb = buscar_columna("Créditos de Liquidez c/garantía del encaje legal en títulos por parte del BCB (millones de $us)")
operaciones_reporto_eif = buscar_columna("Operaciones de Reportos de las EIF con el BCB (millones de $us)")

# Precios y tasas
inflacion_mensual = buscar_columna("Variación mensual inflacion total")
inflacion_acumulada = buscar_columna("Variación acumulada en el año")
inflacion_12m = buscar_columna("Variación a doce meses")
tasa_reporto_total = buscar_columna("Tasas premio de reporto del BCB")
tasa_reporto_mn = buscar_columna("Tasas premio de reporto del BCB en Moneda nacional")
tasa_reporto_me = buscar_columna("Tasas premio de reporto del BCB en Moneda extranjera")

# Sistema financiero
ahorro_sistema = buscar_columna("Ahorro en el Sistema Financiero")
bol_ahorro = buscar_columna("Bolivianización del ahorro del sistema financiero")
depositos = buscar_columna("Depósitos en entidades de intermediación financiera")
bol_dep = buscar_columna("Bolivianización Depósitos")
dep_vista = buscar_columna("Depósitos a la vista en entidades de intermediación financiera")
bol_dep_vista = buscar_columna("Bolivianización depósitos vista")
caja_ahorro = buscar_columna("Caja de ahorro en entidades de intermediación financiera")
bol_caja = buscar_columna("Bolivianización caja de ahorro")
dep_plazo = buscar_columna("Depósitos a Plazo en entidades de intermediación financiera")
bol_dep_plazo = buscar_columna("Bolivianización depósitos a plazo")
otros_depositos = buscar_columna("Otros Depósitos")
bol_otros_dep = buscar_columna("Bolivianización de otros depósitos")
encaje_constituido = buscar_columna("Encaje constituido por el sistema financiero")
encaje_efectivo_mn = buscar_columna("Encaje constituido por el sistema financiero en Efectivo MN")
encaje_titulos_mn = buscar_columna("Encaje constituido por el sistema financiero en Títulos MN")
encaje_efectivo_me = buscar_columna("Encaje constituido por el sistema financiero en Efectivo ME")
encaje_titulos_me = buscar_columna("Encaje constituido por el sistema financiero en Títulos ME")
excedente_encaje_efectivo = buscar_columna("Excedente de Encaje en el BCB del sistema financiero ( en efectivo")
excedente_encaje_mn = buscar_columna("Excedente de Encaje en el BCB del sistema financiero en MN")
excedente_encaje_me = buscar_columna("Excedente de Encaje en el BCB del sistema financiero en ME")
credito_privado = buscar_columna("Crédito del sistema financiero al sector privado")
bol_cred = buscar_columna("Bolivianización Créditos")

# Externo / RIN
reservas_brutas = buscar_columna("Reservas internacionales brutas del BCB")
deg = buscar_columna("DEG")
oro = buscar_columna("Oro")
oro_ley_tn = buscar_columna("d/c Oro según Ley N°1503 en Tn")
oro_ley = buscar_columna("d/c Oro según Ley N°1503")
recursos_alta_liquidez = buscar_columna("Recursos de Alta Liquidez")
oro_convertible = buscar_columna("Oro convertible en divisas")
divisas = buscar_columna("Divisas")
posicion_fmi = buscar_columna("Posición con el FMI")
rin = buscar_columna("Reservas internacionales netas del BCB")
activos_ext_eif = buscar_columna("Activos externos netos de entidades financieras de intermediación")
fondo_ral_me = buscar_columna("Activos externos netos de entidades financieras de intermediación d/c Fondo RAL en ME")
finpro = buscar_columna("Fondo para la Revolución Industrial Productiva")
fondos_proteccion = buscar_columna("Fondos de Protección")
fpieeh = buscar_columna("Fondo Para la Inversión Exploracion y Explotacion de Hidrocarburos")
total_activos_externos = buscar_columna("Total Activos Externos")

adjud_bolsin_sf = buscar_columna("Adjudicación de dólares en el Bolsín - Sistema Financiero")
adjud_bolsin_priv = buscar_columna("Adjudicación de dólares en el Bolsín - Sector Privado")
compras_dolares_bcb = buscar_columna("Compras de dólares al Sistema Financiero por el BCB")
transf_ext_a_sf_bcb = buscar_columna("Transferencias del exterior al Sistema Financiero a través del BCB")
transf_sf_al_ext_bcb = buscar_columna("Transferencias del Sistema Financiero al exterior a través del BCB")
transf_eif_al_ext = buscar_columna("Transferencias al exterior de las EIF a través del sistema financiero")
transf_ext_a_eif = buscar_columna("Transferencias del exterior de las EIF a través del sistema financiero")

# Tipo de cambio
# Histórico hasta mayo 2026 + columna oficial nueva desde junio 2026.
tc_oficial_historico = buscar_columna("Tipo de cambio de venta en el Bolsín (Bs/$us)")
tc_compra_bcb = buscar_columna("Tipo de cambio de compra en el BCB")
tc_oficial_nuevo = buscar_columna("Tipo de Cambio Oficial (Bs/USD)")
tc_venta = buscar_columna("Valor referencial de venta del dólar estadounidense")
tc_compra_ref = buscar_columna("Valor referencial de compra del dólar estadounidense")
tcr = buscar_columna("Índice de tipo de cambio real")
ufv_habil = buscar_columna("UFV (Bs/UFV día hábil")
ufv_mes = buscar_columna("UFV (Bs/UFV último día del mes)")

tc_oficial = None
if tc_oficial_nuevo is not None or tc_oficial_historico is not None:
    tc_oficial = "Tipo de cambio oficial consolidado"
    nueva = df_original[tc_oficial_nuevo] if tc_oficial_nuevo else pd.Series(index=df_original.index, dtype=float)
    vieja = df_original[tc_oficial_historico] if tc_oficial_historico else pd.Series(index=df_original.index, dtype=float)
    df_original[tc_oficial] = nueva.combine_first(vieja)

# Comercio exterior
saldo_comercial = buscar_columna("Saldo Comercial")
exportaciones_valor = buscar_columna("Exportaciones (En millones de dólares)")
exportaciones_peso = buscar_columna("Exportaciones (Peso neto en toneladas)")
importaciones_valor = buscar_columna("Importaciones (Valor CIF en millones dólares)")
importaciones_peso = buscar_columna("Importaciones (Peso Bruto en Toneladas)")
terminos_intercambio = buscar_columna("Índice de Términos del Intercambio")

# Fiscal
ingresos_totales_spnf = buscar_columna("Ingresos Totales SPNF")
ingresos_corrientes_spnf = buscar_columna("Ingresos Corrientes del SPNF")
ingresos_capital_spnf = buscar_columna("Ingresos de Capital del SPNF")
egresos_totales_spnf = buscar_columna("Egresos Totales del SPNF")
egresos_corrientes_spnf = buscar_columna("Egresos Corrientes del SPNF")
egresos_capital_spnf = buscar_columna("Egresos de Capital del SPNF")
resultado_corriente_spnf = buscar_columna("Resultado Fiscal Corriente del SPNF")
resultado_global_spnf = buscar_columna("Resultado Fiscal Global del SPNF")
financiamiento_externo_spnf = buscar_columna("Financiamiento Externo del SPNF")
financiamiento_interno_spnf = buscar_columna("Financiamiento Interno del SPNF")

# Social
pobreza_bolivia = buscar_columna("Bolivia: Incidencia de pobreza")
pobreza_urbana = buscar_columna("Urbano: Incidencia de pobreza")
pobreza_rural = buscar_columna("Rural: Incidencia de pobreza")
pobreza_extrema_bolivia = buscar_columna("Bolivia: Indidencia de pobreza extrema")
pobreza_extrema_urbana = buscar_columna("Urbano: Incidencia de pobreza extrema")
pobreza_extrema_rural = buscar_columna("Rural: Incidencia de pobreza extrema")
gini_bolivia = buscar_columna("Bolivia: Índice de GINI")
gini_urbano = buscar_columna("Urbano: Índice de GINI")
gini_rural = buscar_columna("Rural: Índice de GINI")
desocupacion_nacional = buscar_columna("Tasa de Desocupación Nacional")
pea = buscar_columna("Población Económicamente Activa en número de personas")


# ============================================================
# 7. SERIES DERIVADAS
# ============================================================


def construir_derivadas(df):
    d = df.copy().sort_values("fecha")

    # Ratio crédito / depósitos, comparable aunque ambos estén expresados en USD.
    if credito_privado in d.columns and depositos in d.columns:
        d["Ratio crédito/depósitos (%)"] = np.where(
            d[depositos].notna() & (d[depositos] != 0),
            d[credito_privado] / d[depositos] * 100,
            np.nan,
        )

    # Brecha cambiaria solo en fechas donde ambas series están presentes.
    if tc_venta in d.columns and tc_oficial in d.columns:
        d["Brecha TC referencial-oficial (%)"] = np.where(
            d[tc_oficial].notna() & d[tc_venta].notna() & (d[tc_oficial] != 0),
            (d[tc_venta] / d[tc_oficial] - 1) * 100,
            np.nan,
        )

    # Valor unitario aproximado: millones USD / toneladas * 1e6.
    if exportaciones_valor in d.columns and exportaciones_peso in d.columns:
        d["Valor unitario exportaciones (USD/ton)"] = np.where(
            d[exportaciones_peso].notna() & (d[exportaciones_peso] != 0),
            d[exportaciones_valor] * 1_000_000 / d[exportaciones_peso],
            np.nan,
        )

    if importaciones_valor in d.columns and importaciones_peso in d.columns:
        d["Valor unitario importaciones (USD/ton)"] = np.where(
            d[importaciones_peso].notna() & (d[importaciones_peso] != 0),
            d[importaciones_valor] * 1_000_000 / d[importaciones_peso],
            np.nan,
        )

    # Flujo neto de transferencias por BCB: entradas - salidas.
    if transf_ext_a_sf_bcb in d.columns and transf_sf_al_ext_bcb in d.columns:
        d["Transferencias netas vía BCB (MM USD)"] = d[transf_ext_a_sf_bcb] - d[transf_sf_al_ext_bcb]

    # Flujo neto de EIF por sistema financiero.
    if transf_ext_a_eif in d.columns and transf_eif_al_ext in d.columns:
        d["Transferencias netas EIF (MM USD)"] = d[transf_ext_a_eif] - d[transf_eif_al_ext]

    # Brechas sociales.
    if pobreza_rural in d.columns and pobreza_urbana in d.columns:
        d["Brecha pobreza rural-urbana (p.p.)"] = d[pobreza_rural] - d[pobreza_urbana]
    if pobreza_extrema_rural in d.columns and pobreza_extrema_urbana in d.columns:
        d["Brecha pobreza extrema rural-urbana (p.p.)"] = d[pobreza_extrema_rural] - d[pobreza_extrema_urbana]
    if gini_rural in d.columns and gini_urbano in d.columns:
        d["Brecha Gini rural-urbano"] = d[gini_rural] - d[gini_urbano]

    return d


df_original = construir_derivadas(df_original)

ratio_credito_depositos = "Ratio crédito/depósitos (%)" if "Ratio crédito/depósitos (%)" in df_original.columns else None
brecha_tc_col = "Brecha TC referencial-oficial (%)" if "Brecha TC referencial-oficial (%)" in df_original.columns else None
vu_export = "Valor unitario exportaciones (USD/ton)" if "Valor unitario exportaciones (USD/ton)" in df_original.columns else None
vu_import = "Valor unitario importaciones (USD/ton)" if "Valor unitario importaciones (USD/ton)" in df_original.columns else None
transf_netas_bcb = "Transferencias netas vía BCB (MM USD)" if "Transferencias netas vía BCB (MM USD)" in df_original.columns else None
transf_netas_eif = "Transferencias netas EIF (MM USD)" if "Transferencias netas EIF (MM USD)" in df_original.columns else None
brecha_pobreza = "Brecha pobreza rural-urbana (p.p.)" if "Brecha pobreza rural-urbana (p.p.)" in df_original.columns else None
brecha_pobreza_extrema = "Brecha pobreza extrema rural-urbana (p.p.)" if "Brecha pobreza extrema rural-urbana (p.p.)" in df_original.columns else None
brecha_gini = "Brecha Gini rural-urbano" if "Brecha Gini rural-urbano" in df_original.columns else None


# ============================================================
# 8. ALERTAS Y PULSO MACRO
# ============================================================


def nivel_inflacion(df):
    v, _ = ultimo_valor(df, inflacion_12m)
    if v is None:
        return "Sin dato"
    if v >= 6:
        return "Alto"
    if v >= 3:
        return "Moderado"
    return "Bajo"


def nivel_externo(df):
    v, _ = ultimo_valor(df, rin)
    if v is None:
        return "Sin dato"
    if v < 2000:
        return "Alto"
    if v < 5000:
        return "Moderado"
    return "Bajo"


def nivel_cambiario(df):
    v, _ = ultimo_valor(df, brecha_tc_col)
    if v is None:
        return "Sin dato"
    a = abs(v)
    if a >= 8:
        return "Alto"
    if a >= 2:
        return "Moderado"
    return "Bajo"


def nivel_financiero(df):
    ratio, _ = ultimo_valor(df, ratio_credito_depositos)
    liq, _ = ultimo_valor(df, excedente_encaje_efectivo)
    if ratio is None and liq is None:
        return "Sin dato"
    if ratio is not None and ratio >= 110:
        return "Alto"
    if ratio is not None and ratio >= 100:
        return "Moderado"
    return "Bajo"


def nivel_real(df):
    # Preferir IGAE interanual: evita volver a diferenciar una tasa de PIB.
    v = variacion_interanual(df, igae)
    if v is None:
        v, _ = ultimo_valor(df, pib_pm)
    if v is None:
        return "Sin dato"
    if v < 0:
        return "Alto"
    if v < 2:
        return "Moderado"
    return "Bajo"


def nivel_fiscal(df):
    v, _ = valor_acumulado_anual(df, resultado_global_spnf)
    if v is None:
        return "Sin dato"
    if v < -5000:
        return "Alto"
    if v < 0:
        return "Moderado"
    return "Bajo"


def nivel_social(df):
    p, f = ultimo_valor(df, pobreza_bolivia)
    u, _ = ultimo_valor(df, desocupacion_nacional)
    if p is None and u is None:
        return "Sin dato"
    # Si pobreza es muy antigua, no debe dominar un semáforo coyuntural.
    if p is not None and f is not None and (df_original["fecha"].max() - f).days <= 730 and p >= 35:
        return "Alto"
    if u is not None and u >= 8:
        return "Alto"
    if (p is not None and p >= 25) or (u is not None and u >= 5):
        return "Moderado"
    return "Bajo"


def generar_lectura_ejecutiva(df):
    mejoro, empeoro, vigilar = [], [], []

    infl, _ = ultimo_valor(df, inflacion_12m)
    infl_pp = variacion_interanual_pp(df, inflacion_12m)
    rin_v, _ = ultimo_valor(df, rin)
    rin_yoy = variacion_interanual(df, rin)
    tc_v, _ = ultimo_valor(df, tc_oficial)
    brecha, _ = ultimo_valor(df, brecha_tc_col)
    bm_yoy = variacion_interanual(df, base_monetaria)
    dep_yoy = variacion_interanual(df, depositos)
    cred_yoy = variacion_interanual(df, credito_privado)
    saldo_acum, saldo_f = valor_acumulado_anual(df, saldo_comercial)
    saldo_yoy = variacion_acumulada_interanual(df, saldo_comercial)
    res_fiscal, f_fiscal = valor_acumulado_anual(df, resultado_global_spnf)

    if infl_pp is not None:
        if infl_pp < 0:
            mejoro.append(f"La inflación interanual disminuye {abs(infl_pp):.1f} p.p. frente a un año atrás.")
        else:
            empeoro.append(f"La inflación interanual aumenta {infl_pp:.1f} p.p. frente a un año atrás.")

    if rin_yoy is not None:
        if rin_yoy >= 0:
            mejoro.append(f"Las RIN aumentan {rin_yoy:.1f}% interanual.")
        else:
            empeoro.append(f"Las RIN disminuyen {abs(rin_yoy):.1f}% interanual.")

    if saldo_acum is not None and saldo_yoy is not None:
        if saldo_acum > 0:
            mejoro.append(f"El saldo comercial acumulado es superavitario ({formato_numero(saldo_acum,0)} MM USD).")
        elif saldo_yoy < 0:
            empeoro.append("El saldo comercial acumulado permanece deficitario.")

    if bm_yoy is not None and bm_yoy >= 15:
        vigilar.append(f"La base monetaria crece {bm_yoy:.1f}% interanual; revisar transmisión a precios y mercado cambiario.")

    if dep_yoy is not None and dep_yoy < 0:
        vigilar.append(f"Los depósitos caen {abs(dep_yoy):.1f}% interanual en la unidad reportada por la fuente.")

    if cred_yoy is not None and dep_yoy is not None and cred_yoy > dep_yoy + 5:
        vigilar.append("El crecimiento del crédito supera al de depósitos; monitorear fondeo y liquidez.")

    if brecha is not None and abs(brecha) >= 2:
        vigilar.append(f"La brecha referencial-oficial alcanza {brecha:.1f}% en la última fecha común.")

    if res_fiscal is not None and res_fiscal < 0:
        vigilar.append(f"El resultado fiscal global acumulado es deficitario ({formato_numero(res_fiscal,0)} MM Bs).")

    if tc_v is not None and infl is not None:
        # Mensaje descriptivo; no atribuye causalidad.
        vigilar.append("Conviene seguir conjuntamente tipo de cambio e inflación para detectar transmisión cambiaria a precios.")

    return mejoro[:4], empeoro[:4], vigilar[:5]


# ============================================================
# 9. SIDEBAR, AUTO-REFRESH Y FILTROS
# ============================================================

st.sidebar.markdown("## CENGOB")
st.sidebar.caption("Sistema de Información Económica de Bolivia")

actualizacion_automatica = st.sidebar.toggle(
    "Actualización automática",
    value=True,
    help="Si está disponible streamlit-autorefresh, la app se recarga cada 60 segundos.",
)

if actualizacion_automatica and st_autorefresh is not None:
    st_autorefresh(interval=AUTO_REFRESH_MS, limit=None, key="cengob_autorefresh")
elif actualizacion_automatica and st_autorefresh is None:
    st.sidebar.caption("ℹ️ Para auto-recarga visual instale `streamlit-autorefresh`.")

if st.sidebar.button("🔄 Actualizar ahora", use_container_width=True):
    descargar_excel_drive.clear()
    cargar_datos.clear()
    st.rerun()

fecha_min = df_original["fecha"].min().date()
fecha_max = df_original["fecha"].max().date()

rango = st.sidebar.date_input(
    "Rango visible",
    value=(fecha_min, fecha_max),
    min_value=fecha_min,
    max_value=fecha_max,
)

if isinstance(rango, (tuple, list)) and len(rango) == 2:
    inicio, fin = pd.to_datetime(rango[0]), pd.to_datetime(rango[1])
else:
    inicio, fin = pd.to_datetime(fecha_min), pd.to_datetime(fecha_max)

# La lógica analítica conserva toda la historia para calcular YoY, pero la visualización usa rango.
df_analitico = df_original.copy()
df = df_original[(df_original["fecha"] >= inicio) & (df_original["fecha"] <= fin)].copy()

PAGINAS = [
    "📌 Resumen ejecutivo",
    "🔀 Cruces de variables",
    "🔥 Precios y cambio",
    "🌎 Sector externo",
    "💵 Monetario",
    "🏦 Financiero",
    "🏭 Sector real",
    "🏛️ Fiscal",
    "👥 Social",
    "🚦 Riesgos",
    "🧭 Cobertura y explorador",
]

pagina = st.sidebar.radio("Navegación", PAGINAS, label_visibility="collapsed")

st.sidebar.markdown("---")
st.sidebar.metric("Variables disponibles", len(df_original.columns) - 1)
st.sidebar.metric("Última fecha en la base", pd.Timestamp(fecha_max).strftime("%d/%m/%Y"))

fuente = df_original.attrs.get("fuente_drive", "Fuente")
st.sidebar.caption(f"Fuente cargada: {fuente}")


# ============================================================
# 10. HEADER COMPACTO
# ============================================================

logo_html = ""
if os.path.exists("logo_cengob.png"):
    try:
        with open("logo_cengob.png", "rb") as f:
            logo64 = base64.b64encode(f.read()).decode()
        logo_html = f'<img src="data:image/png;base64,{logo64}" style="height:58px; object-fit:contain; margin-right:18px;">'
    except Exception:
        pass

st.markdown(
    f"""
    <div class="hero">
        <div style="display:flex; align-items:center; justify-content:space-between; gap:18px; flex-wrap:wrap;">
            <div style="display:flex; align-items:center; min-width:280px;">
                {logo_html}
                <div>
                    <div style="font-size:13px; font-weight:800; color:#F2CC5C !important;">CENGOB · MONITOR MACROECONÓMICO</div>
                    <h1 style="margin:2px 0 2px 0; font-size:clamp(24px,2.4vw,38px);">Sistema de Información Económica de Bolivia</h1>
                    <div style="font-size:13px; opacity:0.94;">Seguimiento ejecutivo, cruces de variables y alertas para la toma de decisiones</div>
                </div>
            </div>
            <div style="text-align:right; min-width:190px;">
                <div style="font-size:11px; opacity:0.82;">ÚLTIMA FECHA DETECTADA</div>
                <div style="font-size:20px; font-weight:900; color:#F2CC5C !important;">{pd.Timestamp(fecha_max).strftime('%d/%m/%Y')}</div>
                <div style="font-size:11px; opacity:0.82;">Actualización Drive · caché {DRIVE_CACHE_TTL}s</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 11. PÁGINA: RESUMEN EJECUTIVO
# ============================================================

if pagina == "📌 Resumen ejecutivo":
    st.subheader("Resumen ejecutivo")
    st.caption("Cada tarjeta muestra la fecha propia de la serie. No se asume que todos los indicadores tengan el mismo corte.")

    c1, c2, c3, c4, c5, c6 = st.columns(6)

    with c1:
        kpi_desde_serie(df_analitico, "Inflación interanual", inflacion_12m, "%", delta_tipo="pp", mejora_si_sube=False, dec=2)
    with c2:
        kpi_desde_serie(df_analitico, "Tipo de cambio oficial", tc_oficial, "Bs/USD", delta_tipo="interanual", mejora_si_sube=False, dec=2)
    with c3:
        kpi_desde_serie(df_analitico, "RIN", rin, "MM USD", delta_tipo="interanual", mejora_si_sube=True, dec=1)
    with c4:
        kpi_desde_serie(df_analitico, "Base monetaria", base_monetaria, "MM Bs", delta_tipo="interanual", mejora_si_sube=False, dec=0)
    with c5:
        kpi_desde_serie(df_analitico, "Saldo comercial", saldo_comercial, "MM USD", tipo="acumulado", delta_tipo="acumulado", mejora_si_sube=True, dec=0)
    with c6:
        kpi_desde_serie(df_analitico, "Resultado fiscal global", resultado_global_spnf, "MM Bs", tipo="acumulado", delta_tipo="acumulado", mejora_si_sube=True, dec=0)

    st.markdown("### Pulso macroeconómico CENGOB")
    r1, r2, r3, r4, r5, r6, r7 = st.columns(7)
    niveles = [
        ("Precios", nivel_inflacion(df_analitico)),
        ("Externo", nivel_externo(df_analitico)),
        ("Cambiario", nivel_cambiario(df_analitico)),
        ("Financiero", nivel_financiero(df_analitico)),
        ("Actividad", nivel_real(df_analitico)),
        ("Fiscal", nivel_fiscal(df_analitico)),
        ("Social", nivel_social(df_analitico)),
    ]
    for col, (t, n) in zip([r1, r2, r3, r4, r5, r6, r7], niveles):
        with col:
            tarjeta_riesgo(t, n)

    st.markdown("### Cruces prioritarios")
    a, b, c = st.columns(3)
    with a:
        grafico_cruce(
            df, inflacion_12m, tc_oficial,
            "Inflación vs. tipo de cambio oficial",
            "Inflación 12 meses", "TC oficial",
            "Nivel", "Variación interanual (%)",
            eje1="Inflación (%)", eje2="TC: variación interanual (%)",
            height=350,
        )
    with b:
        grafico_cruce(
            df, rin, tc_oficial,
            "RIN vs. tipo de cambio oficial",
            "RIN", "TC oficial",
            "Nivel", "Nivel",
            eje1="MM USD", eje2="Bs/USD",
            height=350,
        )
    with c:
        grafico_cruce(
            df, base_monetaria, m3,
            "Base monetaria vs. M'3",
            "Base monetaria", "M'3",
            "Variación interanual (%)", "Variación interanual (%)",
            eje1="% interanual", eje2="% interanual",
            height=350,
        )

    st.markdown("### Lectura ejecutiva automática")
    mejoro, empeoro, vigilar = generar_lectura_ejecutiva(df_analitico)
    x1, x2, x3 = st.columns(3)
    with x1:
        caja_senal("Qué mejoró", mejoro, "ok")
    with x2:
        caja_senal("Qué empeoró", empeoro, "danger")
    with x3:
        caja_senal("Qué vigilar", vigilar, "watch")

    st.markdown("### Composición y cobertura")
    y1, y2 = st.columns([1.05, 1.95])
    with y1:
        vals = []
        labs = []
        for lab, col in [("Oro", oro), ("Divisas", divisas), ("Recursos alta liquidez", recursos_alta_liquidez), ("DEG", deg)]:
            v, _ = ultimo_valor(df_analitico, col)
            if v is not None and v >= 0:
                vals.append(v)
                labs.append(lab)
        if vals:
            fig = go.Figure(go.Pie(
                labels=labs, values=vals, hole=0.62,
                marker=dict(colors=[CENGOB_GOLD, CENGOB_GREEN, "#8FAF9A", "#64748B"]),
                textinfo="percent",
            ))
            base_layout(fig, "Composición de activos de reserva", height=350, legend=True)
            fig.update_layout(margin=dict(l=15, r=15, t=60, b=25))
            mostrar_fig(fig)
        else:
            st.info("Sin información suficiente para composición de reservas.")

    with y2:
        indicadores_cov = [
            ("Inflación interanual", inflacion_12m),
            ("TC oficial", tc_oficial),
            ("RIN", rin),
            ("Base monetaria", base_monetaria),
            ("Crédito privado", credito_privado),
            ("Depósitos", depositos),
            ("Saldo comercial", saldo_comercial),
            ("Resultado fiscal", resultado_global_spnf),
            ("IGAE", igae),
            ("Desocupación", desocupacion_nacional),
            ("Pobreza", pobreza_bolivia),
        ]
        filas = []
        fmax = df_original["fecha"].max()
        for n, col in indicadores_cov:
            f = ultima_fecha_serie(df_analitico, col)
            est, _ = estado_frescura(f, fmax)
            filas.append({
                "Indicador": n,
                "Frecuencia aprox.": frecuencia_serie(df_analitico, col),
                "Último dato": formato_fecha(f),
                "Rezago (días)": None if f is None else int((fmax - f).days),
                "Estado": est,
            })
        st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True, height=330)


# ============================================================
# 12. PÁGINA: CRUCES DE VARIABLES
# ============================================================

elif pagina == "🔀 Cruces de variables":
    st.subheader("Cruces de variables y señales analíticas")
    st.caption("Los cruces se sincronizan a frecuencia mensual. La correlación describe co-movimiento; no implica causalidad.")

    bloque = st.selectbox(
        "Bloque analítico",
        [
            "Precios y mercado cambiario",
            "Monetario y financiero",
            "Sector externo",
            "Fiscal y actividad",
            "Social",
            "Matriz de correlaciones",
            "Analizador interactivo",
        ],
    )

    if bloque == "Precios y mercado cambiario":
        a, b = st.columns(2)
        with a:
            grafico_cruce(df, inflacion_12m, tc_oficial,
                          "Inflación vs. tipo de cambio",
                          "Inflación 12m", "TC oficial",
                          "Nivel", "Variación interanual (%)",
                          eje1="Inflación (%)", eje2="TC YoY (%)")
        with b:
            grafico_correlacion_movil(df, inflacion_12m, tc_oficial,
                                      "Correlación inflación–TC",
                                      "Nivel", "Variación interanual (%)", 12)

        c, d = st.columns(2)
        with c:
            grafico_cruce(df, inflacion_12m, base_monetaria,
                          "Inflación vs. expansión de base monetaria",
                          "Inflación 12m", "Base monetaria",
                          "Nivel", "Variación interanual (%)",
                          eje1="Inflación (%)", eje2="BM YoY (%)")
        with d:
            grafico_cruce(df, inflacion_12m, m3,
                          "Inflación vs. crecimiento de M'3",
                          "Inflación 12m", "M'3",
                          "Nivel", "Variación interanual (%)",
                          eje1="Inflación (%)", eje2="M3 YoY (%)")

        e, f = st.columns(2)
        with e:
            grafico_cruce(df, inflacion_12m, tcr,
                          "Inflación vs. tipo de cambio real",
                          "Inflación", "TCR",
                          "Nivel", "Nivel",
                          eje1="%", eje2="Índice")
        with f:
            grafico_cruce(df, brecha_tc_col, inflacion_mensual,
                          "Brecha cambiaria vs. inflación mensual",
                          "Brecha TC", "Inflación mensual",
                          "Nivel", "Nivel",
                          eje1="%", eje2="%")

    elif bloque == "Monetario y financiero":
        a, b = st.columns(2)
        with a:
            grafico_cruce(df, base_monetaria, m3,
                          "Base monetaria vs. M'3",
                          "Base monetaria", "M'3",
                          "Variación interanual (%)", "Variación interanual (%)",
                          eje1="% YoY", eje2="% YoY")
        with b:
            grafico_cruce(df, credito_privado, depositos,
                          "Crédito vs. depósitos",
                          "Crédito", "Depósitos",
                          "Nivel", "Nivel",
                          eje1="MM USD", eje2="MM USD")

        c, d = st.columns(2)
        with c:
            grafico_linea(df, ratio_credito_depositos, "Ratio crédito/depósitos", "%")
        with d:
            grafico_cruce(df, credito_privado, depositos,
                          "Crecimiento crédito vs. depósitos",
                          "Crédito", "Depósitos",
                          "Variación interanual (%)", "Variación interanual (%)",
                          eje1="Crédito YoY (%)", eje2="Depósitos YoY (%)")

        e, f = st.columns(2)
        with e:
            grafico_cruce(df, excedente_encaje_efectivo, tc_oficial,
                          "Liquidez excedente vs. tipo de cambio",
                          "Excedente encaje", "TC oficial",
                          "Nivel", "Nivel",
                          eje1="MM USD", eje2="Bs/USD")
        with f:
            grafico_cruce(df, tasa_reporto_mn, excedente_encaje_efectivo,
                          "Tasa de reporto vs. liquidez excedente",
                          "Reporto MN", "Excedente encaje",
                          "Nivel", "Nivel",
                          eje1="%", eje2="MM USD")

        g, h = st.columns(2)
        with g:
            grafico_cruce(df, bol_dep, bol_cred,
                          "Bolivianización de depósitos vs. créditos",
                          "Bolivianización depósitos", "Bolivianización créditos",
                          "Nivel", "Nivel", eje1="%", eje2="%")
        with h:
            grafico_cruce(df, financiamiento_corto_bcb, operaciones_reporto_eif,
                          "Financiamiento de corto plazo vs. reportos EIF",
                          "Financiamiento corto BCB", "Reportos EIF",
                          "Nivel", "Nivel", eje1="MM USD", eje2="MM USD")

    elif bloque == "Sector externo":
        a, b = st.columns(2)
        with a:
            grafico_cruce(df, rin, tc_oficial,
                          "RIN vs. tipo de cambio oficial",
                          "RIN", "TC oficial",
                          "Nivel", "Nivel", eje1="MM USD", eje2="Bs/USD")
        with b:
            grafico_correlacion_movil(df, rin, tc_oficial,
                                      "Correlación móvil RIN–TC", "Nivel", "Nivel", 12)

        c, d = st.columns(2)
        with c:
            grafico_cruce(df, terminos_intercambio, saldo_comercial,
                          "Términos de intercambio vs. saldo comercial",
                          "Términos de intercambio", "Saldo comercial",
                          "Nivel", "Nivel", how2="sum",
                          eje1="Índice", eje2="MM USD")
        with d:
            grafico_cruce(df, vu_export, vu_import,
                          "Valor unitario implícito: exportaciones vs. importaciones",
                          "Exportaciones", "Importaciones",
                          "Nivel", "Nivel",
                          eje1="USD/ton", eje2="USD/ton")

        e, f = st.columns(2)
        with e:
            grafico_cruce(df, transf_netas_bcb, tc_oficial,
                          "Transferencias netas vía BCB vs. TC",
                          "Transferencias netas", "TC oficial",
                          "Nivel", "Nivel", how1="sum",
                          eje1="MM USD", eje2="Bs/USD")
        with f:
            grafico_cruce(df, saldo_comercial, rin,
                          "Saldo comercial vs. RIN",
                          "Saldo comercial", "RIN",
                          "Nivel", "Nivel", how1="sum",
                          eje1="MM USD", eje2="MM USD")

        g, h = st.columns(2)
        with g:
            grafico_cruce(df, oro, divisas,
                          "Composición de reservas: oro vs. divisas",
                          "Oro", "Divisas", "Nivel", "Nivel",
                          eje1="MM USD", eje2="MM USD")
        with h:
            grafico_cruce(df, total_activos_externos, rin,
                          "Activos externos totales vs. RIN",
                          "Activos externos", "RIN", "Nivel", "Nivel",
                          eje1="MM USD", eje2="MM USD")

    elif bloque == "Fiscal y actividad":
        a, b = st.columns(2)
        with a:
            grafico_cruce(df, ingresos_totales_spnf, egresos_totales_spnf,
                          "Ingresos vs. egresos del SPNF",
                          "Ingresos", "Egresos", "Nivel", "Nivel",
                          how1="sum", how2="sum", eje1="MM Bs", eje2="MM Bs")
        with b:
            grafico_cruce(df, resultado_global_spnf, financiamiento_interno_spnf,
                          "Resultado fiscal vs. financiamiento interno",
                          "Resultado global", "Financiamiento interno",
                          "Nivel", "Nivel", how1="sum", how2="sum",
                          eje1="MM Bs", eje2="MM Bs")

        c, d = st.columns(2)
        with c:
            grafico_cruce(df, resultado_global_spnf, financiamiento_externo_spnf,
                          "Resultado fiscal vs. financiamiento externo",
                          "Resultado global", "Financiamiento externo",
                          "Nivel", "Nivel", how1="sum", how2="sum",
                          eje1="MM Bs", eje2="MM Bs")
        with d:
            grafico_cruce(df, credito_neto_bcb_spnf, base_monetaria,
                          "Crédito neto del BCB al SPNF vs. base monetaria",
                          "Crédito neto BCB-SPNF", "Base monetaria",
                          "Nivel", "Nivel", eje1="MM Bs", eje2="MM Bs")

        e, f = st.columns(2)
        with e:
            grafico_cruce(df, igae, inflacion_12m,
                          "Actividad (IGAE) vs. inflación",
                          "IGAE", "Inflación", "Variación interanual (%)", "Nivel",
                          eje1="IGAE YoY (%)", eje2="Inflación (%)")
        with f:
            grafico_cruce(df, formacion_capital, pib_pm,
                          "Inversión vs. crecimiento del PIB",
                          "FBKF", "PIB", "Nivel", "Nivel",
                          eje1="%", eje2="%")

    elif bloque == "Social":
        a, b = st.columns(2)
        with a:
            grafico_cruce(df, pobreza_urbana, pobreza_rural,
                          "Pobreza urbana vs. rural",
                          "Urbana", "Rural", "Nivel", "Nivel",
                          eje1="%", eje2="%")
        with b:
            grafico_linea(df, brecha_pobreza, "Brecha de pobreza rural–urbana", "p.p.")

        c, d = st.columns(2)
        with c:
            grafico_cruce(df, pobreza_extrema_urbana, pobreza_extrema_rural,
                          "Pobreza extrema urbana vs. rural",
                          "Urbana", "Rural", "Nivel", "Nivel",
                          eje1="%", eje2="%")
        with d:
            grafico_linea(df, brecha_pobreza_extrema, "Brecha de pobreza extrema rural–urbana", "p.p.")

        e, f = st.columns(2)
        with e:
            grafico_cruce(df, gini_urbano, gini_rural,
                          "GINI urbano vs. rural",
                          "GINI urbano", "GINI rural", "Nivel", "Nivel",
                          eje1="Índice", eje2="Índice")
        with f:
            grafico_cruce(df, desocupacion_nacional, pea,
                          "Desocupación vs. PEA",
                          "Desocupación", "PEA", "Nivel", "Nivel",
                          eje1="%", eje2="Personas")

    elif bloque == "Matriz de correlaciones":
        st.markdown("#### Correlaciones mensuales de variables transformadas")
        st.caption("Por defecto se usan transformaciones que hacen más comparables las dinámicas. Puede cambiar la selección.")

        candidatos = {
            "Inflación 12m": (inflacion_12m, "Nivel"),
            "TC oficial YoY": (tc_oficial, "Variación interanual (%)"),
            "RIN YoY": (rin, "Variación interanual (%)"),
            "Base monetaria YoY": (base_monetaria, "Variación interanual (%)"),
            "M3 YoY": (m3, "Variación interanual (%)"),
            "Crédito YoY": (credito_privado, "Variación interanual (%)"),
            "Depósitos YoY": (depositos, "Variación interanual (%)"),
            "Saldo comercial": (saldo_comercial, "Nivel"),
            "IGAE YoY": (igae, "Variación interanual (%)"),
            "Reporto MN": (tasa_reporto_mn, "Nivel"),
        }
        disponibles = [k for k, (c, _) in candidatos.items() if c is not None]
        sel = st.multiselect("Variables", disponibles, default=disponibles[:8])

        if len(sel) >= 2:
            frames = []
            for nombre in sel:
                col, transf = candidatos[nombre]
                s = transformar_serie(serie_mensual(df, col), transf).rename(nombre)
                frames.append(s)
            mat = pd.concat(frames, axis=1).corr(min_periods=6)
            fig = px.imshow(
                mat,
                text_auto=".2f",
                zmin=-1,
                zmax=1,
                color_continuous_scale="RdBu_r",
                aspect="auto",
            )
            base_layout(fig, "Matriz de correlaciones", height=600, legend=False)
            fig.update_layout(coloraxis_colorbar=dict(title="ρ"))
            mostrar_fig(fig)
        else:
            st.info("Seleccione al menos dos variables.")

    elif bloque == "Analizador interactivo":
        st.markdown("#### Analizador de dos variables")
        variables_numericas = [c for c in df.columns if c != "fecha" and pd.api.types.is_numeric_dtype(df[c])]
        if len(variables_numericas) >= 2:
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                xcol = st.selectbox("Variable X", variables_numericas, index=0)
            with c2:
                y_index = 1 if len(variables_numericas) > 1 else 0
                ycol = st.selectbox("Variable Y", variables_numericas, index=y_index)
            transformaciones = ["Nivel", "Variación interanual (%)", "Variación mensual (%)", "Cambio 12 meses (p.p./unidades)", "Índice base 100", "Z-score"]
            with c3:
                tx = st.selectbox("Transformación X", transformaciones, index=0)
            with c4:
                ty = st.selectbox("Transformación Y", transformaciones, index=0)

            grafico_cruce(df, xcol, ycol, f"{xcol} vs. {ycol}", xcol, ycol, tx, ty)
            s1, s2 = st.columns(2)
            with s1:
                grafico_scatter_cruce(df, xcol, ycol, "Dispersión", xcol, ycol, tx, ty)
            with s2:
                ventana = st.slider("Ventana de correlación móvil (meses)", 6, 36, 12)
                grafico_correlacion_movil(df, xcol, ycol, "Correlación móvil", tx, ty, ventana)


# ============================================================
# 13. PÁGINA: PRECIOS Y CAMBIO
# ============================================================

elif pagina == "🔥 Precios y cambio":
    st.subheader("Precios y mercado cambiario")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_desde_serie(df_analitico, "Inflación mensual", inflacion_mensual, "%", delta_tipo="pp", mejora_si_sube=False)
    with c2:
        kpi_desde_serie(df_analitico, "Inflación acumulada", inflacion_acumulada, "%", delta_tipo="pp", mejora_si_sube=False)
    with c3:
        kpi_desde_serie(df_analitico, "Inflación interanual", inflacion_12m, "%", delta_tipo="pp", mejora_si_sube=False)
    with c4:
        kpi_desde_serie(df_analitico, "TC oficial", tc_oficial, "Bs/USD", delta_tipo="interanual", mejora_si_sube=False)

    a, b = st.columns(2)
    with a:
        grafico_lineas_multiples(
            df, [inflacion_mensual, inflacion_acumulada, inflacion_12m],
            "Inflación: mensual, acumulada e interanual", "%",
            nombres={inflacion_mensual:"Mensual", inflacion_acumulada:"Acumulada", inflacion_12m:"12 meses"},
        )
    with b:
        grafico_lineas_multiples(
            df, [tc_oficial, tc_venta, tc_compra_ref],
            "Tipos de cambio", "Bs/USD",
            nombres={tc_oficial:"Oficial", tc_venta:"Referencial venta", tc_compra_ref:"Referencial compra"},
        )

    c, d = st.columns(2)
    with c:
        grafico_linea(df, brecha_tc_col, "Brecha referencial–oficial", "%")
    with d:
        grafico_linea(df, tcr, "Índice de tipo de cambio real", "Índice")

    e, f = st.columns(2)
    with e:
        grafico_cruce(df, inflacion_12m, tc_oficial,
                      "Inflación vs. TC", "Inflación", "TC",
                      "Nivel", "Variación interanual (%)",
                      eje1="%", eje2="TC YoY (%)")
    with f:
        grafico_cruce(df, inflacion_12m, base_monetaria,
                      "Inflación vs. base monetaria", "Inflación", "BM",
                      "Nivel", "Variación interanual (%)",
                      eje1="%", eje2="BM YoY (%)")


# ============================================================
# 14. PÁGINA: SECTOR EXTERNO
# ============================================================

elif pagina == "🌎 Sector externo":
    st.subheader("Sector externo")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_desde_serie(df_analitico, "RIN", rin, "MM USD", mejora_si_sube=True, dec=1)
    with c2:
        kpi_desde_serie(df_analitico, "Divisas", divisas, "MM USD", mejora_si_sube=True, dec=1)
    with c3:
        kpi_desde_serie(df_analitico, "Oro", oro, "MM USD", mejora_si_sube=True, dec=1)
    with c4:
        kpi_desde_serie(df_analitico, "Saldo comercial", saldo_comercial, "MM USD", tipo="acumulado", delta_tipo="acumulado", mejora_si_sube=True, dec=0)

    a, b = st.columns(2)
    with a:
        grafico_linea(df, rin, "Reservas internacionales netas", "MM USD", rango=True)
    with b:
        grafico_lineas_multiples(df, [oro, divisas, recursos_alta_liquidez, deg], "Composición de reservas", "MM USD",
                                 nombres={oro:"Oro", divisas:"Divisas", recursos_alta_liquidez:"Alta liquidez", deg:"DEG"})

    c, d = st.columns(2)
    with c:
        grafico_doble_eje(df, exportaciones_peso, exportaciones_valor,
                          "Exportaciones: volumen y valor", "Toneladas", "Valor",
                          "Toneladas", "MM USD", "ton", "MM USD")
    with d:
        grafico_doble_eje(df, importaciones_peso, importaciones_valor,
                          "Importaciones: volumen y valor", "Toneladas", "Valor CIF",
                          "Toneladas", "MM USD", "ton", "MM USD")

    e, f = st.columns(2)
    with e:
        grafico_lineas_multiples(df, [vu_export, vu_import], "Valor unitario implícito", "USD/ton",
                                 nombres={vu_export:"Exportaciones", vu_import:"Importaciones"})
    with f:
        grafico_cruce(df, terminos_intercambio, saldo_comercial,
                      "Términos de intercambio vs. saldo comercial",
                      "Términos de intercambio", "Saldo comercial",
                      "Nivel", "Nivel", how2="sum", eje1="Índice", eje2="MM USD")

    g, h = st.columns(2)
    with g:
        grafico_lineas_multiples(df, [transf_ext_a_sf_bcb, transf_sf_al_ext_bcb, transf_netas_bcb],
                                 "Transferencias vía BCB", "MM USD",
                                 nombres={transf_ext_a_sf_bcb:"Entradas", transf_sf_al_ext_bcb:"Salidas", transf_netas_bcb:"Neto"})
    with h:
        grafico_cruce(df, transf_netas_bcb, tc_oficial,
                      "Transferencias netas vs. TC", "Transferencias netas", "TC",
                      "Nivel", "Nivel", how1="sum", eje1="MM USD", eje2="Bs/USD")


# ============================================================
# 15. PÁGINA: MONETARIO
# ============================================================

elif pagina == "💵 Monetario":
    st.subheader("Sector monetario y operaciones del BCB")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_desde_serie(df_analitico, "Base monetaria", base_monetaria, "MM Bs", dec=0, mejora_si_sube=False)
    with c2:
        kpi_desde_serie(df_analitico, "M'1", m1, "MM Bs", dec=0, mejora_si_sube=False)
    with c3:
        kpi_desde_serie(df_analitico, "M'2", m2, "MM Bs", dec=0, mejora_si_sube=False)
    with c4:
        kpi_desde_serie(df_analitico, "M'3", m3, "MM Bs", dec=0, mejora_si_sube=False)

    a, b = st.columns(2)
    with a:
        grafico_lineas_multiples(df, [base_monetaria, emision_monetaria, reservas_bancarias_totales],
                                 "Base monetaria, emisión y reservas bancarias", "MM Bs",
                                 nombres={base_monetaria:"Base monetaria", emision_monetaria:"Emisión", reservas_bancarias_totales:"Reservas bancarias"})
    with b:
        grafico_lineas_multiples(df, [m1, m2, m3], "Agregados monetarios", "MM Bs",
                                 nombres={m1:"M'1", m2:"M'2", m3:"M'3"})

    c, d = st.columns(2)
    with c:
        grafico_lineas_multiples(df, [tasa_reporto_mn, tasa_reporto_me], "Tasas de reporto BCB", "%",
                                 nombres={tasa_reporto_mn:"MN", tasa_reporto_me:"ME"})
    with d:
        grafico_lineas_multiples(df, [titulos_bcb_usd, titulos_tgn_usd], "Títulos BCB y TGN", "MM USD",
                                 nombres={titulos_bcb_usd:"BCB", titulos_tgn_usd:"TGN"})

    e, f = st.columns(2)
    with e:
        grafico_cruce(df, credito_neto_bcb_spnf, base_monetaria,
                      "Crédito neto BCB al SPNF vs. base monetaria",
                      "Crédito neto al SPNF", "Base monetaria", "Nivel", "Nivel",
                      eje1="MM Bs", eje2="MM Bs")
    with f:
        grafico_cruce(df, base_monetaria, m3,
                      "Base monetaria vs. M'3", "BM", "M'3",
                      "Variación interanual (%)", "Variación interanual (%)",
                      eje1="YoY %", eje2="YoY %")

    g, h = st.columns(2)
    with g:
        grafico_lineas_multiples(df, [financiamiento_corto_bcb, creditos_liquidez_bcb, operaciones_reporto_eif],
                                 "Facilidades de liquidez del BCB", "MM USD",
                                 nombres={financiamiento_corto_bcb:"Financiamiento corto", creditos_liquidez_bcb:"Créditos liquidez", operaciones_reporto_eif:"Reportos EIF"})
    with h:
        grafico_lineas_multiples(df, [m1_ratio, m2_ratio, m3_ratio], "Participación de MN en agregados monetarios", "%",
                                 nombres={m1_ratio:"M1/M'1", m2_ratio:"M2/M'2", m3_ratio:"M3/M'3"})


# ============================================================
# 16. PÁGINA: FINANCIERO
# ============================================================

elif pagina == "🏦 Financiero":
    st.subheader("Sistema financiero")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_desde_serie(df_analitico, "Crédito privado", credito_privado, "MM USD", dec=0)
    with c2:
        kpi_desde_serie(df_analitico, "Depósitos", depositos, "MM USD", dec=0)
    with c3:
        kpi_desde_serie(df_analitico, "Ratio crédito/depósitos", ratio_credito_depositos, "%", delta_tipo="pp", mejora_si_sube=False)
    with c4:
        kpi_desde_serie(df_analitico, "Excedente de encaje", excedente_encaje_efectivo, "MM USD", mejora_si_sube=True)

    a, b = st.columns(2)
    with a:
        grafico_lineas_multiples(df, [credito_privado, depositos], "Crédito y depósitos", "MM USD",
                                 nombres={credito_privado:"Crédito", depositos:"Depósitos"})
    with b:
        grafico_linea(df, ratio_credito_depositos, "Ratio crédito/depósitos", "%")

    c, d = st.columns(2)
    with c:
        grafico_lineas_multiples(df, [bol_dep, bol_cred, bol_ahorro], "Bolivianización", "%",
                                 nombres={bol_dep:"Depósitos", bol_cred:"Créditos", bol_ahorro:"Ahorro"})
    with d:
        grafico_lineas_multiples(df, [dep_vista, caja_ahorro, dep_plazo, otros_depositos], "Composición de depósitos", "MM USD",
                                 nombres={dep_vista:"Vista", caja_ahorro:"Caja ahorro", dep_plazo:"Plazo", otros_depositos:"Otros"})

    e, f = st.columns(2)
    with e:
        grafico_lineas_multiples(df, [encaje_constituido, excedente_encaje_efectivo, excedente_encaje_me],
                                 "Encaje y liquidez excedente", "MM USD",
                                 nombres={encaje_constituido:"Encaje", excedente_encaje_efectivo:"Excedente total", excedente_encaje_me:"Excedente ME"})
    with f:
        grafico_cruce(df, tasa_reporto_mn, excedente_encaje_efectivo,
                      "Reporto MN vs. liquidez excedente", "Tasa reporto", "Liquidez",
                      "Nivel", "Nivel", eje1="%", eje2="MM USD")

    g, h = st.columns(2)
    with g:
        grafico_cruce(df, credito_privado, depositos,
                      "Crecimiento crédito vs. depósitos", "Crédito", "Depósitos",
                      "Variación interanual (%)", "Variación interanual (%)",
                      eje1="YoY %", eje2="YoY %")
    with h:
        grafico_cruce(df, excedente_encaje_me, tc_oficial,
                      "Liquidez en ME vs. tipo de cambio", "Liquidez ME", "TC",
                      "Nivel", "Nivel", eje1="MM USD", eje2="Bs/USD")


# ============================================================
# 17. PÁGINA: SECTOR REAL
# ============================================================

elif pagina == "🏭 Sector real":
    st.subheader("Sector real")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_desde_serie(df_analitico, "IGAE", igae, "Índice", dec=1)
    with c2:
        kpi_desde_serie(df_analitico, "PIB", pib_pm, "%", delta_tipo="pp", dec=2)
    with c3:
        kpi_desde_serie(df_analitico, "Consumo hogares", consumo_hogares, "%", delta_tipo="pp")
    with c4:
        kpi_desde_serie(df_analitico, "FBKF", formacion_capital, "%", delta_tipo="pp")

    a, b = st.columns(2)
    with a:
        grafico_linea(df, igae, "Índice Global de Actividad Económica", "Índice")
    with b:
        grafico_lineas_multiples(df, [consumo_hogares, consumo_publico, formacion_capital],
                                 "Demanda interna", "%",
                                 nombres={consumo_hogares:"Consumo hogares", consumo_publico:"Consumo público", formacion_capital:"FBKF"})

    c, d = st.columns(2)
    with c:
        grafico_lineas_multiples(df, [expo_bienes_servicios, impo_bienes_servicios], "Sector externo real", "%",
                                 nombres={expo_bienes_servicios:"Exportaciones", impo_bienes_servicios:"Importaciones"})
    with d:
        grafico_cruce(df, igae, inflacion_12m,
                      "IGAE vs. inflación", "IGAE", "Inflación",
                      "Variación interanual (%)", "Nivel",
                      eje1="IGAE YoY (%)", eje2="Inflación (%)")


# ============================================================
# 18. PÁGINA: FISCAL
# ============================================================

elif pagina == "🏛️ Fiscal":
    st.subheader("Sector fiscal - SPNF")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_desde_serie(df_analitico, "Ingresos totales", ingresos_totales_spnf, "MM Bs", tipo="acumulado", delta_tipo="acumulado", dec=0)
    with c2:
        kpi_desde_serie(df_analitico, "Egresos totales", egresos_totales_spnf, "MM Bs", tipo="acumulado", delta_tipo="acumulado", mejora_si_sube=False, dec=0)
    with c3:
        kpi_desde_serie(df_analitico, "Resultado corriente", resultado_corriente_spnf, "MM Bs", tipo="acumulado", delta_tipo="acumulado", dec=0)
    with c4:
        kpi_desde_serie(df_analitico, "Resultado global", resultado_global_spnf, "MM Bs", tipo="acumulado", delta_tipo="acumulado", dec=0)

    a, b = st.columns(2)
    with a:
        grafico_lineas_multiples(df, [ingresos_totales_spnf, egresos_totales_spnf], "Ingresos y egresos del SPNF", "MM Bs",
                                 nombres={ingresos_totales_spnf:"Ingresos", egresos_totales_spnf:"Egresos"})
    with b:
        grafico_lineas_multiples(df, [resultado_corriente_spnf, resultado_global_spnf], "Resultados fiscales", "MM Bs",
                                 nombres={resultado_corriente_spnf:"Corriente", resultado_global_spnf:"Global"})

    c, d = st.columns(2)
    with c:
        grafico_lineas_multiples(df, [financiamiento_interno_spnf, financiamiento_externo_spnf], "Financiamiento del SPNF", "MM Bs",
                                 nombres={financiamiento_interno_spnf:"Interno", financiamiento_externo_spnf:"Externo"})
    with d:
        grafico_cruce(df, resultado_global_spnf, financiamiento_interno_spnf,
                      "Resultado fiscal vs. financiamiento interno",
                      "Resultado global", "Financiamiento interno",
                      "Nivel", "Nivel", how1="sum", how2="sum",
                      eje1="MM Bs", eje2="MM Bs")

    e, f = st.columns(2)
    with e:
        grafico_lineas_multiples(df, [ingresos_corrientes_spnf, ingresos_capital_spnf], "Composición de ingresos", "MM Bs",
                                 nombres={ingresos_corrientes_spnf:"Corrientes", ingresos_capital_spnf:"Capital"})
    with f:
        grafico_lineas_multiples(df, [egresos_corrientes_spnf, egresos_capital_spnf], "Composición de egresos", "MM Bs",
                                 nombres={egresos_corrientes_spnf:"Corrientes", egresos_capital_spnf:"Capital"})


# ============================================================
# 19. PÁGINA: SOCIAL
# ============================================================

elif pagina == "👥 Social":
    st.subheader("Indicadores sociales")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_desde_serie(df_analitico, "Pobreza", pobreza_bolivia, "%", delta_tipo="pp", mejora_si_sube=False)
    with c2:
        kpi_desde_serie(df_analitico, "Pobreza extrema", pobreza_extrema_bolivia, "%", delta_tipo="pp", mejora_si_sube=False)
    with c3:
        kpi_desde_serie(df_analitico, "GINI", gini_bolivia, "", delta_tipo="pp", mejora_si_sube=False, dec=3)
    with c4:
        kpi_desde_serie(df_analitico, "Desocupación", desocupacion_nacional, "%", delta_tipo="pp", mejora_si_sube=False)

    a, b = st.columns(2)
    with a:
        grafico_lineas_multiples(df, [pobreza_bolivia, pobreza_urbana, pobreza_rural], "Incidencia de pobreza", "%",
                                 nombres={pobreza_bolivia:"Bolivia", pobreza_urbana:"Urbano", pobreza_rural:"Rural"})
    with b:
        grafico_lineas_multiples(df, [pobreza_extrema_bolivia, pobreza_extrema_urbana, pobreza_extrema_rural], "Pobreza extrema", "%",
                                 nombres={pobreza_extrema_bolivia:"Bolivia", pobreza_extrema_urbana:"Urbano", pobreza_extrema_rural:"Rural"})

    c, d = st.columns(2)
    with c:
        grafico_lineas_multiples(df, [gini_bolivia, gini_urbano, gini_rural], "Índice de GINI", "Índice",
                                 nombres={gini_bolivia:"Bolivia", gini_urbano:"Urbano", gini_rural:"Rural"})
    with d:
        grafico_linea(df, brecha_pobreza, "Brecha pobreza rural–urbana", "p.p.")

    e, f = st.columns(2)
    with e:
        grafico_linea(df, desocupacion_nacional, "Tasa de desocupación nacional", "%")
    with f:
        grafico_cruce(df, desocupacion_nacional, pea, "Desocupación vs. PEA", "Desocupación", "PEA", "Nivel", "Nivel", eje1="%", eje2="Personas")


# ============================================================
# 20. PÁGINA: RIESGOS
# ============================================================

elif pagina == "🚦 Riesgos":
    st.subheader("Semáforo macroeconómico")
    st.caption("Los umbrales son referenciales y deben calibrarse con criterio técnico e histórico.")

    riesgos = [
        ("Presión inflacionaria", nivel_inflacion(df_analitico), "Inflación interanual."),
        ("Posición externa", nivel_externo(df_analitico), "Nivel de RIN."),
        ("Presión cambiaria", nivel_cambiario(df_analitico), "Brecha en última fecha con TC referencial y oficial simultáneos."),
        ("Liquidez financiera", nivel_financiero(df_analitico), "Ratio crédito/depósitos y liquidez."),
        ("Actividad", nivel_real(df_analitico), "Crecimiento IGAE; PIB solo como respaldo."),
        ("Resultado fiscal", nivel_fiscal(df_analitico), "Resultado global acumulado del SPNF."),
        ("Riesgo social", nivel_social(df_analitico), "Pobreza/desocupación considerando rezago."),
    ]

    cols = st.columns(4)
    for i, r in enumerate(riesgos[:4]):
        with cols[i]:
            tarjeta_riesgo(*r)
    cols2 = st.columns(3)
    for i, r in enumerate(riesgos[4:]):
        with cols2[i]:
            tarjeta_riesgo(*r)

    st.markdown("### Indicadores de riesgo detrás del semáforo")
    a, b = st.columns(2)
    with a:
        grafico_linea(df, brecha_tc_col, "Brecha cambiaria sincronizada", "%")
        grafico_linea(df, ratio_credito_depositos, "Ratio crédito/depósitos", "%")
    with b:
        grafico_cruce(df, excedente_encaje_efectivo, tc_oficial,
                      "Liquidez vs. TC", "Liquidez", "TC", "Nivel", "Nivel", eje1="MM USD", eje2="Bs/USD")
        grafico_cruce(df, resultado_global_spnf, financiamiento_interno_spnf,
                      "Resultado fiscal vs. financiamiento interno",
                      "Resultado", "Financiamiento", "Nivel", "Nivel", how1="sum", how2="sum", eje1="MM Bs", eje2="MM Bs")


# ============================================================
# 21. PÁGINA: COBERTURA Y EXPLORADOR
# ============================================================

elif pagina == "🧭 Cobertura y explorador":
    st.subheader("Cobertura, calidad y explorador de variables")

    st.markdown("### Cobertura de indicadores")
    filas = []
    fmax = df_original["fecha"].max()

    catalogo = [
        ("IGAE", igae, "Real"),
        ("PIB", pib_pm, "Real"),
        ("Inflación 12m", inflacion_12m, "Precios"),
        ("TC oficial", tc_oficial, "Cambiario"),
        ("TC referencial", tc_venta, "Cambiario"),
        ("RIN", rin, "Externo"),
        ("Saldo comercial", saldo_comercial, "Externo"),
        ("Términos de intercambio", terminos_intercambio, "Externo"),
        ("Base monetaria", base_monetaria, "Monetario"),
        ("M3", m3, "Monetario"),
        ("Tasa reporto MN", tasa_reporto_mn, "Monetario"),
        ("Crédito privado", credito_privado, "Financiero"),
        ("Depósitos", depositos, "Financiero"),
        ("Excedente encaje", excedente_encaje_efectivo, "Financiero"),
        ("Resultado fiscal", resultado_global_spnf, "Fiscal"),
        ("Financiamiento interno", financiamiento_interno_spnf, "Fiscal"),
        ("Pobreza", pobreza_bolivia, "Social"),
        ("GINI", gini_bolivia, "Social"),
        ("Desocupación", desocupacion_nacional, "Social"),
    ]

    for nombre, col, sector in catalogo:
        f = ultima_fecha_serie(df_analitico, col)
        estado, _ = estado_frescura(f, fmax)
        filas.append({
            "Sector": sector,
            "Indicador": nombre,
            "Columna detectada": col or "No detectada",
            "Frecuencia aprox.": frecuencia_serie(df_analitico, col),
            "Último dato": formato_fecha(f),
            "Rezago (días)": None if f is None else int((fmax - f).days),
            "Estado": estado,
        })

    cob = pd.DataFrame(filas)
    st.dataframe(cob, use_container_width=True, hide_index=True, height=480)

    st.markdown("### Explorador")
    excluir = {"fecha"}
    variables = [c for c in df.columns if c not in excluir and pd.api.types.is_numeric_dtype(df[c])]

    if variables:
        c1, c2 = st.columns([2, 1])
        with c1:
            sel = st.selectbox("Variable", variables)
        with c2:
            transform = st.selectbox("Transformación", ["Nivel", "Variación interanual (%)", "Variación mensual (%)", "Índice base 100", "Z-score"])

        s = transformar_serie(serie_mensual(df, sel), transform).dropna()
        if not s.empty:
            fig = go.Figure(go.Scatter(x=s.index, y=s.values, mode="lines", line=dict(color=CENGOB_GREEN, width=2.7)))
            base_layout(fig, f"{sel} · {transform}", height=430, legend=False)
            mostrar_fig(fig)
        else:
            st.info("No hay datos suficientes para esa transformación.")

    st.markdown("### Descargas")
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "⬇️ Descargar base filtrada (CSV)",
            data=df.to_csv(index=False).encode("utf-8-sig"),
            file_name="cengob_base_filtrada.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with d2:
        st.download_button(
            "⬇️ Descargar cobertura (CSV)",
            data=cob.to_csv(index=False).encode("utf-8-sig"),
            file_name="cengob_cobertura_indicadores.csv",
            mime="text/csv",
            use_container_width=True,
        )


# ============================================================
# 22. PIE DE PÁGINA
# ============================================================

st.markdown("---")
st.markdown(
    f"""
    <div style="display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap; padding:2px 4px 12px 4px;">
        <div style="color:{MUTED}; font-size:11px;">Fuente: base institucional integrada CENGOB (INE, BCB, ASFI y otras fuentes oficiales según cada serie).</div>
        <div style="color:{CENGOB_GREEN}; font-size:11px; font-weight:800;">Sistema de Información Económica de Bolivia · CENGOB</div>
    </div>
    """,
    unsafe_allow_html=True,
)
