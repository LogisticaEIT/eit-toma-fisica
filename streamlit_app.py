# ==============================================================================
# MÓDULO DE CUSTODIOS & TOMA FÍSICA: TomaFisicaApp.py (V7 gPro - Producción)
# ==============================================================================
# COMANDO DE EJECUCIÓN EN CMD (WINDOWS 7 / PYTHON 3.8):
# cd /d "D:\Bases\ADM LOG"
# python -m streamlit run TomaFisicaApp.py --server.address 192.168.100.47 --server.port 8501 --server.fileWatcherType none --server.enableCORS false --server.enableXsrfProtection false
# ==============================================================================
# -*- coding: utf-8 -*-

import os
import json
import io
import csv
import time
import math
import html
import hashlib
import urllib.parse
import urllib3
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import requests
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

st.set_page_config(
    page_title="Toma Física & Muestreo V7 gPro",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"], .main {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
    }
    [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"] {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
    }
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="stSidebarCollapsedControl"] button,
    [data-testid="stSidebarCollapsedControl"] svg,
    [data-testid="stHeader"] button,
    [data-testid="stHeader"] svg {
        color: #0F172A !important;
        fill: #0F172A !important;
        opacity: 1 !important;
    }
    [data-testid="stSidebarCollapsedControl"] button {
        background-color: #FEF3C7 !important;
        border: 2px solid #D97706 !important;
        border-radius: 8px !important;
    }
    [data-testid="stSidebar"] { background-color: #0F172A !important; color: #F8FAFC !important; }
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] div, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] summary {
        color: #F8FAFC !important;
    }
    [data-testid="stSidebar"] details,
    [data-testid="stSidebar"] [data-testid="stExpander"] details {
        background-color: #1E293B !important;
        border: 1px solid #334155 !important;
        border-radius: 8px !important;
        margin-bottom: 6px !important;
        padding: 2px 4px !important;
    }
    [data-testid="stSidebar"] input {
        background-color: #0F172A !important; color: #FEF3C7 !important; border: 1px solid #F59E0B !important; font-weight: 700 !important;
    }
    [data-testid="stSidebar"] [data-baseweb="select"] > div {
        background-color: #0F172A !important; color: #FEF3C7 !important; border: 1px solid #F59E0B !important; font-weight: 700 !important;
    }
    [data-baseweb="select"] * { color: #451A03 !important; }
    [data-testid="stSidebar"] [data-baseweb="select"] * { color: #FEF3C7 !important; }
    [data-testid="stSidebar"] hr {
        border-color: #475569 !important;
        margin: 10px 0 !important;
    }
    [data-testid="stMain"] [data-testid="stExpander"] summary,
    [data-testid="stMain"] [data-testid="stExpander"] summary *,
    [data-testid="stMain"] [data-testid="stTabs"] button,
    [data-testid="stMain"] [data-testid="stTabs"] button * {
        color: #0F172A !important;
        font-weight: 700 !important;
    }
    [data-testid="stMain"] [data-testid="stExpander"] details {
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
    }
    .main-title {
        background: #0F172A; color: #FFFFFF !important; padding: 14px 18px; border-radius: 10px;
        font-size: 20px !important; font-weight: 800; border-bottom: 4px solid #F59E0B; margin-bottom: 8px;
    }
    .card-box { background: #FFFFFF !important; border: 1px solid #CBD5E1; border-radius: 10px; padding: 14px; margin-bottom: 12px; color: #0F172A !important; max-width: 100%; box-sizing: border-box; }
    .stTextInput input, .stNumberInput input, [data-testid="stMain"] [data-baseweb="select"] > div {
        background-color: #FEF3C7 !important; color: #451A03 !important; font-weight: 700 !important; border: 2px solid #D97706 !important;
    }
    .stButton button, .stDownloadButton button, .stButton button *, .stDownloadButton button * {
        background-color: #0F172A !important; color: #FFFFFF !important; font-weight: 700 !important;
    }
    .saldo-box {
        background: #EFF6FF; border: 2px solid #2563EB; border-radius: 8px; padding: 10px 14px; text-align: center;
    }
    .saldo-lbl { font-size: 13px !important; font-weight: 700 !important; color: #1E3A8A !important; display: block; margin-bottom: 4px; }
    .saldo-val { font-size: 22px !important; font-weight: 900 !important; color: #0F172A !important; display: block; }
    .scan-alert {
        background: #DCFCE7; border: 2px solid #16A34A; color: #14532D !important;
        border-radius: 8px; padding: 10px 12px; margin: 8px 0; font-weight: 800; font-size: 14px;
    }
    .last-tx-banner {
        background: #F0FDF4; border: 1px solid #86EFAC; border-left: 5px solid #16A34A;
        border-radius: 8px; padding: 10px 14px; margin-bottom: 10px; font-size: 13.5px; color: #0F172A !important;
    }
    .chart-card {
        background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 10px; padding: 14px; margin-bottom: 14px;
    }
    .chart-title {
        font-size: 17px !important; font-weight: 800 !important; color: #0F172A !important; margin-bottom: 8px;
    }
    table.tf-table { width: 100%; border-collapse: collapse; background: #FFFFFF; }
    table.tf-table thead th {
        background-color: #0F172A !important; color: #FFFFFF !important; padding: 8px 10px !important;
        text-align: left !important; font-weight: 700 !important; border: 1px solid #334155 !important; white-space: nowrap;
    }
    table.tf-table tbody td {
        padding: 7px 10px !important; border: 1px solid #CBD5E1 !important; color: #0F172A !important;
        background-color: #FFFFFF !important; font-size: 13px !important; white-space: nowrap;
    }
    @media (max-width: 768px) {
        .main-title { font-size: 17px !important; padding: 11px 13px; }
        .card-box { padding: 10px; }
        .saldo-val { font-size: 18px !important; }
        table.tf-table thead th, table.tf-table tbody td { padding: 5px 7px !important; font-size: 12px !important; }
    }
</style>
""",
    unsafe_allow_html=True,
)

components.html(
    """
    <script>
    try {
      window.parent.document.documentElement.setAttribute("translate", "no");
      window.parent.document.documentElement.classList.add("notranslate");
      window.parent.document.body.setAttribute("translate", "no");
      window.parent.document.body.classList.add("notranslate");
    } catch (e) {}
    </script>
    """,
    height=0,
)

RUTA_SCRIPT = os.path.abspath(__file__)
BASE_DIR = os.path.dirname(RUTA_SCRIPT)
st.session_state["errores_ui"] = []
st.session_state.setdefault("cap_v", 0)
st.session_state.setdefault("mem_regs", [])
st.session_state.setdefault("pin_reverso_act", "7031")
st.session_state.setdefault("cola_etiquetas", [])
st.session_state.setdefault("cod_label_act", "LBL-EIT-A4")
st.session_state.setdefault("fuente_gs_url", "")
st.session_state.setdefault("sku_pend_scan", "")
st.session_state.setdefault("scan_confirmado_cant", False)
st.session_state.setdefault("ultimo_disparo_ts", 0.0)
st.session_state.setdefault("ultimo_disparo_firma", "")
st.session_state.setdefault("ultimo_tx_resumen", "")
st.session_state.setdefault("tx_reverso_draft", "")


def norm_mayus_limpio(v):
    return " ".join(str(v or "").strip().upper().split())


def fmt_miles_comas(val, dec=2):
    n = float(val or 0.0)
    if not math.isfinite(n):
        n = 0.0
    signo = "-" if n < 0 else ""
    abs_n = abs(n)
    parte_ent = int(math.floor(abs_n))
    s_ent = f"{parte_ent:,}".replace(",", ".")
    if dec <= 0:
        return f"{signo}{s_ent}"
    factor = 10 ** dec
    parte_dec = int(round((abs_n - parte_ent) * factor))
    if parte_dec >= factor:
        parte_ent += 1
        parte_dec = 0
        s_ent = f"{parte_ent:,}".replace(",", ".")
    s_dec = f"{parte_dec:0{dec}d}"
    return f"{signo}{s_ent},{s_dec}"


def fmt_moneda_ec(val, dec=2):
    return f"USD {fmt_miles_comas(val, dec=dec)}"


def resolver_ruta(win_path, rel_name):
    if os.path.exists(win_path):
        return win_path
    rel_p = os.path.join(BASE_DIR, rel_name)
    if os.path.exists(rel_p):
        return rel_p
    return win_path


def fecha_arch(path):
    if os.path.exists(path):
        try:
            return datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return "No accesible"
    return "No encontrado"


pivot_path = resolver_ruta(r"D:\Bases\MICRO\PRODUCTO - PIVOT.xlsm", "PRODUCTO - PIVOT.xlsm")
nac_path = resolver_ruta(r"D:\Bases\MICRO\NAC.xlsx", "NAC.xlsx")
json_path = resolver_ruta(r"D:\Bases\ADM LOG\toma_fisica_registros.json", "toma_fisica_registros.json")
enviados_path = resolver_ruta(r"D:\Bases\ADM LOG\tx_enviados_gs.json", "tx_enviados_gs.json")
logs_path = resolver_ruta(r"D:\Bases\ADM LOG\toma_fisica_logs.json", "toma_fisica_logs.json")
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxGk8xG7zUy0vu-OUXwgg_qlUS8lmNlU1bh_CygmJkpieTt2jdfOkknoUjAKKis7EiA5g/exec"
PESTANA_LOGS_GS = "Logs_UI"

BODEGAS_OFICIALES = ["01", "03", "04", "05", "3V", "5V"]
MAPEO_BOD_COL = {"01": 5, "03": 6, "04": 7, "05": 8, "3V": 9, "5V": 10}
IDX_TOTAL_NAC = 11
IDX_COL_V = 21
IDX_COL_W = 22
MAX_COSTO_UNIT = 1000000.0

CUSTODIO_BOD = {
    "alexpajaritoruiz@icloud.com": BODEGAS_OFICIALES,
    "jlogistica.eit@gmail.com": BODEGAS_OFICIALES,
    "san14mej@gmail.com": ["01"],
    "bodegagye2026@gmail.com": ["03", "3V"],
    "bodegauio.eit@gmail.com": ["05", "5V"],
}

COLS_REP = [
    "Periodo", "ID_Tx", "Timestamp", "Correo_Usuario", "Ciudad", "Bodega",
    "Ubicacion", "COD_SKU", "Q", "BOD/SKU", "Motivo", "Estado_Reverso",
    "LINEA DE INVENTARIO", "PROVEEDOR", "ALTERNO", "DESCRIPCIÓN", "UNIDAD",
    "DIFERENCIA", "ESTADO", "COSTO UNITARIO", "COSTO TOTAL", "COSTO DIF", "QSA_Teorico",
]

COLS_LOG_UI = ["Timestamp", "Usuario", "Error_UI", "Enviado_GS"]


def norm_bod(v):
    if v is None:
        return ""
    s = str(v).strip().lstrip("'").replace(".0", "")
    return s.zfill(2) if s.isdigit() else s.upper()


# ==============================================================================
# CÓDIGO DE BARRAS CODE 128B (SVG VECTORIAL CENTRADO) PARA ETIQUETAS A4
# ==============================================================================
_CODE128_PATTERNS = [
    "212222", "222122", "222221", "121223", "121322", "131222", "122213", "122312", "132212", "221213",
    "221312", "231212", "112232", "122132", "122231", "113222", "123122", "123221", "223211", "221132",
    "221231", "213212", "223112", "312131", "311222", "321122", "321221", "312212", "322112", "322211",
    "212123", "212321", "232121", "111323", "131123", "131321", "112313", "132113", "132311", "211313",
    "231113", "231311", "112133", "112331", "132131", "113123", "113321", "133121", "313121", "211331",
    "231131", "213113", "213311", "213131", "311123", "311321", "331121", "312113", "312311", "332111",
    "314111", "221411", "431111", "111224", "111422", "121124", "121421", "141122", "141221", "112214",
    "112412", "122114", "122411", "142112", "142211", "241211", "221114", "413111", "241112", "134111",
    "111242", "121142", "121241", "114212", "124112", "124211", "411212", "421112", "421211", "212141",
    "214121", "412121", "111143", "111341", "131141", "114113", "114311", "411113", "411311", "113141",
    "114131", "311141", "411131", "211412", "211214", "211232", "2331112",
]


def generar_svg_code128(texto, alto_px=44):
    txt = "".join(ch for ch in norm_mayus_limpio(texto) if 32 <= ord(ch) <= 126)
    if not txt:
        txt = "000000"
    codes = [104]
    checksum = 104
    for pos, ch in enumerate(txt, start=1):
        val = ord(ch) - 32
        codes.append(val)
        checksum += val * pos
    codes.append(checksum % 103)
    codes.append(106)

    modulos = []
    for c in codes:
        pat = _CODE128_PATTERNS[c]
        for idx, dig in enumerate(pat):
            modulos.append((idx % 2 == 0, int(dig)))

    quiet_px = 18
    total_w = sum(w for _, w in modulos) + (quiet_px * 2)
    rects = [f'<rect x="0" y="0" width="{total_w}" height="{alto_px}" fill="#FFFFFF"/>']
    x = quiet_px
    for es_barra, ancho in modulos:
        if es_barra:
            rects.append(f'<rect x="{x}" y="0" width="{ancho}" height="{alto_px}" fill="#000000"/>')
        x += ancho

    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {total_w} {alto_px}" '
        f'preserveAspectRatio="xMidYMid meet" style="width:auto;max-width:96%;height:12.5mm;display:block;margin:0 auto;background:#FFFFFF;">'
        + "".join(rects)
        + "</svg>"
    )


def agregar_etiqueta_cola(sku, desc="", bod="", ubic="", usuario_lbl="", copias=1):
    sku_n = norm_mayus_limpio(sku)
    if not sku_n:
        return
    n_cop = max(1, min(int(copias or 1), 60))
    st.session_state.setdefault("cola_etiquetas", [])
    fecha_hoy = date.today().strftime("%Y-%m-%d")
    login_corto = str(usuario_lbl or st.session_state.get("usr_act", "")).split("@")[0].upper()[:14]
    for _ in range(n_cop):
        st.session_state["cola_etiquetas"].append({
            "COD_SKU": sku_n,
            "DESCRIPCION": norm_mayus_limpio(desc)[:42],
            "FECHA_LOGIN": f"{fecha_hoy} | {login_corto}".strip(" |"),
            "BODEGA": norm_bod(bod),
            "UBICACION": norm_mayus_limpio(ubic),
        })


def render_hoja_a4_etiquetas(cola, cod_label="LBL-EIT-A4", ancho_mm=58, alto_mm=25):
    if not cola:
        st.info("Aún no hay etiquetas marcadas en la cola.")
        return
    por_hoja = 33
    total_hojas = max(1, math.ceil(len(cola) / por_hoja))
    bloques_html = []
    for h_idx in range(total_hojas):
        sub = cola[h_idx * por_hoja : (h_idx + 1) * por_hoja]
        celdas = []
        for it in sub:
            s_sku = norm_mayus_limpio(it.get("COD_SKU", ""))
            s_desc = html.escape(norm_mayus_limpio(it.get("DESCRIPCION", ""))[:40])
            s_bod = html.escape(norm_bod(it.get("BODEGA", "")))
            svg_bar = generar_svg_code128(s_sku, alto_px=44)
            celdas.append(
                f'<div class="notranslate" translate="no" style="position:relative;width:{ancho_mm}mm;height:{alto_mm}mm;'
                f'border:1px dashed #64748B;box-sizing:border-box;padding:0.8mm 1.5mm;display:flex;flex-direction:column;'
                f'justify-content:space-between;align-items:center;text-align:center;background:#FFFFFF;overflow:hidden;">'
                f'<div style="width:100%;font-size:5.8pt;font-weight:700;color:#0F172A;line-height:1.0;'
                f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;text-align:center;">{s_desc}</div>'
                f'<div style="width:100%;display:flex;justify-content:center;align-items:center;background:#FFFFFF;">{svg_bar}</div>'
                f'<div style="position:relative;width:100%;display:flex;justify-content:center;align-items:center;">'
                f'<span style="position:absolute;left:0;bottom:0;font-size:7.8pt;font-weight:900;color:#000000;line-height:1.0;">{s_bod}</span>'
                f'<span style="font-size:8.8pt;font-weight:900;letter-spacing:0.6px;color:#000000;'
                f'line-height:1.0;text-align:center;font-family:Arial,Helvetica,sans-serif;">{html.escape(s_sku)}</span>'
                f'</div>'
                f'</div>'
            )
        bloques_html.append(
            f'<div class="hoja-a4 notranslate" translate="no" style="width:210mm;min-height:297mm;padding:4.5mm 5.5mm;box-sizing:border-box;'
            f'background:#FFFFFF;border:1px dashed #94A3B8;margin:0 auto 10px auto;display:grid;'
            f'grid-template-columns:repeat(3, {ancho_mm}mm);grid-auto-rows:{alto_mm}mm;gap:1.2mm 2mm;align-content:start;">'
            + "".join(celdas)
            + '</div>'
        )

    html_doc = f"""
    <div class="notranslate" translate="no">
      <div style="margin-bottom:8px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;">
        <button onclick="window.print()" type="button" style="padding:8px 14px;background:#0F172A;color:#FFFFFF;border:none;border-radius:6px;font-weight:700;cursor:pointer;">
          Imprimir Hoja A4 ({len(cola)} etiquetas / {total_hojas} hoja(s))
        </button>
        <span style="font-size:12px;color:#334155;font-weight:700;">A4 márgenes reducidos • {ancho_mm}×{alto_mm} mm con línea punteada • CODE128 con zona quieta</span>
      </div>
      <style>
        @page {{ size: A4 portrait; margin: 4mm; }}
        @media print {{
          body * {{ visibility: hidden; }}
          .hoja-a4, .hoja-a4 * {{ visibility: visible; }}
          .hoja-a4 {{ position: relative; width: 100% !important; min-height: auto !important; padding: 0 !important; border: none !important; margin: 0 !important; page-break-after: always; }}
          button {{ display: none !important; }}
        }}
      </style>
      {''.join(bloques_html)}
    </div>
    """
    components.html(html_doc, height=540, scrolling=True)


def generar_pin_transaccion(pin_base="7031", tx_id="", bod="", cod_sku="", q_val=0):
    tx_n = norm_mayus_limpio(tx_id)
    bod_n = norm_bod(bod)
    sku_n = norm_mayus_limpio(cod_sku)
    q_n = float(q_val or 0.0)
    raw = f"{str(pin_base).strip()}|TX|{tx_n}|{bod_n}|{sku_n}|{q_n}".encode("utf-8")
    dig = int(hashlib.sha256(raw).hexdigest()[:8], 16) % 10000
    return f"{dig:04d}"


def cargar_logs():
    if os.path.exists(logs_path):
        try:
            with open(logs_path, "r", encoding="utf-8") as f:
                d = json.load(f)
                return d if isinstance(d, list) else []
        except Exception:
            return []
    return []


def guardar_logs_lista(lista_logs):
    try:
        os.makedirs(os.path.dirname(logs_path), exist_ok=True)
        with open(logs_path, "w", encoding="utf-8") as f:
            json.dump(lista_logs[-500:], f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False


def registrar_log(detalle, origen="APP", usuario=""):
    det_s = str(detalle or "").strip()
    if not det_s:
        return
    if "InsecureRequestWarning" in det_s or "Chart fallback" in det_s:
        return
    now_s = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    usr_s = str(usuario or st.session_state.get("usr_act", "sistema")).strip().lower()
    err_texto = f"[{str(origen or 'APP').upper()}] {det_s[:280]}"
    item = {"Timestamp": now_s, "Usuario": usr_s, "Error_UI": err_texto, "Enviado_GS": "NO"}
    logs = cargar_logs()
    logs.append(item)
    guardar_logs_lista(logs)
    st.session_state["errores_ui"].append(f"{now_s[11:]} | {err_texto}")
    st.session_state["errores_ui"] = st.session_state["errores_ui"][-5:]


def log_err(msg, origen="APP"):
    registrar_log(msg, origen=origen)


def enviar_logs_a_gs(logs_pendientes):
    if not logs_pendientes:
        return 0, 0
    enviados_ok = 0
    intentos = 0
    for lg in logs_pendientes:
        intentos += 1
        ts_v = str(lg.get("Timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        usr_v = str(lg.get("Usuario", "sistema")).strip().lower()
        err_v = str(lg.get("Error_UI", "")).strip()
        payload_log = {
            "sheet": PESTANA_LOGS_GS,
            "pestana": PESTANA_LOGS_GS,
            "action": "log_ui",
            "Timestamp": ts_v,
            "Usuario": usr_v,
            "Error_UI": err_v,
            "row": [ts_v, usr_v, err_v],
        }
        try:
            r = requests.post(
                WEB_APP_URL,
                json=payload_log,
                headers={"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"},
                verify=False,
                timeout=15,
                allow_redirects=True,
            )
            txt = (r.text or "").strip()
            if r.status_code == 200 and "error" not in txt.lower():
                lg["Enviado_GS"] = "SI"
                enviados_ok += 1
        except Exception:
            pass
    return enviados_ok, intentos


def norm_est(v):
    s = str(v or "").strip().upper()
    if s in ["SOB", "SOBRANTE"]:
        return "SOBRANTE"
    if s in ["FAL", "FALTANTE"]:
        return "FALTANTE"
    if s in ["PENDIENTE_REVERSO", "PENDIENTE_APROBACION"]:
        return "PENDIENTE_REVERSO"
    if s in ["REVERSO_RECHAZADO", "RECHAZADO"]:
        return "REVERSO_RECHAZADO"
    if s in ["OK", "CUADRADO", "CUADRADO (OK)", "DE ACUERDO"]:
        return "CUADRADO (OK)"
    return s or "CUADRADO (OK)"


def norm_estado_rev(v, q_val=0.0):
    s = str(v or "").strip().upper()
    if s in ["PENDIENTE_APROBACION", "PENDIENTE", "APROBADO", "RECHAZADO", "NO_APLICA"]:
        return s
    if float(q_val or 0.0) < 0:
        return "PENDIENTE_APROBACION"
    return "NO_APLICA"


def es_computable_en_consolidado(reg):
    q_v = num_es(reg.get("Q", 0), max_abs=1e12)
    est_rev = norm_estado_rev(reg.get("Estado_Reverso", ""), q_v)
    if q_v < 0 and est_rev != "APROBADO":
        return False
    return True


def num_es(v, max_abs=None):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        val = float(v)
        if not math.isfinite(val):
            return 0.0
        if max_abs is not None and abs(val) > max_abs:
            return 0.0
        return val
    s = str(v).strip().replace("USD", "").replace("$", "").replace(" ", "")
    if not s or s.startswith("#") or s.upper() in ["NAN", "NONE", "NULL", "ERROR", "ERR"]:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        partes = s.split(",")
        if len(partes) == 2 and len(partes[1]) == 3 and partes[0].lstrip("-").isdigit():
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    try:
        val = float(s)
        if not math.isfinite(val):
            return 0.0
        if max_abs is not None and abs(val) > max_abs:
            return 0.0
        return val
    except Exception:
        return 0.0


def limpiar_costo_unit(v):
    c = num_es(v, max_abs=MAX_COSTO_UNIT)
    return round(c, 4) if c > 0 else 0.0


def calcular_dif_estado_costo(q_acum, st_teo, q_actual, c_unit, est_rev="NO_APLICA"):
    c_u = limpiar_costo_unit(c_unit)
    est_r = norm_estado_rev(est_rev, q_actual)
    if float(q_actual) < 0 and est_r == "PENDIENTE_APROBACION":
        return 0.0, "PENDIENTE_REVERSO", 0.0, 0.0, c_u
    if float(q_actual) < 0 and est_r == "RECHAZADO":
        return 0.0, "REVERSO_RECHAZADO", 0.0, 0.0, c_u
    dif_fisica = round(float(q_acum) - float(st_teo), 4)
    if dif_fisica > 0:
        estado = "SOBRANTE"
        signo_costo = 1.0
    elif dif_fisica < 0:
        estado = "FALTANTE"
        signo_costo = -1.0
    else:
        estado = "CUADRADO (OK)"
        signo_costo = 0.0
    costo_total = round(c_u * float(q_actual), 2)
    costo_dif = round(c_u * abs(dif_fisica) * signo_costo, 2)
    return dif_fisica, estado, costo_total, costo_dif, c_u


def fmt_txt(v):
    s = norm_mayus_limpio(v)
    return f"'{s}" if s and not s.startswith("'") else s


def cargar_json():
    regs = []
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                d = json.load(f)
                if isinstance(d, list):
                    regs.extend(d)
        except Exception as e:
            log_err(f"JSON leer: {e}", origen="JSON")
    if isinstance(st.session_state.get("mem_regs"), list):
        vistos = {norm_mayus_limpio(x.get("ID_Tx", "")) for x in regs if isinstance(x, dict) and x.get("ID_Tx")}
        for r in st.session_state["mem_regs"]:
            if isinstance(r, dict):
                tx = norm_mayus_limpio(r.get("ID_Tx", ""))
                if not tx or tx not in vistos:
                    regs.append(r)
                    if tx:
                        vistos.add(tx)
    return regs


def guardar_json(reg):
    st.session_state.setdefault("mem_regs", [])
    st.session_state["mem_regs"].append(reg)
    try:
        carpeta_json = os.path.dirname(json_path)
        if carpeta_json:
            os.makedirs(carpeta_json, exist_ok=True)
        regs_disco = []
        if os.path.exists(json_path):
            with open(json_path, "r", encoding="utf-8") as f:
                d = json.load(f)
                if isinstance(d, list):
                    regs_disco = d
        regs_disco.append(reg)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(regs_disco, f, indent=2, ensure_ascii=False)
    except Exception:
        pass
    return True


def _recalcular_lista_sku(lst, per_s, bod_s, sku_s, st_teo_k):
    q_comp = 0.0
    for r in lst:
        if (
            str(r.get("Periodo", "")).strip() == per_s
            and norm_bod(r.get("Bodega", "")) == bod_s
            and norm_mayus_limpio(r.get("COD_SKU", "")) == sku_s
            and es_computable_en_consolidado(r)
        ):
            q_comp = round(q_comp + num_es(r.get("Q", 0), max_abs=1e12), 4)

    dif_consol = round(q_comp - float(st_teo_k), 4)
    if dif_consol > 0:
        est_consol = "SOBRANTE"
        signo = 1.0
    elif dif_consol < 0:
        est_consol = "FALTANTE"
        signo = -1.0
    else:
        est_consol = "CUADRADO (OK)"
        signo = 0.0

    camb = False
    for r in lst:
        if (
            str(r.get("Periodo", "")).strip() == per_s
            and norm_bod(r.get("Bodega", "")) == bod_s
            and norm_mayus_limpio(r.get("COD_SKU", "")) == sku_s
        ):
            q_r = num_es(r.get("Q", 0), max_abs=1e12)
            er_r = norm_estado_rev(r.get("Estado_Reverso", ""), q_r)
            cu = limpiar_costo_unit(r.get("COSTO UNITARIO", 0))
            r["COD_SKU"] = sku_s
            r["Ubicacion"] = norm_mayus_limpio(r.get("Ubicacion", ""))
            r["Estado_Reverso"] = er_r
            r["QSA_Teorico"] = float(st_teo_k)
            r["COSTO UNITARIO"] = cu
            if q_r < 0 and er_r == "PENDIENTE_APROBACION":
                r["ESTADO"] = "PENDIENTE_REVERSO"
                r["DIFERENCIA"] = 0.0
                r["COSTO TOTAL"] = 0.0
                r["COSTO DIF"] = 0.0
            elif q_r < 0 and er_r == "RECHAZADO":
                r["ESTADO"] = "REVERSO_RECHAZADO"
                r["DIFERENCIA"] = 0.0
                r["COSTO TOTAL"] = 0.0
                r["COSTO DIF"] = 0.0
            else:
                r["ESTADO"] = est_consol
                r["DIFERENCIA"] = float(dif_consol)
                r["COSTO TOTAL"] = round(cu * q_r, 2)
                r["COSTO DIF"] = round(cu * abs(float(dif_consol)) * signo, 2)
            camb = True
    return camb, q_comp, dif_consol, est_consol


def consolidar_estado_sku_en_json(periodo_k, bod_k, sku_k, st_teo_k):
    per_s = str(periodo_k or "").strip()
    bod_s = norm_bod(bod_k)
    sku_s = norm_mayus_limpio(sku_k)
    if not per_s or not bod_s or not sku_s:
        return 0.0, 0.0, "CUADRADO (OK)"

    q_comp, dif_c, est_c = 0.0, 0.0, "CUADRADO (OK)"
    if isinstance(st.session_state.get("mem_regs"), list):
        _, q_comp, dif_c, est_c = _recalcular_lista_sku(st.session_state["mem_regs"], per_s, bod_s, sku_s, st_teo_k)
    try:
        if os.path.exists(json_path):
            with open(json_path, "r", encoding="utf-8") as f:
                d = json.load(f)
            if isinstance(d, list):
                camb, q_d, dif_d, est_d = _recalcular_lista_sku(d, per_s, bod_s, sku_s, st_teo_k)
                q_comp, dif_c, est_c = q_d, dif_d, est_d
                if camb:
                    with open(json_path, "w", encoding="utf-8") as f:
                        json.dump(d, f, indent=2, ensure_ascii=False)
    except Exception:
        pass
    return q_comp, dif_c, est_c


def resolver_reverso_pendiente(tx_id, aprobar=True, admin_usr=""):
    tx_s = norm_mayus_limpio(tx_id)
    if not tx_s:
        return False, None
    nuevo_est_rev = "APROBADO" if aprobar else "RECHAZADO"
    reg_afectado = None

    def _act_rev(lst):
        nonlocal reg_afectado
        toc = False
        for r in lst:
            if norm_mayus_limpio(r.get("ID_Tx", "")) == tx_s:
                r["Estado_Reverso"] = nuevo_est_rev
                mot_prev = str(r.get("Motivo", "")).strip()
                nota = f"[REV {nuevo_est_rev} por {admin_usr} {datetime.now().strftime('%Y-%m-%d %H:%M')}]"
                if nota not in mot_prev:
                    r["Motivo"] = f"{mot_prev} {nota}".strip()
                reg_afectado = dict(r)
                toc = True
        return toc

    if isinstance(st.session_state.get("mem_regs"), list):
        _act_rev(st.session_state["mem_regs"])
    try:
        if os.path.exists(json_path):
            with open(json_path, "r", encoding="utf-8") as f:
                d = json.load(f)
            if isinstance(d, list) and _act_rev(d):
                with open(json_path, "w", encoding="utf-8") as f:
                    json.dump(d, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

    if reg_afectado is not None:
        consolidar_estado_sku_en_json(
            reg_afectado.get("Periodo", ""),
            reg_afectado.get("Bodega", ""),
            reg_afectado.get("COD_SKU", ""),
            num_es(reg_afectado.get("QSA_Teorico", 0), max_abs=1e12),
        )
        for r in cargar_json():
            if norm_mayus_limpio(r.get("ID_Tx", "")) == tx_s:
                reg_afectado = dict(r)
                break
        return True, reg_afectado
    return False, None


def ids_enviados():
    if os.path.exists(enviados_path):
        try:
            with open(enviados_path, "r", encoding="utf-8") as f:
                d = json.load(f)
                if isinstance(d, list):
                    return {norm_mayus_limpio(x) for x in d if norm_mayus_limpio(x)}
        except Exception:
            pass
    return set()


def marcar_enviado(tx_id):
    t = norm_mayus_limpio(tx_id)
    if not t:
        return
    s = ids_enviados()
    if t not in s:
        s.add(t)
        try:
            carpeta_env = os.path.dirname(enviados_path)
            if carpeta_env:
                os.makedirs(carpeta_env, exist_ok=True)
            with open(enviados_path, "w", encoding="utf-8") as f:
                json.dump(sorted(list(s)), f, indent=2, ensure_ascii=False)
        except Exception:
            pass


@st.cache_data(ttl=60, show_spinner=False)
def leer_gs():
    out = []
    try:
        r = requests.get(
            f"{WEB_APP_URL}?action=list&_={int(datetime.now().timestamp() // 60)}",
            headers={"User-Agent": "Mozilla/5.0"},
            verify=False,
            timeout=10,
            allow_redirects=True,
        )
        if r.status_code == 200:
            txt = (r.text or "").strip()
            if txt.startswith("[") or txt.startswith("{"):
                data = json.loads(txt)
                filas = data if isinstance(data, list) else next((data[k] for k in ["registros", "data", "rows", "items"] if isinstance(data.get(k), list)), [])
                for it in filas:
                    if isinstance(it, dict):
                        d = {c: it.get(c, it.get("Q(SA) TEÓRICO", it.get("Q(SA) TEORICO", ""))) for c in COLS_REP}
                        if d.get("ID_Tx") or d.get("COD_SKU"):
                            out.append(d)
                    elif isinstance(it, list) and len(it) >= 8 and str(it[0]).strip().lower() not in ["periodo", "fecha"]:
                        d = {c: (it[i] if i < len(it) else "") for i, c in enumerate(COLS_REP)}
                        if d.get("ID_Tx") or d.get("COD_SKU"):
                            out.append(d)
    except Exception as e:
        log_err(f"GS leer: {e}", origen="GS_GET")
    return out


def armar_payload_gs(reg, tx_id):
    cu = limpiar_costo_unit(reg.get("COSTO UNITARIO", 0))
    q_v = num_es(reg.get("Q", 0))
    er_v = norm_estado_rev(reg.get("Estado_Reverso", ""), q_v)
    dif_v = num_es(reg.get("DIFERENCIA", 0))
    est_v = norm_est(reg.get("ESTADO", ""))
    ct_v = num_es(reg.get("COSTO TOTAL", 0), max_abs=1e11)
    cd_v = num_es(reg.get("COSTO DIF", 0), max_abs=1e11)
    mot_txt = str(reg.get("Motivo", "") or "").strip()
    if q_v < 0 and er_v not in ["", "NO_APLICA"]:
        mot_txt = f"[{er_v}] {mot_txt}".strip()
    bod_n = norm_bod(reg.get("Bodega", ""))
    sku_n = norm_mayus_limpio(reg.get("COD_SKU", ""))
    ubic_n = norm_mayus_limpio(reg.get("Ubicacion", ""))
    return [
        reg.get("Periodo", ""), tx_id, reg.get("Timestamp", ""), reg.get("Correo_Usuario", ""),
        reg.get("Ciudad", ""), fmt_txt(bod_n), ubic_n,
        fmt_txt(sku_n), q_v, f"{bod_n}-{sku_n}", mot_txt,
        reg.get("LINEA DE INVENTARIO", ""), norm_mayus_limpio(reg.get("PROVEEDOR", "")), fmt_txt(reg.get("ALTERNO", "")),
        reg.get("DESCRIPCIÓN", ""), reg.get("UNIDAD", ""), dif_v, est_v,
        cu, ct_v, cd_v, num_es(reg.get("QSA_Teorico", 0)),
    ]


def enviar_gs(reg, conocidos=None, modo_dup="omitir"):
    tx = norm_mayus_limpio(reg.get("ID_Tx", ""))
    if conocidos is None:
        conocidos = {norm_mayus_limpio(x.get("ID_Tx", "")) for x in leer_gs()} | ids_enviados()
    tx_env = tx
    if tx and tx in conocidos:
        if modo_dup == "omitir":
            return True, "omitido", tx
        if modo_dup == "duplicar":
            tx_env = f"{tx}-DUP-{datetime.now().strftime('%H%M%S')}"
    payload = armar_payload_gs(reg, tx_env)
    try:
        r = requests.post(
            WEB_APP_URL,
            json=payload,
            headers={"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"},
            verify=False,
            timeout=20,
            allow_redirects=True,
        )
        txt = (r.text or "").strip()
        ok = (r.status_code == 200) and ("error" not in txt.lower())
        if ok and tx_env:
            conocidos.add(tx_env)
            marcar_enviado(tx_env)
            leer_gs.clear()
        elif not ok:
            log_err(f"GS post HTTP {r.status_code}: {txt[:120]}", origen="GS_POST")
        return ok, ("sobreescrito" if modo_dup == "sobreescribir" else ("duplicado" if modo_dup == "duplicar" else "enviado")), tx_env
    except Exception as e:
        log_err(f"GS post exc: {e}", origen="GS_POST")
        return False, "error", tx_env


def _clave_dedup(r):
    tx = norm_mayus_limpio(r.get("ID_Tx", ""))
    if tx:
        return f"TX:{tx}"
    return (
        f"ROW:{str(r.get('Periodo','')).strip()}|{str(r.get('Timestamp','')).strip()}|"
        f"{norm_bod(r.get('Bodega',''))}|{norm_mayus_limpio(r.get('COD_SKU',''))}|{num_es(r.get('Q',0))}"
    )


def datos_combinados(incluir_gs=False):
    loc = cargar_json()
    res, vistos = [], set()
    for r in loc:
        if isinstance(r, dict):
            k = _clave_dedup(r)
            if k not in vistos:
                vistos.add(k)
                res.append(r)
    if incluir_gs:
        for r in leer_gs():
            if isinstance(r, dict):
                k = _clave_dedup(r)
                if k not in vistos:
                    vistos.add(k)
                    res.append(r)
    return res


def prep_df(regs):
    if not regs:
        return pd.DataFrame(columns=COLS_REP)
    df = pd.DataFrame(regs)
    for c in COLS_REP:
        if c not in df.columns:
            df[c] = ""
    df["Bodega"] = df["Bodega"].apply(norm_bod)
    df["Ubicacion"] = df["Ubicacion"].apply(norm_mayus_limpio)
    df["Ciudad"] = df["Ciudad"].apply(norm_mayus_limpio)
    df["Periodo"] = df["Periodo"].astype(str).str.strip()
    df["Correo_Usuario"] = df["Correo_Usuario"].astype(str).str.strip().str.lower()
    df["COD_SKU"] = df["COD_SKU"].apply(norm_mayus_limpio)
    df["PROVEEDOR"] = df["PROVEEDOR"].apply(norm_mayus_limpio)
    df["ESTADO"] = df["ESTADO"].apply(norm_est)
    for nc in ["Q", "QSA_Teorico", "DIFERENCIA"]:
        df[nc] = df[nc].apply(lambda x: num_es(x, max_abs=1e12))
    df["Estado_Reverso"] = [norm_estado_rev(er, qv) for er, qv in zip(df["Estado_Reverso"], df["Q"])]
    df["COSTO UNITARIO"] = df["COSTO UNITARIO"].apply(limpiar_costo_unit)
    df["COSTO TOTAL"] = df["COSTO TOTAL"].apply(lambda x: num_es(x, max_abs=1e11))
    df["COSTO DIF"] = df["COSTO DIF"].apply(lambda x: num_es(x, max_abs=1e11))
    return df[COLS_REP].copy()


def consolidar_df_por_sku(df_in):
    cols_out = [
        "Periodo", "Ciudad", "Bodega", "COD_SKU", "DESCRIPCIÓN", "PROVEEDOR",
        "Tx_Count", "Q_Neta", "QSA_Teorico", "Diferencia_Neta", "Estado_Neto",
        "Costo_Unitario", "Costo_Total_Neto", "Costo_Dif_Neto",
    ]
    if df_in is None or df_in.empty:
        return pd.DataFrame(columns=cols_out)
    df_of = df_in[~df_in["ESTADO"].isin(["PENDIENTE_REVERSO", "REVERSO_RECHAZADO"])].copy()
    if df_of.empty:
        return pd.DataFrame(columns=cols_out)

    filas = []
    for (per_k, bod_k, sku_k), grp in df_of.groupby(["Periodo", "Bodega", "COD_SKU"], as_index=False):
        q_neta = round(float(grp["Q"].sum()), 4)
        q_teo = float(grp["QSA_Teorico"].max())
        cu = float(grp["COSTO UNITARIO"].max())
        dif_neta, est_neto, c_tot_neto, c_dif_neto, cu_limp = calcular_dif_estado_costo(
            q_neta, q_teo, q_neta, cu, est_rev="NO_APLICA"
        )
        filas.append({
            "Periodo": per_k,
            "Ciudad": str(grp["Ciudad"].iloc[-1] or ""),
            "Bodega": bod_k,
            "COD_SKU": sku_k,
            "DESCRIPCIÓN": str(grp["DESCRIPCIÓN"].iloc[-1] or ""),
            "PROVEEDOR": norm_mayus_limpio(grp["PROVEEDOR"].iloc[-1] or ""),
            "Tx_Count": int(len(grp)),
            "Q_Neta": q_neta,
            "QSA_Teorico": q_teo,
            "Diferencia_Neta": dif_neta,
            "Estado_Neto": est_neto,
            "Costo_Unitario": cu_limp,
            "Costo_Total_Neto": c_tot_neto,
            "Costo_Dif_Neto": c_dif_neto,
        })
    return pd.DataFrame(filas)


def _formatear_celda_texto(val, col_name=""):
    if isinstance(val, (int, float)):
        cn = col_name.upper()
        if "UNITARIO" in cn:
            return fmt_moneda_ec(val, dec=4)
        if "COSTO" in cn or "USD" in cn:
            return fmt_moneda_ec(val, dec=2)
        if "PCT" in cn:
            return f"{fmt_miles_comas(val, dec=1)}%"
        if float(val).is_integer():
            return fmt_miles_comas(val, dec=0)
        return fmt_miles_comas(val, dec=2)
    return str(val or "")


def resolver_col_bodega(nac_rows, bod):
    bod_n = norm_bod(bod)
    if nac_rows and len(nac_rows) > 0:
        for i, h in enumerate(nac_rows[0]):
            if norm_bod(h) == bod_n:
                return i
    return MAPEO_BOD_COL.get(bod_n, 5)


def sku_habilitado_movimiento(row, muestreo=True):
    if not muestreo:
        return True
    if IDX_COL_V >= len(row):
        return True
    val_v = norm_mayus_limpio(row[IDX_COL_V])
    if not val_v or val_v in ["SIN MOV", "SIN_MOV", "NONE", "NAN", "0"]:
        return False
    return True


def obtener_costo_nac_col_w(row):
    if IDX_COL_W < len(row):
        return limpiar_costo_unit(row[IDX_COL_W])
    return 0.0


def obtener_skus_teoricos_bodega(nac_rows, dict_maestro, bod, muestreo=True, regs_respaldo=None):
    skus_set = set()
    desc_map = {}
    bod_n = norm_bod(bod)
    idx_b = resolver_col_bodega(nac_rows, bod_n)
    if nac_rows and len(nac_rows) > 1:
        for row in nac_rows[1:]:
            if not row or len(row) == 0 or row[0] is None:
                continue
            sv = norm_mayus_limpio(row[0])
            if not sv or sv in ["NONE", "NAN"]:
                continue
            if not sku_habilitado_movimiento(row, muestreo=muestreo):
                continue
            saldo_bod = num_es(row[idx_b], max_abs=1e12) if idx_b < len(row) else 0.0
            if saldo_bod > 0:
                skus_set.add(sv)
                if sv not in desc_map:
                    desc_map[sv] = next((v for k, v in dict_maestro.get(sv, {}).items() if "DESC" in k), "")
    if not skus_set and dict_maestro:
        for sv, info in dict_maestro.items():
            skus_set.add(sv)
            desc_map[sv] = next((v for k, v in info.items() if "DESC" in k), "")
    if not skus_set and regs_respaldo:
        for r in regs_respaldo:
            if norm_bod(r.get("Bodega", "")) == bod_n:
                sv = norm_mayus_limpio(r.get("COD_SKU", ""))
                if sv:
                    skus_set.add(sv)
                    if sv not in desc_map:
                        desc_map[sv] = str(r.get("DESCRIPCIÓN", "") or "").strip()
    return sorted(skus_set), desc_map


def contar_teoricos_bodega(nac_rows, bod, muestreo=True):
    skus, _ = obtener_skus_teoricos_bodega(nac_rows, {}, bod, muestreo=muestreo)
    return len(skus)


def construir_metricas_avance(df, nac_rows, muestreo=True):
    if df.empty:
        return []
    filas = []
    for (per_k, bod_k, cus_k), grp in df.groupby(["Periodo", "Bodega", "Correo_Usuario"]):
        grp_ok = grp[~grp["ESTADO"].isin(["PENDIENTE_REVERSO", "REVERSO_RECHAZADO"])]
        df_sku_u = consolidar_df_por_sku(grp_ok)
        teo_skus = contar_teoricos_bodega(nac_rows, bod_k, muestreo=muestreo)
        cont_skus = int(df_sku_u["COD_SKU"].nunique()) if not df_sku_u.empty else 0
        pend_skus = max(teo_skus - cont_skus, 0) if teo_skus > 0 else 0
        pct_avance = round((cont_skus / teo_skus) * 100.0, 2) if teo_skus > 0 else 0.0
        filas.append([
            str(per_k), str(bod_k), str(cus_k), int(len(grp_ok)), int(teo_skus), int(cont_skus),
            int(pend_skus), float(pct_avance), round(float(df_sku_u["Q_Neta"].sum()), 2) if not df_sku_u.empty else 0.0,
            int((df_sku_u["Estado_Neto"] == "SOBRANTE").sum()) if not df_sku_u.empty else 0,
            int((df_sku_u["Estado_Neto"] == "FALTANTE").sum()) if not df_sku_u.empty else 0,
            int((df_sku_u["Estado_Neto"] == "CUADRADO (OK)").sum()) if not df_sku_u.empty else 0,
            round(float(df_sku_u["Costo_Total_Neto"].sum()), 2) if not df_sku_u.empty else 0.0,
            round(float(df_sku_u["Costo_Dif_Neto"].sum()), 2) if not df_sku_u.empty else 0.0,
        ])
    return filas


def construir_resumen_bodega(df, nac_rows, muestreo=True, bodega_filtro="TODAS"):
    bodegas_iter = [norm_bod(bodega_filtro)] if bodega_filtro != "TODAS" else BODEGAS_OFICIALES
    filas = []
    for bod_k in bodegas_iter:
        grp_all = df[df["Bodega"] == bod_k] if not df.empty else pd.DataFrame(columns=COLS_REP)
        grp = grp_all[~grp_all["ESTADO"].isin(["PENDIENTE_REVERSO", "REVERSO_RECHAZADO"])] if not grp_all.empty else grp_all
        df_sku_u = consolidar_df_por_sku(grp)
        teo_skus = contar_teoricos_bodega(nac_rows, bod_k, muestreo=muestreo)
        cont_skus = int(df_sku_u["COD_SKU"].nunique()) if not df_sku_u.empty else 0
        pend_skus = max(teo_skus - cont_skus, 0) if teo_skus > 0 else 0
        pct_avance = round((cont_skus / teo_skus) * 100.0, 1) if teo_skus > 0 else 0.0
        if not grp_all.empty or bodega_filtro != "TODAS":
            filas.append({
                "Bodega": bod_k,
                "SKU_Teoricos": teo_skus,
                "SKU_Contados": cont_skus,
                "SKU_Pendientes": pend_skus,
                "Avance_Pct": pct_avance,
                "Tx": int(len(grp)),
                "Cantidad_Q": round(float(df_sku_u["Q_Neta"].sum()), 2) if not df_sku_u.empty else 0.0,
                "QSA_Teorico_Tot": round(float(df_sku_u["QSA_Teorico"].sum()), 2) if not df_sku_u.empty else 0.0,
                "Diferencia_Q_Tot": round(float(df_sku_u["Diferencia_Neta"].sum()), 2) if not df_sku_u.empty else 0.0,
                "Sobrantes": int((df_sku_u["Estado_Neto"] == "SOBRANTE").sum()) if not df_sku_u.empty else 0,
                "Faltantes": int((df_sku_u["Estado_Neto"] == "FALTANTE").sum()) if not df_sku_u.empty else 0,
                "Cuadrados": int((df_sku_u["Estado_Neto"] == "CUADRADO (OK)").sum()) if not df_sku_u.empty else 0,
                "Costo_Total_USD": round(float(df_sku_u["Costo_Total_Neto"].sum()), 2) if not df_sku_u.empty else 0.0,
                "Costo_Dif_USD": round(float(df_sku_u["Costo_Dif_Neto"].sum()), 2) if not df_sku_u.empty else 0.0,
            })
    return pd.DataFrame(filas)


def render_svg_barras_comparativas(df_plot, label_col, serie_a, serie_b, titulo, leyenda_a, leyenda_b, color_a="#2563EB", color_b="#94A3B8", es_moneda=False):
    if df_plot is None or df_plot.empty:
        return
    filas = df_plot.head(10).to_dict(orient="records")
    max_v = max(
        [abs(float(r.get(serie_a, 0) or 0)) for r in filas]
        + [abs(float(r.get(serie_b, 0) or 0)) for r in filas]
        + [1.0]
    )
    row_h = 56
    top_pad = 36
    left_w = 180
    bar_max_w = 540
    svg_w = left_w + bar_max_w + 220
    svg_h = top_pad + len(filas) * row_h + 14

    elems = [
        f'<rect x="12" y="8" width="16" height="16" rx="3" fill="{color_a}"/>',
        f'<text x="34" y="21" font-size="15" font-weight="800" fill="#0F172A">{html.escape(leyenda_a)}</text>',
        f'<rect x="270" y="8" width="16" height="16" rx="3" fill="{color_b}"/>',
        f'<text x="292" y="21" font-size="15" font-weight="800" fill="#0F172A">{html.escape(leyenda_b)}</text>',
    ]
    for i, r in enumerate(filas):
        y = top_pad + i * row_h
        lbl = html.escape(str(r.get(label_col, ""))[:22])
        va = float(r.get(serie_a, 0) or 0)
        vb = float(r.get(serie_b, 0) or 0)
        wa = max(3, int((abs(va) / max_v) * bar_max_w)) if va != 0 else 0
        wb = max(3, int((abs(vb) / max_v) * bar_max_w)) if vb != 0 else 0
        fa = fmt_moneda_ec(va, dec=2) if es_moneda else fmt_miles_comas(va, dec=0 if float(va).is_integer() else 2)
        fb = fmt_moneda_ec(vb, dec=2) if es_moneda else fmt_miles_comas(vb, dec=0 if float(vb).is_integer() else 2)
        elems.append(f'<text x="10" y="{y + 24}" font-size="16" font-weight="800" fill="#0F172A">{lbl}</text>')
        elems.append(f'<rect x="{left_w}" y="{y + 4}" width="{wa}" height="17" rx="4" fill="{color_a}"/>')
        elems.append(f'<text x="{left_w + wa + 10}" y="{y + 18}" font-size="15" font-weight="800" fill="#0F172A">{html.escape(fa)}</text>')
        elems.append(f'<rect x="{left_w}" y="{y + 25}" width="{wb}" height="17" rx="4" fill="{color_b}"/>')
        elems.append(f'<text x="{left_w + wb + 10}" y="{y + 39}" font-size="15" font-weight="700" fill="#1E293B">{html.escape(fb)}</text>')
        elems.append(f'<line x1="10" y1="{y + 50}" x2="{svg_w - 10}" y2="{y + 50}" stroke="#E2E8F0" stroke-width="1"/>')

    svg_html = (
        f'<div class="chart-card notranslate" translate="no"><div class="chart-title">{html.escape(titulo)}</div>'
        f'<svg width="100%" viewBox="0 0 {svg_w} {svg_h}" style="max-width:100%;height:auto;display:block;">'
        + "".join(elems)
        + "</svg></div>"
    )
    st.markdown(svg_html, unsafe_allow_html=True)


def render_svg_barras_simples(df_plot, label_col, val_col, titulo, color_pos="#16A34A", color_neg="#DC2626", es_moneda=False):
    if df_plot is None or df_plot.empty:
        return
    filas = df_plot.head(10).to_dict(orient="records")
    max_v = max([abs(float(r.get(val_col, 0) or 0)) for r in filas] + [1.0])
    row_h = 42
    left_w = 210
    bar_max_w = 510
    svg_w = left_w + bar_max_w + 220
    svg_h = len(filas) * row_h + 14
    elems = []
    for i, r in enumerate(filas):
        y = 8 + i * row_h
        lbl = html.escape(str(r.get(label_col, ""))[:24])
        v = float(r.get(val_col, 0) or 0)
        w = max(3, int((abs(v) / max_v) * bar_max_w)) if v != 0 else 0
        col_b = color_neg if v < 0 else color_pos
        fv = fmt_moneda_ec(v, dec=2) if es_moneda else fmt_miles_comas(v, dec=0 if float(v).is_integer() else 2)
        elems.append(f'<text x="10" y="{y + 22}" font-size="15.5" font-weight="800" fill="#0F172A">{lbl}</text>')
        elems.append(f'<rect x="{left_w}" y="{y + 6}" width="{w}" height="22" rx="4" fill="{col_b}"/>')
        elems.append(f'<text x="{left_w + w + 10}" y="{y + 22}" font-size="15" font-weight="800" fill="#0F172A">{html.escape(fv)}</text>')
    st.markdown(
        f'<div class="chart-card notranslate" translate="no"><div class="chart-title">{html.escape(titulo)}</div>'
        f'<svg width="100%" viewBox="0 0 {svg_w} {svg_h}" style="max-width:100%;height:auto;display:block;">'
        + "".join(elems)
        + "</svg></div>",
        unsafe_allow_html=True,
    )


def render_graficos_avance_operativo(df_bod_av):
    st.markdown("#### Gráficos Ejecutivos Totales (Escalables para 5.000+ SKUs)")
    if df_bod_av is None or df_bod_av.empty:
        return

    render_svg_barras_comparativas(
        df_bod_av, "Bodega", "SKU_Contados", "SKU_Teoricos",
        "Totales de SKUs: Auditados vs. Teóricos por Bodega",
        "SKUs Contados", "SKUs Teóricos",
        color_a="#16A34A", color_b="#1E3A8A", es_moneda=False,
    )
    render_svg_barras_comparativas(
        df_bod_av, "Bodega", "Cantidad_Q", "QSA_Teorico_Tot",
        "Totales en Cantidades (Unidades): Q Contada vs. QSA Teórico",
        "Q Física Contada", "QSA Teórico Total",
        color_a="#2563EB", color_b="#64748B", es_moneda=False,
    )
    render_svg_barras_comparativas(
        df_bod_av, "Bodega", "Costo_Total_USD", "Costo_Dif_USD",
        "Totales en Valores (USD): Valor Contado vs. Impacto Diferencia",
        "Valor Total Contado", "Impacto Dif. Neto",
        color_a="#0284C7", color_b="#F59E0B", es_moneda=True,
    )


def render_tabla_segura(df_in, cols, max_filas=50):
    if df_in is None or df_in.empty:
        st.info("No hay registros disponibles con los filtros actuales.")
        return
    sub = df_in[cols].head(max_filas)
    ths = "".join(f"<th>{html.escape(str(c))}</th>" for c in cols)
    trs = []
    for row in sub.to_dict(orient="records"):
        tds = [f"<td>{html.escape(_formatear_celda_texto(row.get(c, ''), c))}</td>" for c in cols]
        trs.append(f"<tr>{''.join(tds)}</tr>")
    st.markdown(
        f'<div class="notranslate" translate="no" style="overflow-x:auto;border-radius:8px;border:1px solid #CBD5E1;margin-top:6px;margin-bottom:10px;"><table class="tf-table"><thead><tr>{ths}</tr></thead><tbody>{"".join(trs)}</tbody></table></div>',
        unsafe_allow_html=True,
    )
    if len(df_in) > max_filas:
        st.caption(f"Mostrando las {max_filas} filas más recientes de {len(df_in)} registros filtrados.")


def _inyectar_sku_y_enfocar_cantidad(codigo):
    limpio = norm_mayus_limpio(codigo)
    if not limpio:
        return ""
    js_code = f"""
    <script>
    (function() {{
      const cod = {json.dumps(limpio)};
      try {{
        if (navigator.clipboard && navigator.clipboard.writeText) {{
          navigator.clipboard.writeText(cod).catch(() => {{}});
        }}
        const inputs = window.parent.document.querySelectorAll('input[type="text"], input[type="number"]');
        let cantInput = null;
        for (const inp of inputs) {{
          const lbl = (inp.getAttribute("aria-label") || "").toUpperCase();
          if (lbl.includes("COD_SKU") && inp.type === "text") {{
            const nativeSetter = Object.getOwnPropertyDescriptor(window.parent.HTMLInputElement.prototype, "value").set;
            nativeSetter.call(inp, cod);
            inp.dispatchEvent(new Event("input", {{ bubbles: true }}));
            inp.dispatchEvent(new Event("change", {{ bubbles: true }}));
          }} else if (lbl.includes("CANTIDAD (Q)") || lbl.includes("CANTIDAD")) {{
            cantInput = inp;
          }}
        }}
        if (cantInput) {{
          cantInput.scrollIntoView({{ behavior: "smooth", block: "center" }});
          cantInput.focus();
          cantInput.select();
        }}
      }} catch (e) {{}}
    }})();
    </script>
    """
    components.html(js_code, height=0)
    return limpio


def decodificar_sku_desde_imagen_bytes(img_bytes):
    if not img_bytes:
        return ""
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(img_bytes)).convert("L")
        try:
            from pyzbar.pyzbar import decode as pyzbar_decode
            res = pyzbar_decode(im)
            for item in res:
                raw = norm_mayus_limpio(item.data.decode("utf-8", errors="ignore"))
                if raw:
                    return raw
        except Exception:
            pass
    except Exception:
        pass
    return ""


def render_lector_codigos_barras(comp_key):
    html_code = f"""
    <div class="notranslate" translate="no" style="font-family:sans-serif;padding:10px;background:#FFFFFF;border:1px solid #94A3B8;border-radius:8px;color:#0F172A;">
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px;">
        <button id="btnStartScan_{comp_key}" type="button" style="padding:9px 13px;background:#0F172A;color:#FFF;border:none;border-radius:6px;font-weight:700;cursor:pointer;">Activar cámara</button>
        <button id="btnSnapScan_{comp_key}" type="button" style="padding:9px 13px;background:#D97706;color:#FFF;border:none;border-radius:6px;font-weight:700;cursor:pointer;">Capturar cuadro ahora</button>
        <button id="btnStopScan_{comp_key}" type="button" style="padding:9px 13px;background:#475569;color:#FFF;border:none;border-radius:6px;font-weight:700;cursor:pointer;">Detener</button>
      </div>
      <div style="position:relative;width:100%;min-height:230px;background:#0F172A;border-radius:8px;overflow:hidden;display:flex;align-items:center;justify-content:center;">
        <video id="vidScan_{comp_key}" playsinline muted style="width:100%;min-height:230px;max-height:310px;object-fit:cover;display:none;"></video>
        <span id="phScan_{comp_key}" style="color:#E2E8F0;font-size:13px;font-weight:700;padding:12px;text-align:center;">Pulsa "Activar cámara" para ver el código de barras en pantalla amplia.</span>
      </div>
      <canvas id="cvsScan_{comp_key}" style="display:none;"></canvas>
      <div style="display:flex;gap:6px;margin-top:8px;align-items:center;">
        <input id="inpManualScan_{comp_key}" type="text" placeholder="Si prefieres, digita/pega SKU y pulsa Usar SKU" style="flex:1;padding:8px 10px;border:1.5px solid #D97706;background:#FEF3C7;color:#451A03;border-radius:6px;font-weight:700;font-size:13px;" />
        <button id="btnUseManual_{comp_key}" type="button" style="padding:8px 12px;background:#16A34A;color:#FFF;border:none;border-radius:6px;font-weight:700;cursor:pointer;">Usar SKU</button>
      </div>
      <div id="resScan_{comp_key}" style="margin-top:8px;font-size:13px;font-weight:800;color:#1E3A8A;">Listo para escanear. Centra el código y usa "Capturar cuadro ahora" si tarda en leer.</div>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/@zxing/library@0.21.3/umd/index.min.js"></script>
    <script>
    (function() {{
      const btnStart = document.getElementById("btnStartScan_{comp_key}");
      const btnSnap = document.getElementById("btnSnapScan_{comp_key}");
      const btnStop = document.getElementById("btnStopScan_{comp_key}");
      const btnUseManual = document.getElementById("btnUseManual_{comp_key}");
      const inpManual = document.getElementById("inpManualScan_{comp_key}");
      const video = document.getElementById("vidScan_{comp_key}");
      const canvas = document.getElementById("cvsScan_{comp_key}");
      const phScan = document.getElementById("phScan_{comp_key}");
      const statusEl = document.getElementById("resScan_{comp_key}");
      let reader = null;
      let stream = null;
      let activo = false;

      function aplicarCodigo(codigo, detenerCamara) {{
        const limpio = String(codigo || "").trim().toUpperCase();
        if (!limpio) return;
        statusEl.textContent = "LEÍDO: " + limpio + " -> Ingresa Cantidad (Q) y pulsa REGISTRAR.";
        try {{
          if (navigator.clipboard && navigator.clipboard.writeText) {{
            navigator.clipboard.writeText(limpio).catch(() => {{}});
          }}
          const inputs = window.parent.document.querySelectorAll('input[type="text"], input[type="number"]');
          let cantInput = null;
          for (const inp of inputs) {{
            const lbl = (inp.getAttribute("aria-label") || "").toUpperCase();
            if (lbl.includes("COD_SKU") && inp.type === "text") {{
              const nativeSetter = Object.getOwnPropertyDescriptor(window.parent.HTMLInputElement.prototype, "value").set;
              nativeSetter.call(inp, limpio);
              inp.dispatchEvent(new Event("input", {{ bubbles: true }}));
              inp.dispatchEvent(new Event("change", {{ bubbles: true }}));
            }} else if (lbl.includes("CANTIDAD (Q)") || lbl.includes("CANTIDAD")) {{
              cantInput = inp;
            }}
          }}
          if (cantInput) {{
            cantInput.scrollIntoView({{ behavior: "smooth", block: "center" }});
            cantInput.focus();
            cantInput.select();
          }}
        }} catch(e) {{}}
        if (detenerCamara !== false) {{
          detener();
        }}
      }}

      function detener() {{
        activo = false;
        try {{ if (reader) {{ reader.reset(); }} }} catch(e) {{}}
        if (stream) {{
          stream.getTracks().forEach(t => t.stop());
          stream = null;
        }}
        video.style.display = "none";
        phScan.style.display = "block";
      }}

      async function capturarCuadro() {{
        if (!stream || video.readyState < 2) {{
          statusEl.textContent = "Primero activa la cámara para capturar un cuadro.";
          return;
        }}
        const w = video.videoWidth || 1280;
        const h = video.videoHeight || 720;
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(video, 0, 0, w, h);
        const dataUrl = canvas.toDataURL("image/png");
        let detectado = "";
        try {{
          if ("BarcodeDetector" in window) {{
            const det = new window.BarcodeDetector({{ formats: ["code_128", "code_39", "ean_13", "ean_8", "upc_a", "upc_e", "qr_code", "itf", "codabar"] }});
            const codes = await det.detect(canvas);
            if (codes && codes.length > 0 && codes[0].rawValue) {{
              detectado = codes[0].rawValue;
            }}
          }}
        }} catch(e) {{}}
        if (!detectado && window.ZXing && window.ZXing.BrowserMultiFormatReader) {{
          try {{
            const tmpReader = new window.ZXing.BrowserMultiFormatReader();
            const res = await tmpReader.decodeFromImageUrl(dataUrl);
            if (res && res.text) {{
              detectado = res.text;
            }}
          }} catch(e) {{}}
        }}
        if (detectado) {{
          aplicarCodigo(detectado, true);
        }} else {{
          const sug = (inpManual.value || "").trim().toUpperCase();
          if (sug) {{
            aplicarCodigo(sug, true);
          }} else {{
            statusEl.textContent = "Cuadro capturado sin lectura automática: acerca más el código o escríbelo arriba y pulsa Usar SKU.";
            inpManual.focus();
          }}
        }}
      }}

      async function iniciar() {{
        detener();
        activo = true;
        if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {{
          statusEl.textContent = "Sin acceso directo a cámara. Digita el SKU arriba o usa Foto Evidencia.";
          return;
        }}
        try {{
          stream = await navigator.mediaDevices.getUserMedia({{
            video: {{
              facingMode: {{ ideal: "environment" }},
              width: {{ ideal: 1280 }},
              height: {{ ideal: 720 }}
            }},
            audio: false
          }});
          video.srcObject = stream;
          phScan.style.display = "none";
          video.style.display = "block";
          await video.play();
          statusEl.textContent = "Cámara activa: enfoca el código o pulsa Capturar cuadro ahora.";

          if (window.ZXing && window.ZXing.BrowserMultiFormatReader) {{
            reader = new window.ZXing.BrowserMultiFormatReader();
            reader.decodeFromVideoElementContinuously(video, (result, err) => {{
              if (!activo) return;
              if (result && result.text) {{
                aplicarCodigo(result.text, true);
              }}
            }});
          }} else if ("BarcodeDetector" in window) {{
            const det = new window.BarcodeDetector({{ formats: ["code_128", "code_39", "ean_13", "ean_8", "upc_a", "upc_e", "qr_code", "itf"] }});
            const iv = setInterval(async () => {{
              if (!activo || video.readyState < 2) return;
              try {{
                const codes = await det.detect(video);
                if (codes && codes.length > 0 && codes[0].rawValue) {{
                  clearInterval(iv);
                  aplicarCodigo(codes[0].rawValue, true);
                }}
              }} catch(e) {{}}
            }}, 300);
          }}
        }} catch(err) {{
          statusEl.textContent = "No se pudo abrir cámara. Digita el SKU arriba o usa Foto Evidencia.";
        }}
      }}

      btnStart.addEventListener("click", iniciar);
      btnSnap.addEventListener("click", capturarCuadro);
      btnStop.addEventListener("click", detener);
      btnUseManual.addEventListener("click", () => aplicarCodigo(inpManual.value, false));
    }})();
    </script>
    """
    components.html(html_code, height=430)


def excel_bytes(df, nac_rows=None, muestreo=True):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DetalleTomaFisica"
    fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    font = Font(color="FFFFFF", bold=True, size=11)
    ws.append(list(df.columns))
    for i in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=i)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r_vals in df.itertuples(index=False, name=None):
        ws.append(list(r_vals))
        for c_idx, cname in enumerate(df.columns, start=1):
            c = ws.cell(row=ws.max_row, column=c_idx)
            if cname in {"Bodega", "COD_SKU", "ALTERNO", "ID_Tx", "Ubicacion", "PROVEEDOR"}:
                c.number_format = "@"
                if c.value is not None:
                    c.value = norm_mayus_limpio(c.value)
            elif cname == "COSTO UNITARIO":
                c.number_format = "#,##0.0000"
            elif cname in {"COSTO TOTAL", "COSTO DIF", "Q", "QSA_Teorico", "DIFERENCIA"}:
                c.number_format = "#,##0.00"
    for col in ws.columns:
        ml = max((len(str(cell.value or "")) for cell in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max(ml + 3, 12), 45)
    ws.freeze_panes = "A2"

    ws_sku = wb.create_sheet(title="ConsolidadoSKU")
    df_sku_c = consolidar_df_por_sku(df)
    ws_sku.append(list(df_sku_c.columns))
    for i in range(1, len(df_sku_c.columns) + 1):
        cell = ws_sku.cell(row=1, column=i)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r_vals in df_sku_c.itertuples(index=False, name=None):
        ws_sku.append(list(r_vals))
    ws_sku.freeze_panes = "A2"

    ws_av = wb.create_sheet(title="MetricasAvance")
    av_cols = [
        "Periodo", "Bodega", "Custodio", "Registros_Tx", "SKU_Teoricos",
        "SKU_Contados", "SKU_Pendientes", "Avance_Pct", "Cantidad_Q",
        "Sobrantes", "Faltantes", "Cuadrados", "Costo_Total", "Costo_Dif_Total"
    ]
    ws_av.append(av_cols)
    for i in range(1, len(av_cols) + 1):
        cell = ws_av.cell(row=1, column=i)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r_vals in construir_metricas_avance(df, nac_rows or [], muestreo=muestreo):
        ws_av.append(r_vals)
    for r_i in range(2, ws_av.max_row + 1):
        ws_av.cell(row=r_i, column=2).number_format = "@"
        for col_num in [8, 9, 13, 14]:
            ws_av.cell(row=r_i, column=col_num).number_format = "#,##0.00"
    for col in ws_av.columns:
        ml = max((len(str(cell.value or "")) for cell in col), default=12)
        ws_av.column_dimensions[col[0].column_letter].width = min(max(ml + 3, 14), 30)
    ws_av.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _parsear_filas_a_maestro_y_nac(filas):
    dm = {}
    nac = []
    if not filas:
        return dm, nac
    cab = [norm_mayus_limpio(c) if c is not None else "" for c in filas[0]]
    ck = next((i for i, h in enumerate(cab) if any(k in h for k in ["COD_SKU", "CODSKU", "SKU", "CODIGO", "CÓDIGO", "ART"])), 0)
    for row in filas[1:]:
        if not row or len(row) <= ck:
            continue
        sv = norm_mayus_limpio(row[ck])
        if not sv or sv in ["NONE", "NAN"]:
            continue
        dm.setdefault(sv, {})
        for i, val in enumerate(row):
            if i < len(cab) and cab[i] and val is not None:
                vs = str(val).strip()
                if vs and not vs.startswith("#") and vs.upper() not in ["NONE", "NAN"]:
                    dm[sv][cab[i]] = vs
        fila_nac = [""] * max(len(row), 23)
        fila_nac[0] = sv
        for i, val in enumerate(row):
            if i < len(fila_nac) and val is not None:
                fila_nac[i] = val
        for i, h in enumerate(cab):
            if i < len(row) and row[i] is not None:
                nb = norm_bod(h)
                if nb in MAPEO_BOD_COL:
                    fila_nac[MAPEO_BOD_COL[nb]] = row[i]
                if "COSTO" in h and ("UNIT" in h or h == "COSTO" or "W" == h):
                    fila_nac[IDX_COL_W] = row[i]
                if h in ["V", "MOV", "HAB", "MUESTREO"]:
                    fila_nac[IDX_COL_V] = row[i]
        nac.append(fila_nac)
    if nac:
        enc_nac = [""] * 23
        enc_nac[0] = "COD_SKU"
        for b_k, col_i in MAPEO_BOD_COL.items():
            enc_nac[col_i] = b_k
        enc_nac[IDX_COL_V] = "V"
        enc_nac[IDX_COL_W] = "COSTO"
        nac = [enc_nac] + nac
    return dm, nac


@st.cache_data(ttl=120, show_spinner=False)
def cargar_fuente_gs_externa(url_gs):
    u = str(url_gs or "").strip()
    if not u:
        return {}, []
    try:
        r = requests.get(u, headers={"User-Agent": "Mozilla/5.0"}, verify=False, timeout=15, allow_redirects=True)
        if r.status_code != 200:
            return {}, []
        txt = (r.text or "").strip()
        if txt.startswith("[") or txt.startswith("{"):
            data = json.loads(txt)
            filas_raw = data if isinstance(data, list) else next((data[k] for k in ["catalogo", "maestro", "nac", "data", "rows", "items"] if isinstance(data.get(k), list)), [])
            if filas_raw and isinstance(filas_raw[0], dict):
                cols = list(filas_raw[0].keys())
                filas = [cols] + [[d.get(c, "") for c in cols] for d in filas_raw]
                return _parsear_filas_a_maestro_y_nac(filas)
            if filas_raw and isinstance(filas_raw[0], list):
                return _parsear_filas_a_maestro_y_nac(filas_raw)
        reader = csv.reader(io.StringIO(txt))
        filas_csv = [row for row in reader if any(str(c or "").strip() for c in row)]
        return _parsear_filas_a_maestro_y_nac(filas_csv)
    except Exception as e:
        log_err(f"GS Catálogo: {e}", origen="GS_CAT")
        return {}, []


@st.cache_data(show_spinner=False)
def cargar_pivot(p, bytes_subidos=None):
    dm = {}
    wb = None
    try:
        if bytes_subidos:
            wb = openpyxl.load_workbook(io.BytesIO(bytes_subidos), data_only=True)
        elif os.path.exists(p):
            wb = openpyxl.load_workbook(p, data_only=True)
        if wb is None:
            return dm
        for sh in wb.sheetnames:
            rows = list(wb[sh].iter_rows(values_only=True))
            for idx_r, r_try in enumerate(rows[:6]):
                cand = [norm_mayus_limpio(c) if c is not None else "" for c in r_try]
                if any("SKU" in h or "COD" in h or "ART" in h for h in cand):
                    ck = next((i for i, h in enumerate(cand) if "SKU" in h or "COD" in h or "ART" in h), None)
                    if ck is not None:
                        for row in rows[idx_r + 1:]:
                            if len(row) > ck and row[ck] is not None:
                                sv = norm_mayus_limpio(row[ck])
                                if sv and sv not in ["NONE", "NAN"]:
                                    dm.setdefault(sv, {})
                                    for i, val in enumerate(row):
                                        if i < len(cand) and cand[i] and val is not None:
                                            vs = str(val).strip()
                                            if vs and not vs.startswith("#") and vs.upper() not in ["NONE", "NAN"]:
                                                dm[sv][cand[i]] = vs
                    break
    except Exception as e:
        log_err(f"PIVOT: {e}", origen="PIVOT")
    return dm


@st.cache_data(show_spinner=False)
def cargar_nac(p, bytes_subidos=None):
    try:
        if bytes_subidos:
            wb = openpyxl.load_workbook(io.BytesIO(bytes_subidos), data_only=True)
            return list(wb.active.iter_rows(values_only=True))
        if os.path.exists(p):
            wb = openpyxl.load_workbook(p, data_only=True)
            return list(wb.active.iter_rows(values_only=True))
        return []
    except Exception as e:
        log_err(f"NAC: {e}", origen="NAC")
        return []


# ==============================================================================
# PANEL IZQUIERDO LIMPIO (SOLO LO OPERATIVO Y NECESARIO)
# ==============================================================================
def limpiar_controles_bodega():
    st.session_state["cap_v"] = st.session_state.get("cap_v", 0) + 1
    st.session_state["sku_pend_scan"] = ""
    st.session_state["scan_confirmado_cant"] = False
    st.session_state["tx_reverso_draft"] = ""
    st.session_state.pop("flt_bod", None)
    st.session_state.pop("flt_ciu", None)


with st.sidebar.expander("1. Operación, Usuario & Bodega", expanded=True):
    usuario = st.text_input("Correo Usuario / Custodio:", value="alexpajaritoruiz@icloud.com", key="usr_in").strip().lower()
    st.session_state["usr_act"] = usuario
    es_admin = usuario in {"alexpajaritoruiz@icloud.com", "jlogistica.eit@gmail.com"} or "ALEX" in usuario.upper() or "PAJARI" in usuario.upper()
    bods_perm = BODEGAS_OFICIALES if es_admin else CUSTODIO_BOD.get(usuario, ["01"])

    tipo_proc = st.radio("Modo:", ["Control Mensual de Inventario (Muestreo)", "Toma Física Completa"], key="modo_proc_in")
    if "bod_act" not in st.session_state or st.session_state["bod_act"] not in bods_perm:
        st.session_state["bod_act"] = bods_perm[0]

    bod_norm = norm_bod(st.radio("Bodega:", bods_perm, key="bod_act", on_change=limpiar_controles_bodega))
    ciudades = ["Quito"] if bod_norm in ["01", "05", "5V"] else (["Guayaquil"] if bod_norm in ["03", "3V"] else (["Cuenca"] if bod_norm == "04" else ["Quito", "Guayaquil", "Cuenca", "Otras"]))
    ciudad_def = ciudades[0].upper()

    f_ini = st.date_input("Desde:", value=date.today(), key="f_ini_in")
    f_fin = st.date_input("Hasta:", value=date.today(), key="f_fin_in")
    periodo_act = f"{f_ini.strftime('%Y-%m-%d')} al {f_fin.strftime('%Y-%m-%d')}"

st.sidebar.markdown("---")

if "per_abierto" not in st.session_state:
    st.session_state["per_abierto"] = False

with st.sidebar.expander("2. Periodo, PIN & Enlace Custodios", expanded=es_admin):
    if es_admin:
        st.session_state["per_abierto"] = st.toggle("Habilitar periodo de auditoría (Admin)", value=st.session_state["per_abierto"], key="per_abierto_tgl")
        st.session_state["pin_reverso_act"] = st.text_input("PIN Base Reversos (Admin):", value=st.session_state["pin_reverso_act"], key="pin_rev_admin_in").strip()
    else:
        st.caption("Estado periodo: " + ("HABILITADO" if st.session_state["per_abierto"] else "INHABILITADO"))

    url_wa_raw = st.text_input("URL Custodios:", value="https://eit-toma-fisica-4wrzxgv3pgzfdr8agv9mtm.streamlit.app/", key="url_wa_in")
    url_wa_limpia = "".join(str(url_wa_raw or "").split()).strip()
    if url_wa_limpia:
        msg_wa = f"CONTROL DE INVENTARIO EIT\n\nEnlace para registro de tomas físicas:\n{url_wa_limpia}"
        st.markdown(
            f'<a href="https://api.whatsapp.com/send?text={urllib.parse.quote(msg_wa)}" target="_blank" style="display:block;text-align:center;padding:8px;background:#16A34A;color:#FFF!important;font-weight:700;border-radius:6px;text-decoration:none;">Enviar por WhatsApp</a>',
            unsafe_allow_html=True,
        )

st.sidebar.markdown("---")

with st.sidebar.expander("3. Filtros de Estadística & Reportes", expanded=False):
    incluir_gs_rep = st.checkbox("Incluir registros de Google Sheets", value=False, key="chk_inc_gs")
    origen_label = "Local + Google Sheets" if incluir_gs_rep else "Local"

    def resetear_filtros_estadistica():
        st.session_state["flt_per"] = "TODOS"
        st.session_state["flt_bod"] = "TODAS"
        st.session_state["flt_ciu"] = "TODAS"
        st.session_state["flt_cus"] = "TODOS"
        st.session_state["flt_sku"] = "TODOS"
        st.session_state["flt_prv"] = "TODOS"

    def sincronizar_filtros_operacion():
        st.session_state["flt_per"] = periodo_act
        st.session_state["flt_bod"] = bod_norm
        st.session_state["flt_ciu"] = ciudad_def
        st.session_state["flt_cus"] = "TODOS"
        st.session_state["flt_sku"] = "TODOS"
        st.session_state["flt_prv"] = "TODOS"

    f_col1, f_col2 = st.columns(2)
    f_col1.button("Periodo/Bod act.", on_click=sincronizar_filtros_operacion, use_container_width=True, key="btn_flt_act")
    f_col2.button("Ver todo", on_click=resetear_filtros_estadistica, use_container_width=True, key="btn_flt_all")

    df_base_rep = prep_df(datos_combinados(incluir_gs=incluir_gs_rep))
    muestreo_rep = "Muestreo" in tipo_proc

    periodos_disp = ["TODOS"] + sorted(set([periodo_act] + [p for p in df_base_rep["Periodo"].unique() if p]), reverse=True)
    if "flt_per" not in st.session_state or st.session_state["flt_per"] not in periodos_disp:
        st.session_state["flt_per"] = periodo_act if periodo_act in periodos_disp else "TODOS"
    f_per = st.selectbox("Periodo:", periodos_disp, key="flt_per")

    bodegas_disp = ["TODAS"] + sorted(set(BODEGAS_OFICIALES + [b for b in df_base_rep["Bodega"].unique() if b]))
    if "flt_bod" not in st.session_state or st.session_state["flt_bod"] not in bodegas_disp:
        st.session_state["flt_bod"] = bod_norm if bod_norm in bodegas_disp else "TODAS"
    f_bod = st.selectbox("Bodega:", bodegas_disp, key="flt_bod")

    ciudades_disp = ["TODAS"] + sorted(set(["QUITO", "GUAYAQUIL", "CUENCA", ciudad_def] + [c for c in df_base_rep["Ciudad"].unique() if c]))
    if "flt_ciu" not in st.session_state or st.session_state["flt_ciu"] not in ciudades_disp:
        st.session_state["flt_ciu"] = ciudad_def if ciudad_def in ciudades_disp else "TODAS"
    f_ciu = st.selectbox("Ciudad:", ciudades_disp, key="flt_ciu")

    custodios_disp = ["TODOS"] + sorted([c for c in df_base_rep["Correo_Usuario"].unique() if c]) if not df_base_rep.empty else ["TODOS"]
    if "flt_cus" not in st.session_state or st.session_state["flt_cus"] not in custodios_disp:
        st.session_state["flt_cus"] = "TODOS"
    f_cus = st.selectbox("Custodio:", custodios_disp, key="flt_cus")

    skus_disp = ["TODOS"] + sorted([s for s in df_base_rep["COD_SKU"].unique() if s]) if not df_base_rep.empty else ["TODOS"]
    if "flt_sku" not in st.session_state or st.session_state["flt_sku"] not in skus_disp:
        st.session_state["flt_sku"] = "TODOS"
    f_sku = st.selectbox("COD_SKU:", skus_disp, key="flt_sku")

    provs_disp = ["TODOS"] + sorted([p for p in df_base_rep["PROVEEDOR"].unique() if p]) if not df_base_rep.empty else ["TODOS"]
    if "flt_prv" not in st.session_state or st.session_state["flt_prv"] not in provs_disp:
        st.session_state["flt_prv"] = "TODOS"
    f_prv = st.selectbox("Proveedor:", provs_disp, key="flt_prv")

df_rep = df_base_rep.copy()
if not df_rep.empty:
    if f_per != "TODOS":
        df_rep = df_rep[df_rep["Periodo"] == str(f_per).strip()]
    else:
        ts_dt = pd.to_datetime(df_rep["Timestamp"], errors="coerce")
        mask_f = ts_dt.isna() | ((ts_dt.dt.date >= f_ini) & (ts_dt.dt.date <= f_fin))
        df_rep = df_rep[mask_f]
    if f_bod != "TODAS":
        df_rep = df_rep[df_rep["Bodega"] == norm_bod(f_bod)]
    if f_ciu != "TODAS":
        df_rep = df_rep[df_rep["Ciudad"] == norm_mayus_limpio(f_ciu)]
    if f_cus != "TODOS":
        df_rep = df_rep[df_rep["Correo_Usuario"] == str(f_cus).strip().lower()]
    if f_sku != "TODOS":
        df_rep = df_rep[df_rep["COD_SKU"] == norm_mayus_limpio(f_sku)]
    if f_prv != "TODOS":
        df_rep = df_rep[df_rep["PROVEEDOR"] == norm_mayus_limpio(f_prv)]

df_sku_rep = consolidar_df_por_sku(df_rep)

st.sidebar.markdown("---")

with st.sidebar.expander("4. Sincronizar datos fuente", expanded=True):
    fecha_nac = fecha_arch(nac_path)
    fecha_pivot = fecha_arch(pivot_path)
    st.caption(
        f"Los teóricos se obtienen de **NAC.xlsx** (`{fecha_nac}`) y **PRODUCTO - PIVOT.xlsm** (`{fecha_pivot}`). "
        "No requiere ejecutar `.bat` si las fuentes ya están actualizadas o cargadas en Google Sheets."
    )
    if st.button("Sincronizar datos fuente", use_container_width=True, key="btn_sync_fuentes"):
        leer_gs.clear()
        cargar_pivot.clear()
        cargar_nac.clear()
        cargar_fuente_gs_externa.clear()
        st.rerun()

    if es_admin and not df_rep.empty:
        st.download_button(
            f"Exportar Excel ({len(df_rep)})",
            data=excel_bytes(df_rep, nac_rows=[], muestreo=muestreo_rep),
            file_name=f"reporte_toma_fisica_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="btn_dl_side",
        )

dm_gs, nac_gs = cargar_fuente_gs_externa(st.session_state["fuente_gs_url"])
dm_loc = cargar_pivot(pivot_path, None)
nac_loc = cargar_nac(nac_path, None)

dict_maestro = dict(dm_gs)
for k_s, v_d in dm_loc.items():
    dict_maestro.setdefault(k_s, {}).update(v_d)

nac_rows = nac_gs if nac_gs else nac_loc
regs_json = cargar_json()
logs_json = cargar_logs()
f_script = fecha_arch(RUTA_SCRIPT)
cap_v = st.session_state["cap_v"]

# ==============================================================================
# CABECERA + PESTAÑAS
# ==============================================================================
st.markdown('<div class="main-title notranslate" translate="no">Toma Física & Muestreo V7 gPro</div>', unsafe_allow_html=True)
st.caption(
    f"Script: `{RUTA_SCRIPT}` | Modificado: `{f_script}` | NAC: `{fecha_nac}` | PIVOT: `{fecha_pivot}` | "
    f"Periodo: `{'HABILITADO' if st.session_state['per_abierto'] else 'INHABILITADO'}`"
)

if st.session_state["errores_ui"]:
    st.warning("Incidencias registradas en esta sesión: " + " || ".join(st.session_state["errores_ui"][-2:]))

reversos_pend_all = [
    r for r in regs_json
    if num_es(r.get("Q", 0), max_abs=1e12) < 0 and norm_estado_rev(r.get("Estado_Reverso", ""), r.get("Q", 0)) == "PENDIENTE_APROBACION"
]
if es_admin and reversos_pend_all:
    st.warning(
        f"Notificación en bandeja Admin: tienes **{len(reversos_pend_all)}** reverso(s) pendiente(s) de autorización. "
        "Revísalos en la pestaña *Dashboard Estadístico & Avance Visual*."
    )

tab_op, tab_dash, tab_etq, tab_exp = st.tabs([
    "Operación y Captura de Campo",
    "Dashboard Estadístico & Avance Visual",
    f"Etiquetas A4 (~5,8×2,5 cm) [{len(st.session_state['cola_etiquetas'])}]",
    f"Exportación, Detalle & {PESTANA_LOGS_GS}",
])

with tab_op:
    st.markdown('<div class="card-box notranslate" translate="no">', unsafe_allow_html=True)
    if not st.session_state["per_abierto"]:
        st.warning("Periodo INHABILITADO: habilítalo desde la carpeta 2 del panel izquierdo cuando corresponda.")
    if not nac_rows and not dict_maestro:
        st.info("Sin fuente `NAC` / `PIVOT` conectada en este entorno; puedes digitar o escanear `COD_SKU` directo.")

    c1, c2, c3 = st.columns(3)
    ciudad = c1.selectbox("Ciudad:", ciudades, key="ciu_in")
    c2.markdown(f"**Bodega:** `{bod_norm}`")
    ubic_raw = c3.text_input("Ubicación (MAYÚSCULAS):", placeholder="EJ: N2611", key=f"ubic_in_{cap_v}")
    ubic = norm_mayus_limpio(ubic_raw)

    idx_b = resolver_col_bodega(nac_rows, bod_norm)
    muestreo = "Muestreo" in tipo_proc
    ya_aud = {
        norm_mayus_limpio(r.get("COD_SKU", ""))
        for r in regs_json
        if str(r.get("Periodo", "")).strip() == periodo_act and norm_bod(r.get("Bodega", "")) == bod_norm and es_computable_en_consolidado(r)
    }

    skus_unicos, desc_map = obtener_skus_teoricos_bodega(nac_rows, dict_maestro, bod_norm, muestreo=muestreo, regs_respaldo=regs_json)
    tot_teo_bod = len(skus_unicos)
    aud_bod_cnt = len([s for s in skus_unicos if s in ya_aud])
    pend_bod_cnt = max(tot_teo_bod - aud_bod_cnt, 0)
    st.caption(f"Guía SKU Bodega `{bod_norm}` (SKUs con movimiento en periodo y saldo > 0) -> Teóricos: **{tot_teo_bod}** | Auditados: **{aud_bod_cnt}** | Pendientes: **{pend_bod_cnt}**")

    opts = ["-- SELECCIONE O ESCANEE SKU --"] + [
        f"{s} | {'[AUDITADO]' if s in ya_aud else '[PENDIENTE]'} {desc_map.get(s, '')}" for s in skus_unicos
    ]
    sel_combo = st.selectbox("Catálogo Teórico:", opts, key=f"cb_{bod_norm}_{cap_v}")
    pre_sku = norm_mayus_limpio(sel_combo.split(" | ")[0]) if sel_combo != "-- SELECCIONE O ESCANEE SKU --" else ""

    components.html(
        """
        <script>
        (function() {
          function aplicarNegrillaAuditado() {
            try {
              const doc = window.parent.document;
              const items = doc.querySelectorAll('li[role="option"], [data-baseweb="menu"] li, [data-baseweb="popover"] li, option');
              items.forEach(el => {
                const txt = (el.textContent || "");
                if (txt.includes("[AUDITADO]")) {
                  el.style.fontWeight = "900";
                  el.style.color = "#0F172A";
                  el.querySelectorAll("span, div").forEach(ch => {
                    ch.style.fontWeight = "900";
                  });
                }
              });
            } catch (e) {}
          }
          aplicarNegrillaAuditado();
          const obs = new MutationObserver(aplicarNegrillaAuditado);
          try { obs.observe(window.parent.document.body, { childList: true, subtree: true }); } catch (e) {}
        })();
        </script>
        """,
        height=0,
    )

    sku_detectado_foto = ""
    with st.expander("Cámara: Lector de Código de Barras & Foto Evidencia", expanded=False):
        t_scan, t_foto = st.tabs(["1. Lector Código de Barras", "2. Foto Evidencia"])
        with t_scan:
            render_lector_codigos_barras(f"{bod_norm}_{cap_v}")
        with t_foto:
            foto_cap = st.camera_input("Tomar fotografía de la etiqueta / código de barras:", key=f"cam_{bod_norm}_{cap_v}")
            foto_up = st.file_uploader("O seleccionar imagen de etiqueta desde galería:", type=["png", "jpg", "jpeg", "webp"], key=f"fup_etq_{bod_norm}_{cap_v}")
            img_fuente = foto_cap.getvalue() if foto_cap is not None else (foto_up.getvalue() if foto_up is not None else None)
            if img_fuente:
                sku_detectado_foto = decodificar_sku_desde_imagen_bytes(img_fuente)
                if sku_detectado_foto:
                    if st.session_state.get("sku_pend_scan") != sku_detectado_foto:
                        st.session_state["sku_pend_scan"] = sku_detectado_foto
                        st.session_state["scan_confirmado_cant"] = False
                    st.success(f"Código leído en foto: `{sku_detectado_foto}` — ingresa la cantidad antes de REGISTRAR.")
                else:
                    st.caption("Foto guardada. Si el código no se leyó automáticamente, escríbelo o pégalo en `COD_SKU`.")

    if st.session_state.get("sku_pend_scan") and not pre_sku:
        pre_sku = st.session_state["sku_pend_scan"]

    c4, c5, c6 = st.columns([2.5, 1.5, 1])
    sku_raw = c4.text_input("COD_SKU (MAYÚSCULAS):", value=pre_sku, key=f"sku_{bod_norm}_{cap_v}")
    cod_sku = norm_mayus_limpio(sku_raw)

    if cod_sku and st.session_state.get("sku_pend_scan") == cod_sku and not st.session_state.get("scan_confirmado_cant"):
        st.markdown(
            f'<div class="scan-alert notranslate" translate="no">Código capturado: {html.escape(cod_sku)} — Ahora digita o ajusta la <u>Cantidad (Q)</u> antes de pulsar REGISTRAR.</div>',
            unsafe_allow_html=True,
        )

    if sku_detectado_foto:
        _inyectar_sku_y_enfocar_cantidad(sku_detectado_foto)

    q_key = f"q_{bod_norm}_{cap_v}"
    valor_inicial_q = 0 if (cod_sku and st.session_state.get("sku_pend_scan") == cod_sku) else 1
    cant = c5.number_input("Cantidad (Q):", step=1, value=valor_inicial_q, key=q_key)
    cant_modificada = (float(cant) != float(valor_inicial_q)) or (q_key in st.session_state and float(st.session_state[q_key]) != float(valor_inicial_q))
    if cant != 0 and cod_sku and st.session_state.get("sku_pend_scan") == cod_sku:
        st.session_state["scan_confirmado_cant"] = True

    motivo = ""
    pin_ingresado = ""
    pin_valido_cust = False
    tx_borrador_pin = ""
    pin_tx_esperado = ""

    if cant < 0:
        if not st.session_state.get("tx_reverso_draft"):
            st.session_state["tx_reverso_draft"] = f"TX-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        tx_borrador_pin = st.session_state["tx_reverso_draft"]
        pin_tx_esperado = generar_pin_transaccion(
            st.session_state["pin_reverso_act"],
            tx_id=tx_borrador_pin,
            bod=bod_norm,
            cod_sku=cod_sku,
            q_val=cant,
        )
        r_m1, r_m2 = st.columns([2.2, 1.3])
        motivo = norm_mayus_limpio(r_m1.text_input("Motivo Reverso (Col K obligatorio si Q < 0):", key=f"mot_{bod_norm}_{cap_v}"))
        if not es_admin:
            pin_ingresado = r_m2.text_input(f"PIN único Tx ({tx_borrador_pin}):", max_chars=6, key=f"pin_{bod_norm}_{cap_v}").strip()
            pin_valido_cust = bool(pin_ingresado) and (pin_ingresado == pin_tx_esperado)
            if pin_ingresado and not pin_valido_cust:
                st.warning(f"PIN incorrecto para `{tx_borrador_pin}`: el reverso quedará **PENDIENTE_APROBACION** para revisión del admin.")
            elif pin_valido_cust:
                st.success(f"PIN válido para `{tx_borrador_pin}`: reverso autorizado y aplicado al consolidado.")
            else:
                st.info(f"Sin PIN ingresado para `{tx_borrador_pin}`: el reverso queda **PENDIENTE_APROBACION** en la bandeja del administrador.")
        else:
            st.caption(f"PIN único para esta transacción (`{tx_borrador_pin}`): **`{pin_tx_esperado}`**")
    else:
        st.session_state["tx_reverso_draft"] = ""

    c6.write("")
    c6.write("")
    ok_rev = (cant >= 0) or bool(motivo)
    ubic_valida = bool(ubic) and ubic not in ["UBICACIÓN", "UBICACION", "ESPECIFIQUE"]
    es_scan_sin_cant = bool(cod_sku and st.session_state.get("sku_pend_scan") == cod_sku and cant == 0)
    puede = st.session_state["per_abierto"] and ubic_valida and bool(cod_sku) and (cant != 0) and ok_rev and not es_scan_sin_cant
    btn_reg = c6.button("REGISTRAR", use_container_width=True, disabled=not puede, key="btn_reg_tx")

    e_col1, e_col2, e_col3 = st.columns([2.2, 1.0, 1.8])
    marcar_etq = e_col1.checkbox("Marcar este SKU para etiqueta (~5,8×2,5 cm)", value=False, key=f"chk_etq_{bod_norm}_{cap_v}")
    copias_etq = e_col2.number_input("Copias:", min_value=1, max_value=30, value=1, step=1, key=f"cop_etq_{bod_norm}_{cap_v}")
    if e_col3.button("Añadir SKU a cola etiqueta", use_container_width=True, disabled=not bool(cod_sku), key="btn_add_etq_rapido"):
        d_etq = dict_maestro.get(cod_sku, {}).get("DESCRIPCIÓN") or dict_maestro.get(cod_sku, {}).get("DESCRIPCION") or desc_map.get(cod_sku, "")
        agregar_etiqueta_cola(cod_sku, d_etq, bod_norm, ubic, usuario_lbl=usuario, copias=copias_etq)
        st.session_state["ultimo_tx_resumen"] = (
            f"<b>Cola de etiquetas actualizada:</b> {html.escape(cod_sku)} ({int(copias_etq)} copia(s)) agregado para impresión A4."
        )
        st.rerun()

    faltan = []
    if not st.session_state["per_abierto"]:
        faltan.append("Periodo habilitado")
    if not ubic_valida:
        faltan.append("Ubicación")
    if not cod_sku:
        faltan.append("COD_SKU")
    if cant == 0:
        faltan.append("Anotar Cantidad (Q) distinta de 0")
    if not ok_rev:
        faltan.append("Motivo Reverso")
    if faltan:
        st.caption("Falta para habilitar REGISTRAR: " + ", ".join(faltan) + ".")
    st.markdown("</div>", unsafe_allow_html=True)

    st_teo, st_nac, c_unit_vista = 0.0, 0.0, 0.0
    if cod_sku:
        art = dict_maestro.get(cod_sku, {})

        def flex_exacto(*ks):
            for k in ks:
                for ak, v in art.items():
                    if ak.strip().upper() == k.upper() and v:
                        return str(v).strip()
            for k in ks:
                for ak, v in art.items():
                    if k.upper() in ak.upper() and v:
                        return str(v).strip()
            return ""

        if len(nac_rows) > 1:
            for row in nac_rows[1:]:
                if row and row[0] and norm_mayus_limpio(row[0]) == cod_sku:
                    st_nac = num_es(row[IDX_TOTAL_NAC], max_abs=1e12) if IDX_TOTAL_NAC < len(row) else 0.0
                    st_teo = num_es(row[idx_b], max_abs=1e12) if (idx_b < len(row) and sku_habilitado_movimiento(row, muestreo=muestreo)) else 0.0
                    c_unit_vista = obtener_costo_nac_col_w(row)
                    break
        if c_unit_vista <= 0:
            c_unit_vista = next(
                (limpiar_costo_unit(v) for k, v in art.items() if any(t in k.upper() for t in ["COSTO", "PRECIO", "VAL_UNIT", "V_UNIT", "P_UNIT"]) and limpiar_costo_unit(v) > 0),
                0.0,
            )

        prov_cod = norm_mayus_limpio(flex_exacto("PROVEEDOR", "COD_PROV", "PROV"))
        prev_comp = sum(
            num_es(r.get("Q", 0), max_abs=1e12)
            for r in regs_json
            if str(r.get("Periodo", "")).strip() == periodo_act
            and norm_bod(r.get("Bodega", "")) == bod_norm
            and norm_mayus_limpio(r.get("COD_SKU", "")) == cod_sku
            and es_computable_en_consolidado(r)
        )
        dif_actual_sin_q = round(float(prev_comp) - float(st_teo), 4)
        linea_1 = (
            f"<b>COD_SKU:</b> {html.escape(cod_sku)} &nbsp;|&nbsp; "
            f"<b>Desc:</b> {html.escape(flex_exacto('DESCRIPCIÓN', 'DESCRIPCION', 'DESC'))} &nbsp;|&nbsp; "
            f"<b>Und:</b> {html.escape(flex_exacto('UNIDAD', 'UND'))} &nbsp;|&nbsp; "
            f"<b>Cód. Proveedor:</b> <code>{html.escape(prov_cod)}</code> &nbsp;|&nbsp; "
            f"<b>Costo Unit:</b> {fmt_moneda_ec(c_unit_vista, 4)}"
        )
        if cant_modificada:
            aplica_rev_ya = (cant >= 0) or es_admin or pin_valido_cust
            aporte_proy = float(cant) if aplica_rev_ya else 0.0
            q_acum_calc = round(float(prev_comp) + aporte_proy, 4)
            dif_proy = round(q_acum_calc - float(st_teo), 4)
            nota_proy = " (Reverso pendiente de aprobación)" if (cant < 0 and not aplica_rev_ya) else ""
            linea_2 = (
                f"<b>Acum. Anterior:</b> {fmt_miles_comas(prev_comp, 2)} &nbsp;|&nbsp; "
                f"<b>Dif. Actual:</b> {fmt_miles_comas(dif_actual_sin_q, 2)} &nbsp;|&nbsp; "
                f"<b>Proyectado Q:</b> {fmt_miles_comas(q_acum_calc, 2)} &nbsp;|&nbsp; "
                f"<b>Dif. Proy:</b> {fmt_miles_comas(dif_proy, 2)}{html.escape(nota_proy)}"
            )
        else:
            linea_2 = (
                f"<b>Acum. Anterior:</b> {fmt_miles_comas(prev_comp, 2)} &nbsp;|&nbsp; "
                f"<b>Dif. Actual:</b> {fmt_miles_comas(dif_actual_sin_q, 2)} &nbsp;|&nbsp; "
                f"<span style='color:#475569;'><i>Proyección oculta hasta ajustar Cantidad (Q)</i></span>"
            )
        st.markdown(
            f'<div class="notranslate" translate="no" style="background:#EFF6FF;border:1px solid #93C5FD;border-radius:8px;padding:10px 14px;margin-bottom:10px;font-size:14px;line-height:1.5;">'
            f'{linea_1}<br>{linea_2}'
            f'</div>',
            unsafe_allow_html=True,
        )
        m1, m2 = st.columns(2)
        m1.markdown(f'<div class="saldo-box notranslate" translate="no"><span class="saldo-lbl">QSA_Teorico ({bod_norm})</span><span class="saldo-val">{fmt_miles_comas(st_teo, 0 if float(st_teo).is_integer() else 2)}</span></div>', unsafe_allow_html=True)
        m2.markdown(f'<div class="saldo-box notranslate" translate="no"><span class="saldo-lbl">Nacional</span><span class="saldo-val">{fmt_miles_comas(st_nac, 0 if float(st_nac).is_integer() else 2)}</span></div>', unsafe_allow_html=True)

    if btn_reg:
        firma_actual = f"{periodo_act}|{bod_norm}|{cod_sku}|{float(cant)}"
        ahora_epoch = time.time()
        if (
            st.session_state.get("ultimo_disparo_firma") == firma_actual
            and (ahora_epoch - float(st.session_state.get("ultimo_disparo_ts", 0.0))) < 3.0
        ):
            st.warning("Disparo repetido bloqueado automáticamente (menos de 3 segundos con el mismo SKU y cantidad).")
        else:
            st.session_state["ultimo_disparo_firma"] = firma_actual
            st.session_state["ultimo_disparo_ts"] = ahora_epoch

            art = dict_maestro.get(cod_sku, {})

            def flex_r(*ks):
                for k in ks:
                    for ak, v in art.items():
                        if ak.strip().upper() == k.upper() and v:
                            return str(v).strip()
                for k in ks:
                    for ak, v in art.items():
                        if k.upper() in ak.upper() and v:
                            return str(v).strip()
                return ""

            st_teo_r, c_unit_r = 0.0, 0.0
            if len(nac_rows) > 1:
                for row in nac_rows[1:]:
                    if row and row[0] and norm_mayus_limpio(row[0]) == cod_sku:
                        st_teo_r = num_es(row[idx_b], max_abs=1e12) if (idx_b < len(row) and sku_habilitado_movimiento(row, muestreo=muestreo)) else 0.0
                        c_unit_r = obtener_costo_nac_col_w(row)
                        break
            if c_unit_r <= 0:
                c_unit_r = next(
                    (limpiar_costo_unit(v) for k, v in art.items() if any(t in k.upper() for t in ["COSTO", "PRECIO", "VAL_UNIT", "V_UNIT", "P_UNIT"]) and limpiar_costo_unit(v) > 0),
                    0.0,
                )

            prev_comp = sum(
                num_es(r.get("Q", 0), max_abs=1e12)
                for r in regs_json
                if str(r.get("Periodo", "")).strip() == periodo_act
                and norm_bod(r.get("Bodega", "")) == bod_norm
                and norm_mayus_limpio(r.get("COD_SKU", "")) == cod_sku
                and es_computable_en_consolidado(r)
            )
            now = datetime.now()
            tx_id_nuevo = st.session_state.get("tx_reverso_draft") or f"TX-{now.strftime('%Y%m%d%H%M%S')}"
            pin_tx_actual = generar_pin_transaccion(
                st.session_state["pin_reverso_act"],
                tx_id=tx_id_nuevo,
                bod=bod_norm,
                cod_sku=cod_sku,
                q_val=cant,
            )
            motivo_final = motivo
            if cant < 0:
                if es_admin:
                    est_rev_ini = "APROBADO"
                elif pin_valido_cust:
                    est_rev_ini = "APROBADO"
                    motivo_final = f"{motivo} [PIN_TX_{pin_tx_actual}_AUTORIZADO]".strip()
                else:
                    est_rev_ini = "PENDIENTE_APROBACION"
            else:
                est_rev_ini = "NO_APLICA"

            q_acum_final = round(float(prev_comp) + (float(cant) if est_rev_ini != "PENDIENTE_APROBACION" else 0.0), 4)
            dif_f, est_f, costo_total, costo_dif, c_unit_limpio = calcular_dif_estado_costo(
                q_acum_final, st_teo_r, cant, c_unit_r, est_rev=est_rev_ini
            )
            desc_r = flex_r("DESCRIPCIÓN", "DESCRIPCION", "DESC")
            prov_r = norm_mayus_limpio(flex_r("PROVEEDOR", "COD_PROV", "PROV"))
            nuevo = {
                "Periodo": periodo_act,
                "ID_Tx": tx_id_nuevo,
                "Timestamp": now.strftime("%Y-%m-%d %H:%M:%S"),
                "Correo_Usuario": usuario,
                "Ciudad": norm_mayus_limpio(ciudad),
                "Bodega": bod_norm,
                "Ubicacion": ubic,
                "COD_SKU": cod_sku,
                "Q": float(cant),
                "BOD/SKU": f"{bod_norm}-{cod_sku}",
                "Motivo": motivo_final,
                "Estado_Reverso": est_rev_ini,
                "LINEA DE INVENTARIO": norm_mayus_limpio(flex_r("LÍNEA DE INVENTARIO", "LINEA", "INVENTARIO")),
                "PROVEEDOR": prov_r,
                "ALTERNO": norm_mayus_limpio(flex_r("ALTERNO", "ALT")),
                "DESCRIPCIÓN": desc_r,
                "UNIDAD": norm_mayus_limpio(flex_r("UNIDAD", "UND")),
                "DIFERENCIA": float(dif_f),
                "ESTADO": est_f,
                "COSTO UNITARIO": float(c_unit_limpio),
                "COSTO TOTAL": float(costo_total),
                "COSTO DIF": float(costo_dif),
                "QSA_Teorico": float(st_teo_r),
            }
            if guardar_json(nuevo):
                if marcar_etq:
                    agregar_etiqueta_cola(cod_sku, desc_r, bod_norm, ubic, usuario_lbl=usuario, copias=copias_etq)
                consolidar_estado_sku_en_json(periodo_act, bod_norm, cod_sku, st_teo_r)
                ok_g, est_g, tx_g = enviar_gs(nuevo)
                st.session_state["cap_v"] += 1
                st.session_state["sku_pend_scan"] = ""
                st.session_state["scan_confirmado_cant"] = False
                st.session_state["tx_reverso_draft"] = ""
                if est_rev_ini == "PENDIENTE_APROBACION":
                    st.session_state["ultimo_tx_resumen"] = (
                        f"<b>Último registro guardado:</b> <code>{html.escape(tx_id_nuevo)}</code> | "
                        f"<b>SKU:</b> {html.escape(cod_sku)} | <b>Q:</b> {fmt_miles_comas(cant, 0)} | "
                        f"<b>Estado:</b> PENDIENTE_REVERSO | <b>GS:</b> {html.escape(est_g)} ({html.escape(tx_g)})"
                    )
                else:
                    st.session_state["ultimo_tx_resumen"] = (
                        f"<b>Último registro guardado:</b> <code>{html.escape(tx_id_nuevo)}</code> | "
                        f"<b>SKU:</b> {html.escape(cod_sku)} | <b>Q:</b> {fmt_miles_comas(cant, 0)} | "
                        f"<b>Acum:</b> {fmt_miles_comas(q_acum_final, 2)} | <b>Dif:</b> {fmt_miles_comas(dif_f, 2)} | "
                        f"<b>Estado:</b> {html.escape(est_f)} | <b>GS:</b> {html.escape(est_g)} ({html.escape(tx_g)})"
                    )
                st.rerun()

    st.markdown("---")
    if st.session_state.get("ultimo_tx_resumen"):
        st.markdown(
            f'<div class="last-tx-banner notranslate" translate="no">{st.session_state["ultimo_tx_resumen"]}</div>',
            unsafe_allow_html=True,
        )
    st.markdown("### Últimos Registros")
    prev_df = prep_df(datos_combinados(incluir_gs=False))
    render_tabla_segura(prev_df.iloc[::-1], ["Timestamp", "Bodega", "Ubicacion", "PROVEEDOR", "COD_SKU", "Q", "QSA_Teorico", "DIFERENCIA", "ESTADO", "Estado_Reverso", "COSTO UNITARIO", "COSTO DIF"], max_filas=6)

with tab_dash:
    st.markdown("### Panorama General de Toma Física & Avance Nacional")
    st.caption(
        f"**Fuente de datos:** `{origen_label}` | **Filas base:** `{len(df_base_rep)}` -> **Filtradas:** `{len(df_rep)}` | "
        f"**Filtros:** Periodo `{f_per}` | Bodega `{f_bod}` | Ciudad `{f_ciu}` | Custodio `{f_cus}` | SKU `{f_sku}` | Prov `{f_prv}`"
    )

    if es_admin:
        pend_rev_df = df_base_rep[
            (df_base_rep["Q"] < 0) & (df_base_rep["Estado_Reverso"] == "PENDIENTE_APROBACION")
        ] if not df_base_rep.empty else pd.DataFrame(columns=COLS_REP)
        st.markdown("#### Autorización de Reversos Pendientes (Superusuario)")
        if pend_rev_df.empty:
            st.caption("No hay reversos pendientes de aprobación en este momento.")
        else:
            render_tabla_segura(
                pend_rev_df.iloc[::-1],
                ["ID_Tx", "Timestamp", "Periodo", "Bodega", "COD_SKU", "Q", "QSA_Teorico", "COSTO UNITARIO", "Correo_Usuario", "Motivo"],
                max_filas=20,
            )
            lista_tx_pend = [norm_mayus_limpio(x) for x in pend_rev_df["ID_Tx"].tolist() if norm_mayus_limpio(x)]
            sel_tx_rev = st.selectbox("Seleccionar reverso pendiente:", lista_tx_pend, key="sel_tx_rev_admin")
            reg_sel_rev = next(
                (r for r in pend_rev_df.to_dict(orient="records") if norm_mayus_limpio(r.get("ID_Tx", "")) == sel_tx_rev),
                None,
            )
            pin_tx_sel = generar_pin_transaccion(
                st.session_state["pin_reverso_act"],
                tx_id=sel_tx_rev,
                bod=reg_sel_rev.get("Bodega", "") if reg_sel_rev else "",
                cod_sku=reg_sel_rev.get("COD_SKU", "") if reg_sel_rev else "",
                q_val=reg_sel_rev.get("Q", 0) if reg_sel_rev else 0,
            )
            st.caption(f"PIN único para `{sel_tx_rev}`: **`{pin_tx_sel}`** (PIN base: `{st.session_state['pin_reverso_act']}`)")
            r_col1, r_col2 = st.columns(2)
            if r_col1.button("Aprobar Reverso", use_container_width=True, key="btn_aprobar_rev"):
                ok_r, reg_mod = resolver_reverso_pendiente(sel_tx_rev, aprobar=True, admin_usr=usuario)
                if ok_r and reg_mod is not None:
                    enviar_gs(reg_mod, modo_dup="sobreescribir")
                    st.session_state["ultimo_tx_resumen"] = f"<b>Reverso aprobado:</b> <code>{html.escape(sel_tx_rev)}</code> y consolidado aplicado."
                    st.rerun()
            if r_col2.button("Rechazar Reverso", use_container_width=True, key="btn_rechazar_rev"):
                ok_r, reg_mod = resolver_reverso_pendiente(sel_tx_rev, aprobar=False, admin_usr=usuario)
                if ok_r and reg_mod is not None:
                    enviar_gs(reg_mod, modo_dup="sobreescribir")
                    st.session_state["ultimo_tx_resumen"] = f"<b>Reverso rechazado:</b> <code>{html.escape(sel_tx_rev)}</code>; no altera el consolidado oficial."
                    st.rerun()
        st.divider()

    if df_rep.empty:
        st.info("No hay registros que coincidan con los filtros seleccionados en la carpeta 3 del panel izquierdo.")
        if not df_base_rep.empty:
            st.warning(f"Hay {len(df_base_rep)} fila(s) base disponibles pero 0 con el filtro actual (`Periodo: {f_per}`, `Bodega: {f_bod}`, `Ciudad: {f_ciu}`).")
            if st.button("Mostrar todos los registros (Ver todo)", key="btn_ver_todo_dash"):
                resetear_filtros_estadistica()
                st.rerun()
    else:
        sob_cnt = int((df_sku_rep["Estado_Neto"] == "SOBRANTE").sum()) if not df_sku_rep.empty else 0
        fal_cnt = int((df_sku_rep["Estado_Neto"] == "FALTANTE").sum()) if not df_sku_rep.empty else 0
        ok_cnt = int((df_sku_rep["Estado_Neto"] == "CUADRADO (OK)").sum()) if not df_sku_rep.empty else 0
        pend_r_cnt = int((df_rep["ESTADO"] == "PENDIENTE_REVERSO").sum())
        q_neta_sum = float(df_sku_rep["Q_Neta"].sum()) if not df_sku_rep.empty else 0.0
        val_neto_sum = float(df_sku_rep["Costo_Total_Neto"].sum()) if not df_sku_rep.empty else 0.0
        dif_neto_sum = float(df_sku_rep["Costo_Dif_Neto"].sum()) if not df_sku_rep.empty else 0.0

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Transacciones (Tx)", fmt_miles_comas(len(df_rep), 0))
        k2.metric("SKUs Únicos Contados", fmt_miles_comas(df_sku_rep["COD_SKU"].nunique() if not df_sku_rep.empty else 0, 0))
        k3.metric("Cantidad Física Neta (Q)", fmt_miles_comas(q_neta_sum, 2))
        k4.metric("Valor Contado Neto", fmt_moneda_ec(val_neto_sum, 2))
        k5.metric("Impacto Dif. Neto", fmt_moneda_ec(dif_neto_sum, 2))

        st.markdown(
            f"**Estados Netos por SKU:** &nbsp; **SOBRANTE:** `{sob_cnt}` &nbsp;|&nbsp; **FALTANTE:** `{fal_cnt}` "
            f"&nbsp;|&nbsp; **CUADRADO (OK):** `{ok_cnt}` &nbsp;|&nbsp; **Tx PENDIENTE_REVERSO:** `{pend_r_cnt}`"
        )
        df_bod_av = construir_resumen_bodega(df_rep, nac_rows, muestreo=muestreo_rep, bodega_filtro=f_bod)
        st.divider()
        render_graficos_avance_operativo(df_bod_av)
        st.divider()

        df_cus_g = (
            df_rep[~df_rep["ESTADO"].isin(["PENDIENTE_REVERSO", "REVERSO_RECHAZADO"])]
            .groupby("Correo_Usuario", as_index=False)
            .agg(Transacciones=("ID_Tx", "count"), Unidades_Q=("Q", "sum"))
            .sort_values("Transacciones", ascending=False)
        )
        render_svg_barras_simples(df_cus_g, "Correo_Usuario", "Unidades_Q", "Totales de Unidades Contadas (Q) por Custodio", color_pos="#2563EB", es_moneda=False)

        df_est_g = (
            df_sku_rep.groupby("Estado_Neto", as_index=False)
            .agg(SKUs=("COD_SKU", "count"), Costo_Dif_Neto=("Costo_Dif_Neto", "sum"))
            if not df_sku_rep.empty else pd.DataFrame()
        )
        render_svg_barras_simples(df_est_g, "Estado_Neto", "Costo_Dif_Neto", "Impacto Económico Total por Estado de SKU", color_pos="#16A34A", color_neg="#DC2626", es_moneda=True)

        st.divider()
        st.markdown("#### Matriz de Avance y Valoración Total por Bodega")
        if not df_bod_av.empty:
            render_tabla_segura(
                df_bod_av,
                ["Bodega", "SKU_Teoricos", "SKU_Contados", "SKU_Pendientes", "Avance_Pct", "Tx", "Cantidad_Q", "QSA_Teorico_Tot", "Diferencia_Q_Tot", "Sobrantes", "Faltantes", "Cuadrados", "Costo_Total_USD", "Costo_Dif_USD"],
                max_filas=20,
            )
        st.markdown("#### Consolidado Neto por SKU (Una fila por SKU sin duplicar diferencias)")
        if not df_sku_rep.empty:
            render_tabla_segura(
                df_sku_rep,
                ["Periodo", "Ciudad", "Bodega", "COD_SKU", "DESCRIPCIÓN", "PROVEEDOR", "Tx_Count", "Q_Neta", "QSA_Teorico", "Diferencia_Neta", "Estado_Neto", "Costo_Unitario", "Costo_Total_Neto", "Costo_Dif_Neto"],
                max_filas=40,
            )
        st.markdown("#### Bitácora de Transacciones Filtradas")
        render_tabla_segura(
            df_rep.iloc[::-1],
            ["ID_Tx", "Timestamp", "Correo_Usuario", "Ciudad", "Bodega", "PROVEEDOR", "COD_SKU", "Q", "QSA_Teorico", "DIFERENCIA", "ESTADO", "Estado_Reverso", "COSTO UNITARIO", "COSTO TOTAL", "COSTO DIF"],
            max_filas=30,
        )

with tab_etq:
    st.markdown("### Cola de Etiquetas (~5,8 × 2,5 cm) & Armado de Hoja A4")
    st.caption("Diseño centrado: `DESCRIPCIÓN` arriba + código de barras `CODE 128` al centro + `COD_SKU` abajo + bodega en negrita inferior izquierda + borde punteado para recorte.")
    ec1, ec2, ec3, ec4 = st.columns([2.2, 1.0, 1.0, 1.4])
    sku_etq_man = norm_mayus_limpio(ec1.text_input("Agregar COD_SKU a cola:", key="sku_etq_man_in"))
    cop_etq_man = ec2.number_input("Copias:", min_value=1, max_value=33, value=1, step=1, key="cop_etq_man_in")
    if ec3.button("Agregar", use_container_width=True, disabled=not bool(sku_etq_man), key="btn_add_etq_man"):
        desc_m = dict_maestro.get(sku_etq_man, {}).get("DESCRIPCIÓN") or dict_maestro.get(sku_etq_man, {}).get("DESCRIPCION") or ""
        agregar_etiqueta_cola(sku_etq_man, desc_m, bod_norm, "", usuario_lbl=usuario, copias=cop_etq_man)
        st.rerun()
    if ec4.button("Vaciar cola", use_container_width=True, disabled=len(st.session_state["cola_etiquetas"]) == 0, key="btn_vaciar_etq"):
        st.session_state["cola_etiquetas"] = []
        st.rerun()

    if not df_rep.empty:
        skus_filt = sorted([s for s in df_rep["COD_SKU"].unique() if s])
        if skus_filt:
            f_et1, f_et2 = st.columns([3, 1.5])
            sel_skus_etq = f_et1.multiselect("Marcar SKUs del filtro actual para etiqueta:", skus_filt, key="ms_skus_etq")
            if f_et2.button("Sumar seleccionados a cola", use_container_width=True, disabled=not bool(sel_skus_etq), key="btn_sum_ms_etq"):
                for s_item in sel_skus_etq:
                    d_s = dict_maestro.get(s_item, {}).get("DESCRIPCIÓN") or dict_maestro.get(s_item, {}).get("DESCRIPCION") or ""
                    agregar_etiqueta_cola(s_item, d_s, bod_norm, "", usuario_lbl=usuario, copias=1)
                st.rerun()

    cola_act = st.session_state.get("cola_etiquetas", [])
    st.caption(f"Etiquetas acumuladas en cola: **{len(cola_act)}** (hasta 33 etiquetas por hoja A4 en grilla 3×11 de ~5,8×2,5 cm).")
    if cola_act:
        render_tabla_segura(pd.DataFrame(cola_act), ["DESCRIPCION", "COD_SKU", "BODEGA"], max_filas=60)
    render_hoja_a4_etiquetas(cola_act, cod_label=st.session_state["cod_label_act"], ancho_mm=58, alto_mm=25)

with tab_exp:
    st.markdown("### Detalle Filtrado & Exportación a Excel")
    if df_rep.empty:
        st.info("No hay registros disponibles con los filtros actuales.")
    else:
        st.download_button(
            f"Descargar Excel Completo ({len(df_rep)} registros)",
            data=excel_bytes(df_rep, nac_rows=nac_rows, muestreo=muestreo_rep),
            file_name=f"TomaFisica_V7gPro_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="btn_dl_tab",
        )
        render_tabla_segura(df_rep.iloc[::-1], COLS_REP, max_filas=50)

    st.markdown("---")
    st.markdown(f"### Reporte de LOGS (`{PESTANA_LOGS_GS}` -> Columnas `A:C`)")
    logs_act = cargar_logs()
    pend_logs = [lg for lg in logs_act if str(lg.get("Enviado_GS", "NO")).upper() != "SI"]
    lcol1, lcol2, lcol3 = st.columns([1.6, 1.4, 2])
    if lcol1.button(f"Enviar a {PESTANA_LOGS_GS} ({len(pend_logs)})", use_container_width=True, key="btn_send_logs"):
        if pend_logs:
            ok_cnt_l, tot_cnt_l = enviar_logs_a_gs(pend_logs)
            guardar_logs_lista(logs_act)
            st.session_state["ultimo_tx_resumen"] = f"<b>Logs enviados a {html.escape(PESTANA_LOGS_GS)}:</b> {ok_cnt_l}/{tot_cnt_l} registros."
            st.rerun()
        else:
            st.info(f"No hay logs pendientes por enviar a {PESTANA_LOGS_GS}.")
    if lcol2.button("Limpiar LOGS locales", use_container_width=True, key="btn_clear_logs"):
        if logs_act:
            guardar_logs_lista([])
            st.session_state["ultimo_tx_resumen"] = "<b>Historial de LOGS locales limpiado.</b>"
            st.rerun()
        else:
            st.info("No hay logs locales para limpiar.")
    lcol3.caption(f"Pestaña destino: `{PESTANA_LOGS_GS}` (`A:C`) | Total: `{len(logs_act)}` | Pendientes: `{len(pend_logs)}`")
    if logs_act:
        df_logs = pd.DataFrame(logs_act).iloc[::-1]
        for c_l in COLS_LOG_UI:
            if c_l not in df_logs.columns:
                df_logs[c_l] = ""
        render_tabla_segura(df_logs, COLS_LOG_UI, max_filas=40)
    else:
        st.info("Sin errores registrados en este momento.")
