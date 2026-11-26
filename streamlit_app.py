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
import math
import html
import shutil
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
    html, body, [class*="css"], .stMarkdown, p, span, label, div, h1, h2, h3, h4, li, td, th { color: #0F172A !important; }
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"], .main { background-color: #F8FAFC !important; }
    [data-testid="stHeader"] { background-color: #F8FAFC !important; }
    [data-testid="stSidebar"] { background-color: #0F172A !important; }
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] div, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: #F8FAFC !important;
    }
    [data-testid="stSidebar"] input {
        background-color: #1E293B !important; color: #FEF3C7 !important; border: 1px solid #F59E0B !important; font-weight: 700 !important;
    }
    [data-testid="stSidebar"] [data-baseweb="select"] > div {
        background-color: #1E293B !important; color: #FEF3C7 !important; border: 1px solid #F59E0B !important; font-weight: 700 !important;
    }
    [data-baseweb="select"] * { color: #451A03 !important; }
    [data-testid="stSidebar"] [data-baseweb="select"] * { color: #FEF3C7 !important; }
    .main-title {
        background: #0F172A; color: #FFFFFF !important; padding: 14px 18px; border-radius: 10px;
        font-size: 20px !important; font-weight: 800; border-bottom: 4px solid #F59E0B; margin-bottom: 8px;
    }
    .card-box { background: #FFFFFF !important; border: 1px solid #CBD5E1; border-radius: 10px; padding: 14px; margin-bottom: 12px; }
    .stTextInput input, .stNumberInput input, [data-testid="stMain"] [data-baseweb="select"] > div {
        background-color: #FEF3C7 !important; color: #451A03 !important; font-weight: 700 !important; border: 2px solid #D97706 !important;
    }
    .stButton button, .stDownloadButton button, .stButton button *, .stDownloadButton button * {
        background-color: #0F172A !important; color: #FFFFFF !important; font-weight: 700 !important;
    }
    .stAlert, .stAlert * { color: #0F172A !important; }
    .saldo-box {
        background: #EFF6FF; border: 2px solid #2563EB; border-radius: 8px; padding: 10px 14px; text-align: center;
    }
    .saldo-lbl { font-size: 13px !important; font-weight: 700 !important; color: #1E3A8A !important; display: block; margin-bottom: 4px; }
    .saldo-val { font-size: 22px !important; font-weight: 900 !important; color: #0F172A !important; display: block; }
</style>
""",
    unsafe_allow_html=True,
)

RUTA_SCRIPT = os.path.abspath(__file__)
BASE_DIR = os.path.dirname(RUTA_SCRIPT)
st.session_state["errores_ui"] = []
st.session_state.setdefault("cap_v", 0)
st.session_state.setdefault("mem_regs", [])
st.session_state.setdefault("pin_reverso_act", "7031")
st.session_state.setdefault("cola_etiquetas", [])
st.session_state.setdefault("cod_label_act", "LBL-EIT-A4")


def norm_mayus_limpio(v):
    return " ".join(str(v or "").strip().upper().split())


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
            "DESCRIPCION": norm_mayus_limpio(desc)[:38],
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
            s_sku = html.escape(norm_mayus_limpio(it.get("COD_SKU", "")))
            s_desc = html.escape(norm_mayus_limpio(it.get("DESCRIPCION", ""))[:36])
            s_fl = html.escape(norm_mayus_limpio(it.get("FECHA_LOGIN", ""))[:26])
            celdas.append(
                f'<div style="width:{ancho_mm}mm;height:{alto_mm}mm;border:1px dashed #334155;box-sizing:border-box;'
                f'padding:1.5mm 2mm;display:flex;flex-direction:column;justify-content:space-between;background:#FFFFFF;overflow:hidden;">'
                f'<div style="font-size:9.5pt;font-weight:900;letter-spacing:0.4px;color:#0F172A;line-height:1.1;">{s_sku}</div>'
                f'<div style="font-size:6.5pt;font-weight:700;color:#1E293B;line-height:1.15;max-height:7mm;overflow:hidden;">{s_desc}</div>'
                f'<div style="font-size:6pt;font-weight:700;color:#475569;line-height:1.05;">{s_fl}</div>'
                f'</div>'
            )
        bloques_html.append(
            f'<div class="hoja-a4" style="width:210mm;min-height:297mm;padding:10mm 11mm;box-sizing:border-box;'
            f'background:#FFFFFF;border:1px solid #94A3B8;margin:0 auto 14px auto;display:grid;'
            f'grid-template-columns:repeat(3, {ancho_mm}mm);grid-auto-rows:{alto_mm}mm;gap:2mm 3mm;align-content:start;">'
            + "".join(celdas)
            + '</div>'
        )

    html_doc = f"""
    <div>
      <div style="margin-bottom:10px;display:flex;gap:10px;align-items:center;">
        <button onclick="window.print()" type="button" style="padding:8px 14px;background:#0F172A;color:#FFFFFF;border:none;border-radius:6px;font-weight:700;cursor:pointer;">
          Imprimir Hoja A4 ({len(cola)} etiquetas / {total_hojas} hoja(s))
        </button>
        <span style="font-size:12px;color:#334155;font-weight:700;">COD_LABEL: {html.escape(cod_label)} • A4 • ~{ancho_mm}×{alto_mm} mm • COD_SKU / DESCRIPCIÓN / FECHA-LOGIN</span>
      </div>
      <style>
        @media print {{
          body * {{ visibility: hidden; }}
          .hoja-a4, .hoja-a4 * {{ visibility: visible; }}
          .hoja-a4 {{ position: relative; border: none !important; margin: 0 !important; page-break-after: always; }}
          button {{ display: none !important; }}
        }}
      </style>
      {''.join(bloques_html)}
    </div>
    """
    components.html(html_doc, height=540, scrolling=True)


def generar_pin_diario(pin_base="7031", fecha_ref=None):
    f_str = (fecha_ref or date.today()).strftime("%Y-%m-%d")
    raw = f"{str(pin_base).strip()}|DIA|{f_str}".encode("utf-8")
    dig = int(hashlib.sha256(raw).hexdigest()[:6], 16) % 10000
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


def respaldar_y_limpiar_pruebas():
    carpeta_respaldo = os.path.join(os.path.dirname(json_path) or BASE_DIR, "Respaldos_Pruebas")
    try:
        os.makedirs(carpeta_respaldo, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        for arch in [json_path, enviados_path]:
            if os.path.exists(arch):
                destino = os.path.join(carpeta_respaldo, f"{os.path.splitext(os.path.basename(arch))[0]}_{stamp}.json")
                shutil.copy2(arch, destino)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump([], f, indent=2, ensure_ascii=False)
        with open(enviados_path, "w", encoding="utf-8") as f:
            json.dump([], f, indent=2, ensure_ascii=False)
    except Exception:
        pass
    st.session_state["mem_regs"] = []
    leer_gs.clear()
    return carpeta_respaldo


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
        reg.get("LINEA DE INVENTARIO", ""), reg.get("PROVEEDOR", ""), fmt_txt(reg.get("ALTERNO", "")),
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


def datos_combinados(incluir_gs=False):
    loc = cargar_json()
    if not incluir_gs:
        return loc
    rem = leer_gs()
    res, vistos = [], set()
    for r in loc + rem:
        if not isinstance(r, dict):
            continue
        tx = norm_mayus_limpio(r.get("ID_Tx", ""))
        k = tx if tx else f"{r.get('Timestamp','')}|{norm_bod(r.get('Bodega',''))}|{norm_mayus_limpio(r.get('COD_SKU',''))}|{num_es(r.get('Q',0))}"
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


def tabla_md(df, cols, n=8):
    if df is None or df.empty:
        return "_Sin registros._"
    sub = df[cols].head(n)
    cab = "| " + " | ".join(cols) + " |"
    sep = "| " + " | ".join(["---"] * len(cols)) + " |"
    lines = [cab, sep]
    for row in sub.to_dict(orient="records"):
        vals = []
        for c in cols:
            v = row.get(c, "")
            if isinstance(v, float) and "UNITARIO" in c.upper():
                s = f"{v:,.4f}"
            elif isinstance(v, float) and "COSTO" in c.upper():
                s = f"{v:,.2f}"
            elif isinstance(v, float):
                s = f"{v:g}"
            else:
                s = str(v or "")
            vals.append(s.replace("|", "/").replace("\n", " ").strip())
        lines.append("| " + " | ".join(vals) + " |")
    return "\n".join(lines)


def resolver_col_bodega(nac_rows, bod):
    bod_n = norm_bod(bod)
    if nac_rows and len(nac_rows) > 0:
        for i, h in enumerate(nac_rows[0]):
            if norm_bod(h) == bod_n:
                return i
    return MAPEO_BOD_COL.get(bod_n, 5)


def sku_habilitado_col_v(row, muestreo=True):
    if not muestreo:
        return True
    val_v = norm_mayus_limpio(row[IDX_COL_V]) if IDX_COL_V < len(row) else ""
    if not val_v or val_v in ["SIN MOV", "SIN_MOV", "NONE", "NAN", "0"]:
        return False
    return True


def obtener_costo_nac_col_w(row):
    if IDX_COL_W < len(row):
        return limpiar_costo_unit(row[IDX_COL_W])
    return 0.0


def obtener_skus_teoricos_bodega(nac_rows, dict_maestro, bod, muestreo=True):
    skus_set = set()
    desc_map = {}
    if not nac_rows or len(nac_rows) <= 1:
        return [], desc_map
    idx_b = resolver_col_bodega(nac_rows, bod)
    for row in nac_rows[1:]:
        if not row or len(row) == 0 or row[0] is None:
            continue
        sv = norm_mayus_limpio(row[0])
        if not sv or sv in ["NONE", "NAN"]:
            continue
        if not sku_habilitado_col_v(row, muestreo=muestreo):
            continue
        saldo_bod = num_es(row[idx_b], max_abs=1e12) if idx_b < len(row) else 0.0
        if saldo_bod > 0:
            skus_set.add(sv)
            if sv not in desc_map:
                desc_map[sv] = next((v for k, v in dict_maestro.get(sv, {}).items() if "DESC" in k), "")
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
        teo_skus = contar_teoricos_bodega(nac_rows, bod_k, muestreo=muestreo)
        cont_skus = int(grp_ok["COD_SKU"].nunique())
        pend_skus = max(teo_skus - cont_skus, 0) if teo_skus > 0 else 0
        pct_avance = round((cont_skus / teo_skus) * 100.0, 2) if teo_skus > 0 else 0.0
        filas.append([
            str(per_k), str(bod_k), str(cus_k), int(len(grp_ok)), int(teo_skus), int(cont_skus),
            int(pend_skus), float(pct_avance), round(float(grp_ok["Q"].sum()), 2),
            int((grp_ok["ESTADO"] == "SOBRANTE").sum()), int((grp_ok["ESTADO"] == "FALTANTE").sum()),
            int((grp_ok["ESTADO"] == "CUADRADO (OK)").sum()), round(float(grp_ok["COSTO TOTAL"].sum()), 2),
            round(float(grp_ok["COSTO DIF"].sum()), 2),
        ])
    return filas


def construir_resumen_bodega(df, nac_rows, muestreo=True, bodega_filtro="TODAS"):
    bodegas_iter = [norm_bod(bodega_filtro)] if bodega_filtro != "TODAS" else BODEGAS_OFICIALES
    filas = []
    for bod_k in bodegas_iter:
        grp_all = df[df["Bodega"] == bod_k] if not df.empty else pd.DataFrame(columns=COLS_REP)
        grp = grp_all[~grp_all["ESTADO"].isin(["PENDIENTE_REVERSO", "REVERSO_RECHAZADO"])] if not grp_all.empty else grp_all
        teo_skus = contar_teoricos_bodega(nac_rows, bod_k, muestreo=muestreo)
        cont_skus = int(grp["COD_SKU"].nunique()) if not grp.empty else 0
        pend_skus = max(teo_skus - cont_skus, 0) if teo_skus > 0 else 0
        pct_avance = round((cont_skus / teo_skus) * 100.0, 1) if teo_skus > 0 else 0.0
        if not grp_all.empty or bodega_filtro != "TODAS":
            filas.append({
                "Bodega": bod_k, "SKU_Teoricos": teo_skus, "SKU_Contados": cont_skus,
                "SKU_Pendientes": pend_skus, "Avance_Pct": pct_avance, "Tx": int(len(grp)),
                "Cantidad_Q": round(float(grp["Q"].sum()), 2) if not grp.empty else 0.0,
                "Sobrantes": int((grp["ESTADO"] == "SOBRANTE").sum()) if not grp.empty else 0,
                "Faltantes": int((grp["ESTADO"] == "FALTANTE").sum()) if not grp.empty else 0,
                "Cuadrados": int((grp["ESTADO"] == "CUADRADO (OK)").sum()) if not grp.empty else 0,
                "Costo_Total_USD": round(float(grp["COSTO TOTAL"].sum()), 2) if not grp.empty else 0.0,
                "Costo_Dif_USD": round(float(grp["COSTO DIF"].sum()), 2) if not grp.empty else 0.0,
            })
    return pd.DataFrame(filas)


def render_grafica_control_y_matriz(df_src, titulo="Control de Tomas & Proyección (Q(PRY) / LCS / x / cT / LCI)"):
    st.markdown(f"#### {titulo}")
    if df_src is None or df_src.empty:
        st.info("Sin datos para graficar control y matriz.")
        return

    df_ord = df_src.sort_values("Timestamp").reset_index(drop=True).tail(12).copy()
    etiquetas = [f"{i + 1}" for i in range(len(df_ord))]
    ct_vals = [round(float(x), 2) for x in df_ord["Q"].tolist()]
    media_x = round(sum(ct_vals) / len(ct_vals), 2) if ct_vals else 0.0
    max_q = max(ct_vals) if ct_vals else 0.0
    lcs_val = round(max(media_x * 2.0, max_q * 1.1, 30.0), 2)
    lci_val = 0.0

    q_pry_vals = [None] * len(ct_vals)
    for idx_p in range(max(0, len(ct_vals) - 3), len(ct_vals)):
        q_pry_vals[idx_p] = round(max_q + (idx_p + 1) * 5.0, 2)

    fig_html = f"""
    <div style="background:#FFFFFF;border:1px solid #CBD5E1;border-radius:8px;padding:10px;overflow-x:auto;">
      <canvas id="canvasControlV7" width="860" height="250"></canvas>
      <table style="width:100%;border-collapse:collapse;margin-top:6px;font-size:12px;">
        <thead>
          <tr style="background:#F1F5F9;">
            <th style="border:1px solid #CBD5E1;padding:4px 6px;text-align:left;">Serie</th>
            {''.join(f'<th style="border:1px solid #CBD5E1;padding:4px 6px;text-align:center;">{e}</th>' for e in etiquetas)}
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="border:1px solid #CBD5E1;padding:4px 6px;font-weight:700;"><span style="display:inline-block;width:12px;height:10px;background:#F59E0B;margin-right:4px;"></span>Q(PRY)</td>
            {''.join(f'<td style="border:1px solid #CBD5E1;padding:4px 6px;text-align:center;font-weight:700;">{"" if v is None else f"{v:g}"}</td>' for v in q_pry_vals)}
          </tr>
          <tr>
            <td style="border:1px solid #CBD5E1;padding:4px 6px;font-weight:700;color:#2563EB;">● LCS</td>
            {''.join(f'<td style="border:1px solid #CBD5E1;padding:4px 6px;text-align:center;">{lcs_val:g}</td>' for _ in etiquetas)}
          </tr>
          <tr>
            <td style="border:1px solid #CBD5E1;padding:4px 6px;font-weight:700;color:#EA580C;">● x</td>
            {''.join(f'<td style="border:1px solid #CBD5E1;padding:4px 6px;text-align:center;">{media_x:g}</td>' for _ in etiquetas)}
          </tr>
          <tr>
            <td style="border:1px solid #CBD5E1;padding:4px 6px;font-weight:700;color:#64748B;">● cT</td>
            {''.join(f'<td style="border:1px solid #CBD5E1;padding:4px 6px;text-align:center;">{v:g}</td>' for v in ct_vals)}
          </tr>
          <tr>
            <td style="border:1px solid #CBD5E1;padding:4px 6px;font-weight:700;color:#1D4ED8;">● LCI</td>
            {''.join(f'<td style="border:1px solid #CBD5E1;padding:4px 6px;text-align:center;">{lci_val:g}</td>' for _ in etiquetas)}
          </tr>
        </tbody>
      </table>
    </div>
    <script>
    (function() {{
      const cv = document.getElementById("canvasControlV7");
      if (!cv) return;
      const ctx = cv.getContext("2d");
      const labels = {json.dumps(etiquetas)};
      const ct = {json.dumps(ct_vals)};
      const qpry = {json.dumps(q_pry_vals)};
      const lcs = labels.map(() => {lcs_val});
      const xmed = labels.map(() => {media_x});
      const lci = labels.map(() => {lci_val});

      const padL = 45, padR = 45, padT = 20, padB = 28;
      const w = cv.width - padL - padR;
      const h = cv.height - padT - padB;
      const allNums = ct.concat(lcs, xmed, lci).concat(qpry.filter(v => v !== null));
      const yMax = Math.max(10, ...allNums) * 1.12;

      ctx.clearRect(0, 0, cv.width, cv.height);
      ctx.strokeStyle = "#E2E8F0";
      ctx.lineWidth = 1;
      for (let g = 0; g <= 5; g++) {{
        const y = padT + h - (g / 5) * h;
        ctx.beginPath(); ctx.moveTo(padL, y); ctx.lineTo(padL + w, y); ctx.stroke();
        ctx.fillStyle = "#475569"; ctx.font = "11px sans-serif";
        ctx.fillText(Math.round((g / 5) * yMax), 6, y + 4);
      }}

      const stepX = w / Math.max(1, labels.length);
      labels.forEach((lbl, i) => {{
        const cx = padL + i * stepX + stepX / 2;
        ctx.fillStyle = "#334155"; ctx.font = "11px sans-serif"; ctx.textAlign = "center";
        ctx.fillText(lbl, cx, cv.height - 8);
        if (qpry[i] !== null && qpry[i] !== undefined) {{
          const bh = (qpry[i] / yMax) * h;
          const bw = Math.min(28, stepX * 0.45);
          ctx.fillStyle = "#F59E0B";
          ctx.fillRect(cx - bw / 2, padT + h - bh, bw, bh);
          ctx.fillStyle = "#991B1B"; ctx.font = "bold 11px sans-serif";
          ctx.fillText(String(qpry[i]), cx, padT + h - bh - 4);
        }}
      }});

      function drawLine(arr, color) {{
        ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.fillStyle = color;
        ctx.beginPath();
        arr.forEach((v, i) => {{
          const x = padL + i * stepX + stepX / 2;
          const y = padT + h - (Number(v || 0) / yMax) * h;
          if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
        }});
        ctx.stroke();
        arr.forEach((v, i) => {{
          const x = padL + i * stepX + stepX / 2;
          const y = padT + h - (Number(v || 0) / yMax) * h;
          ctx.beginPath(); ctx.arc(x, y, 3.5, 0, Math.PI * 2); ctx.fill();
        }});
      }}

      drawLine(lcs, "#2563EB");
      drawLine(xmed, "#EA580C");
      drawLine(ct, "#64748B");
      drawLine(lci, "#1D4ED8");
    }})();
    </script>
    """
    components.html(fig_html, height=420, scrolling=True)


def render_barras_peso_inventario(df_src, titulo="Peso Inventario vs. Peso Costo vs. Neumática"):
    st.markdown(f"#### {titulo}")
    if df_src is None or df_src.empty:
        st.info("Sin datos para calcular pesos porcentuales.")
        return

    grp = df_src.groupby("COD_SKU", as_index=False).agg(
        Q_TOT=("Q", "sum"),
        COSTO_TOT=("COSTO TOTAL", "sum"),
        LINEA=("LINEA DE INVENTARIO", "first"),
    ).head(8)

    sum_q = float(grp["Q_TOT"].sum()) or 1.0
    sum_c = float(grp["COSTO_TOT"].sum()) or 1.0
    neum_df = grp[grp["LINEA"].astype(str).str.upper().str.contains("NEUM")]
    sum_n = float(neum_df["Q_TOT"].sum()) or 1.0

    rows_html = []
    for r in grp.to_dict(orient="records"):
        p_inv = min(100.0, max(0.0, (float(r["Q_TOT"]) / sum_q) * 100.0))
        p_cos = min(100.0, max(0.0, (float(r["COSTO_TOT"]) / sum_c) * 100.0))
        es_neum = "NEUM" in str(r.get("LINEA", "")).upper()
        p_neu = min(100.0, max(0.0, (float(r["Q_TOT"]) / sum_n) * 100.0)) if es_neum else 0.0

        def barra(pct, grad, borde, texto):
            return (
                f'<div style="position:relative;width:100%;height:20px;background:#F8FAFC;border:1px solid #E2E8F0;overflow:hidden;">'
                f'<div style="width:{pct:.1f}%;height:100%;background:{grad};border-right:1px solid {borde};"></div>'
                f'<span style="position:absolute;right:6px;top:1px;font-size:11px;font-weight:800;color:#0F172A;">{texto}</span>'
                f'</div>'
            )

        rows_html.append(
            f'<tr style="border-bottom:1px solid #EA580C;">'
            f'<td style="padding:4px 6px;font-weight:700;font-size:12px;">{html.escape(str(r["COD_SKU"]))}</td>'
            f'<td style="padding:4px;">{barra(p_inv, "linear-gradient(90deg,#86EFAC,#16A34A)", "#15803D", f"{p_inv:.0f}%")}</td>'
            f'<td style="padding:4px;">{barra(p_cos, "linear-gradient(90deg,#DBEAFE,#2563EB)", "#1D4ED8", f"{p_cos:.0f}%")}</td>'
            f'<td style="padding:4px;">{barra(p_neu, "linear-gradient(90deg,#FFEDD5,#EA580C)", "#C2410C", f"{p_neu:.0f}%" if es_neum else "")}</td>'
            f'</tr>'
        )

    tabla_peso = f"""
    <div style="background:#FFFFFF;border:1px solid #CBD5E1;border-radius:8px;padding:8px;overflow-x:auto;">
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:#EA580C;color:#FFFFFF;font-size:12px;">
            <th style="padding:6px;text-align:left;">COD_SKU</th>
            <th style="padding:6px;text-align:center;">PESO INV</th>
            <th style="padding:6px;text-align:center;">PESO COSTO INV</th>
            <th style="padding:6px;text-align:center;">NEUMÁTICA</th>
          </tr>
        </thead>
        <tbody>
          {''.join(rows_html)}
          <tr style="background:#E2E8F0;font-weight:800;font-size:12px;">
            <td style="padding:6px;">TOTAL</td>
            <td style="padding:6px;text-align:center;">100%</td>
            <td style="padding:6px;text-align:center;">100%</td>
            <td style="padding:6px;text-align:center;">100%</td>
          </tr>
        </tbody>
      </table>
    </div>
    """
    components.html(tabla_peso, height=320, scrolling=True)


def render_graficos_tomas_y_auditoria(df_of, df_bod_av):
    st.markdown("#### Gráficos Visuales: Auditado vs. Por Auditar y Evolución de Tomas")
    if df_bod_av is not None and not df_bod_av.empty:
        render_lista_visual(df_bod_av, "Bodega", ["SKU_Contados", "SKU_Pendientes", "Avance_Pct"], "Auditado vs. Por Auditar por Bodega")
    if df_of is not None and not df_of.empty:
        df_t = df_of.sort_values("Timestamp").reset_index(drop=True).tail(15).copy()
        df_t["Toma_Etq"] = [f"Toma {i + 1} ({s})" for i, s in enumerate(df_t["COD_SKU"].tolist())]
        tot_q = float(df_t["Q"].sum())
        df_t["Pct_Acum"] = ((df_t["Q"].cumsum() / tot_q) * 100.0).round(1) if tot_q != 0 else 0.0
        render_lista_visual(df_t, "Toma_Etq", ["Q", "Pct_Acum"], "Evolución de Tomas y % Acumulado")

    st.divider()
    render_grafica_control_y_matriz(df_of)
    st.divider()
    render_barras_peso_inventario(df_of)


def render_lista_visual(df_plot, label_col, val_cols, titulo=""):
    if titulo:
        st.markdown(f"#### {titulo}")
    if df_plot is None or df_plot.empty:
        st.info("Sin datos para mostrar con los filtros actuales.")
        return
    col_main = val_cols[0]
    max_val = max((abs(float(v)) for v in df_plot[col_main]), default=1.0)
    if max_val <= 0:
        max_val = 1.0
    for r in df_plot.to_dict(orient="records"):
        etq = str(r.get(label_col, "")).strip() or "SIN DATO"
        v_main = float(r.get(col_main, 0.0) or 0.0)
        ratio = min(max(abs(v_main) / max_val, 0.0), 1.0)
        if "USD" in col_main.upper() or "COSTO" in col_main.upper():
            main_fmt = f"${v_main:,.2f}"
        elif "PCT" in col_main.upper():
            main_fmt = f"{v_main:,.1f}%"
        else:
            main_fmt = f"{v_main:,.0f}" if float(v_main).is_integer() else f"{v_main:,.2f}"
        extras = []
        for c_sec in val_cols[1:]:
            v_sec = float(r.get(c_sec, 0.0) or 0.0)
            if "USD" in c_sec.upper() or "COSTO" in c_sec.upper():
                val_sec_fmt = f"${v_sec:,.2f}"
            elif "PCT" in c_sec.upper():
                val_sec_fmt = f"{v_sec:,.1f}%"
            else:
                val_sec_fmt = f"{v_sec:,.0f}" if float(v_sec).is_integer() else f"{v_sec:,.2f}"
            extras.append(f"**{c_sec}:** {val_sec_fmt}")
        st.markdown(f"**{etq}** — **{col_main}:** `{main_fmt}`")
        if extras:
            st.caption(" | ".join(extras))
        st.progress(ratio)


def render_tabla_segura(df_in, cols, max_filas=50, key_suf="tbl"):
    if df_in is None or df_in.empty:
        st.info("No hay registros disponibles con los filtros actuales.")
        return
    sub = df_in[cols].head(max_filas)
    ths = "".join(
        f'<th style="background:#0F172A;color:#FFFFFF!important;padding:8px 10px;text-align:left;font-weight:700;border:1px solid #334155;white-space:nowrap;">{html.escape(str(c))}</th>'
        for c in cols
    )
    trs = []
    for row in sub.to_dict(orient="records"):
        tds = []
        for c in cols:
            val = row.get(c, "")
            if isinstance(val, float):
                if "UNITARIO" in c.upper():
                    txt = f"${val:,.4f}"
                elif "COSTO" in c.upper() or "USD" in c.upper():
                    txt = f"${val:,.2f}"
                elif "PCT" in c.upper():
                    txt = f"{val:,.1f}%"
                else:
                    txt = f"{val:,.0f}" if float(val).is_integer() else f"{val:,.2f}"
            else:
                txt = str(val or "")
            tds.append(f'<td style="padding:7px 10px;border:1px solid #CBD5E1;color:#0F172A!important;font-size:13px;white-space:nowrap;">{html.escape(txt)}</td>')
        trs.append(f'<tr style="background:#FFFFFF;">{"".join(tds)}</tr>')
    st.markdown(
        f'<div style="overflow-x:auto;border-radius:8px;border:1px solid #CBD5E1;margin-top:6px;margin-bottom:10px;"><table style="width:100%;border-collapse:collapse;"><thead><tr>{ths}</tr></thead><tbody>{"".join(trs)}</tbody></table></div>',
        unsafe_allow_html=True,
    )
    if len(df_in) > max_filas:
        st.caption(f"Mostrando las {max_filas} filas más recientes de {len(df_in)} registros filtrados.")


def render_lector_codigos_barras(comp_key):
    html_code = f"""
    <div style="font-family:sans-serif;padding:8px;background:#F8FAFC;border:1px solid #CBD5E1;border-radius:8px;">
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px;">
        <button id="btnStartScan_{comp_key}" type="button" style="padding:8px 12px;background:#0F172A;color:#FFF;border:none;border-radius:6px;font-weight:700;cursor:pointer;">Activar lector de barras</button>
        <button id="btnStopScan_{comp_key}" type="button" style="padding:8px 12px;background:#475569;color:#FFF;border:none;border-radius:6px;font-weight:700;cursor:pointer;">Detener cámara</button>
      </div>
      <video id="vidScan_{comp_key}" playsinline muted style="width:100%;max-height:220px;background:#0F172A;border-radius:6px;display:none;"></video>
      <div id="resScan_{comp_key}" style="margin-top:6px;font-size:13px;font-weight:700;color:#1E3A8A;">Apunta la cámara trasera al código de barras.</div>
    </div>
    <script>
    (function() {{
      const btnStart = document.getElementById("btnStartScan_{comp_key}");
      const btnStop = document.getElementById("btnStopScan_{comp_key}");
      const video = document.getElementById("vidScan_{comp_key}");
      const statusEl = document.getElementById("resScan_{comp_key}");
      let stream = null;
      let timerId = null;
      function aplicarCodigo(codigo) {{
        const limpio = String(codigo || "").trim().toUpperCase();
        if (!limpio) return;
        statusEl.textContent = "Código detectado: " + limpio;
        try {{
          if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(limpio).catch(() => {{}});
          const inputs = window.parent.document.querySelectorAll('input[type="text"]');
          for (const inp of inputs) {{
            if ((inp.getAttribute("aria-label") || "").toUpperCase().includes("COD_SKU")) {{
              const nativeSetter = Object.getOwnPropertyDescriptor(window.parent.HTMLInputElement.prototype, "value").set;
              nativeSetter.call(inp, limpio);
              inp.dispatchEvent(new Event("input", {{ bubbles: true }}));
              inp.dispatchEvent(new Event("change", {{ bubbles: true }}));
              break;
            }}
          }}
        }} catch(e) {{}}
        detener();
      }}
      function detener() {{
        if (timerId) {{ clearInterval(timerId); timerId = null; }}
        if (stream) {{ stream.getTracks().forEach(t => t.stop()); stream = null; }}
        video.style.display = "none";
      }}
      async function iniciar() {{
        detener();
        if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {{
          statusEl.textContent = "Sin acceso directo a cámara. Digita o pega COD_SKU.";
          return;
        }}
        try {{
          stream = await navigator.mediaDevices.getUserMedia({{ video: {{ facingMode: {{ ideal: "environment" }} }}, audio: false }});
          video.srcObject = stream;
          video.style.display = "block";
          await video.play();
          statusEl.textContent = "Cámara activa: enfocando código de barras...";
          if ("BarcodeDetector" in window) {{
            const detector = new window.BarcodeDetector({{ formats: ["code_128", "code_39", "ean_13", "ean_8", "upc_a", "upc_e", "qr_code", "itf"] }});
            timerId = setInterval(async () => {{
              if (video.readyState >= 2) {{
                try {{
                  const codes = await detector.detect(video);
                  if (codes && codes.length > 0 && codes[0].rawValue) aplicarCodigo(codes[0].rawValue);
                }} catch(e) {{}}
              }}
            }}, 350);
          }} else {{
            statusEl.textContent = "Cámara abierta. Usa Foto Evidencia o digita COD_SKU.";
          }}
        }} catch(err) {{
          statusEl.textContent = "No se pudo abrir cámara. Digita o pega COD_SKU.";
        }}
      }}
      btnStart.addEventListener("click", iniciar);
      btnStop.addEventListener("click", detener);
    }})();
    </script>
    """
    components.html(html_code, height=290)


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
            if cname in {"Bodega", "COD_SKU", "ALTERNO", "ID_Tx", "Ubicacion"}:
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


@st.cache_data(show_spinner=False)
def cargar_pivot(p):
    dm = {}
    if not os.path.exists(p):
        return dm
    try:
        wb = openpyxl.load_workbook(p, data_only=True)
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
def cargar_nac(p):
    if not os.path.exists(p):
        return []
    try:
        wb = openpyxl.load_workbook(p, data_only=True)
        return list(wb.active.iter_rows(values_only=True))
    except Exception as e:
        log_err(f"NAC: {e}", origen="NAC")
        return []


dict_maestro = cargar_pivot(pivot_path)
nac_rows = cargar_nac(nac_path)
regs_json = cargar_json()
logs_json = cargar_logs()
f_script = fecha_arch(RUTA_SCRIPT)
cap_v = st.session_state["cap_v"]

# ==============================================================================
# PANEL IZQUIERDO
# ==============================================================================
usuario = st.sidebar.text_input("Correo Usuario / Custodio:", value="alexpajaritoruiz@icloud.com", key="usr_in").strip().lower()
st.session_state["usr_act"] = usuario
es_admin = usuario in {"alexpajaritoruiz@icloud.com", "jlogistica.eit@gmail.com"} or "ALEX" in usuario.upper() or "PAJARI" in usuario.upper()
bods_perm = BODEGAS_OFICIALES if es_admin else CUSTODIO_BOD.get(usuario, ["01"])

tipo_proc = st.sidebar.radio("Modo:", ["Control Mensual de Inventario (Muestreo)", "Toma Física Completa"], key="modo_proc_in")
if "bod_act" not in st.session_state or st.session_state["bod_act"] not in bods_perm:
    st.session_state["bod_act"] = bods_perm[0]


def limpiar_controles_bodega():
    st.session_state["cap_v"] = st.session_state.get("cap_v", 0) + 1
    st.session_state.pop("flt_bod", None)
    st.session_state.pop("flt_ciu", None)


bod_norm = norm_bod(st.sidebar.radio("Bodega:", bods_perm, key="bod_act", on_change=limpiar_controles_bodega))
ciudades = ["Quito"] if bod_norm in ["01", "05", "5V"] else (["Guayaquil"] if bod_norm in ["03", "3V"] else (["Cuenca"] if bod_norm == "04" else ["Quito", "Guayaquil", "Cuenca", "Otras"]))
ciudad_def = ciudades[0].upper()

# Periodo de auditoría INHABILITADO por defecto
if "per_abierto" not in st.session_state:
    st.session_state["per_abierto"] = False
if es_admin:
    st.session_state["per_abierto"] = st.sidebar.toggle("Habilitar periodo de auditoría (Admin)", value=st.session_state["per_abierto"], key="per_abierto_tgl")
    st.session_state["pin_reverso_act"] = st.sidebar.text_input("PIN Base Reversos (Admin):", value=st.session_state["pin_reverso_act"], key="pin_rev_admin_in").strip()
    st.session_state["cod_label_act"] = norm_mayus_limpio(st.sidebar.text_input("COD_LABEL Etiqueta:", value=st.session_state["cod_label_act"], key="cod_lbl_admin_in")) or "LBL-EIT-A4"
else:
    st.sidebar.caption("Estado periodo: " + ("HABILITADO" if st.session_state["per_abierto"] else "INHABILITADO"))

f_ini = st.sidebar.date_input("Desde:", value=date.today(), key="f_ini_in")
f_fin = st.sidebar.date_input("Hasta:", value=date.today(), key="f_fin_in")
periodo_act = f"{f_ini.strftime('%Y-%m-%d')} al {f_fin.strftime('%Y-%m-%d')}"

url_wa_raw = st.sidebar.text_input("URL Custodios:", value="https://eit-toma-fisica-4wrzxgv3pgzfdr8agv9mtm.streamlit.app/", key="url_wa_in")
url_wa_limpia = "".join(str(url_wa_raw or "").split()).strip()
if url_wa_limpia:
    msg_wa = f"CONTROL DE INVENTARIO EIT\n\nEnlace para registro de tomas físicas:\n{url_wa_limpia}"
    st.sidebar.markdown(
        f'<a href="https://api.whatsapp.com/send?text={urllib.parse.quote(msg_wa)}" target="_blank" style="display:block;text-align:center;padding:8px;background:#16A34A;color:#FFF!important;font-weight:700;border-radius:6px;text-decoration:none;">Enviar por WhatsApp</a>',
        unsafe_allow_html=True,
    )

if st.sidebar.button("Refrescar datos", use_container_width=True, key="btn_ref_top"):
    leer_gs.clear()
    cargar_pivot.clear()
    cargar_nac.clear()
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### Filtros de Análisis & Reportes")
incluir_gs_rep = st.sidebar.checkbox("Incluir registros de Google Sheets", value=False, key="chk_inc_gs")
origen_label = "JSON local + Google Sheets" if incluir_gs_rep else "Solo JSON local (toma_fisica_registros.json)"
st.sidebar.caption(f"Origen activo: {origen_label}")


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


f_col1, f_col2 = st.sidebar.columns(2)
f_col1.button("Periodo/Bod act.", on_click=sincronizar_filtros_operacion, use_container_width=True, key="btn_flt_act")
f_col2.button("Ver todo", on_click=resetear_filtros_estadistica, use_container_width=True, key="btn_flt_all")

df_base_rep = prep_df(datos_combinados(incluir_gs=incluir_gs_rep))
muestreo_rep = "Muestreo" in tipo_proc

periodos_disp = ["TODOS"] + sorted(set([periodo_act] + [p for p in df_base_rep["Periodo"].unique() if p]), reverse=True)
if "flt_per" not in st.session_state or st.session_state["flt_per"] not in periodos_disp:
    st.session_state["flt_per"] = periodo_act if periodo_act in periodos_disp else "TODOS"
f_per = st.sidebar.selectbox("Periodo:", periodos_disp, key="flt_per")

bodegas_disp = ["TODAS"] + sorted(set(BODEGAS_OFICIALES + [b for b in df_base_rep["Bodega"].unique() if b]))
if "flt_bod" not in st.session_state or st.session_state["flt_bod"] not in bodegas_disp:
    st.session_state["flt_bod"] = bod_norm if bod_norm in bodegas_disp else "TODAS"
f_bod = st.sidebar.selectbox("Bodega:", bodegas_disp, key="flt_bod")

ciudades_disp = ["TODAS"] + sorted(set(["QUITO", "GUAYAQUIL", "CUENCA", ciudad_def] + [c for c in df_base_rep["Ciudad"].unique() if c]))
if "flt_ciu" not in st.session_state or st.session_state["flt_ciu"] not in ciudades_disp:
    st.session_state["flt_ciu"] = ciudad_def if ciudad_def in ciudades_disp else "TODAS"
f_ciu = st.sidebar.selectbox("Ciudad:", ciudades_disp, key="flt_ciu")

custodios_disp = ["TODOS"] + sorted([c for c in df_base_rep["Correo_Usuario"].unique() if c]) if not df_base_rep.empty else ["TODOS"]
if "flt_cus" not in st.session_state or st.session_state["flt_cus"] not in custodios_disp:
    st.session_state["flt_cus"] = "TODOS"
f_cus = st.sidebar.selectbox("Custodio:", custodios_disp, key="flt_cus")

skus_disp = ["TODOS"] + sorted([s for s in df_base_rep["COD_SKU"].unique() if s]) if not df_base_rep.empty else ["TODOS"]
if "flt_sku" not in st.session_state or st.session_state["flt_sku"] not in skus_disp:
    st.session_state["flt_sku"] = "TODOS"
f_sku = st.sidebar.selectbox("COD_SKU:", skus_disp, key="flt_sku")

provs_disp = ["TODOS"] + sorted([p for p in df_base_rep["PROVEEDOR"].unique() if p]) if not df_base_rep.empty else ["TODOS"]
if "flt_prv" not in st.session_state or st.session_state["flt_prv"] not in provs_disp:
    st.session_state["flt_prv"] = "TODOS"
f_prv = st.sidebar.selectbox("Proveedor:", provs_disp, key="flt_prv")

if st.sidebar.button("Refrescar estadísticas", use_container_width=True, key="btn_ref_stats_bottom"):
    leer_gs.clear()
    cargar_pivot.clear()
    cargar_nac.clear()
    st.rerun()

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

if es_admin:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### Métricas & Sync (Superusuario)")
    if not df_rep.empty:
        bod_m = bod_norm if f_bod == "TODAS" else norm_bod(f_bod)
        df_rep_ok = df_rep[~df_rep["ESTADO"].isin(["PENDIENTE_REVERSO", "REVERSO_RECHAZADO"])]
        teo_b = contar_teoricos_bodega(nac_rows, bod_m, muestreo=muestreo_rep)
        cont_b = int(df_rep_ok[df_rep_ok["Bodega"] == bod_m]["COD_SKU"].nunique())
        pend_b = max(teo_b - cont_b, 0) if teo_b > 0 else 0
        q_tot = float(df_rep_ok["Q"].sum())
        val_tot = float(df_rep_ok["COSTO TOTAL"].sum())
        dif_tot = float(df_rep_ok["COSTO DIF"].sum())
        st.sidebar.caption(
            f"Bod {bod_m} -> Teór:{teo_b} | Cont:{cont_b} | Pend:{pend_b}\n"
            f"Tx:{len(df_rep)} | Q:{q_tot:g} | Val:USD {val_tot:,.2f} | Dif:USD {dif_tot:,.2f}"
        )
        st.sidebar.download_button(
            f"Generar Reporte Excel ({len(df_rep)})",
            data=excel_bytes(df_rep, nac_rows=nac_rows, muestreo=muestreo_rep),
            file_name=f"reporte_toma_fisica_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="btn_dl_side",
        )

    modo_dup_ui = st.sidebar.radio("Si hay duplicados en Google:", ["Omitir duplicados", "Sobrescribir", "Duplicar"], key="modo_dup_in")
    modo_dup_map = {"Omitir duplicados": "omitir", "Sobrescribir": "sobreescribir", "Duplicar": "duplicar"}
    if st.sidebar.button("Sincronizar JSON a Google", use_container_width=True, key="btn_sync_gs"):
        pend = cargar_json()
        if pend:
            con = {norm_mayus_limpio(x.get("ID_Tx", "")) for x in leer_gs()} | ids_enviados()
            dups = sum(1 for r in pend if norm_mayus_limpio(r.get("ID_Tx", "")) in con)
            cnt_ok, cnt_om, cnt_dup, cnt_ov = 0, 0, 0, 0
            for r in pend:
                ok_s, est_s, _ = enviar_gs(r, con, modo_dup=modo_dup_map[modo_dup_ui])
                if ok_s:
                    if est_s == "omitido":
                        cnt_om += 1
                    elif est_s == "duplicado":
                        cnt_dup += 1
                    elif est_s == "sobreescrito":
                        cnt_ov += 1
                    else:
                        cnt_ok += 1
            st.sidebar.success(f"Dups:{dups} | OK:{cnt_ok} | Omit:{cnt_om} | Dup:{cnt_dup} | Sobr:{cnt_ov}")
        else:
            st.sidebar.info("Sin registros locales.")

    with st.sidebar.expander("Limpieza de Pruebas (Pre-Campo)"):
        confirmar_limpieza = st.checkbox("Respaldar y limpiar JSON de pruebas", value=False, key="chk_limp_pruebas")
        if st.button("Limpiar registros de prueba", use_container_width=True, disabled=not confirmar_limpieza, key="btn_limp_pruebas"):
            carp = respaldar_y_limpiar_pruebas()
            st.session_state["cap_v"] += 1
            st.success(f"Respaldo creado en: {carp}.")
            st.rerun()

pend_logs_cnt = sum(1 for lg in logs_json if str(lg.get("Enviado_GS", "NO")).upper() != "SI")
st.sidebar.markdown("---")
st.sidebar.markdown(
    f"**PIVOT:** {fecha_arch(pivot_path)}<br>"
    f"**NAC:** {fecha_arch(nac_path)}<br>"
    f"**JSON:** {fecha_arch(json_path)} ({len(regs_json)})<br>"
    f"**Etiquetas Cola:** {len(st.session_state['cola_etiquetas'])}<br>"
    f"**{PESTANA_LOGS_GS}:** {len(logs_json)} ({pend_logs_cnt} sin enviar)",
    unsafe_allow_html=True,
)

# ==============================================================================
# CABECERA + PESTAÑAS
# ==============================================================================
st.markdown('<div class="main-title">Toma Física & Muestreo V7 gPro</div>', unsafe_allow_html=True)
st.caption(
    f"Script: `{RUTA_SCRIPT}` | Modificado: `{f_script}` | Periodo: `{'HABILITADO' if st.session_state['per_abierto'] else 'INHABILITADO'}` | "
    f"COD_LABEL: `{st.session_state['cod_label_act']}` | Cola etiquetas: `{len(st.session_state['cola_etiquetas'])}`"
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
        "Revísalos en la pestaña *Dashboard Estadístico & Avance Visual* -> *Autorización de Reversos Pendientes*."
    )

tab_op, tab_dash, tab_etq, tab_exp = st.tabs([
    "Operación y Captura de Campo",
    "Dashboard Estadístico & Avance Visual",
    f"Etiquetas A4 (~5,8×2,5 cm) [{len(st.session_state['cola_etiquetas'])}]",
    f"Exportación, Detalle & {PESTANA_LOGS_GS}",
])

with tab_op:
    st.markdown('<div class="card-box">', unsafe_allow_html=True)
    if not st.session_state["per_abierto"]:
        st.warning("Periodo INHABILITADO: habilítalo desde el panel izquierdo cuando corresponda.")
    if (not os.path.exists(nac_path)) and (not os.path.exists(pivot_path)):
        st.info("Sin `NAC.xlsx` ni `PIVOT` locales: puedes digitar o escanear `COD_SKU` y registrar directo.")

    c1, c2, c3 = st.columns(3)
    ciudad = c1.selectbox("Ciudad:", ciudades, key="ciu_in")
    c2.markdown(f"**Bodega:** `{bod_norm}`")
    ubic_raw = c3.text_input("Ubicación (MAYÚSCULAS):", placeholder="EJ: 5N2611", key=f"ubic_in_{cap_v}")
    ubic = norm_mayus_limpio(ubic_raw)

    idx_b = resolver_col_bodega(nac_rows, bod_norm)
    muestreo = "Muestreo" in tipo_proc
    ya_aud = {
        norm_mayus_limpio(r.get("COD_SKU", ""))
        for r in regs_json
        if str(r.get("Periodo", "")).strip() == periodo_act and norm_bod(r.get("Bodega", "")) == bod_norm and es_computable_en_consolidado(r)
    }

    skus_unicos, desc_map = obtener_skus_teoricos_bodega(nac_rows, dict_maestro, bod_norm, muestreo=muestreo)
    tot_teo_bod = len(skus_unicos)
    aud_bod_cnt = len([s for s in skus_unicos if s in ya_aud])
    pend_bod_cnt = max(tot_teo_bod - aud_bod_cnt, 0)
    st.caption(f"Guía SKU Bodega `{bod_norm}` (Col V + saldo >0) -> SKUs Teóricos: **{tot_teo_bod}** | Auditados: **{aud_bod_cnt}** | Pendientes: **{pend_bod_cnt}**")

    opts = ["-- SELECCIONE O ESCANEE SKU --"] + [
        f"{s} | {'[AUDITADO]' if s in ya_aud else '[PENDIENTE]'} {desc_map.get(s, '')}" for s in skus_unicos
    ]
    sel_combo = st.selectbox("Catálogo Teórico:", opts, key=f"cb_{bod_norm}_{cap_v}")
    pre_sku = norm_mayus_limpio(sel_combo.split(" | ")[0]) if sel_combo != "-- SELECCIONE O ESCANEE SKU --" else ""

    with st.expander("Cámara: Lector de Código de Barras (Principal) & Foto Evidencia (Secundaria)"):
        t_scan, t_foto = st.tabs(["1. Lector Código de Barras", "2. Foto Evidencia"])
        with t_scan:
            render_lector_codigos_barras(f"{bod_norm}_{cap_v}")
        with t_foto:
            foto_cap = st.camera_input("Tomar fotografía de evidencia (etiqueta / ubicación):", key=f"cam_{bod_norm}_{cap_v}")
            if foto_cap is not None:
                st.caption("Foto de evidencia guardada en la sesión actual.")

    c4, c5, c6 = st.columns([2.5, 1.5, 1])
    sku_raw = c4.text_input("COD_SKU (MAYÚSCULAS):", value=pre_sku, key=f"sku_{bod_norm}_{cap_v}")
    cod_sku = norm_mayus_limpio(sku_raw)

    cant = c5.number_input("Cantidad (Q):", step=1, value=1, key=f"q_{bod_norm}_{cap_v}")
    motivo = ""
    pin_ingresado = ""
    pin_valido_cust = False

    if cant < 0:
        r_m1, r_m2 = st.columns([2.2, 1.3])
        motivo = norm_mayus_limpio(r_m1.text_input("Motivo Reverso (Col K obligatorio si Q < 0):", key=f"mot_{bod_norm}_{cap_v}"))
        if not es_admin:
            pin_dia = generar_pin_diario(st.session_state["pin_reverso_act"], date.today())
            pin_ingresado = r_m2.text_input("PIN Día / Autorización (opcional):", max_chars=6, key=f"pin_{bod_norm}_{cap_v}").strip()
            pin_valido_cust = bool(pin_ingresado) and (pin_ingresado == pin_dia or pin_ingresado == str(st.session_state["pin_reverso_act"]).strip())
            if pin_ingresado and not pin_valido_cust:
                st.warning("PIN incorrecto: el reverso quedará **PENDIENTE_APROBACION** para revisión del admin.")
            elif pin_valido_cust:
                st.success("PIN válido: reverso autorizado y aplicado al consolidado.")
            else:
                st.info("Sin PIN ingresado: el reverso queda **PENDIENTE_APROBACION** en la bandeja del administrador.")

    c6.write("")
    c6.write("")
    ok_rev = (cant >= 0) or bool(motivo)
    ubic_valida = bool(ubic) and ubic not in ["UBICACIÓN", "UBICACION", "ESPECIFIQUE"]
    puede = st.session_state["per_abierto"] and ubic_valida and bool(cod_sku) and (cant != 0) and ok_rev
    btn_reg = c6.button("REGISTRAR", use_container_width=True, disabled=not puede, key="btn_reg_tx")

    e_col1, e_col2, e_col3 = st.columns([2.2, 1.0, 1.8])
    marcar_etq = e_col1.checkbox("Marcar este SKU para etiqueta (~5,8×2,5 cm)", value=False, key=f"chk_etq_{bod_norm}_{cap_v}")
    copias_etq = e_col2.number_input("Copias:", min_value=1, max_value=30, value=1, step=1, key=f"cop_etq_{bod_norm}_{cap_v}")
    if e_col3.button("Añadir SKU a cola etiqueta", use_container_width=True, disabled=not bool(cod_sku), key="btn_add_etq_rapido"):
        d_etq = dict_maestro.get(cod_sku, {}).get("DESCRIPCIÓN") or dict_maestro.get(cod_sku, {}).get("DESCRIPCION") or desc_map.get(cod_sku, "")
        agregar_etiqueta_cola(cod_sku, d_etq, bod_norm, ubic, usuario_lbl=usuario, copias=copias_etq)
        st.success(f"Añadido {cod_sku} ({copias_etq} copia(s)) a la cola de etiquetas.")
        st.rerun()

    faltan = []
    if not st.session_state["per_abierto"]:
        faltan.append("Periodo habilitado")
    if not ubic_valida:
        faltan.append("Ubicación")
    if not cod_sku:
        faltan.append("COD_SKU")
    if cant == 0:
        faltan.append("Cantidad (Q) distinta de 0")
    if not ok_rev:
        faltan.append("Motivo Reverso")
    if faltan:
        st.caption("Falta para habilitar REGISTRAR: " + ", ".join(faltan) + ".")
    st.markdown("</div>", unsafe_allow_html=True)

    st_teo, st_nac, c_unit_vista = 0.0, 0.0, 0.0
    if cod_sku:
        art = dict_maestro.get(cod_sku, {})

        def flex(*ks):
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
                    st_teo = num_es(row[idx_b], max_abs=1e12) if (idx_b < len(row) and sku_habilitado_col_v(row, muestreo=muestreo)) else 0.0
                    c_unit_vista = obtener_costo_nac_col_w(row)
                    break
        if c_unit_vista <= 0:
            c_unit_vista = next(
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
        aplica_rev_ya = (cant >= 0) or es_admin or pin_valido_cust
        aporte_proy = float(cant) if aplica_rev_ya else 0.0
        q_acum_calc = round(float(prev_comp) + aporte_proy, 4)
        dif_proy = round(q_acum_calc - float(st_teo), 4)
        nota_proy = " (Reverso pendiente de aprobación)" if (cant < 0 and not aplica_rev_ya) else ""
        st.info(
            f"**SKU:** {cod_sku} | **Desc:** {flex('DESCRIPCIÓN', 'DESCRIPCION', 'DESC')} | **Und:** {flex('UNIDAD', 'UND')} | "
            f"**Prov:** {flex('PROVEEDOR', 'SUP', 'FABRICANTE')} | **Costo Unit (W):** ${c_unit_vista:,.4f} | "
            f"**Acum. Oficial Prev:** {prev_comp:g} | **Proyectado Q:** {q_acum_calc:g} | **Dif. Proy:** {dif_proy:g}{nota_proy}"
        )
        m1, m2 = st.columns(2)
        m1.markdown(f'<div class="saldo-box"><span class="saldo-lbl">QSA_Teorico ({bod_norm})</span><span class="saldo-val">{st_teo:g}</span></div>', unsafe_allow_html=True)
        m2.markdown(f'<div class="saldo-box"><span class="saldo-lbl">Nacional</span><span class="saldo-val">{st_nac:g}</span></div>', unsafe_allow_html=True)

    if btn_reg:
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
                    st_teo_r = num_es(row[idx_b], max_abs=1e12) if (idx_b < len(row) and sku_habilitado_col_v(row, muestreo=muestreo)) else 0.0
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
        motivo_final = motivo
        if cant < 0:
            if es_admin:
                est_rev_ini = "APROBADO"
            elif pin_valido_cust:
                est_rev_ini = "APROBADO"
                motivo_final = f"{motivo} [PIN_AUTORIZADO]".strip()
            else:
                est_rev_ini = "PENDIENTE_APROBACION"
        else:
            est_rev_ini = "NO_APLICA"

        q_acum_final = round(float(prev_comp) + (float(cant) if est_rev_ini != "PENDIENTE_APROBACION" else 0.0), 4)
        dif_f, est_f, costo_total, costo_dif, c_unit_limpio = calcular_dif_estado_costo(
            q_acum_final, st_teo_r, cant, c_unit_r, est_rev=est_rev_ini
        )
        desc_r = flex_r("DESCRIPCIÓN", "DESCRIPCION", "DESC")
        now = datetime.now()
        nuevo = {
            "Periodo": periodo_act,
            "ID_Tx": f"TX-{now.strftime('%Y%m%d%H%M%S')}",
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
            "PROVEEDOR": norm_mayus_limpio(flex_r("PROVEEDOR", "SUP", "FABRICANTE")),
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
            if est_rev_ini == "PENDIENTE_APROBACION":
                st.success(f"Reverso registrado en cola de aprobación ({cod_sku}, Q {cant:g}) -> PENDIENTE_REVERSO | GS: {est_g} ({tx_g})" if ok_g else f"Reverso guardado en cola local | GS: {est_g}")
            else:
                st.success(f"Guardado {cod_sku} ({cant:g} | Acum {q_acum_final:g}) -> {est_f} (Dif {dif_f:g} | Costo U ${c_unit_limpio:,.4f}) | GS: {est_g} ({tx_g})" if ok_g else f"Guardado local OK | GS: {est_g}")
            st.rerun()

    st.markdown("---")
    st.markdown("### Últimos Registros")
    prev_df = prep_df(datos_combinados(incluir_gs=False))
    st.markdown(tabla_md(prev_df.iloc[::-1], ["Timestamp", "Bodega", "Ubicacion", "COD_SKU", "Q", "QSA_Teorico", "DIFERENCIA", "ESTADO", "Estado_Reverso", "COSTO UNITARIO", "COSTO DIF"], 6))

with tab_dash:
    st.markdown("### Panorama General de Toma Física & Avance Nacional")
    st.caption(
        f"**Origen:** `{origen_label}` | **Filas base:** `{len(df_base_rep)}` -> **Filtradas:** `{len(df_rep)}` | "
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
                key_suf="rev_pend_tbl",
            )
            lista_tx_pend = [norm_mayus_limpio(x) for x in pend_rev_df["ID_Tx"].tolist() if norm_mayus_limpio(x)]
            sel_tx_rev = st.selectbox("Seleccionar reverso pendiente:", lista_tx_pend, key="sel_tx_rev_admin")
            pin_dia_act = generar_pin_diario(st.session_state["pin_reverso_act"], date.today())
            st.caption(f"PIN diario activo ({date.today().strftime('%Y-%m-%d')}): **`{pin_dia_act}`** (PIN base: `{st.session_state['pin_reverso_act']}`)")
            r_col1, r_col2 = st.columns(2)
            if r_col1.button("Aprobar Reverso", use_container_width=True, key="btn_aprobar_rev"):
                ok_r, reg_mod = resolver_reverso_pendiente(sel_tx_rev, aprobar=True, admin_usr=usuario)
                if ok_r and reg_mod is not None:
                    enviar_gs(reg_mod, modo_dup="sobreescribir")
                    st.success(f"Reverso {sel_tx_rev} aprobado y consolidado aplicado.")
                    st.rerun()
            if r_col2.button("Rechazar Reverso", use_container_width=True, key="btn_rechazar_rev"):
                ok_r, reg_mod = resolver_reverso_pendiente(sel_tx_rev, aprobar=False, admin_usr=usuario)
                if ok_r and reg_mod is not None:
                    enviar_gs(reg_mod, modo_dup="sobreescribir")
                    st.info(f"Reverso {sel_tx_rev} rechazado; no altera el consolidado oficial.")
                    st.rerun()
        st.divider()

    if df_rep.empty:
        st.info("No hay registros que coincidan con los filtros seleccionados en el panel izquierdo.")
    else:
        df_rep_of = df_rep[~df_rep["ESTADO"].isin(["PENDIENTE_REVERSO", "REVERSO_RECHAZADO"])]
        sob_cnt = int((df_rep_of["ESTADO"] == "SOBRANTE").sum())
        fal_cnt = int((df_rep_of["ESTADO"] == "FALTANTE").sum())
        ok_cnt = int((df_rep_of["ESTADO"] == "CUADRADO (OK)").sum())
        pend_r_cnt = int((df_rep["ESTADO"] == "PENDIENTE_REVERSO").sum())

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Transacciones (Tx)", f"{len(df_rep):,}")
        k2.metric("SKUs Únicos Contados", f"{df_rep_of['COD_SKU'].nunique():,}")
        k3.metric("Cantidad Física (Q)", f"{df_rep_of['Q'].sum():,.2f}")
        k4.metric("Valor Contado (USD)", f"${df_rep_of['COSTO TOTAL'].sum():,.2f}")
        k5.metric("Impacto Dif. (USD)", f"${df_rep_of['COSTO DIF'].sum():,.2f}")

        st.markdown(
            f"**Estados en filtro:** &nbsp; **SOBRANTE:** `{sob_cnt}` &nbsp;|&nbsp; **FALTANTE:** `{fal_cnt}` "
            f"&nbsp;|&nbsp; **CUADRADO (OK):** `{ok_cnt}` &nbsp;|&nbsp; **PENDIENTE_REVERSO:** `{pend_r_cnt}`"
        )
        df_bod_av = construir_resumen_bodega(df_rep, nac_rows, muestreo=muestreo_rep, bodega_filtro=f_bod)
        st.divider()
        render_graficos_tomas_y_auditoria(df_rep_of, df_bod_av)
        st.divider()
        g1, g2 = st.columns(2)
        with g1:
            render_lista_visual(df_bod_av, "Bodega", ["SKU_Contados", "SKU_Pendientes", "SKU_Teoricos", "Avance_Pct"], "Avance de SKUs por Bodega (Contados vs Pendientes)")
        with g2:
            df_cus_g = df_rep_of.groupby("Correo_Usuario", as_index=False).agg(Transacciones=("ID_Tx", "count"), Unidades_Q=("Q", "sum"), Costo_Dif_USD=("COSTO DIF", "sum")).sort_values("Transacciones", ascending=False)
            render_lista_visual(df_cus_g, "Correo_Usuario", ["Transacciones", "Unidades_Q", "Costo_Dif_USD"], "Actividad por Custodio (Tx y Unidades Q)")
        st.divider()
        g3, g4 = st.columns(2)
        with g3:
            df_ciu_g = df_rep_of.groupby("Ciudad", as_index=False).agg(Transacciones=("ID_Tx", "count"), Unidades_Q=("Q", "sum"), Costo_Dif_USD=("COSTO DIF", "sum")).sort_values("Unidades_Q", ascending=False)
            render_lista_visual(df_ciu_g, "Ciudad", ["Transacciones", "Unidades_Q", "Costo_Dif_USD"], "Distribución por Ciudad (Tx y Unidades Q)")
        with g4:
            df_est_g = df_rep.groupby("ESTADO", as_index=False).agg(Registros=("ID_Tx", "count"), Unidades_Q=("Q", "sum"), Costo_Dif_USD=("COSTO DIF", "sum"))
            render_lista_visual(df_est_g, "ESTADO", ["Registros", "Unidades_Q", "Costo_Dif_USD"], "Desglose por Estado (Registros y Costo Dif USD)")
        st.divider()
        st.markdown("#### Matriz de Avance por Bodega")
        if not df_bod_av.empty:
            render_tabla_segura(
                df_bod_av,
                ["Bodega", "SKU_Teoricos", "SKU_Contados", "SKU_Pendientes", "Avance_Pct", "Tx", "Cantidad_Q", "Sobrantes", "Faltantes", "Cuadrados", "Costo_Dif_USD"],
                max_filas=20,
                key_suf="matriz_bod",
            )
        st.markdown("#### Detalle Filtrado (SKU, Proveedor, Cantidades y Diferencia)")
        render_tabla_segura(
            df_rep.iloc[::-1],
            ["Timestamp", "Correo_Usuario", "Ciudad", "Bodega", "PROVEEDOR", "COD_SKU", "Q", "QSA_Teorico", "DIFERENCIA", "ESTADO", "Estado_Reverso", "COSTO UNITARIO", "COSTO DIF"],
            max_filas=30,
            key_suf="dash_det_filt",
        )

with tab_etq:
    st.markdown("### Cola de Etiquetas (~5,8 × 2,5 cm) & Armado de Hoja A4")
    st.caption(f"Campos impresos: `COD_SKU`, `DESCRIPCIÓN` y `FECHA | LOGIN` • `COD_LABEL` activo: `{st.session_state['cod_label_act']}`")
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
        render_tabla_segura(pd.DataFrame(cola_act), ["COD_SKU", "DESCRIPCION", "FECHA_LOGIN", "BODEGA"], max_filas=60, key_suf="tbl_cola_etq")
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
        render_tabla_segura(df_rep.iloc[::-1], COLS_REP, max_filas=50, key_suf="detalle_rep")

    st.markdown("---")
    st.markdown(f"### Reporte de LOGS (`{PESTANA_LOGS_GS}` -> Columnas `A:C`)")
    logs_act = cargar_logs()
    pend_logs = [lg for lg in logs_act if str(lg.get("Enviado_GS", "NO")).upper() != "SI"]
    lcol1, lcol2, lcol3 = st.columns([1.6, 1.4, 2])
    if lcol1.button(f"Enviar a {PESTANA_LOGS_GS} ({len(pend_logs)})", use_container_width=True, key="btn_send_logs"):
        if pend_logs:
            ok_cnt_l, tot_cnt_l = enviar_logs_a_gs(pend_logs)
            guardar_logs_lista(logs_act)
            st.success(f"Enviados a {PESTANA_LOGS_GS}: {ok_cnt_l}/{tot_cnt_l} registros.")
            st.rerun()
        else:
            st.info(f"No hay logs pendientes por enviar a {PESTANA_LOGS_GS}.")
    if lcol2.button("Limpiar LOGS locales", use_container_width=True, key="btn_clear_logs"):
        if logs_act:
            guardar_logs_lista([])
            st.success("Historial de LOGS locales limpiado.")
            st.rerun()
        else:
            st.info("No hay logs locales para limpiar.")
    lcol3.caption(f"Pestaña destino: `{PESTANA_LOGS_GS}` (`A:C`) | Total: `{len(logs_act)}` | Pendientes: `{len(pend_logs)}`")
    if logs_act:
        df_logs = pd.DataFrame(logs_act).iloc[::-1]
        for c_l in COLS_LOG_UI:
            if c_l not in df_logs.columns:
                df_logs[c_l] = ""
        render_tabla_segura(df_logs, COLS_LOG_UI, max_filas=40, key_suf="logs_ui")
    else:
        st.info("Sin errores registrados en este momento.")
