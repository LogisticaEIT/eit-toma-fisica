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
import urllib.parse
import urllib3
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import requests
import streamlit as st
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
    .sob { background: #DCFCE7; color: #14532D !important; padding: 10px 14px; border-radius: 8px; font-size: 18px !important; font-weight: 900; text-align: center; border: 2px solid #22C55E; }
    .fal { background: #FEE2E2; color: #7F1D1D !important; padding: 10px 14px; border-radius: 8px; font-size: 18px !important; font-weight: 900; text-align: center; border: 2px solid #EF4444; }
    .ok  { background: #FEF9C3; color: #713F12 !important; padding: 10px 14px; border-radius: 8px; font-size: 18px !important; font-weight: 900; text-align: center; border: 2px solid #EAB308; }
</style>
""",
    unsafe_allow_html=True,
)

RUTA_SCRIPT = os.path.abspath(__file__)
BASE_DIR = os.path.dirname(RUTA_SCRIPT)
st.session_state.setdefault("errores_ui", [])
st.session_state.setdefault("cap_v", 0)


def log_err(msg):
    st.session_state["errores_ui"].append(f"{datetime.now().strftime('%H:%M:%S')} | {msg}")
    st.session_state["errores_ui"] = st.session_state["errores_ui"][-10:]


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
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxGk8xG7zUy0vu-OUXwgg_qlUS8lmNlU1bh_CygmJkpieTt2jdfOkknoUjAKKis7EiA5g/exec"

BODEGAS_OFICIALES = ["01", "03", "04", "05", "3V", "5V"]
MAPEO_BOD_COL = {"01": 5, "03": 6, "04": 7, "05": 8, "3V": 9, "5V": 10}
IDX_TOTAL_NAC = 11
IDX_MOV_V = 21

CUSTODIO_BOD = {
    "alexpajaritoruiz@icloud.com": BODEGAS_OFICIALES,
    "jlogistica.eit@gmail.com": BODEGAS_OFICIALES,
    "san14mej@gmail.com": ["01"],
    "bodegagye2026@gmail.com": ["03", "3V"],
    "bodegauio.eit@gmail.com": ["05", "5V"],
}

# Estructura oficial de 22 columnas: A-U intactas + QSA_Teorico al final (Columna V)
COLS_REP = [
    "Periodo", "ID_Tx", "Timestamp", "Correo_Usuario", "Ciudad", "Bodega",
    "Ubicacion", "COD_SKU", "Q", "BOD/SKU", "Motivo", "LINEA DE INVENTARIO",
    "PROVEEDOR", "ALTERNO", "DESCRIPCIÓN", "UNIDAD", "DIFERENCIA", "ESTADO",
    "COSTO UNITARIO", "COSTO TOTAL", "COSTO DIF", "QSA_Teorico",
]


def norm_bod(v):
    if v is None:
        return ""
    s = str(v).strip().lstrip("'").replace(".0", "")
    return s.zfill(2) if s.isdigit() else s.upper()


def norm_est(v):
    s = str(v or "").strip().upper()
    if s in ["SOB", "SOBRANTE"]:
        return "SOBRANTE"
    if s in ["FAL", "FALTANTE"]:
        return "FALTANTE"
    if s in ["OK", "CUADRADO", "CUADRADO (OK)", "DE ACUERDO"]:
        return "CUADRADO (OK)"
    return s or "CUADRADO (OK)"


def num_es(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("USD", "").replace("$", "").strip()
    if not s:
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def calcular_dif_estado_costo(q_acum, st_teo, q_actual, c_unit):
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
    costo_total = round(float(c_unit) * float(q_actual), 2)
    costo_dif = round(float(c_unit) * abs(dif_fisica) * signo_costo, 2)
    return dif_fisica, estado, costo_total, costo_dif


def fmt_txt(v):
    s = str(v or "").strip()
    return f"'{s}" if s and not s.startswith("'") else s


def cargar_json():
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                d = json.load(f)
                return d if isinstance(d, list) else []
        except Exception as e:
            log_err(f"JSON leer: {e}")
            return []
    return []


def guardar_json(reg):
    regs = cargar_json()
    regs.append(reg)
    try:
        os.makedirs(os.path.dirname(json_path), exist_ok=True)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(regs, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        log_err(f"JSON guardar: {e}")
        return False


def ids_enviados():
    if os.path.exists(enviados_path):
        try:
            with open(enviados_path, "r", encoding="utf-8") as f:
                d = json.load(f)
                if isinstance(d, list):
                    return {str(x).strip().upper() for x in d if str(x).strip()}
        except Exception:
            pass
    return set()


def marcar_enviado(tx_id):
    t = str(tx_id or "").strip().upper()
    if not t:
        return
    s = ids_enviados()
    if t not in s:
        s.add(t)
        try:
            os.makedirs(os.path.dirname(enviados_path), exist_ok=True)
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
        log_err(f"GS leer: {e}")
    return out


def armar_payload_gs(reg, tx_id):
    return [
        reg.get("Periodo", ""), tx_id, reg.get("Timestamp", ""), reg.get("Correo_Usuario", ""),
        reg.get("Ciudad", ""), fmt_txt(norm_bod(reg.get("Bodega", ""))), reg.get("Ubicacion", ""),
        fmt_txt(reg.get("COD_SKU", "")), num_es(reg.get("Q", 0)), reg.get("BOD/SKU", ""), reg.get("Motivo", ""),
        reg.get("LINEA DE INVENTARIO", ""), reg.get("PROVEEDOR", ""), fmt_txt(reg.get("ALTERNO", "")),
        reg.get("DESCRIPCIÓN", ""), reg.get("UNIDAD", ""), num_es(reg.get("DIFERENCIA", 0)),
        norm_est(reg.get("ESTADO", "")), num_es(reg.get("COSTO UNITARIO", 0)),
        num_es(reg.get("COSTO TOTAL", 0)), num_es(reg.get("COSTO DIF", 0)),
        num_es(reg.get("QSA_Teorico", 0)),
    ]


def enviar_gs(reg, conocidos=None, modo_dup="omitir"):
    tx = str(reg.get("ID_Tx", "")).strip().upper()
    if conocidos is None:
        conocidos = {str(x.get("ID_Tx", "")).strip().upper() for x in leer_gs()} | ids_enviados()
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
            log_err(f"GS post HTTP {r.status_code}: {txt[:120]}")
        return ok, ("sobreescrito" if modo_dup == "sobreescribir" else ("duplicado" if modo_dup == "duplicar" else "enviado")), tx_env
    except Exception as e:
        log_err(f"GS post exc: {e}")
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
        tx = str(r.get("ID_Tx", "")).strip().upper()
        k = tx if tx else f"{r.get('Timestamp','')}|{norm_bod(r.get('Bodega',''))}|{str(r.get('COD_SKU','')).strip().upper()}|{num_es(r.get('Q',0))}"
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
    df["Ciudad"] = df["Ciudad"].astype(str).str.strip()
    df["Periodo"] = df["Periodo"].astype(str).str.strip()
    df["Correo_Usuario"] = df["Correo_Usuario"].astype(str).str.strip().str.lower()
    df["COD_SKU"] = df["COD_SKU"].astype(str).str.strip().str.upper()
    df["PROVEEDOR"] = df["PROVEEDOR"].astype(str).str.strip().str.upper()
    df["ESTADO"] = df["ESTADO"].apply(norm_est)
    for nc in ["Q", "QSA_Teorico", "DIFERENCIA", "COSTO UNITARIO", "COSTO TOTAL", "COSTO DIF"]:
        df[nc] = df[nc].apply(num_es)
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
            s = f"{v:,.2f}" if isinstance(v, float) and "COSTO" in c.upper() else (f"{v:g}" if isinstance(v, float) else str(v or ""))
            vals.append(s.replace("|", "/").replace("\n", " ").strip())
        lines.append("| " + " | ".join(vals) + " |")
    return "\n".join(lines)


def resolver_col_bodega(nac_rows, bod):
    idx_b = MAPEO_BOD_COL.get(bod, None)
    if idx_b is not None:
        return idx_b
    if nac_rows and len(nac_rows) > 0:
        for i, h in enumerate(nac_rows[0]):
            if norm_bod(h) == bod:
                return i
    return 5


def contar_teoricos_bodega(nac_rows, bod, muestreo=False):
    idx_b = resolver_col_bodega(nac_rows, bod)
    skus = set()
    if len(nac_rows) > 1:
        for row in nac_rows[1:]:
            if not row or not row[0]:
                continue
            sv = str(row[0]).strip().upper()
            if not sv or sv in ["NONE", "NAN"]:
                continue
            stb = num_es(row[idx_b]) if idx_b < len(row) else 0.0
            if stb > 0:
                if muestreo:
                    mv = str(row[IDX_MOV_V] or "").strip().upper() if IDX_MOV_V < len(row) else ""
                    if mv in ["", "SIN MOV"]:
                        continue
                skus.add(sv)
    return len(skus)


def construir_metricas_avance(df, nac_rows, muestreo=False):
    if df.empty:
        return []
    filas = []
    for (per_k, bod_k, cus_k), grp in df.groupby(["Periodo", "Bodega", "Correo_Usuario"]):
        teo_skus = contar_teoricos_bodega(nac_rows, bod_k, muestreo=muestreo)
        cont_skus = int(grp["COD_SKU"].nunique())
        pend_skus = max(teo_skus - cont_skus, 0) if teo_skus > 0 else 0
        pct_avance = round((cont_skus / teo_skus) * 100.0, 2) if teo_skus > 0 else 0.0
        filas.append([
            str(per_k),
            str(bod_k),
            str(cus_k),
            int(len(grp)),
            int(teo_skus),
            int(cont_skus),
            int(pend_skus),
            float(pct_avance),
            round(float(grp["Q"].sum()), 2),
            int((grp["ESTADO"] == "SOBRANTE").sum()),
            int((grp["ESTADO"] == "FALTANTE").sum()),
            int((grp["ESTADO"] == "CUADRADO (OK)").sum()),
            round(float(grp["COSTO TOTAL"].sum()), 2),
            round(float(grp["COSTO DIF"].sum()), 2),
        ])
    return filas


def construir_resumen_bodega(df, nac_rows, muestreo=False):
    filas = []
    for bod_k in BODEGAS_OFICIALES:
        grp = df[df["Bodega"] == bod_k] if not df.empty else pd.DataFrame(columns=COLS_REP)
        teo_skus = contar_teoricos_bodega(nac_rows, bod_k, muestreo=muestreo)
        cont_skus = int(grp["COD_SKU"].nunique()) if not grp.empty else 0
        pend_skus = max(teo_skus - cont_skus, 0) if teo_skus > 0 else 0
        pct_avance = round((cont_skus / teo_skus) * 100.0, 1) if teo_skus > 0 else 0.0
        if not grp.empty or teo_skus > 0:
            filas.append({
                "Bodega": bod_k,
                "SKU_Teoricos": teo_skus,
                "SKU_Contados": cont_skus,
                "SKU_Pendientes": pend_skus,
                "Avance_Pct": pct_avance,
                "Tx": int(len(grp)),
                "Cantidad_Q": round(float(grp["Q"].sum()), 2) if not grp.empty else 0.0,
                "Sobrantes": int((grp["ESTADO"] == "SOBRANTE").sum()) if not grp.empty else 0,
                "Faltantes": int((grp["ESTADO"] == "FALTANTE").sum()) if not grp.empty else 0,
                "Cuadrados": int((grp["ESTADO"] == "CUADRADO (OK)").sum()) if not grp.empty else 0,
                "Costo_Total_USD": round(float(grp["COSTO TOTAL"].sum()), 2) if not grp.empty else 0.0,
                "Costo_Dif_USD": round(float(grp["COSTO DIF"].sum()), 2) if not grp.empty else 0.0,
            })
    return pd.DataFrame(filas)


def excel_bytes(df, nac_rows=None, muestreo=False):
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
    for r_idx, r_vals in enumerate(df.itertuples(index=False, name=None), start=2):
        ws.append(list(r_vals))
        for c_idx, cname in enumerate(df.columns, start=1):
            c = ws.cell(row=r_idx, column=c_idx)
            if cname in {"Bodega", "COD_SKU", "ALTERNO", "ID_Tx", "Ubicacion"}:
                c.number_format = "@"
                if c.value is not None:
                    c.value = str(c.value)
            elif cname in {"COSTO UNITARIO", "COSTO TOTAL", "COSTO DIF", "Q", "QSA_Teorico", "DIFERENCIA"}:
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

    metricas_rows = construir_metricas_avance(df, nac_rows or [], muestreo=muestreo)
    for r_vals in metricas_rows:
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
                cand = [str(c).strip().upper() if c is not None else "" for c in r_try]
                if any("SKU" in h or "COD" in h or "ART" in h for h in cand):
                    ck = next((i for i, h in enumerate(cand) if "SKU" in h or "COD" in h or "ART" in h), None)
                    if ck is not None:
                        for row in rows[idx_r + 1:]:
                            if len(row) > ck and row[ck] is not None:
                                sv = str(row[ck]).strip().upper()
                                if sv and sv not in ["NONE", "NAN"]:
                                    dm.setdefault(sv, {})
                                    for i, val in enumerate(row):
                                        if i < len(cand) and cand[i] and val is not None:
                                            vs = str(val).strip()
                                            if vs and vs.upper() not in ["NONE", "NAN"]:
                                                dm[sv][cand[i]] = vs
                    break
    except Exception as e:
        log_err(f"PIVOT: {e}")
    return dm


@st.cache_data(show_spinner=False)
def cargar_nac(p):
    if not os.path.exists(p):
        return []
    try:
        wb = openpyxl.load_workbook(p, data_only=True)
        return list(wb.active.iter_rows(values_only=True))
    except Exception as e:
        log_err(f"NAC: {e}")
        return []


dict_maestro = cargar_pivot(pivot_path)
nac_rows = cargar_nac(nac_path)
regs_json = cargar_json()
f_script = fecha_arch(RUTA_SCRIPT)
cap_v = st.session_state["cap_v"]

st.markdown('<div class="main-title">Toma Física & Muestreo V7 gPro</div>', unsafe_allow_html=True)
st.caption(f"Script: `{RUTA_SCRIPT}` | Modificado: `{f_script}`")

if st.session_state["errores_ui"]:
    st.error("Errores: " + " || ".join(st.session_state["errores_ui"][-3:]))

# ==============================================================================
# PANEL IZQUIERDO: CONTROL OPERATIVO + FILTROS MAESTROS DE ESTADÍSTICAS
# ==============================================================================
usuario = st.sidebar.text_input("Correo Usuario / Custodio:", value="alexpajaritoruiz@icloud.com").strip().lower()
es_admin = usuario in {"alexpajaritoruiz@icloud.com", "jlogistica.eit@gmail.com"} or "ALEX" in usuario.upper() or "PAJARI" in usuario.upper()
bods_perm = BODEGAS_OFICIALES if es_admin else CUSTODIO_BOD.get(usuario, ["01"])

tipo_proc = st.sidebar.radio("Modo:", ["Control Mensual de Inventario (Muestreo)", "Toma Física Completa"])
if "bod_act" not in st.session_state or st.session_state["bod_act"] not in bods_perm:
    st.session_state["bod_act"] = bods_perm[0]


def limpiar_controles_bodega():
    st.session_state["cap_v"] = st.session_state.get("cap_v", 0) + 1


bod_norm = norm_bod(st.sidebar.radio("Bodega:", bods_perm, key="bod_act", on_change=limpiar_controles_bodega))

ciudades = ["Quito"] if bod_norm in ["01", "05", "5V"] else (["Guayaquil"] if bod_norm in ["03", "3V"] else (["Cuenca"] if bod_norm == "04" else ["Quito", "Guayaquil", "Cuenca", "Otras"]))

if "per_abierto" not in st.session_state:
    st.session_state["per_abierto"] = False
if es_admin:
    st.session_state["per_abierto"] = st.sidebar.toggle("Periodo abierto (Admin)", value=st.session_state["per_abierto"])
else:
    st.sidebar.caption("Estado periodo: " + ("ABIERTO" if st.session_state["per_abierto"] else "CERRADO"))

f_ini = st.sidebar.date_input("Desde:", value=date.today())
f_fin = st.sidebar.date_input("Hasta:", value=date.today())
periodo_act = f"{f_ini.strftime('%Y-%m-%d')} al {f_fin.strftime('%Y-%m-%d')}"

url_wa = st.sidebar.text_input("URL Custodios:", value="https://eit-toma-fisica-4wrzxgv3pgzfdr8agv9mtm.streamlit.app/").strip()
if url_wa:
    msg_wa = f"CONTROL DE INVENTARIO EIT\n\nEnlace para registro de tomas físicas:\n{url_wa.splitlines()[0].strip()}"
    st.sidebar.markdown(
        f'<a href="https://api.whatsapp.com/send?text={urllib.parse.quote(msg_wa)}" target="_blank" style="display:block;text-align:center;padding:8px;background:#16A34A;color:#FFF!important;font-weight:700;border-radius:6px;text-decoration:none;">Enviar por WhatsApp</a>',
        unsafe_allow_html=True,
    )

if st.sidebar.button("Refrescar datos", use_container_width=True):
    leer_gs.clear()
    cargar_pivot.clear()
    cargar_nac.clear()
    st.rerun()

# Filtros globales en el panel izquierdo para Dashboard, Estadísticas y Exportación
st.sidebar.markdown("---")
st.sidebar.markdown("### Filtros de Análisis & Reportes")
incluir_gs_rep = st.sidebar.checkbox("Incluir registros de Google Sheets", value=False)
df_base_rep = prep_df(datos_combinados(incluir_gs=incluir_gs_rep))
muestreo_rep = "Muestreo" in tipo_proc

f_per = st.sidebar.selectbox("Periodo:", ["TODOS"] + sorted([p for p in df_base_rep["Periodo"].unique() if p], reverse=True)) if not df_base_rep.empty else "TODOS"
f_bod = st.sidebar.selectbox("Bodega:", ["TODAS"] + sorted([b for b in df_base_rep["Bodega"].unique() if b])) if not df_base_rep.empty else "TODAS"
f_ciu = st.sidebar.selectbox("Ciudad:", ["TODAS"] + sorted([c for c in df_base_rep["Ciudad"].unique() if c])) if not df_base_rep.empty else "TODAS"
f_cus = st.sidebar.selectbox("Custodio:", ["TODOS"] + sorted([c for c in df_base_rep["Correo_Usuario"].unique() if c])) if not df_base_rep.empty else "TODOS"
f_sku = st.sidebar.text_input("COD_SKU:", value="").strip().upper()
f_prv = st.sidebar.text_input("Proveedor:", value="").strip().upper()

df_rep = df_base_rep.copy()
if not df_rep.empty:
    if f_per != "TODOS":
        df_rep = df_rep[df_rep["Periodo"] == f_per]
    if f_bod != "TODAS":
        df_rep = df_rep[df_rep["Bodega"] == f_bod]
    if f_ciu != "TODAS":
        df_rep = df_rep[df_rep["Ciudad"] == f_ciu]
    if f_cus != "TODOS":
        df_rep = df_rep[df_rep["Correo_Usuario"] == f_cus]
    if f_sku:
        df_rep = df_rep[df_rep["COD_SKU"].str.contains(f_sku, na=False, regex=False)]
    if f_prv:
        df_rep = df_rep[df_rep["PROVEEDOR"].str.contains(f_prv, na=False, regex=False)]

if es_admin:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### Métricas & Sync (Superusuario)")
    if not df_rep.empty:
        bod_m = bod_norm if f_bod == "TODAS" else f_bod
        teo_b = contar_teoricos_bodega(nac_rows, bod_m, muestreo=muestreo_rep)
        cont_b = int(df_rep[df_rep["Bodega"] == bod_m]["COD_SKU"].nunique())
        pend_b = max(teo_b - cont_b, 0) if teo_b > 0 else 0
        q_tot = float(df_rep["Q"].sum())
        val_tot = float(df_rep["COSTO TOTAL"].sum())
        dif_tot = float(df_rep["COSTO DIF"].sum())
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
        )

    modo_dup_ui = st.sidebar.radio("Si hay duplicados en Google:", ["Omitir duplicados", "Sobrescribir", "Duplicar"])
    modo_dup_map = {"Omitir duplicados": "omitir", "Sobrescribir": "sobreescribir", "Duplicar": "duplicar"}
    if st.sidebar.button("Sincronizar JSON a Google", use_container_width=True):
        pend = cargar_json()
        if pend:
            con = {str(x.get("ID_Tx", "")).strip().upper() for x in leer_gs()} | ids_enviados()
            dups = sum(1 for r in pend if str(r.get("ID_Tx", "")).strip().upper() in con)
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

st.sidebar.markdown("---")
st.sidebar.markdown(
    f"**PIVOT:** {fecha_arch(pivot_path)}<br>"
    f"**NAC:** {fecha_arch(nac_path)}<br>"
    f"**JSON:** {fecha_arch(json_path)} ({len(regs_json)})",
    unsafe_allow_html=True,
)

# ==============================================================================
# PESTAÑAS PRINCIPALES: OPERACIÓN | DASHBOARD VISUAL | EXPORTACIÓN
# ==============================================================================
tab_op, tab_dash, tab_exp = st.tabs([
    "Operación y Captura de Campo",
    "Dashboard Estadístico & Avance Visual",
    "Exportación & Detalle Filtrado",
])

with tab_op:
    st.markdown('<div class="card-box">', unsafe_allow_html=True)
    if not st.session_state["per_abierto"]:
        st.warning("Periodo CERRADO: el superusuario debe abrir el periodo para registrar.")
    if not os.path.exists(nac_path):
        st.warning("Sin `NAC.xlsx` en este host: digita o escanea `COD_SKU` y registra.")

    c1, c2, c3 = st.columns(3)
    ciudad = c1.selectbox("Ciudad:", ciudades, key="ciu_in")
    c2.markdown(f"**Bodega:** `{bod_norm}`")
    ubic = c3.text_input("Ubicación:", placeholder="EJ: N1311", key=f"ubic_in_{cap_v}").strip().upper()

    idx_b = resolver_col_bodega(nac_rows, bod_norm)
    muestreo = "Muestreo" in tipo_proc
    skus_bod, desc_map, ya_aud = [], {}, set()

    for r in regs_json:
        if str(r.get("Periodo", "")).strip() == periodo_act and norm_bod(r.get("Bodega", "")) == bod_norm:
            ya_aud.add(str(r.get("COD_SKU", "")).strip().upper())

    if len(nac_rows) > 1:
        for row in nac_rows[1:]:
            if not row or not row[0]:
                continue
            sv = str(row[0]).strip().upper()
            if not sv or sv in ["NONE", "NAN"]:
                continue
            stb = num_es(row[idx_b]) if idx_b < len(row) else 0.0
            if stb > 0:
                if muestreo:
                    mv = str(row[IDX_MOV_V] or "").strip().upper() if IDX_MOV_V < len(row) else ""
                    if mv in ["", "SIN MOV"]:
                        continue
                d_it = next((v for k, v in dict_maestro.get(sv, {}).items() if "DESC" in k), "")
                skus_bod.append(sv)
                desc_map[sv] = d_it

    if not skus_bod and dict_maestro:
        skus_bod = sorted(dict_maestro.keys())
        for sv in skus_bod:
            desc_map[sv] = next((v for k, v in dict_maestro.get(sv, {}).items() if "DESC" in k), "")

    skus_unicos = sorted(set(skus_bod))
    tot_teo_bod = len(skus_unicos)
    aud_bod_cnt = len([s for s in skus_unicos if s in ya_aud])
    pend_bod_cnt = max(tot_teo_bod - aud_bod_cnt, 0)
    st.caption(f"Guía Bodega `{bod_norm}` -> Teóricos: **{tot_teo_bod}** | Auditados: **{aud_bod_cnt}** | Pendientes: **{pend_bod_cnt}**")

    opts = ["-- SELECCIONE O ESCANEE SKU --"] + [
        f"{s} | {'[AUDITADO]' if s in ya_aud else '[PENDIENTE]'} {desc_map.get(s, '')}" for s in skus_unicos
    ]
    sel_combo = st.selectbox("Catálogo Teórico:", opts, key=f"cb_{bod_norm}_{cap_v}")
    pre_sku = sel_combo.split(" | ")[0].strip().upper() if sel_combo != "-- SELECCIONE O ESCANEE SKU --" else ""

    with st.expander("Cámara / Evidencia (Custodios)"):
        foto_cap = st.camera_input("Capturar etiqueta / ubicación:", key=f"cam_{bod_norm}_{cap_v}")
        if foto_cap is not None:
            st.caption("Foto capturada en sesión.")

    c4, c5, c6 = st.columns([2.5, 1.5, 1])
    cod_sku = c4.text_input("COD_SKU:", value=pre_sku, key=f"sku_{bod_norm}_{cap_v}").strip().upper()
    cant = c5.number_input("Cantidad (Q):", step=1, value=0, key=f"q_{bod_norm}_{cap_v}")
    motivo = st.text_input("Motivo Reverso (Col K obligatorio si Q < 0):", key=f"mot_{bod_norm}_{cap_v}").strip().upper() if cant < 0 else ""

    c6.write("")
    c6.write("")
    ok_rev = (cant >= 0) or bool(motivo)
    puede = st.session_state["per_abierto"] and bool(ubic) and ubic not in ["UBICACIÓN", "ESPECIFIQUE"] and bool(cod_sku) and (cant != 0) and ok_rev
    btn_reg = c6.button("REGISTRAR", use_container_width=True, disabled=not puede)
    st.markdown("</div>", unsafe_allow_html=True)

    st_teo, st_nac = 0.0, 0.0
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
                if row and row[0] and str(row[0]).strip().upper() == cod_sku:
                    st_teo = num_es(row[idx_b]) if idx_b < len(row) else 0.0
                    st_nac = num_es(row[IDX_TOTAL_NAC]) if IDX_TOTAL_NAC < len(row) else 0.0
                    break
        prev = sum(
            num_es(r.get("Q", 0))
            for r in regs_json
            if str(r.get("Periodo", "")).strip() == periodo_act and norm_bod(r.get("Bodega", "")) == bod_norm and str(r.get("COD_SKU", "")).strip().upper() == cod_sku
        )
        c_unit_prev = next(
            (num_es(v) for k, v in art.items() if any(t in k.upper() for t in ["COSTO", "PRECIO", "VAL_UNIT", "V_UNIT", "P_UNIT"]) and num_es(v) > 0),
            0.0,
        )
        q_act_val = float(cant)
        q_acum_calc = float(prev) + q_act_val
        dif, est_vista, _, c_dif_vista = calcular_dif_estado_costo(q_acum_calc, st_teo, q_act_val, c_unit_prev)
        st.info(
            f"**SKU:** {cod_sku} | **Desc:** {flex('DESCRIPCIÓN', 'DESCRIPCION', 'DESC')} | **Und:** {flex('UNIDAD', 'UND')} | **Prov:** {flex('PROVEEDOR', 'SUP', 'FABRICANTE')}"
        )
        m1, m2, m3 = st.columns(3)
        m1.markdown(f'<div class="saldo-box"><span class="saldo-lbl">QSA_Teorico ({bod_norm})</span><span class="saldo-val">{st_teo:g}</span></div>', unsafe_allow_html=True)
        m2.markdown(f'<div class="saldo-box"><span class="saldo-lbl">Nacional</span><span class="saldo-val">{st_nac:g}</span></div>', unsafe_allow_html=True)
        if est_vista == "SOBRANTE":
            m3.markdown(f'<div class="sob">SOBRANTE: {abs(dif):g} un.<br>USD {c_dif_vista:,.2f}</div>', unsafe_allow_html=True)
        elif est_vista == "FALTANTE":
            m3.markdown(f'<div class="fal">FALTANTE: {abs(dif):g} un.<br>USD {c_dif_vista:,.2f}</div>', unsafe_allow_html=True)
        else:
            m3.markdown('<div class="ok">CUADRADO (OK)<br>USD 0.00</div>', unsafe_allow_html=True)

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

        c_unit = next(
            (num_es(v) for k, v in art.items() if any(t in k.upper() for t in ["COSTO", "PRECIO", "VAL_UNIT", "V_UNIT", "P_UNIT"]) and num_es(v) > 0),
            0.0,
        )
        prev = sum(
            num_es(r.get("Q", 0))
            for r in regs_json
            if str(r.get("Periodo", "")).strip() == periodo_act and norm_bod(r.get("Bodega", "")) == bod_norm and str(r.get("COD_SKU", "")).strip().upper() == cod_sku
        )
        dif_f, est_f, costo_total, costo_dif = calcular_dif_estado_costo(prev + float(cant), st_teo, cant, c_unit)
        now = datetime.now()
        nuevo = {
            "Periodo": periodo_act,
            "ID_Tx": f"TX-{now.strftime('%Y%m%d%H%M%S')}",
            "Timestamp": now.strftime("%Y-%m-%d %H:%M:%S"),
            "Correo_Usuario": usuario,
            "Ciudad": ciudad,
            "Bodega": bod_norm,
            "Ubicacion": ubic,
            "COD_SKU": cod_sku,
            "Q": float(cant),
            "BOD/SKU": f"{bod_norm}-{cod_sku}",
            "Motivo": motivo,
            "LINEA DE INVENTARIO": flex_r("LÍNEA DE INVENTARIO", "LINEA", "INVENTARIO"),
            "PROVEEDOR": flex_r("PROVEEDOR", "SUP", "FABRICANTE"),
            "ALTERNO": flex_r("ALTERNO", "ALT"),
            "DESCRIPCIÓN": flex_r("DESCRIPCIÓN", "DESCRIPCION", "DESC"),
            "UNIDAD": flex_r("UNIDAD", "UND"),
            "DIFERENCIA": float(dif_f),
            "ESTADO": est_f,
            "COSTO UNITARIO": float(c_unit),
            "COSTO TOTAL": float(costo_total),
            "COSTO DIF": float(costo_dif),
            "QSA_Teorico": float(st_teo),
        }
        if guardar_json(nuevo):
            ok_g, est_g, tx_g = enviar_gs(nuevo)
            st.session_state["cap_v"] += 1
            st.success(f"Guardado {cod_sku} ({cant:g}) -> {est_f} | GS: {est_g} ({tx_g})" if ok_g else f"Guardado local OK | GS: {est_g}")
            st.rerun()

    st.markdown("---")
    st.markdown("### Últimos Registros")
    prev_df = prep_df(datos_combinados(incluir_gs=False))
    st.markdown(tabla_md(prev_df.iloc[::-1], ["Timestamp", "Bodega", "Ubicacion", "COD_SKU", "Q", "QSA_Teorico", "DIFERENCIA", "ESTADO", "COSTO DIF"], 6))

with tab_dash:
    st.markdown("### Panorama General de Toma Física & Avance Nacional")
    st.caption("Los gráficos y métricas reflejan los filtros activos en el panel izquierdo.")

    if df_rep.empty:
        st.info("No hay registros que coincidan con los filtros seleccionados en el panel izquierdo.")
    else:
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Transacciones (Tx)", f"{len(df_rep):,}")
        k2.metric("SKUs Únicos Contados", f"{df_rep['COD_SKU'].nunique():,}")
        k3.metric("Cantidad Física (Q)", f"{df_rep['Q'].sum():,.2f}")
        k4.metric("Valor Contado (USD)", f"${df_rep['COSTO TOTAL'].sum():,.2f}")
        k5.metric("Impacto Dif. (USD)", f"${df_rep['COSTO DIF'].sum():,.2f}")

        st.markdown("---")
        g1, g2 = st.columns(2)
        with g1:
            st.markdown("**Avance de SKUs por Bodega (Contados vs Pendientes)**")
            df_bod_av = construir_resumen_bodega(df_rep, nac_rows, muestreo=muestreo_rep)
            if not df_bod_av.empty:
                st.bar_chart(df_bod_av.set_index("Bodega")[["SKU_Contados", "SKU_Pendientes"]], use_container_width=True)
            else:
                st.info("Sin datos por bodega.")

        with g2:
            st.markdown("**Actividad por Custodio (Transacciones y Unidades Q)**")
            df_cus_g = (
                df_rep.groupby("Correo_Usuario", as_index=False)
                .agg(Transacciones=("ID_Tx", "count"), Unidades_Q=("Q", "sum"))
                .sort_values("Transacciones", ascending=False)
            )
            st.bar_chart(df_cus_g.set_index("Correo_Usuario")[["Transacciones", "Unidades_Q"]], use_container_width=True)

        g3, g4 = st.columns(2)
        with g3:
            st.markdown("**Distribución por Ciudad (Unidades Q Contadas)**")
            df_ciu_g = (
                df_rep.groupby("Ciudad", as_index=False)
                .agg(Transacciones=("ID_Tx", "count"), Unidades_Q=("Q", "sum"), Costo_Dif_USD=("COSTO DIF", "sum"))
                .sort_values("Unidades_Q", ascending=False)
            )
            st.bar_chart(df_ciu_g.set_index("Ciudad")[["Transacciones", "Unidades_Q"]], use_container_width=True)

        with g4:
            st.markdown("**Desglose por Estado (SOBRANTE / FALTANTE / CUADRADO)**")
            df_est_g = (
                df_rep.groupby("ESTADO", as_index=False)
                .agg(Registros=("ID_Tx", "count"), Costo_Dif_USD=("COSTO DIF", "sum"))
            )
            st.bar_chart(df_est_g.set_index("ESTADO")[["Registros"]], use_container_width=True)

        st.markdown("---")
        st.markdown("**Matriz de Avance por Bodega**")
        df_bod_tabla = construir_resumen_bodega(df_rep, nac_rows, muestreo=muestreo_rep)
        if not df_bod_tabla.empty:
            st.dataframe(df_bod_tabla, use_container_width=True, hide_index=True)

with tab_exp:
    st.markdown("### Detalle Filtrado & Exportación a Excel")
    st.caption("Descarga el detalle transaccional (22 columnas con QSA_Teorico) y la hoja de avance consolidado.")

    if df_rep.empty:
        st.info("No hay registros disponibles con los filtros actuales.")
    else:
        st.download_button(
            f"Descargar Excel Completo ({len(df_rep)} registros)",
            data=excel_bytes(df_rep, nac_rows=nac_rows, muestreo=muestreo_rep),
            file_name=f"TomaFisica_V7gPro_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.dataframe(df_rep.iloc[::-1], use_container_width=True, hide_index=True)
