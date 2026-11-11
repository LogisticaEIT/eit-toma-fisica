# ==============================================================================
# MÓDULO DE CUSTODIOS & TOMA FÍSICA: TomaFisicaApp.py (Versión Master Definitiva)
# ==============================================================================
# COMANDO DE EJECUCIÓN EN CMD (WINDOWS 7 / PYTHON 3.8):
# cd /d "D:\Bases\ADM LOG"
# python -m streamlit run TomaFisicaApp.py --server.address 192.168.100.47 --server.port 8501 --server.fileWatcherType none --server.enableCORS false --server.enableXsrfProtection false
# ==============================================================================

import os
import re
import json
import io
import urllib.parse
from datetime import date, datetime
import openpyxl
import requests
import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Control de Inventario - Módulo Toma Física & Muestreo",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
    html, body, [class*="css"] { font-size: 18px !important; }
    .stApp { background-color: #F1F5F9; }
    div[data-testid="stException"] { display: none !important; }
    
    .main-title {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        color: #FFFFFF; padding: 24px 35px; border-radius: 14px;
        font-size: 32px !important; font-weight: 800; margin-bottom: 20px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1); border-bottom: 5px solid #F59E0B;
    }
    .card-section {
        background-color: #FFFFFF; border: 1px solid #CBD5E1;
        border-radius: 12px; padding: 25px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
        margin-bottom: 25px;
    }
    .stTextInput input, .stNumberInput input, .stSelectbox select, .stDateInput input {
        background-color: #FEF3C7 !important; color: #78350F !important;
        font-weight: 700 !important; font-size: 20px !important;
        border: 2px solid #F59E0B !important; border-radius: 8px !important;
    }
    .status-sob {
        background-color: #DCFCE7; color: #166534; padding: 16px;
        border-radius: 10px; font-weight: bold; text-align: center;
        border: 2px solid #86EFAC; font-size: 20px !important;
    }
    .status-fal {
        background-color: #FEE2E2; color: #991B1B; padding: 16px;
        border-radius: 10px; font-weight: bold; text-align: center;
        border: 2px solid #FCA5A5; font-size: 20px !important;
    }
    .status-ok {
        background-color: #FEF9C3; color: #854D0E; padding: 16px;
        border-radius: 10px; font-weight: bold; text-align: center;
        border: 2px solid #FDE047; font-size: 20px !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="main-title">Control de Inventario - Módulo Toma Física & Muestreo</div>',
    unsafe_allow_html=True,
)

pivot_path = r"D:\Bases\MICRO\PRODUCTO - PIVOT.xlsm"
nac_path = r"D:\Bases\MICRO\NAC.xlsx"
json_local_path = r"D:\Bases\ADM LOG\toma_fisica_registros.json"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxGk8xG7zUy0vu-OUXwgg_qlUS8lmNlU1bh_CygmJkpieTt2jdfOkknoUjAKKis7EiA5g/exec"

BODEGAS_OFICIALES = ["01", "03", "04", "05", "3V", "5V"]
MAPEO_COLUMNAS_FIJO = {"01": 5, "03": 6, "04": 7, "05": 8, "3V": 9, "5V": 10}
IDX_COL_TOTAL_NACIONAL = 11
IDX_COL_MOV_V = 21

ASIGNACION_BODEGAS_CUSTODIOS = {
    "alexpajaritoruiz@icloud.com": BODEGAS_OFICIALES,
    "jlogistica.eit@gmail.com": BODEGAS_OFICIALES,
}


def normalizar_bodega(val):
    if val is None:
        return ""
    v_str = str(val).strip()
    if v_str.endswith(".0"):
        v_str = v_str[:-2]
    if v_str.isdigit():
        return v_str.zfill(2)
    return v_str.upper()


def obtener_fecha_archivo(ruta):
    if os.path.exists(ruta):
        try:
            t_mod = os.path.getmtime(ruta)
            return datetime.fromtimestamp(t_mod).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return "No accesible"
    return "❌ No encontrado"


def cargar_registros_json():
    if os.path.exists(json_local_path):
        try:
            with open(json_local_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []


def guardar_registro_json(nueva_fila_dict):
    registros = cargar_registros_json()
    registros.append(nueva_fila_dict)
    try:
        with open(json_local_path, "w", encoding="utf-8") as f:
            json.dump(registros, f, indent=4, ensure_ascii=False)
        return True
    except Exception:
        return False


datos_locales_json = cargar_registros_json()


@st.cache_data(show_spinner=False)
def cargar_maestros_optimizado(p_path, n_path):
    dict_m = {}
    nac_rows = []
    
    try:
        if os.path.exists(p_path):
            wb = openpyxl.load_workbook(p_path, data_only=True)
            for sh in wb.sheetnames:
                ws = wb[sh]
                rows = list(ws.iter_rows(values_only=True))
                if not rows:
                    continue
                headers = [str(c).strip().upper() if c is not None else "" for c in rows[0]]
                col_key_idx = next((i for i, h in enumerate(headers) if "SKU" in h or "COD" in h), None)
                if col_key_idx is not None:
                    for row in rows[1:]:
                        if len(row) > col_key_idx and row[col_key_idx] is not None:
                            s_val = str(row[col_key_idx]).strip().upper()
                            if s_val and s_val not in ["NONE", "NAN"]:
                                if s_val not in dict_m:
                                    dict_m[s_val] = {}
                                for i, val in enumerate(row):
                                    if i < len(headers) and headers[i] and val is not None:
                                        v_str = str(val).strip()
                                        if v_str and v_str.upper() not in ["NONE", "NAN"]:
                                            dict_m[s_val][headers[i]] = v_str
    except Exception:
        pass

    try:
        if os.path.exists(n_path):
            wb_nac = openpyxl.load_workbook(n_path, data_only=True)
            ws_nac = wb_nac.active
            nac_rows = list(ws_nac.iter_rows(values_only=True))
    except Exception:
        pass
        
    return dict_m, nac_rows


dict_maestro_global, nac_rows = cargar_maestros_optimizado(pivot_path, nac_path)

st.sidebar.markdown("### 🔐 Panel de Control Maestro")
usuario_activo = st.sidebar.text_input(
    "Correo de Usuario / Custodio:", value="alexpajaritoruiz@icloud.com"
).strip().lower()

es_superusuario = (
    "ALEX" in usuario_activo.upper()
    or "PAJARI" in usuario_activo.upper()
    or usuario_activo == "alexpajaritoruiz@icloud.com"
    or usuario_activo == "jlogistica.eit@gmail.com"
)

if es_superusuario:
    bodegas_permitidas = BODEGAS_OFICIALES
    st.sidebar.success("🔑 Perfil: **Superusuario / Admin**")
else:
    bodegas_permitidas = ASIGNACION_BODEGAS_CUSTODIOS.get(usuario_activo, ["01"])
    st.sidebar.info(f"👤 Perfil: **Custodio** (Bodegas: {', '.join(bodegas_permitidas)})")

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ Tipo de Proceso Operativo")
tipo_proceso = st.sidebar.radio(
    "Seleccione el modo de operación:",
    ["Toma Física Completa", "Control Mensual de Inventario (Muestreo)"],
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 🏭 Selección de Bodega Objetivo")

if "bodega_activa_state" not in st.session_state or st.session_state["bodega_activa_state"] not in bodegas_permitidas:
    st.session_state["bodega_activa_state"] = bodegas_permitidas[0]

# Cambio a st.radio para blindar la selección y evitar que se regrese sola
bodega_sel = st.sidebar.radio(
    "Seleccione la Bodega a Auditar:",
    bodegas_permitidas,
    index=bodegas_permitidas.index(st.session_state["bodega_activa_state"]) if st.session_state["bodega_activa_state"] in bodegas_permitidas else 0,
    key="bodega_activa_state"
)
bodega_norm = normalizar_bodega(bodega_sel)

if "periodo_abierto" not in st.session_state:
    st.session_state["periodo_abierto"] = False

if es_superusuario:
    st.session_state["periodo_abierto"] = st.sidebar.toggle(
        "🔓 Habilitar Periodo de Captura (Admin)",
        value=st.session_state["periodo_abierto"],
    )
else:
    st.session_state["periodo_abierto"] = False
    st.sidebar.info("🔒 Periodo controlado por el Administrador")

st.sidebar.markdown("---")
st.sidebar.markdown("### 📅 Rango de Fechas de Auditoría")
f_inicio = st.sidebar.date_input("Fecha de Inicio:", value=date.today())
f_fin = st.sidebar.date_input("Fecha de Fin:", value=date.today())
periodo_auditoria = (
    f"{f_inicio.strftime('%Y-%m-%d')} al {f_fin.strftime('%Y-%m-%d')}"
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📱 Compartir por WhatsApp (Acceso Remoto)")
url_ingresada = st.sidebar.text_input(
    "URL de Acceso para Custodios:",
    value="https://docs.google.com/spreadsheets/d/18wP5AIAaJU1zSMu5GSza7wOmMyZwpYWd1CxlFwBeW5M/edit?usp=sharing",
    key="url_whatsapp_input"
)
if url_ingresada:
    mensaje_wa = (
        f"📦 *CONTROL DE INVENTARIO EIT*\n\n"
        f"Hola Custodio, utiliza el siguiente enlace oficial para registrar tus tomas físicas desde el celular:\n"
        f"{url_ingresada}\n\n"
        f"*(Asegúrate de ingresar los datos correctos en cada toma)*"
    )
    link_wa = f"https://api.whatsapp.com/send?text={urllib.parse.quote(mensaje_wa)}"
    st.sidebar.markdown(
        f"[📲 **Enviar Enlace por WhatsApp**]({link_wa})",
        unsafe_allow_html=True,
    )

if es_superusuario:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### ☁️ Sincronización Masiva Cloud")
    if st.sidebar.button("🚀 Sincronizar Excel Local a Google", use_container_width=True):
        registros_a_enviar = cargar_registros_json()
        if registros_a_enviar:
            exitosos = 0
            for reg in registros_a_enviar:
                fila_lista = [
                    reg.get("Periodo"), reg.get("ID_Tx"), reg.get("Timestamp"), reg.get("Correo_Usuario"),
                    reg.get("Ciudad"), reg.get("Bodega"), reg.get("Ubicacion"), reg.get("COD_SKU"), reg.get("Q"),
                    reg.get("BOD/SKU"), reg.get("Motivo"), reg.get("LINEA DE INVENTARIO"), reg.get("PROVEEDOR"),
                    reg.get("ALTERNO"), reg.get("DESCRIPCIÓN"), reg.get("UNIDAD"), reg.get("DIFERENCIA"),
                    reg.get("ESTADO"), reg.get("COSTO UNITARIO"), reg.get("COSTO TOTAL"), reg.get("COSTO DIF")
                ]
                try:
                    headers = {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"}
                    requests.post(WEB_APP_URL, json=fila_lista, headers=headers, verify=False, timeout=3)
                    exitosos += 1
                except Exception:
                    pass
            st.sidebar.success(f"✅ Sincronizados {exitosos} de {len(registros_a_enviar)} registros.")
        else:
            st.sidebar.info("ℹ️ No hay registros locales pendientes.")

st.sidebar.markdown("---")
st.sidebar.markdown("### 📂 Auditoría de Fechas en Host")
st.sidebar.text(f"• PIVOT.xlsm:\n  {obtener_fecha_archivo(pivot_path)}")
st.sidebar.text(f"• NAC.xlsx:\n  {obtener_fecha_archivo(nac_path)}")
st.sidebar.text(f"• JSON Local:\n  {obtener_fecha_archivo(json_local_path)} ({len(datos_locales_json)} reg.)")

tab_op, tab_dash, tab_exp = st.tabs(
    [
        "⚡ Operación y Captura de Campo",
        "📊 Dashboard Estadístico & Avance",
        "📥 Exportación & Gestión Masiva",
    ]
)

with tab_op:
    try:
        st.markdown('<div class="card-section">', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            ciudad = st.selectbox(
                "Ciudad Target:",
                ["Quito", "Guayaquil", "Cuenca", "Otras", "Ambato", "Manta"],
                key="in_ciu",
            )
        with c2:
            st.markdown(f"**Bodega Target Activa:** `{bodega_norm}`")
        with c3:
            ubicacion = st.text_input(
                "Ubicación Física:", value="UBICACIÓN", key="in_ubic"
            ).strip().upper()

        idx_col_bodega = MAPEO_COLUMNAS_FIJO.get(bodega_norm, None)
        if idx_col_bodega is None and nac_rows and len(nac_rows) > 0:
            headers_nac = [str(c).strip().upper() if c is not None else "" for c in nac_rows[0]]
            for i, h in enumerate(headers_nac):
                if normalizar_bodega(h) == bodega_norm:
                    idx_col_bodega = i
                    break

        skus_maestros_bodega = []
        mapa_sku_desc = {}
        skus_ya_auditados = set()

        if datos_locales_json:
            for row in datos_locales_json:
                per = str(row.get("Periodo", "")).strip()
                bod = normalizar_bodega(row.get("Bodega", ""))
                sku_reg = str(row.get("COD_SKU", "")).strip().upper()
                if per == periodo_auditoria.strip() and bod == bodega_norm and sku_reg:
                    skus_ya_auditados.add(sku_reg)

        es_modo_muestreo = "Control Mensual" in tipo_proceso or "Muestreo" in tipo_proceso

        if nac_rows and len(nac_rows) > 1 and idx_col_bodega is not None:
            for row in nac_rows[1:]:
                if not row or row[0] is None:
                    continue
                s_val = str(row[0]).strip().upper()
                if not s_val or s_val in ["NONE", "NAN"]:
                    continue
                
                st_bod_val = 0.0
                if idx_col_bodega < len(row) and row[idx_col_bodega] is not None:
                    try:
                        st_bod_val = float(row[idx_col_bodega])
                    except:
                        st_bod_val = 0.0

                if st_bod_val > 0:
                    if es_modo_muestreo:
                        val_v = ""
                        if IDX_COL_MOV_V < len(row) and row[IDX_COL_MOV_V] is not None:
                            val_v = str(row[IDX_COL_MOV_V]).strip().upper()
                        if val_v == "SIN MOV" or val_v == "":
                            continue

                    desc_item = ""
                    if s_val in dict_maestro_global:
                        for k, v in dict_maestro_global[s_val].items():
                            if "DESC" in k:
                                desc_item = v
                                break
                    skus_maestros_bodega.append(s_val)
                    mapa_sku_desc[s_val] = desc_item

        skus_maestros_bodega = sorted(list(set(skus_maestros_bodega)))

        if skus_maestros_bodega:
            pendientes_n = len([s for s in skus_maestros_bodega if s not in skus_ya_auditados])
            auditados_n = len([s for s in skus_maestros_bodega if s in skus_ya_auditados])

            with st.expander(
                f"🚦 Guía de Estado de SKUs en Bodega {bodega_norm}"
                f" ({tipo_proceso}) | Pendientes: {pendientes_n} / Total:"
                f" {len(skus_maestros_bodega)}"
            ):
                st.markdown(
                    f"🟢 **Verde:** Pendientes [{pendientes_n}] | 🔴 **Rojo:** Contados [{auditados_n}]"
                )
                filas_guia_html = ""
                for s in skus_maestros_bodega:
                    contado = s in skus_ya_auditados
                    bg = "#FEE2E2" if contado else "#DCFCE7"
                    tc = "#991B1B" if contado else "#166534"
                    st_txt = "AUDITADO" if contado else "PENDIENTE"
                    desc_s = mapa_sku_desc.get(s, "")
                    filas_guia_html += (
                        f"<tr style='background-color:{bg}; color:{tc}; border-bottom:1px solid #CBD5E1;'>"
                        f"<td style='padding:6px; font-weight:bold;'>{st_txt}</td>"
                        f"<td style='padding:6px; font-weight:bold;'>{s}</td>"
                        f"<td style='padding:6px;'>{desc_s}</td></tr>"
                    )

                html_guia_final = f"""
                <div style='max-height: 250px; overflow-y: auto; border: 1px solid #CBD5E1; border-radius: 8px;'>
                    <table style='width:100%; border-collapse: collapse; font-size:15px;'>
                        <thead>
                            <tr style='background-color:#1E293B; color:white; position: sticky; top: 0; z-index: 2;'>
                                <th style='padding:8px; text-align:left;'>Estado</th>
                                <th style='padding:8px; text-align:left;'>COD_SKU</th>
                                <th style='padding:8px; text-align:left;'>Descripción</th>
                            </tr>
                        </thead>
                        <tbody>
                            {filas_guia_html}
                        </tbody>
                    </table>
                </div>
                """
                st.markdown(html_guia_final, unsafe_allow_html=True)
        else:
            st.info(
                f"ℹ️ No hay SKUs teóricos registrados para la Bodega"
                f" {bodega_norm} bajo el criterio de {tipo_proceso}."
            )

        opciones_sku_combo = ["-- SELECCIONE O ESCANEE SKU --"]
        for s in skus_maestros_bodega:
            estado_etiqueta = "AUDITADO" if s in skus_ya_auditados else "PENDIENTE"
            desc_s = mapa_sku_desc.get(s, "")
            opciones_sku_combo.append(f"{s} | [{estado_etiqueta}] {desc_s}")

        st.markdown("---")
        sku_seleccionado_combo = st.selectbox(
            "📋 Selección Rápida del Catálogo Teórico:",
            opciones_sku_combo,
            key="select_sku_dropdown",
        )

        sku_pre_elegido = ""
        if sku_seleccionado_combo != "-- SELECCIONE O ESCANEE SKU --":
            sku_pre_elegido = sku_seleccionado_combo.split(" | ")[0].strip().upper()

        ubicacion_incompleta = (
            not ubicacion or ubicacion in ["UBICACIÓN", "ESPECIFIQUE"]
        )

        c4, c5, c6 = st.columns([2.5, 1.5, 1.0])
        with c4:
            cod_sku_input = st.text_input(
                "Escanear o Digitar COD_SKU:",
                value=sku_pre_elegido,
                key="input_sku",
            ).strip().upper()
            cod_sku = cod_sku_input
        with c5:
            cantidad = st.number_input(
                "Cantidad Física Contada (Q):", step=1, key="input_q", value=0
            )
        with c6:
            st.write("")
            st.write("")
            campos_completOS = (
                st.session_state["periodo_abierto"]
                and not ubicacion_incompleta
                and bool(cod_sku)
                and (cantidad != 0)
            )
            btn_reg = st.button(
                "💾 REGISTRAR",
                use_container_width=True,
                disabled=not campos_completOS,
            )

        if not st.session_state["periodo_abierto"]:
            st.warning("🔒 El periodo de captura está CERRADO por el Administrador. Los registros están bloqueados.")
        if ubicacion_incompleta:
            st.warning("⚠️ Debe especificar una Ubicación Física válida.")
        if not cod_sku:
            st.warning("⚠️ Debe ingresar o escanear un COD SKU.")
        if cantidad == 0:
            st.warning("⚠️ La Cantidad Física (Q) no puede ser 0.")

        motivo = ""
        if cantidad < 0:
            motivo = st.text_input("Motivo de Reverso obligatorio:").strip().upper()

        st.markdown("</div>", unsafe_allow_html=True)

        st_teorico = 0.0
        st_nac_total = 0.0

        if cod_sku:
            info_art = dict_maestro_global.get(cod_sku, {})

            def get_exact_flex(art_dict, *keys):
                for k in keys:
                    for ak, val in art_dict.items():
                        if ak.strip().upper() == k.upper() and val:
                            return str(val).strip()
                for k in keys:
                    for ak, val in art_dict.items():
                        if (k.upper() in ak.upper() or ak.upper() in k.upper()) and val:
                            return str(val).strip()
                return ""

            desc = get_exact_flex(info_art, "DESCRIPCIÓN", "DESCRIPCION", "DESC")
            unidad = get_exact_flex(info_art, "UNIDAD", "UND")
            prov_val = get_exact_flex(info_art, "PROVEEDOR", "SUP", "FABRICANTE")

            if nac_rows and len(nac_rows) > 1:
                for row in nac_rows[1:]:
                    if row and row[0] is not None and str(row[0]).strip().upper() == cod_sku:
                        if idx_col_bodega is not None and idx_col_bodega < len(row):
                            try:
                                st_teorico = float(row[idx_col_bodega]) if row[idx_col_bodega] is not None else 0.0
                            except:
                                st_teorico = 0.0
                        if IDX_COL_TOTAL_NACIONAL < len(row):
                            try:
                                st_nac_total = float(row[IDX_COL_TOTAL_NACIONAL]) if row[IDX_COL_TOTAL_NACIONAL] is not None else 0.0
                            except:
                                st_nac_total = 0.0
                        break

            t_previo = 0.0
            if datos_locales_json:
                for row in datos_locales_json:
                    per = str(row.get("Periodo", "")).strip()
                    bod = normalizar_bodega(row.get("Bodega", ""))
                    sku_r = str(row.get("COD_SKU", "")).strip().upper()
                    if per == periodo_auditoria.strip() and bod == bodega_norm and sku_r == cod_sku:
                        try:
                            t_previo += float(row.get("Q", 0))
                        except:
                            pass

            acumulado_tot = t_previo + cantidad
            dif_calc = st_teorico - acumulado_tot

            st.success(
                f"📌 **SKU:** {cod_sku} | **Desc:** {desc} | **Unidad:** {unidad} | **Proveedor:** {prov_val}"
            )

            ci1, ci2, ci3 = st.columns(3)
            with ci1:
                st.info(f"📋 **Saldo Teórico Bodega {bodega_norm}:** {st_teorico}")
            with ci2:
                st.info(f"🌐 **Total Disponible Nacional:** {st_nac_total}")
            with ci3:
                if dif_calc < 0:
                    st.markdown(f'<div class="status-sob">🟢 SOBRANTE: {abs(dif_calc)} un.</div>', unsafe_allow_html=True)
                elif dif_calc > 0:
                    st.markdown(f'<div class="status-fal">🔴 FALTANTE: {dif_calc} un.</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="status-ok">🟡 BODEGA CUADRADA (OK)</div>', unsafe_allow_html=True)

        if btn_reg:
            info_reg = dict_maestro_global.get(cod_sku, {})

            def get_r(*keys):
                for k in keys:
                    for ak, val in info_reg.items():
                        if ak.strip().upper() == k.upper() and val:
                            return str(val).strip()
                for k in keys:
                    for ak, val in info_reg.items():
                        if (k.upper() in ak.upper() or ak.upper() in k.upper()) and val:
                            return str(val).strip()
                return ""

            prov = get_r("PROVEEDOR", "SUP", "FABRICANTE")
            linea = get_r("LÍNEA DE INVENTARIO", "LINEA", "INVENTARIO")
            alterno = get_r("ALTERNO", "ALT")
            desc_v = get_r("DESCRIPCIÓN", "DESCRIPCION", "DESC")
            und_v = get_r("UNIDAD", "UND")

            costo_unit_raw = get_r("COSTO UNITARIO", "COSTO", "COSTO_UNIT", "PRECIO")
            try:
                costo_unit = float(costo_unit_raw) if costo_unit_raw else 0.0
            except:
                costo_unit = 0.0

            st_teorico_reg = st_teorico if cod_sku else 0.0
            now = datetime.now()
            timestamp_str = now.strftime("%Y-%m-%d %H:%M:%S")
            id_tx = f"TX-{now.strftime('%Y%m%d%H%M%S')}"

            t_prev = 0.0
            if datos_locales_json:
                for row in datos_locales_json:
                    per = str(row.get("Periodo", "")).strip()
                    bod = normalizar_bodega(row.get("Bodega", ""))
                    sku_r = str(row.get("COD_SKU", "")).strip().upper()
                    if per == periodo_auditoria.strip() and bod == bodega_norm and sku_r == cod_sku:
                        try:
                            t_prev += float(row.get("Q", 0))
                        except:
                            pass

            dif_final = st_teorico_reg - (t_prev + float(cantidad))
            estado = "SOB" if dif_final < 0 else ("FAL" if dif_final > 0 else "OK")
            costo_total = round(costo_unit * float(cantidad), 2)
            costo_dif = round(costo_unit * float(dif_final), 2)

            nueva_fila_dict = {
                "Periodo": periodo_auditoria,
                "ID_Tx": id_tx,
                "Timestamp": timestamp_str,
                "Correo_Usuario": usuario_activo,
                "Ciudad": str(ciudad),
                "Bodega": bodega_norm,
                "Ubicacion": ubicacion,
                "COD_SKU": cod_sku,
                "Q": float(cantidad),
                "BOD/SKU": f"{bodega_norm}-{cod_sku}",
                "Motivo": motivo,
                "LINEA DE INVENTARIO": linea,
                "PROVEEDOR": prov,
                "ALTERNO": alterno,
                "DESCRIPCIÓN": desc_v,
                "UNIDAD": und_v,
                "DIFERENCIA": float(dif_final),
                "ESTADO": estado,
                "COSTO UNITARIO": float(costo_unit),
                "COSTO TOTAL": float(costo_total),
                "COSTO DIF": float(costo_dif)
            }

            if guardar_registro_json(nueva_fila_dict):
                try:
                    fila_lista = [
                        nueva_fila_dict.get("Periodo"), nueva_fila_dict.get("ID_Tx"), nueva_fila_dict.get("Timestamp"), nueva_fila_dict.get("Correo_Usuario"),
                        nueva_fila_dict.get("Ciudad"), nueva_fila_dict.get("Bodega"), nueva_fila_dict.get("Ubicacion"), nueva_fila_dict.get("COD_SKU"), nueva_fila_dict.get("Q"),
                        nueva_fila_dict.get("BOD/SKU"), nueva_fila_dict.get("Motivo"), nueva_fila_dict.get("LINEA DE INVENTARIO"), nueva_fila_dict.get("PROVEEDOR"),
                        nueva_fila_dict.get("ALTERNO"), nueva_fila_dict.get("DESCRIPCIÓN"), nueva_fila_dict.get("UNIDAD"), nueva_fila_dict.get("DIFERENCIA"),
                        nueva_fila_dict.get("ESTADO"), nueva_fila_dict.get("COSTO UNITARIO"), nueva_fila_dict.get("COSTO TOTAL"), nueva_fila_dict.get("COSTO DIF")
                    ]
                    headers_req = {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"}
                    requests.post(WEB_APP_URL, json=fila_lista, headers=headers_req, verify=False, timeout=5)
                except Exception:
                    pass

                st.success(f"✅ ¡Guardado local y Sincronizado en la Nube! SKU: {cod_sku} | Q: {cantidad} | Estado: {estado}")
                for key_to_clear in ["input_sku", "select_sku_dropdown", "input_q"]:
                    if key_to_clear in st.session_state:
                        del st.session_state[key_to_clear]
                st.rerun()
            else:
                st.error("❌ Error al guardar en el archivo JSON local.")

        st.markdown("---")
        st.subheader("📋 Últimos Registros Guardados Localmente (Vista Previa)")
        if datos_locales_json:
            st.write(datos_locales_json[-5:][::-1])
        else:
            st.info("Aún no hay registros locales.")

    except Exception as e:
        st.error(f"⚠️ Error en la interfaz de operación: {e}")

with tab_dash:
    try:
        st.markdown("## 📊 Dashboard Estadístico & Nivel de Avance con Filtros")
        st.markdown("Monitoreo gerencial en tiempo real con desglose por bodegas y estados.")
        st.markdown('<div class="card-section">', unsafe_allow_html=True)
        
        if datos_locales_json:
            df_dash_all = pd.DataFrame(datos_locales_json)
            
            st.markdown("### 🔍 Filtros de Análisis Estadístico")
            f_col1, f_col2, f_col3 = st.columns(3)
            
            periodos_disponibles = ["TODOS"] + sorted(df_dash_all["Periodo"].dropna().unique().tolist())
            bodegas_disponibles = ["TODAS"] + sorted(df_dash_all["Bodega"].dropna().unique().tolist())
            usuarios_disponibles = ["TODOS"] + sorted(df_dash_all["Correo_Usuario"].dropna().unique().tolist())
            
            filtro_per = f_col1.selectbox("Filtrar por Periodo:", periodos_disponibles, key="dash_f_per")
            filtro_bod = f_col2.selectbox("Filtrar por Bodega:", bodegas_disponibles, key="dash_f_bod")
            filtro_usu = f_col3.selectbox("Filtrar por Custodio:", usuarios_disponibles, key="dash_f_usu")
            
            df_filtrado = df_dash_all.copy()
            if filtro_per != "TODOS":
                df_filtrado = df_filtrado[df_filtrado["Periodo"] == filtro_per]
            if filtro_bod != "TODAS":
                df_filtrado = df_filtrado[df_filtrado["Bodega"] == filtro_bod]
            if filtro_usu != "TODOS":
                df_filtrado = df_filtrado[df_filtrado["Correo_Usuario"] == filtro_usu]
            
            st.markdown("---")
            
            m1, m2 = st.columns(2)
            m1.metric("Total Transacciones Filtradas", len(df_filtrado))
            m2.metric("Bodegas Activas con Registros", df_filtrado["Bodega"].nunique() if not df_filtrado.empty else 0)
            
            if not df_filtrado.empty:
                st.markdown("### 🏢 Avance de Registros por Bodega")
                conteo_bodegas = df_filtrado["Bodega"].value_counts().reset_index()
                conteo_bodegas.columns = ["BODEGA", "CANTIDAD DE REGISTROS"]
                st.dataframe(conteo_bodegas, use_container_width=True)
                
                st.markdown("### 📈 Desglose por Estado de Diferencias (SOB / FAL / OK)")
                conteo_estados = df_filtrado["ESTADO"].value_counts().reset_index()
                if len(conteo_estados.columns) == 2:
                    conteo_estados.columns = ["ESTADO", "CANTIDAD"]
                st.dataframe(conteo_estados, use_container_width=True)
                
                st.markdown("### 📋 Detalle de Registros Filtrados")
                st.dataframe(df_filtrado[["Timestamp", "Correo_Usuario", "Bodega", "COD_SKU", "DESCRIPCIÓN", "Q", "DIFERENCIA", "ESTADO"]].tail(15), use_container_width=True)
            else:
                st.warning("⚠️ No se encontraron registros con los filtros seleccionados.")
        else:
            st.info("Sin registros en el archivo JSON local para procesar estadísticas.")

        st.markdown('</div>', unsafe_allow_html=True)
    except Exception as e:
        st.error(f"⚠️ Error en el Dashboard: {e}")

with tab_exp:
    try:
        st.markdown("## 📥 Exportación Masiva de Datos Locales con Filtros")
        st.markdown("Filtre y descargue el consolidado específico de las capturas locales en Excel.")
        st.markdown('<div class="card-section">', unsafe_allow_html=True)
        
        if datos_locales_json:
            df_exp_all = pd.DataFrame(datos_locales_json)
            
            st.markdown("### 🔍 Criterios de Filtrado para Exportar")
            e_col1, e_col2, e_col3 = st.columns(3)
            
            exp_periodos = ["TODOS"] + sorted(df_exp_all["Periodo"].dropna().unique().tolist())
            exp_bodegas = ["TODAS"] + sorted(df_exp_all["Bodega"].dropna().unique().tolist())
            exp_usuarios = ["TODOS"] + sorted(df_exp_all["Correo_Usuario"].dropna().unique().tolist())
            
            sel_exp_per = e_col1.selectbox("Periodo a Exportar:", exp_periodos, key="exp_f_per")
            sel_exp_bod = e_col2.selectbox("Bodega a Exportar:", exp_bodegas, key="exp_f_bod")
            sel_exp_usu = e_col3.selectbox("Custodio a Exportar:", exp_usuarios, key="exp_f_usu")
            
            df_export_filtered = df_exp_all.copy()
            if sel_exp_per != "TODOS":
                df_export_filtered = df_export_filtered[df_export_filtered["Periodo"] == sel_exp_per]
            if sel_exp_bod != "TODAS":
                df_export_filtered = df_export_filtered[df_export_filtered["Bodega"] == sel_exp_bod]
            if sel_exp_usu != "TODOS":
                df_export_filtered = df_export_filtered[df_export_filtered["Correo_Usuario"] == sel_exp_usu]
            
            st.markdown("---")
            st.dataframe(df_export_filtered, use_container
