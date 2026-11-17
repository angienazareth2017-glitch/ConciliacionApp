# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import io
import json
import urllib.request
import urllib.error
from datetime import datetime

st.set_page_config(
    page_title="TESORIQ | Inteligencia y Control Financiero",
    page_icon="☘️",
    layout="wide",
    initial_sidebar_state="expanded"
)

LOGO_TESORIQ_DARK_SVG = (
    '<svg width="225" height="56" viewBox="0 0 265 64" '
    'fill="none" xmlns="http://www.w3.org/2000/svg">'
    '<path d="M30 29C30 29 11 31 7 18.5C3 6 18 3.5 25.5 10C32.5 16 30 29 30 29Z" '
    'fill="#238458"/>'
    '<path d="M34 29C34 29 53 31 57 18.5C61 6 46 3.5 38.5 10C31.5 16 34 29 34 29Z" '
    'fill="#8CE0B5"/>'
    '<path d="M32 31C22.5 36 17.5 45.5 25 53.5C28 56.5 30 57.5 30 62H34'
    'C34 57.5 36 56.5 39 53.5C46.5 45.5 41.5 36 32 31Z" '
    'fill="#238458" stroke="#8CE0B5" stroke-width="1.2"/>'
    '<text x="72" y="44" font-family="Manrope, sans-serif" '
    'font-weight="800" font-size="31" fill="#F4F6F4" '
    'letter-spacing="1.5">TESORIQ</text></svg>'
)

def aplicar_identidad_tesoriq():
    css_rules = [
        "<style>",
        "@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Manrope:wght@600;700;800&display=swap');",
        ":root {",
        "  --tesoriq-bosque: #10251D;",
        "  --tesoriq-tecnologico: #238458;",
        "  --tesoriq-menta: #8CE0B5;",
        "  --tesoriq-carbon: #171A18;",
        "  --tesoriq-blanco-suave: #F4F6F4;",
        "  --tesoriq-blanco-puro: #FFFFFF;",
        "}",
        "html, body, [class*='css'], .stDataFrame, input, select {",
        "  font-family: 'Inter', sans-serif !important;",
        "  color: var(--tesoriq-carbon);",
        "}",
        "h1, h2, h3, h4, h5, h6 {",
        "  font-family: 'Manrope', sans-serif !important;",
        "  color: var(--tesoriq-bosque) !important;",
        "  font-weight: 700 !important;",
        "}",
        ".stApp { background-color: var(--tesoriq-blanco-suave) !important; }",
        "section[data-testid='stSidebar'] {",
        "  background: linear-gradient(180deg, #10251D 0%, #171A18 100%) !important;",
        "  border-right: 1px solid rgba(140, 224, 181, 0.18);",
        "}",
        "section[data-testid='stSidebar'] p, section[data-testid='stSidebar'] span, section[data-testid='stSidebar'] label, section[data-testid='stSidebar'] div {",
        "  color: var(--tesoriq-blanco-suave) !important;",
        "}",
        "section[data-testid='stSidebar'] input {",
        "  color: var(--tesoriq-carbon) !important;",
        "  background-color: var(--tesoriq-blanco-suave) !important;",
        "}",
        ".sidebar-section {",
        "  font-family: 'Manrope', sans-serif;",
        "  font-size: 0.72rem;",
        "  text-transform: uppercase;",
        "  letter-spacing: 1.4px;",
        "  color: #8CE0B5 !important;",
        "  margin-top: 1.2rem;",
        "  margin-bottom: 0.4rem;",
        "  font-weight: 700;",
        "}",
        ".stButton > button, .stDownloadButton > button {",
        "  background-color: var(--tesoriq-tecnologico) !important;",
        "  color: var(--tesoriq-blanco-suave) !important;",
        "  font-family: 'Manrope', sans-serif !important;",
        "  font-weight: 600 !important;",
        "  border-radius: 8px !important;",
        "}",
        ".stButton > button:hover, .stDownloadButton > button:hover {",
        "  background-color: var(--tesoriq-bosque) !important;",
        "}",
        ".tesoriq-kpi-card {",
        "  background-color: var(--tesoriq-blanco-puro);",
        "  border: 1px solid #E1E7E4;",
        "  border-left: 5px solid var(--tesoriq-tecnologico);",
        "  border-radius: 12px;",
        "  padding: 18px 20px;",
        "  margin-bottom: 12px;",
        "}",
        ".tesoriq-kpi-title { font-size: 0.84rem; color: #55625C; font-weight: 500; }",
        ".tesoriq-kpi-value { font-family: 'Manrope', sans-serif; font-size: 1.8rem; font-weight: 800; color: var(--tesoriq-bosque); margin: 0; }",
        ".tesoriq-kpi-tag { display: inline-block; margin-top: 8px; padding: 2px 10px; font-size: 0.74rem; font-weight: 600; border-radius: 99px; background-color: rgba(140, 224, 181, 0.28); color: var(--tesoriq-bosque); }",
        ".tesoriq-hero-banner { background: linear-gradient(135deg, #10251D 0%, #238458 100%); color: #F4F6F4; padding: 20px 26px; border-radius: 14px; margin-bottom: 22px; display: flex; justify-content: space-between; align-items: center; }",
        ".tesoriq-hero-banner h2 { color: #F4F6F4 !important; margin: 0 0 4px 0 !important; font-size: 1.35rem !important; }",
        ".tesoriq-hero-banner p { color: #8CE0B5 !important; margin: 0 !important; font-size: 0.92rem !important; }",
        ".tesoriq-ia-panel { background-color: #FFFFFF; border: 1px solid #8CE0B5; border-left: 6px solid #10251D; border-radius: 12px; padding: 20px 24px; margin-top: 16px; }",
        "</style>"
    ]
    st.markdown(chr(10).join(css_rules), unsafe_allow_html=True)

def renderizar_logo_sidebar():
    html_logo = "<div style='padding: 6px 0 14px 0;'>" + LOGO_TESORIQ_DARK_SVG + "</div>"
    st.markdown(html_logo, unsafe_allow_html=True)

def renderizar_banner_tesoriq():
    html_banner = (
        "<div class='tesoriq-hero-banner'>"
        "<div><h2>Panel de Control e Inteligencia de Tesorería</h2>"
        "<p>Deja de buscar diferencias. Empieza a resolverlas con trazabilidad total.</p></div>"
        "<div style='text-align:right; font-weight:700; color:#8CE0B5;'>"
        "TESORIQ • REGLAS C1-C4 ACTIVAS</div></div>"
    )
    st.markdown(html_banner, unsafe_allow_html=True)

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "tasa_usd" not in st.session_state:
    st.session_state.tasa_usd = 45.50
if "tasa_eur" not in st.session_state:
    st.session_state.tasa_eur = 48.20
if "lista_conciliaciones" not in st.session_state:
    st.session_state.lista_conciliaciones = [
        {"id": "CONC-001", "fecha": "2026-11-14", "cuenta": "Banco Mercantil - USD/VES", "estado": "Abierta (Retomable)", "movimientos": 5},
        {"id": "CONC-002", "fecha": "2026-11-10", "cuenta": "Banesco - Corriente", "estado": "Cerrada", "movimientos": 12}
    ]
if "overrides_c1_c4" not in st.session_state:
    st.session_state.overrides_c1_c4 = {}

if not st.session_state.autenticado:
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        st.write("")
        st.write("")
        st.markdown("<div style='text-align:center; padding-top: 2rem;'>" + LOGO_TESORIQ_DARK_SVG + "</div>", unsafe_allow_html=True)
        st.markdown(
            "<p style='text-align:center; color:#5A6560; font-size:0.95rem; margin-bottom: 2rem;'>Plataforma de Inteligencia y Conciliación Financiera</p>",
            unsafe_allow_html=True
        )
        with st.container(border=True):
            st.markdown("#### Acceso Corporativo Segregado")
            usuario_input = st.text_input("Usuario o Correo", placeholder="analista.tesoreria@empresa.com")
            clave_input = st.text_input("Contraseña", type="password", placeholder="••••••••")
            st.write("")
            if st.button("Ingresar a TESORIQ", use_container_width=True):
                if clave_input.strip() != "":
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error("Por favor ingrese su contraseña para continuar.")
    st.stop()

def evaluar_reglas_c1_c4(row, tolerancia=0.05):
    merge = row.get("_merge", "both")
    dif = row.get("Diferencia_Absoluta_VES", 0.0)
    if merge == "both":
        if dif <= tolerancia:
            return "C1: Conciliado Exacto", "Aprobado Automático"
        elif dif <= 15.00:
            return "C2: Diferencia por Redondeo / Céntimos", "Advertencia Leve (Aceptable)"
        else:
            return "C3: Posible Retención SENIAT o Comisión Bancaria", "Advertencia Moderada (Requiere Validación)"
    elif merge == "left_only":
        return "C4: Pendiente en ERP (Solo en Banco)", "Alerta: Partida No Registrada en Libros"
    else:
        return "C4: Pendiente en Banco (Solo en ERP)", "Alerta: Partida en Tránsito o No Efectiva"

def ejecutar_conciliacion(df_b, df_e, tolerancia=0.05):
    b = df_b.copy()
    e = df_e.copy()
    b["Referencia_Clean"] = b["Referencia"].astype(str).str.strip()
    e["Referencia_Clean"] = e["Referencia"].astype(str).str.strip()
    cruce = pd.merge(b, e, on="Referencia_Clean", how="outer", suffixes=("_Banco", "_ERP"), indicator=True)
    cruce["Monto_Banco"] = pd.to_numeric(cruce["Monto_Banco"], errors="coerce").fillna(0.0)
    cruce["Monto_ERP"] = pd.to_numeric(cruce["Monto_ERP"], errors="coerce").fillna(0.0)
    cruce["Diferencia_Absoluta_VES"] = (cruce["Monto_Banco"] - cruce["Monto_ERP"]).abs().round(2)
    resultados_reglas = []
    mensajes_alerta = []
    for _, r in cruce.iterrows():
        regla, alerta = evaluar_reglas_c1_c4(r, tolerancia)
        resultados_reglas.append(regla)
        mensajes_alerta.append(alerta)
    cruce["Regla_Aplicada"] = resultados_reglas
    cruce["Estatus_Alerta"] = mensajes_alerta
    estados_finales = []
    for idx, r in cruce.iterrows():
        ref_key = str(r.get("Referencia_Clean", idx))
        if ref_key in st.session_state.overrides_c1_c4:
            estados_finales.append(st.session_state.overrides_c1_c4[ref_key])
        else:
            regla_txt = r["Regla_Aplicada"]
            if "C1" in regla_txt or "C2" in regla_txt:
                estados_finales.append("Conciliado")
            else:
                estados_finales.append("Pendiente por Revisar")
    cruce["Estado_Conciliacion"] = estados_finales
    return cruce

def generar_datos_demo():
    df_b = pd.DataFrame({
        "Fecha_Banco": ["2026-11-10", "2026-11-11", "2026-11-12", "2026-11-13", "2026-11-14"],
        "Referencia": ["REF-1001", "REF-1002", "REF-1003", "REF-1004", "REF-1005"],
        "Concepto_Banco": ["Pago Proveedor Vet", "Cobro Distribuidora", "Comisión + IGTF 3%", "Pago Servicios", "Abono Clientes"],
        "Monto_Banco": [1500.00, 3200.50, 45.00, 890.00, 2100.00]
    })
    df_e = pd.DataFrame({
        "Fecha_ERP": ["2026-11-10", "2026-11-11", "2026-11-13", "2026-11-14", "2026-11-15"],
        "Referencia": ["REF-1001", "REF-1002", "REF-1004", "REF-1005", "REF-1006"],
        "Concepto_ERP": ["Factura Quimio-Vet", "Cobro Distribuidora", "Servicio Eléctrico", "Cobro Lote B", "Pago en Tránsito"],
        "Monto_ERP": [1500.00, 3200.00, 890.00, 2100.00, 740.00]
    })
    return df_b, df_e

def tarjeta_kpi_html(titulo, valor, etiqueta):
    return (
        "<div class='tesoriq-kpi-card'>"
        "<div class='tesoriq-kpi-title'>" + str(titulo) + "</div>"
        "<p class='tesoriq-kpi-value'>" + str(valor) + "</p>"
        "<span class='tesoriq-kpi-tag'>" + str(etiqueta) + "</span>"
        "</div>"
    )

aplicar_identidad_tesoriq()

if "df_banco" not in st.session_state:
    st.session_state.df_banco, st.session_state.df_erp = generar_datos_demo()
if "df_resultado" not in st.session_state:
    st.session_state.df_resultado = ejecutar_conciliacion(st.session_state.df_banco, st.session_state.df_erp)

with st.sidebar:
    renderizar_logo_sidebar()
    st.markdown("<div class='sidebar-section'>Navegación Principal</div>", unsafe_allow_html=True)
    menu_activo = st.radio(
        "Ir a:",
        options=[
            "📊 Dashboard Ejecutivo",
            "📥 Importar Extracto y ERP",
            "🔄 Conciliación y Cruce",
            "⚠️ Reglas C1-C4 y Excepciones",
            "🏛️ Tasas Multimoneda (USD / EUR)"
        ],
        label_visibility="collapsed"
    )
    st.markdown("<div class='sidebar-section'>Control de Sesión</div>", unsafe_allow_html=True)
    if st.button("🔒 Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

renderizar_banner_tesoriq()

df_res = st.session_state.df_resultado
conciliados = len(df_res[df_res["Estado_Conciliacion"] == "Conciliado"]) if df_res is not None else 0
pendientes = len(df_res[df_res["Estado_Conciliacion"] == "Pendiente por Revisar"]) if df_res is not None else 0

if menu_activo == "📊 Dashboard Ejecutivo":
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(tarjeta_kpi_html("Partidas Pendientes", pendientes, "Requieren decisión"), unsafe_allow_html=True)
    with c2:
        st.markdown(tarjeta_kpi_html("Movimientos Conciliados", conciliados, "Aprobados"), unsafe_allow_html=True)
    with c3:
        st.markdown(tarjeta_kpi_html("Total Movimientos", len(df_res) if df_res is not None else 0, "Cruce activo"), unsafe_allow_html=True)
    with c4:
        st.markdown(tarjeta_kpi_html("Tasa BCV USD", "{:.2f}".format(st.session_state.tasa_usd), "VES / USD"), unsafe_allow_html=True)

    st.write("")
    st.markdown("#### 📁 Conciliaciones Recientes (Retomar o Cerrar)")
    st.caption("Selecciona una conciliación del historial para gestionar su estado operativo.")
    df_historial = pd.DataFrame(st.session_state.lista_conciliaciones)
    for i, row in df_historial.iterrows():
        cols = st.columns([1.5, 2, 2, 1.2, 1.2])
        cols[0].write(f"**{row['id']}**")
        cols[1].write(row['fecha'])
        cols[2].write(row['cuenta'])
        if "Abierta" in row['estado']:
            cols[3].markdown(f"<span style='color:#238458; font-weight:650;'>{row['estado']}</span>", unsafe_allow_html=True)
            if cols[4].button("Cerrar", key=f"cerrar_{i}"):
                st.session_state.lista_conciliaciones[i]['estado'] = "Cerrada"
                st.rerun()
        else:
            cols[3].markdown(f"<span style='color:#777;'>{row['estado']}</span>", unsafe_allow_html=True)
            if cols[4].button("Retomar", key=f"retomar_{i}"):
                st.session_state.lista_conciliaciones[i]['estado'] = "Abierta (Retomable)"
                st.rerun()

    st.write("")
    st.markdown("#### Resumen del Cruce Actual")
    st.dataframe(df_res, use_container_width=True, hide_index=True)

elif menu_activo == "📥 Importar Extracto y ERP":
    col_b, col_e = st.columns(2)
    with col_b:
        up_b = st.file_uploader("1. Subir Extracto Bancario (CSV / Excel)", type=["csv", "xlsx"])
        if up_b is not None:
            st.session_state.df_banco = pd.read_csv(up_b) if up_b.name.endswith(".csv") else pd.read_excel(up_b)
            st.success("Extracto bancario actualizado.")
    with col_e:
        up_e = st.file_uploader("2. Subir Mayor Auxiliar ERP (CSV / Excel)", type=["csv", "xlsx"])
        if up_e is not None:
            st.session_state.df_erp = pd.read_csv(up_e) if up_e.name.endswith(".csv") else pd.read_excel(up_e)
            st.success("Mayor ERP actualizado.")

elif menu_activo == "🔄 Conciliación y Cruce":
    tol = st.slider("Tolerancia por redondeo (VES)", 0.00, 5.00, 0.05, 0.01)
    if st.button("🚀 Ejecutar Cruce Inteligente"):
        st.session_state.df_resultado = ejecutar_conciliacion(st.session_state.df_banco, st.session_state.df_erp, tolerancia=tol)
        st.success("Cruce ejecutado aplicando las reglas de negocio C1-C4.")
    st.dataframe(st.session_state.df_resultado, use_container_width=True, hide_index=True)

elif menu_activo == "⚠️ Reglas C1-C4 y Excepciones":
    st.markdown("### ⚠️ Auditoría de Reglas y Decisión Humana (C1 - C4)")
    st.caption("Inspeccione las advertencias generadas por las reglas del sistema y decida manualmente si desea conciliar la partida (Override) o mantenerla pendiente.")
    if df_res is not None:
        for idx, row in df_res.iterrows():
            ref = str(row.get("Referencia_Clean", "Item-" + str(idx)))
            regla = str(row.get("Regla_Aplicada", "C1"))
            alerta = str(row.get("Estatus_Alerta", "Aprobado"))
            estado_actual = str(row.get("Estado_Conciliacion", "Pendiente"))
            with st.container(border=True):
                c_info1, c_info2, c_btn = st.columns([2.5, 2, 1.5])
                mb = row.get("Monto_Banco", 0)
                me = row.get("Monto_ERP", 0)
                c_info1.markdown("**Ref:** `" + ref + "` | **Banco:** " + "{:,.2f}".format(mb) + " | **ERP:** " + "{:,.2f}".format(me))
                c_info2.markdown("**Regla:** `" + regla + "`\n\n**Estatus:** *" + alerta + "*")
                is_conciliado = (estado_actual == "Conciliado")
                if is_conciliado:
                    c_btn.success("✅ Conciliado")
                    if c_btn.button("Rechazar / Reabrir", key="reabrir_" + str(idx)):
                        st.session_state.overrides_c1_c4[ref] = "Pendiente por Revisar"
                        st.session_state.df_resultado = ejecutar_conciliacion(st.session_state.df_banco, st.session_state.df_erp)
                        st.rerun()
                else:
                    c_btn.warning("⚠️ Pendiente")
                    if c_btn.button("Forzar Conciliación", key="forzar_" + str(idx)):
                        st.session_state.overrides_c1_c4[ref] = "Conciliado"
                        st.session_state.df_resultado = ejecutar_conciliacion(st.session_state.df_banco, st.session_state.df_erp)
                        st.rerun()
    else:
        st.info("Ejecute primero un cruce en el módulo de Conciliación.")

elif menu_activo == "🏛️ Tasas Multimoneda (USD / EUR)":
    st.markdown("### Configuración de Tasas de Cambio Oficiales (BCV)")
    with st.container(border=True):
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            st.session_state.tasa_usd = st.number_input("Tasa Oficial Dólar (USD / VES)", min_value=1.0, value=float(st.session_state.tasa_usd), step=0.10, format="%.2f")
        with col_t2:
            st.session_state.tasa_eur = st.number_input("Tasa Oficial Euro (EUR / VES)", min_value=1.0, value=float(st.session_state.tasa_eur), step=0.10, format="%.2f")
        st.success("Tasas actualizadas. USD: " + "{:.2f}".format(st.session_state.tasa_usd) + " | EUR: " + "{:.2f}".format(st.session_state.tasa_eur))
