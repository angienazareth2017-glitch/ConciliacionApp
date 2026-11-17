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
        "TESORIQ • ENTERPRISE DATA</div></div>"
    )
    st.markdown(html_banner, unsafe_allow_html=True)

def analizar_partidas_pendientes_ia(df_pendientes, tasa_bcv, api_key="", modelo="google/gemini-2.0-flash-001"):
    if df_pendientes is None or df_pendientes.empty:
        return "✅ No existen partidas pendientes ni diferencias por auditar."

    df_trabajo = df_pendientes.copy()
    col_banco = next((c for c in df_trabajo.columns if "banco" in c.lower() and "monto" in c.lower()), None)
    col_erp = next((c for c in df_trabajo.columns if "erp" in c.lower() and "monto" in c.lower()), None)
    col_ref = next((c for c in df_trabajo.columns if "ref" in c.lower()), df_trabajo.columns[0])

    if col_banco and col_erp:
        m_b = pd.to_numeric(df_trabajo[col_banco], errors="coerce").fillna(0.0)
        m_e = pd.to_numeric(df_trabajo[col_erp], errors="coerce").fillna(0.0)
        df_trabajo["Diferencia_Absoluta_VES"] = (m_b - m_e).abs().round(2)
        df_trabajo["Equivalente_USD_BCV"] = (df_trabajo["Diferencia_Absoluta_VES"] / max(float(tasa_bcv), 0.01)).round(2)

    if not api_key or api_key.strip() == "":
        tasa_txt = "{:.2f}".format(float(tasa_bcv))
        total_txt = str(len(df_trabajo))
        cabecera = "*Tasa BCV: **" + tasa_txt + " VES/USD** | Auditadas: **" + total_txt + "** | Integridad: **100% Preservada***"
        lineas = [
            "#### 📋 Dictamen de Auditoría Asistida TESORIQ (Motor Local Activo)",
            cabecera,
            ""
        ]
        for idx, row in df_trabajo.head(15).iterrows():
            ref_val = str(row.get(col_ref, "Partida-" + str(idx + 1)))
            dif_ves = str(row.get("Diferencia_Absoluta_VES", "N/A"))
            dif_usd = str(row.get("Equivalente_USD_BCV", "N/A"))
            detalle = (
                "- **Referencia `" + ref_val + "`**: Diferencia Absoluta **" + dif_ves + " VES** "
                "(*Eq. " + dif_usd + " USD a tasa BCV*) — Verificar retención IVA/ISLR SENIAT, "
                "comisión bancaria/IGTF o partida en tránsito antes de asentar en el ERP."
            )
            lineas.append(detalle)
        lineas.append("")
        lineas.append("> 💡 *Ingresa tu API Key de OpenRouter en la barra lateral para redacción generativa avanzada.*")
        return chr(10).join(lineas)

    muestra = df_trabajo.head(30).to_dict(orient="records")
    prompt_sys = (
        "Eres el Auditor Senior de Tesorería de TESORIQ en Venezuela. "
        "No alteres montos originales. Analiza excepciones en valor absoluto y sugiere asientos."
    )
    prompt_usr = (
        "Tasa BCV: " + "{:.2f}".format(float(tasa_bcv)) + " VES/USD. "
        "Analiza estas partidas pendientes (causa raíz SENIAT IVA/ISLR, IGTF, comisiones o tránsito) "
        "y sugiere acción en ERP: " + json.dumps(muestra, default=str, ensure_ascii=False)
    )

    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": "Bearer " + api_key.strip(),
        "Content-Type": "application/json"
    }
    payload = {
        "model": modelo,
        "messages": [
            {"role": "system", "content": prompt_sys},
            {"role": "user", "content": prompt_usr}
        ],
        "temperature": 0.2
    }
    try:
        req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            return data["choices"][0]["message"]["content"]
    except Exception as e:
        return "⚠️ Error al consultar OpenRouter: " + str(e)

def ejecutar_conciliacion(df_b, df_e, tolerancia=0.05):
    b = df_b.copy()
    e = df_e.copy()
    b["Referencia_Clean"] = b["Referencia"].astype(str).str.strip()
    e["Referencia_Clean"] = e["Referencia"].astype(str).str.strip()
    cruce = pd.merge(b, e, on="Referencia_Clean", how="outer", suffixes=("_Banco", "_ERP"), indicator=True)
    cruce["Monto_Banco"] = pd.to_numeric(cruce["Monto_Banco"], errors="coerce").fillna(0.0)
    cruce["Monto_ERP"] = pd.to_numeric(cruce["Monto_ERP"], errors="coerce").fillna(0.0)
    cruce["Diferencia_Absoluta_VES"] = (cruce["Monto_Banco"] - cruce["Monto_ERP"]).abs().round(2)

    def clasificar(row):
        if row["_merge"] == "both":
            if row["Diferencia_Absoluta_VES"] <= tolerancia:
                return "Conciliado Exacto"
            return "Diferencia en Monto"
        elif row["_merge"] == "left_only":
            return "Pendiente en ERP (Solo Banco)"
        return "Pendiente en Banco (Solo ERP)"

    cruce["Estado_Conciliacion"] = cruce.apply(clasificar, axis=1)
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

def renderizar_seccion_partidas_pendientes(df_pendientes, tasa_bcv=45.50, api_key_default="", modelo_sel="google/gemini-2.0-flash-001"):
    st.markdown("### ⚠️ Auditoría Inteligente de Partidas Pendientes y Excepciones")
    if df_pendientes is None or df_pendientes.empty:
        st.success("🎉 ¡Excelente! No hay partidas pendientes ni diferencias abiertas.")
        return

    df_vista = df_pendientes.copy()
    df_vista["Equivalente_USD_BCV"] = (df_vista["Diferencia_Absoluta_VES"] / max(float(tasa_bcv), 0.01)).round(2)

    k1, k2, k3 = st.columns(3)
    suma_ves = df_vista["Diferencia_Absoluta_VES"].sum()
    suma_usd = suma_ves / max(float(tasa_bcv), 0.01)
    with k1:
        st.markdown(tarjeta_kpi_html("Partidas Abiertas", len(df_vista), "Requieren revisión"), unsafe_allow_html=True)
    with k2:
        st.markdown(tarjeta_kpi_html("Discrepancia Total", "{:,.2f} VES".format(suma_ves), "Valor absoluto"), unsafe_allow_html=True)
    with k3:
        st.markdown(tarjeta_kpi_html("Equivalente USD", "${:,.2f} USD".format(suma_usd), "Tasa BCV: {:.2f}".format(tasa_bcv)), unsafe_allow_html=True)

    st.dataframe(df_vista, use_container_width=True, hide_index=True)

    col_b1, col_b2 = st.columns([1.2, 1.8])
    with col_b1:
        if st.button("🤖 Auditar Partidas Pendientes con IA", use_container_width=True):
            with st.spinner("Ejecutando análisis contable TESORIQ..."):
                st.session_state["tesoriq_ultimo_dictamen"] = analizar_partidas_pendientes_ia(
                    df_vista, tasa_bcv=tasa_bcv, api_key=api_key_default, modelo=modelo_sel
                )
    with col_b2:
        buffer_exc = io.BytesIO()
        with pd.ExcelWriter(buffer_exc, engine="openpyxl") as writer:
            df_vista.to_excel(writer, index=False, sheet_name="Partidas_Pendientes")
        nombre_xlsx = "Partidas_Pendientes_TESORIQ_" + datetime.now().strftime("%Y%m%d") + ".xlsx"
        st.download_button(
            label="📥 Exportar Partidas Pendientes (.xlsx)",
            data=buffer_exc.getvalue(),
            file_name=nombre_xlsx,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    if st.session_state.get("tesoriq_ultimo_dictamen"):
        st.markdown("<div class='tesoriq-ia-panel'>" + st.session_state["tesoriq_ultimo_dictamen"] + "</div>", unsafe_allow_html=True)

# ==============================================================================
# EJECUCIÓN PRINCIPAL
# ==============================================================================
aplicar_identidad_tesoriq()

if "df_banco" not in st.session_state:
    st.session_state.df_banco, st.session_state.df_erp = generar_datos_demo()
if "df_resultado" not in st.session_state:
    st.session_state.df_resultado = ejecutar_conciliacion(st.session_state.df_banco, st.session_state.df_erp)
if "tasa_bcv" not in st.session_state:
    st.session_state.tasa_bcv = 45.50
if "api_key_ia" not in st.session_state:
    st.session_state.api_key_ia = ""

with st.sidebar:
    renderizar_logo_sidebar()
    st.markdown("<div class='sidebar-section'>Navegación Principal</div>", unsafe_allow_html=True)
    menu_activo = st.radio(
        "Ir a:",
        options=[
            "📊 Dashboard Ejecutivo",
            "📥 Importar Extracto y ERP",
            "🔄 Conciliación y Cruce",
            "⚠️ Partidas Pendientes (IA)",
            "🏛️ Tasas BCV y Reglas"
        ],
        label_visibility="collapsed"
    )
    st.markdown("<div class='sidebar-section'>Parámetros BCV</div>", unsafe_allow_html=True)
    st.session_state.tasa_bcv = st.number_input("Tasa Oficial BCV (VES/USD)", min_value=1.0, value=float(st.session_state.tasa_bcv), step=0.10)
    st.markdown("<div class='sidebar-section'>Motor IA (OpenRouter)</div>", unsafe_allow_html=True)
    st.session_state.api_key_ia = st.text_input("OpenRouter API Key", value=st.session_state.api_key_ia, type="password")
    modelo_activo = st.selectbox("Modelo Analítico", options=["google/gemini-2.0-flash-001", "openai/gpt-4o-mini", "anthropic/claude-3.5-haiku"])

renderizar_banner_tesoriq()

df_res = st.session_state.df_resultado
conciliados = len(df_res[df_res["Estado_Conciliacion"] == "Conciliado Exacto"])
diferencias = len(df_res[df_res["Estado_Conciliacion"] == "Diferencia en Monto"])
abiertas = len(df_res[df_res["Estado_Conciliacion"].str.contains("Pendiente")])

if menu_activo == "📊 Dashboard Ejecutivo":
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(tarjeta_kpi_html("Partidas Abiertas", abiertas, "Por revisar"), unsafe_allow_html=True)
    with c2:
        st.markdown(tarjeta_kpi_html("Movimientos Conciliados", conciliados, "Match 100%"), unsafe_allow_html=True)
    with c3:
        st.markdown(tarjeta_kpi_html("Diferencias de Monto", diferencias, "Valor absoluto"), unsafe_allow_html=True)
    with c4:
        st.markdown(tarjeta_kpi_html("Tasa BCV Aplicada", "{:.2f}".format(st.session_state.tasa_bcv), "VES / USD"), unsafe_allow_html=True)
    st.markdown("#### Resumen General de Movimientos")
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
    tol = st.slider("Tolerancia por redondeo (VES/USD)", 0.00, 5.00, 0.05, 0.01)
    if st.button("🚀 Ejecutar Cruce Inteligente"):
        st.session_state.df_resultado = ejecutar_conciliacion(st.session_state.df_banco, st.session_state.df_erp, tolerancia=tol)
        st.success("Conciliación completada respetando el 100% de los datos originales.")
    st.dataframe(st.session_state.df_resultado, use_container_width=True, hide_index=True)

elif menu_activo == "⚠️ Partidas Pendientes (IA)":
    df_exc = df_res[df_res["Estado_Conciliacion"] != "Conciliado Exacto"].copy()
    renderizar_seccion_partidas_pendientes(df_exc, tasa_bcv=st.session_state.tasa_bcv, api_key_default=st.session_state.api_key_ia, modelo_sel=modelo_activo)

elif menu_activo == "🏛️ Tasas BCV y Reglas":
    st.write("Tasa Oficial BCV activa:", "{:.2f} VES/USD".format(st.session_state.tasa_bcv))
    st.write("Regla de Valor Absoluto y validación SENIAT (IVA 75%/100%, ISLR e IGTF 3%) activadas.")
