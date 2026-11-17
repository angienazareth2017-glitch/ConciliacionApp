import streamlit as st
import pandas as pd
import io
import json
import urllib.request
import urllib.error
from datetime import datetime

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA E IDENTIDAD
# ==========================================
st.set_page_config(
    page_title="TESORIQ | Inteligencia y Control Financiero",
    page_icon="☘️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Logotipo Opción 3: Vectorial Mate Suave (SVG Nativo)
LOGO_TESORIQ_SVG = """
<svg width="220" height="54" viewBox="0 0 260 64" fill="none" xmlns="http://www.w3.org/2000/svg">
  <path d="M30 29C30 29 12 31 8 19C4 7 18 4 25 10C32 16 30 29 30 29Z" fill="#238458"/>
  <path d="M34 29C34 29 52 31 56 19C60 7 46 4 39 10C32 16 34 29 34 29Z" fill="#8CE0B5"/>
  <path d="M32 31C23 36 18 45 25 53C28 56 30 57 30 62H34C34 57 36 56 39 53C46 45 41 36 32 31Z" fill="#10251D"/>
  <text x="72" y="44" font-family="'Manrope', sans-serif" font-weight="800" font-size="31" fill="#F4F6F4" letter-spacing="1.5">TESORIQ</text>
</svg>
"""

LOGO_HEADER_LIGHT_SVG = """
<svg width="210" height="50" viewBox="0 0 260 64" fill="none" xmlns="http://www.w3.org/2000/svg">
  <path d="M30 29C30 29 12 31 8 19C4 7 18 4 25 10C32 16 30 29 30 29Z" fill="#238458"/>
  <path d="M34 29C34 29 52 31 56 19C60 7 46 4 39 10C32 16 34 29 34 29Z" fill="#8CE0B5"/>
  <path d="M32 31C23 36 18 45 25 53C28 56 30 57 30 62H34C34 57 36 56 39 53C46 45 41 36 32 31Z" fill="#10251D"/>
  <text x="72" y="44" font-family="'Manrope', sans-serif" font-weight="800" font-size="31" fill="#10251D" letter-spacing="1.5">TESORIQ</text>
</svg>
"""

# ==========================================
# 2. INYECCIÓN DE ESTILOS CSS (PALETA TESORIQ)
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Manrope:wght@600;700;800&display=swap');

    :root {
        --verde-bosque: #10251D;
        --verde-tecnologico: #238458;
        --menta-inteligente: #8CE0B5;
        --carbon: #171A18;
        --blanco-suave: #F4F6F4;
        --blanco-puro: #FFFFFF;
    }

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        color: var(--carbon);
        background-color: var(--blanco-suave);
    }

    h1, h2, h3, h4, .tesoriq-title {
        font-family: 'Manrope', sans-serif !important;
        color: var(--verde-bosque) !important;
        font-weight: 700 !important;
    }

    .stApp {
        background-color: var(--blanco-suave);
    }

    /* Barra Lateral (Sidebar) en Verde Bosque y Carbón */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #10251D 0%, #171A18 100%) !important;
        border-right: 1px solid rgba(140, 224, 181, 0.15);
    }
    section[data-testid="stSidebar"] * {
        color: #F4F6F4 !important;
    }
    section[data-testid="stSidebar"] input {
        color: #171A18 !important;
        background-color: #F4F6F4 !important;
    }
    section[data-testid="stSidebar"] .stRadio label {
        padding: 8px 12px;
        border-radius: 8px;
        transition: all 0.2s ease;
    }
    section[data-testid="stSidebar"] .stRadio label:hover {
        background-color: rgba(35, 132, 88, 0.35);
    }

    .sidebar-section {
        font-family: 'Manrope', sans-serif;
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 1.4px;
        color: #8CE0B5 !important;
        margin-top: 1.2rem;
        margin-bottom: 0.4rem;
        font-weight: 700;
    }

    /* Tarjetas de KPI (Dashboard) */
    .kpi-card {
        background-color: var(--blanco-puro);
        border: 1px solid #E2E8E5;
        border-left: 5px solid var(--verde-tecnologico);
        border-radius: 12px;
        padding: 20px 22px;
        box-shadow: 0 4px 12px rgba(16, 37, 29, 0.04);
    }
    .kpi-label {
        font-size: 0.85rem;
        color: #5A6560;
        font-weight: 500;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-family: 'Manrope', sans-serif;
        font-size: 1.85rem;
        font-weight: 800;
        color: var(--verde-bosque);
        margin: 0;
    }
    .kpi-badge {
        display: inline-block;
        margin-top: 8px;
        padding: 3px 10px;
        font-size: 0.75rem;
        font-weight: 600;
        border-radius: 20px;
        background-color: rgba(140, 224, 181, 0.25);
        color: var(--verde-bosque);
    }

    /* Botones Corporativos */
    .stButton > button {
        background-color: var(--verde-tecnologico) !important;
        color: #F4F6F4 !important;
        font-family: 'Manrope', sans-serif !important;
        font-weight: 600 !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.55rem 1.25rem !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button:hover {
        background-color: var(--verde-bosque) !important;
        box-shadow: 0 4px 12px rgba(35, 132, 88, 0.25) !important;
    }

    /* Banner Superior */
    .tesoriq-banner {
        background: linear-gradient(135deg, #10251D 0%, #238458 100%);
        color: #F4F6F4;
        padding: 22px 28px;
        border-radius: 14px;
        margin-bottom: 24px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .tesoriq-banner h3 {
        color: #F4F6F4 !important;
        margin: 0 0 4px 0;
        font-size: 1.25rem;
    }
    .tesoriq-banner p {
        color: #8CE0B5;
        margin: 0;
        font-size: 0.9rem;
    }

    /* Caja de Dictamen IA */
    .ia-box {
        background-color: #FFFFFF;
        border: 1px solid #8CE0B5;
        border-left: 6px solid #10251D;
        border-radius: 10px;
        padding: 18px 22px;
        margin-top: 14px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. GESTIÓN DE ESTADO E INICIO DE SESIÓN
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "df_banco" not in st.session_state:
    st.session_state.df_banco = None
if "df_erp" not in st.session_state:
    st.session_state.df_erp = None
if "df_resultado" not in st.session_state:
    st.session_state.df_resultado = None
if "tasa_bcv" not in st.session_state:
    st.session_state.tasa_bcv = 45.50
if "analisis_ia_texto" not in st.session_state:
    st.session_state.analisis_ia_texto = None

def obtener_api_key_defecto():
    try:
        if "OPENROUTER_API_KEY" in st.secrets:
            return st.secrets["OPENROUTER_API_KEY"]
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass
    return ""

if "api_key_ia" not in st.session_state:
    st.session_state.api_key_ia = obtener_api_key_defecto()

def mostrar_login():
    col1, col2, col3 = st.columns([1, 1.3, 1])
    with col2:
        st.write("")
        st.write("")
        st.markdown(f"<div style='text-align:center;'>{LOGO_HEADER_LIGHT_SVG}</div>", unsafe_allow_html=True)
        st.markdown(
            "<p style='text-align:center; color:#5A6560; font-size:0.95rem;'>"
            "Automatización y Control de Conciliaciones Financieras</p>",
            unsafe_allow_html=True
        )
        with st.container(border=True):
            st.markdown("#### Acceso Corporativo")
            usuario = st.text_input("Usuario / Correo", placeholder="analista@empresa.com")
            clave = st.text_input("Contraseña de acceso", type="password", placeholder="••••••••")
            if st.button("Ingresar a TESORIQ", use_container_width=True):
                if clave.strip() != "":
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error("Por favor ingresa tu contraseña para continuar.")

if not st.session_state.autenticado:
    mostrar_login()
    st.stop()

# ==========================================
# 4. MOTOR DE CONCILIACIÓN E INTELIGENCIA IA
# ==========================================
def ejecutar_conciliacion(df_b, df_e, tolerancia=0.05):
    """Cruza extracto bancario vs mayor ERP preservando 100% los montos de origen."""
    b = df_b.copy()
    e = df_e.copy()

    b["Referencia_Clean"] = b["Referencia"].astype(str).str.strip()
    e["Referencia_Clean"] = e["Referencia"].astype(str).str.strip()

    cruce = pd.merge(
        b, e,
        on="Referencia_Clean",
        how="outer",
        suffixes=("_Banco", "_ERP"),
        indicator=True
    )

    cruce["Monto_Banco"] = pd.to_numeric(cruce["Monto_Banco"], errors="coerce").fillna(0.0)
    cruce["Monto_ERP"] = pd.to_numeric(cruce["Monto_ERP"], errors="coerce").fillna(0.0)
    cruce["Diferencia_Absoluta"] = (cruce["Monto_Banco"] - cruce["Monto_ERP"]).abs().round(2)

    def clasificar_partida(row):
        if row["_merge"] == "both":
            if row["Diferencia_Absoluta"] <= tolerancia:
                return "Conciliado Exacto"
            return "Diferencia en Monto"
        elif row["_merge"] == "left_only":
            return "Pendiente en ERP (Solo Banco)"
        else:
            return "Pendiente en Banco (Solo ERP)"

    cruce["Estado_Conciliacion"] = cruce.apply(clasificar_partida, axis=1)

    def sugerir_causa_local(row):
        estado = row["Estado_Conciliacion"]
        dif = row["Diferencia_Absoluta"]
        if estado == "Conciliado Exacto":
            return "Sin novedad. Cruce exacto."
        elif estado == "Diferencia en Monto":
            if dif <= 5.00:
                return "Posible diferencial por redondeo o céntimos en tasa BCV."
            return "Verificar retención SENIAT (IVA 75%/100% o ISLR) o comisión bancaria descontada."
        elif estado == "Pendiente en ERP (Solo Banco)":
            return "Cargo/abono bancario no contabilizado (comisión, IGTF o cobranza por identificar)."
        else:
            return "Partida en tránsito registrada en libros pero aún no efectiva en banco."

    cruce["Diagnostico_Preliminar"] = cruce.apply(sugerir_causa_local, axis=1)
    return cruce

def analizar_excepciones_con_ia(df_excepciones, tasa_bcv, api_key, modelo="openai/gpt-4o-mini"):
    """Envía las partidas abiertas a OpenRouter / OpenAI para dictamen contable asistido."""
    resumen_json = df_excepciones[[
        "Referencia_Clean", "Estado_Conciliacion",
        "Monto_Banco", "Monto_ERP",
        "Diferencia_Absoluta", "Equivalente_USD_BCV"
    ]].to_dict(orient="records")

    prompt_sistema = (
        "Eres el auditor financiero y analista experto de tesorería de la plataforma TESORIQ en Venezuela. "
        "REGLA DE GOBERNANZA: Nunca alteres los montos originales ni apruebes ajustes autónomos; tu rol es "
        "analizar las partidas pendientes usando valor absoluto y sugerir al analista humano la acción contable."
    )

    prompt_usuario = f"""
    Tasa oficial BCV aplicada: {tasa_bcv:.2f} VES/USD.
    Analiza las siguientes partidas pendientes de conciliación bancaria:
    {json.dumps(resumen_json, ensure_ascii=False)}

    Para cada referencia indica de forma ejecutiva y estructurada:
    1. **Causa raíz probable** (ej. comisión bancaria, IGTF, retención de IVA 75%/100% o ISLR SENIAT, partida en tránsito o diferencial por redondeo).
    2. **Acción contable / Asiento sugerido** en el ERP para que el analista de cuentas por pagar / tesorería lo valide.
    """

    # Si no hay API Key configurada, genera un dictamen estructurado inteligente local sin fallar
    if not api_key or api_key.strip() == "":
        lineas = ["#### 📋 Dictamen Técnico Asistido TESORIQ (Modo Local Activo)"]
        lineas.append(f"*Tasa de referencia BCV: **{tasa_bcv:.2f} VES/USD** — Integridad de datos originales: **100% Intacta***\n")
        for item in resumen_json:
            ref = item["Referencia_Clean"]
            est = item["Estado_Conciliacion"]
            mb = item["Monto_Banco"]
            me = item["Monto_ERP"]
            dif = item["Diferencia_Absoluta"]
            usd = item["Equivalente_USD_BCV"]
            if est == "Diferencia en Monto":
                accion = "Revisar comprobante de retención IVA/ISLR o ajuste por céntimos de tasa BCV antes de contabilizar diferencia."
            elif "Solo Banco" in est:
                accion = "Registrar nota de débito/comisión bancaria o solicitar soporte de pago al departamento correspondiente."
            else:
                accion = "Mantener en control de partidas en tránsito hasta confirmar efectividad en el próximo extracto bancario."
            lineas.append(
                f"- **Referencia `{ref}`** ({est}): Banco `{mb:.2f}` vs ERP `{me:.2f}` | "
                f"Diferencia Absoluta: **{dif:.2f} VES** (*{usd:.2f} USD*).  \n"
                f"  ↳ **Sugerencia para el Analista:** {accion}"
            )
        lineas.append("\n*(Nota: Ingresa tu API Key de OpenRouter en la barra lateral para activar la redacción generativa avanzada).*")
        return "\n".join(lineas)

    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key.strip()}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://tesoriq.app",
        "X-Title": "TESORIQ Financial Reconciliation"
    }
    payload = {
        "model": modelo,
        "messages": [
            {"role": "system", "content": prompt_sistema},
            {"role": "user", "content": prompt_usuario}
        ],
        "temperature": 0.2
    }

    try:
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers=headers,
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=25) as resp:
            respuesta_data = json.loads(resp.read().decode("utf-8"))
            return respuesta_data["choices"][0]["message"]["content"]
    except urllib.error.HTTPError as e:
        return f"⚠️ Error de autenticación o límite en OpenRouter (Código {e.code}). Verifica tu API Key en la barra lateral."
    except Exception as e:
        return f"⚠️ No se pudo conectar con OpenRouter ({str(e)}). Revisa tu conexión a internet."

def generar_datos_ejemplo():
    df_b = pd.DataFrame({
        "Fecha_Banco": ["2026-11-10", "2026-11-11", "2026-11-12", "2026-11-13", "2026-11-14"],
        "Referencia": ["REF-1001", "REF-1002", "REF-1003", "REF-1004", "REF-1005"],
        "Descripcion_Banco": ["Pago Proveedor Vet", "Transferencia Recibida", "Comisión Mantenimiento e IGTF", "Pago Servicios", "Abono Clientes"],
        "Monto_Banco": [1500.00, 3200.50, 45.00, 890.00, 2100.00]
    })
    df_e = pd.DataFrame({
        "Fecha_ERP": ["2026-11-10", "2026-11-11", "2026-11-13", "2026-11-14", "2026-11-15"],
        "Referencia": ["REF-1001", "REF-1002", "REF-1004", "REF-1005", "REF-1006"],
        "Descripcion_ERP": ["Factura Quimio-Vet", "Cobro Distribuidora", "Servicio Eléctrico", "Cobro Lote B", "Pago en Tránsito Proveedor"],
        "Monto_ERP": [1500.00, 3200.00, 890.00, 2100.00, 740.00]
    })
    return df_b, df_e

# ==========================================
# 5. BARRA LATERAL (MENÚ ESTRUCTURADO)
# ==========================================
with st.sidebar:
    st.markdown(LOGO_TESORIQ_SVG, unsafe_allow_html=True)
    st.markdown("<div class='sidebar-section'>Proceso Principal</div>", unsafe_allow_html=True)
    menu = st.radio(
        "Navegación",
        options=[
            "📊 Dashboard",
            "📥 Importar Archivos",
            "🔄 Conciliación y Cruce",
            "⚠️ Partidas Pendientes (IA)",
            "🏛️ Tasas BCV y Reglas"
        ],
        label_visibility="collapsed"
    )

    st.markdown("<div class='sidebar-section'>Parámetros del Día</div>", unsafe_allow_html=True)
    st.session_state.tasa_bcv = st.number_input(
        "Tasa Oficial BCV (VES/USD)",
        min_value=1.0,
        value=float(st.session_state.tasa_bcv),
        step=0.10,
        format="%.2f"
    )

    st.markdown("<div class='sidebar-section'>Motor IA (OpenRouter)</div>", unsafe_allow_html=True)
    st.session_state.api_key_ia = st.text_input(
        "OpenRouter API Key",
        value=st.session_state.api_key_ia,
        type="password",
        placeholder="sk-or-v1-..."
    )
    modelo_ia = st.selectbox(
        "Modelo LLM Activo",
        options=[
            "openai/gpt-4o-mini",
            "google/gemini-2.0-flash-001",
            "anthropic/claude-3.5-haiku"
        ]
    )

    st.caption(f"Fecha activa: {datetime.now().strftime('%d/%m/%Y')}")
    st.divider()
    if st.button("Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

# ==========================================
# 6. VISTAS DE LA APLICACIÓN
# ==========================================
df_res = st.session_state.df_resultado

if df_res is not None:
    total_mov = len(df_res)
    conciliados = len(df_res[df_res["Estado_Conciliacion"] == "Conciliado Exacto"])
    diferencias = len(df_res[df_res["Estado_Conciliacion"] == "Diferencia en Monto"])
    abiertas = len(df_res[df_res["Estado_Conciliacion"].str.contains("Pendiente")])
else:
    total_mov, conciliados, diferencias, abiertas = 0, 0, 0, 0

if menu == "📊 Dashboard":
    st.markdown("""
    <div class="tesoriq-banner">
        <div>
            <h3>Panel de Control e Inteligencia de Tesorería</h3>
            <p>Deja de buscar diferencias. Empieza a resolverlas con trazabilidad total.</p>
        </div>
        <div style="text-align:right; font-family:'Manrope',sans-serif; font-weight:700; color:#8CE0B5;">
            TESORIQ v2.5 • IA Activa
        </div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Partidas Abiertas</div>
            <p class="kpi-value">{abiertas}</p>
            <span class="kpi-badge">Por revisar</span>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Movimientos Conciliados</div>
            <p class="kpi-value">{conciliados}</p>
            <span class="kpi-badge">Match 100%</span>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Diferencias de Monto</div>
            <p class="kpi-value">{diferencias}</p>
            <span class="kpi-badge">Valor absoluto</span>
        </div>
        """, unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Tasa BCV Aplicada</div>
            <p class="kpi-value">{st.session_state.tasa_bcv:.2f}</p>
            <span class="kpi-badge">VES / USD</span>
        </div>
        """, unsafe_allow_html=True)

    st.write("")
    col_left, col_right = st.columns([1.8, 1])

    with col_left:
        with st.container(border=True):
            st.markdown("#### Conciliaciones Recientes")
            if df_res is None:
                st.info("Aún no se ha ejecutado una conciliación en esta sesión. Usa el botón de prueba rápida o importa tus archivos.")
                if st.button("⚡ Cargar Datos Demo y Conciliar Ahora"):
                    b_demo, e_demo = generar_datos_ejemplo()
                    st.session_state.df_banco = b_demo
                    st.session_state.df_erp = e_demo
                    st.session_state.df_resultado = ejecutar_conciliacion(b_demo, e_demo)
                    st.rerun()
            else:
                st.dataframe(
                    df_res[[
                        "Referencia_Clean", "Monto_Banco", "Monto_ERP",
                        "Diferencia_Absoluta", "Estado_Conciliacion", "Diagnostico_Preliminar"
                    ]],
                    use_container_width=True,
                    hide_index=True
                )

    with col_right:
        with st.container(border=True):
            st.markdown("#### Gobernanza y Supervisión")
            st.write(
                "• **Integridad garantizada:** Los montos originales de Banco y ERP permanecen inmutables.\n"
                "• **Supervisión humana:** TESORIQ clasifica y sugiere acciones sobre excepciones con IA, manteniendo la aprobación final en el analista."
            )
            if df_res is not None:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_res.to_excel(writer, index=False, sheet_name="Conciliacion_Tesoriq")
                st.download_button(
                    label="📥 Descargar Reporte Excel (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"Conciliacion_TESORIQ_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

elif menu == "📥 Importar Archivos":
    st.markdown("### Importación de Estados de Cuenta y Mayor Contable")
    col_b, col_e = st.columns(2)

    with col_b:
        with st.container(border=True):
            st.markdown("#### 1. Extracto Bancario (CSV / Excel)")
            archivo_b = st.file_uploader("Subir archivo de Banco", type=["csv", "xlsx"], key="up_b")
            if archivo_b:
                st.session_state.df_banco = pd.read_csv(archivo_b) if archivo_b.name.endswith(".csv") else pd.read_excel(archivo_b)
                st.success(f"Extracto cargado: {len(st.session_state.df_banco)} registros.")

    with col_e:
        with st.container(border=True):
            st.markdown("#### 2. Mayor Auxiliar ERP (CSV / Excel)")
            archivo_e = st.file_uploader("Subir archivo ERP (SAP / Busis)", type=["csv", "xlsx"], key="up_e")
            if archivo_e:
                st.session_state.df_erp = pd.read_csv(archivo_e) if archivo_e.name.endswith(".csv") else pd.read_excel(archivo_e)
                st.success(f"Mayor ERP cargado: {len(st.session_state.df_erp)} registros.")

    st.write("")
    if st.button("Cargar Plantilla de Ejemplo Lista para Prueba"):
        b_demo, e_demo = generar_datos_ejemplo()
        st.session_state.df_banco = b_demo
        st.session_state.df_erp = e_demo
        st.success("Archivos de prueba cargados correctamente. Dirígete a 'Conciliación y Cruce'.")

elif menu == "🔄 Conciliación y Cruce":
    st.markdown("### Motor de Conciliación Asistida")
    if st.session_state.df_banco is None or st.session_state.df_erp is None:
        st.warning("Primero debes cargar los archivos en el módulo 'Importar Archivos'.")
    else:
        tolerancia = st.slider("Tolerancia para diferencias por redondeo (VES/USD)", 0.00, 5.00, 0.05, 0.01)
        if st.button("🚀 Ejecutar Cruce Inteligente"):
            st.session_state.df_resultado = ejecutar_conciliacion(
                st.session_state.df_banco,
                st.session_state.df_erp,
                tolerancia=tolerancia
            )
            st.session_state.analisis_ia_texto = None
            st.success("Conciliación completada respetando el 100% de la integridad de los datos de origen.")

        if st.session_state.df_resultado is not None:
            st.dataframe(st.session_state.df_resultado, use_container_width=True, hide_index=True)

elif menu == "⚠️ Partidas Pendientes (IA)":
    st.markdown("### Auditoría Inteligente de Excepciones y Partidas Abiertas")
    if df_res is None:
        st.info("Ejecuta primero una conciliación en el Dashboard o en 'Conciliación y Cruce' para auditar las partidas pendientes.")
    else:
        excepciones = df_res[df_res["Estado_Conciliacion"] != "Conciliado Exacto"].copy()
        excepciones["Equivalente_USD_BCV"] = (excepciones["Diferencia_Absoluta"] / st.session_state.tasa_bcv).round(2)

        st.markdown(f"**Total de excepciones detectadas:** `{len(excepciones)}` partidas requieren validación del analista.")
        st.dataframe(
            excepciones[[
                "Referencia_Clean", "Estado_Conciliacion",
                "Monto_Banco", "Monto_ERP",
                "Diferencia_Absoluta", "Equivalente_USD_BCV",
                "Diagnostico_Preliminar"
            ]],
            use_container_width=True,
            hide_index=True
        )

        st.write("")
        col_btn1, col_btn2 = st.columns([1.2, 2])
        with col_btn1:
            if st.button("🤖 Generar Dictamen Contable con IA", use_container_width=True):
                with st.spinner("Analizando excepciones bajo reglas de tesorería y normativa SENIAT..."):
                    st.session_state.analisis_ia_texto = analizar_excepciones_con_ia(
                        excepciones,
                        st.session_state.tasa_bcv,
                        st.session_state.api_key_ia,
                        modelo=modelo_ia
                    )

        if st.session_state.analisis_ia_texto:
            st.markdown("<div class='ia-box'>", unsafe_allow_html=True)
            st.markdown(st.session_state.analisis_ia_texto)
            st.markdown("</div>", unsafe_allow_html=True)

elif menu == "🏛️ Tasas BCV y Reglas":
    st.markdown("### Configuración de Reglas de Negocio y Multimoneda")
    with st.container(border=True):
        st.write(f"**Tasa Oficial BCV activa en sesión:** {st.session_state.tasa_bcv:.2f} VES/USD")
        st.write("**Regla de Valor Absoluto:** Activada (`|Monto_Banco - Monto_ERP|`) para evitar errores de interpretación en diferencias negativas.")
        st.write("**Retenciones SENIAT (IVA / ISLR) e IGTF:** Reglas de auditoría incorporadas en el motor de análisis de excepciones.")