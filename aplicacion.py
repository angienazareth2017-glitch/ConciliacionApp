# -*- coding: utf-8 -*-
"""
================================================================================
TESORIQ | Plataforma de Inteligencia y Conciliación Financiera
Módulo de Identidad Visual, Diseño UI/UX y Auditoría de Excepciones con IA
================================================================================
INSTRUCCIONES DE INTEGRACIÓN PARA TU ARCHIVO DE 2.036 LÍNEAS:
1. NO borres tus 2.036 líneas de lógica actual.
2. Importa o copia las funciones de este módulo al inicio de tu aplicacion.py:
   - aplicar_identidad_tesoriq() -> Aplica paleta (#10251D, #238458, #8CE0B5, #171A18, #F4F6F4),
     tipografías (Manrope e Inter) y el logo Opción 3 (Trébol en T Vectorial Mate Suave).
   - renderizar_logo_sidebar() -> Muestra el logo oficial en el menú lateral.
   - renderizar_banner_tesoriq() -> Muestra el encabezado ejecutivo.
   - modulo_partidas_pendientes_ia(df_excepciones, tasa_bcv) -> Integra el análisis de IA
     para partidas pendientes respetando tu estructura actual.
================================================================================
"""

import streamlit as st
import pandas as pd
import io
import json
import urllib.request
import urllib.error
from datetime import datetime

# ==============================================================================
# 1. LOGOTIPO OFICIAL TESORIQ - OPCIÓN 3: VECTORIAL MATE SUAVE (FLAT MINIMALIST)
# ==============================================================================
LOGO_TESORIQ_DARK_SVG = """
<svg width="225" height="56" viewBox="0 0 265 64" fill="none" xmlns="http://www.w3.org/2000/svg">
  <!-- Hoja Superior Izquierda: Verde Tecnológico (#238458) -->
  <path d="M30 29C30 29 11 31 7 18.5C3 6 18 3.5 25.5 10C32.5 16 30 29 30 29Z" fill="#238458"/>
  <!-- Hoja Superior Derecha: Menta Inteligente (#8CE0B5) -->
  <path d="M34 29C34 29 53 31 57 18.5C61 6 46 3.5 38.5 10C31.5 16 34 29 34 29Z" fill="#8CE0B5"/>
  <!-- Hoja Central Inferior + Tallo en T: Verde Bosque (#10251D) con borde sutil para fondo oscuro -->
  <path d="M32 31C22.5 36 17.5 45.5 25 53.5C28 56.5 30 57.5 30 62H34C34 57.5 36 56.5 39 53.5C46.5 45.5 41.5 36 32 31Z" fill="#238458" stroke="#8CE0B5" stroke-width="1.2"/>
  <!-- Tipografía Principal: Manrope ExtraBold -->
  <text x="72" y="44" font-family="'Manrope', sans-serif" font-weight="800" font-size="31" fill="#F4F6F4" letter-spacing="1.5">TESORIQ</text>
</svg>
"""

LOGO_TESORIQ_LIGHT_SVG = """
<svg width="225" height="56" viewBox="0 0 265 64" fill="none" xmlns="http://www.w3.org/2000/svg">
  <!-- Hoja Superior Izquierda: Verde Tecnológico (#238458) -->
  <path d="M30 29C30 29 11 31 7 18.5C3 6 18 3.5 25.5 10C32.5 16 30 29 30 29Z" fill="#238458"/>
  <!-- Hoja Superior Derecha: Menta Inteligente (#8CE0B5) -->
  <path d="M34 29C34 29 53 31 57 18.5C61 6 46 3.5 38.5 10C31.5 16 34 29 34 29Z" fill="#8CE0B5"/>
  <!-- Hoja Central Inferior + Tallo en T: Verde Bosque (#10251D) -->
  <path d="M32 31C22.5 36 17.5 45.5 25 53.5C28 56.5 30 57.5 30 62H34C34 57.5 36 56.5 39 53.5C46.5 45.5 41.5 36 32 31Z" fill="#10251D"/>
  <!-- Tipografía Principal: Manrope ExtraBold -->
  <text x="72" y="44" font-family="'Manrope', sans-serif" font-weight="800" font-size="31" fill="#10251D" letter-spacing="1.5">TESORIQ</text>
</svg>
"""


# ==============================================================================
# 2. SISTEMA DE DISEÑO Y PALETA CROMÁTICA TESORIQ (CSS INJECTION)
# ==============================================================================
def aplicar_identidad_tesoriq():
    """
    Inyecta la paleta cromática oficial de TESORIQ y las fuentes Manrope + Inter
    sin alterar las funciones ni cálculos existentes en tu código principal.
    """
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Manrope:wght@600;700;800&display=swap');

        :root {
            --tesoriq-bosque: #10251D;       /* Color principal: confianza y solidez */
            --tesoriq-tecnologico: #238458;  /* Acciones: innovación y crecimiento */
            --tesoriq-menta: #8CE0B5;        /* Detalles: datos destacados y acentos */
            --tesoriq-carbon: #171A18;       /* Textos: contraste y elegancia */
            --tesoriq-blanco-suave: #F4F6F4; /* Fondos: limpieza y legibilidad */
            --tesoriq-blanco-puro: #FFFFFF;
        }

        /* Aplicación de fuente secundaria (Inter) para textos, tablas y formularios */
        html, body, [class*="css"], .stDataFrame, .stTable, input, select, textarea {
            font-family: 'Inter', sans-serif !important;
            color: var(--tesoriq-carbon);
        }

        /* Aplicación de fuente principal (Manrope) para títulos y encabezados */
        h1, h2, h3, h4, h5, h6, .tesoriq-heading {
            font-family: 'Manrope', sans-serif !important;
            color: var(--tesoriq-bosque) !important;
            font-weight: 700 !important;
            letter-spacing: -0.01em;
        }

        /* Fondo general limpio (#F4F6F4) */
        .stApp {
            background-color: var(--tesoriq-blanco-suave) !important;
        }

        /* Menú Lateral (Sidebar) en Verde Bosque (#10251D) y Carbón (#171A18) */
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #10251D 0%, #171A18 100%) !important;
            border-right: 1px solid rgba(140, 224, 181, 0.18);
        }
        section[data-testid="stSidebar"] p,
        section[data-testid="stSidebar"] span,
        section[data-testid="stSidebar"] label,
        section[data-testid="stSidebar"] div {
            color: var(--tesoriq-blanco-suave) !important;
        }
        section[data-testid="stSidebar"] input {
            color: var(--tesoriq-carbon) !important;
            background-color: var(--tesoriq-blanco-suave) !important;
            border-radius: 6px !important;
        }

        /* Botones de Acción Principal en Verde Tecnológico (#238458) */
        .stButton > button, .stDownloadButton > button {
            background-color: var(--tesoriq-tecnologico) !important;
            color: var(--tesoriq-blanco-suave) !important;
            font-family: 'Manrope', sans-serif !important;
            font-weight: 600 !important;
            border: 1px solid transparent !important;
            border-radius: 8px !important;
            padding: 0.55rem 1.25rem !important;
            transition: all 0.2s ease-in-out !important;
        }
        .stButton > button:hover, .stDownloadButton > button:hover {
            background-color: var(--tesoriq-bosque) !important;
            border-color: var(--tesoriq-menta) !important;
            box-shadow: 0 4px 12px rgba(35, 132, 88, 0.28) !important;
        }

        /* Tarjetas Ejecutivas de KPIs */
        .tesoriq-kpi-card {
            background-color: var(--tesoriq-blanco-puro);
            border: 1px solid #E1E7E4;
            border-left: 5px solid var(--tesoriq-tecnologico);
            border-radius: 12px;
            padding: 18px 20px;
            box-shadow: 0 3px 10px rgba(16, 37, 29, 0.04);
            margin-bottom: 12px;
        }
        .tesoriq-kpi-title {
            font-family: 'Inter', sans-serif;
            font-size: 0.84rem;
            color: #55625C;
            font-weight: 500;
            margin-bottom: 6px;
        }
        .tesoriq-kpi-value {
            font-family: 'Manrope', sans-serif;
            font-size: 1.8rem;
            font-weight: 800;
            color: var(--tesoriq-bosque);
            margin: 0;
        }
        .tesoriq-kpi-tag {
            display: inline-block;
            margin-top: 8px;
            padding: 2px 10px;
            font-size: 0.74rem;
            font-weight: 600;
            border-radius: 99px;
            background-color: rgba(140, 224, 181, 0.28);
            color: var(--tesoriq-bosque);
        }

        /* Banner Corporativo Superior */
        .tesoriq-hero-banner {
            background: linear-gradient(135deg, #10251D 0%, #238458 100%);
            color: #F4F6F4;
            padding: 20px 26px;
            border-radius: 14px;
            margin-bottom: 22px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 6px 18px rgba(16, 37, 29, 0.12);
        }
        .tesoriq-hero-banner h2 {
            color: #F4F6F4 !important;
            margin: 0 0 4px 0 !important;
            font-size: 1.35rem !important;
        }
        .tesoriq-hero-banner p {
            color: #8CE0B5 !important;
            margin: 0 !important;
            font-size: 0.92rem !important;
        }

        /* Contenedor de Dictamen de Auditoría IA */
        .tesoriq-ia-panel {
            background-color: #FFFFFF;
            border: 1px solid #8CE0B5;
            border-left: 6px solid #10251D;
            border-radius: 12px;
            padding: 20px 24px;
            margin-top: 16px;
            box-shadow: 0 4px 14px rgba(16, 37, 29, 0.05);
        }
    </style>
    """, unsafe_allow_html=True)


def renderizar_logo_sidebar(modo="oscuro"):
    """Muestra el logotipo vectorial mate suave de TESORIQ."""
    svg = LOGO_TESORIQ_DARK_SVG if modo == "oscuro" else LOGO_TESORIQ_LIGHT_SVG
    st.markdown(f"<div style='padding: 6px 0 14px 0;'>{svg}</div>", unsafe_allow_html=True)


def renderizar_banner_tesoriq(titulo="Panel de Control e Inteligencia de Tesorería", subtitulo="Deja de buscar diferencias. Empieza a resolverlas con trazabilidad total."):
    """Renderiza el encabezado ejecutivo con la identidad TESORIQ."""
    st.markdown(f"""
    <div class="tesoriq-hero-banner">
        <div>
            <h2>{titulo}</h2>
            <p>{subtitulo}</p>
        </div>
        <div style="text-align:right; font-family:'Manrope',sans-serif; font-weight:700; color:#8CE0B5; font-size:0.88rem;">
            TESORIQ • ENTERPRISE DATA
        </div>
    </div>
    """, unsafe_allow_html=True)


# ==============================================================================
# 3. MÓDULO DE PARTIDAS PENDIENTES Y AUDITORÍA ASISTIDA CON IA (OPENROUTER)
# ==============================================================================
def analizar_partidas_pendientes_ia(
    df_pendientes: pd.DataFrame,
    tasa_bcv: float,
    api_key: str = "",
    modelo: str = "openai/gpt-4o-mini"
) -> str:
    """
    Analiza las partidas pendientes/excepciones utilizando valor absoluto matemático
    y genera recomendaciones contables sin alterar los datos originales de Banco o ERP.
    Compatible con cualquier DataFrame de tu código de 2.036 líneas.
    """
    if df_pendientes is None or df_pendientes.empty:
        return "✅ No existen partidas pendientes ni diferencias por auditar en este cruce."

    df_trabajo = df_pendientes.copy()

    # Detección flexible de columnas de montos para adaptarse a tu código de 2.036 líneas
    col_banco = next((c for c in df_trabajo.columns if "banco" in c.lower() and ("monto" in c.lower() or "saldo" in c.lower() or "importe" in c.lower())), None)
    col_erp = next((c for c in df_trabajo.columns if ("erp" in c.lower() or "libro" in c.lower() or "mayor" in c.lower()) and ("monto" in c.lower() or "saldo" in c.lower() or "importe" in c.lower())), None)
    col_ref = next((c for c in df_trabajo.columns if "ref" in c.lower() or "doc" in c.lower() or "num" in c.lower()), df_trabajo.columns[0])

    if col_banco and col_erp:
        m_banco = pd.to_numeric(df_trabajo[col_banco], errors="coerce").fillna(0.0)
        m_erp = pd.to_numeric(df_trabajo[col_erp], errors="coerce").fillna(0.0)
        df_trabajo["Diferencia_Absoluta_VES"] = (m_banco - m_erp).abs().round(2)
        df_trabajo["Equivalente_USD_BCV"] = (df_trabajo["Diferencia_Absoluta_VES"] / max(float(tasa_bcv), 0.01)).round(2)

    # Tomamos hasta 30 registros para enviar al prompt sin exceder contexto
    muestra_registros = df_trabajo.head(30).to_dict(orient="records")

    prompt_sistema = (
        "Eres el Auditor Senior de Tesorería e Inteligencia Financiera de TESORIQ en Venezuela. "
        "REGLA DE GOBERNANZA INQUEBRANTABLE: Jamás modifiques los montos originales del Banco ni del ERP, "
        "ni ejecutes ajustes contables autónomos. Analiza las diferencias usando siempre el valor absoluto "
        "matemático y entrega un dictamen técnico claro para que el analista humano tome la decisión final."
    )

    prompt_usuario = f"""
    Tasa de Cambio Oficial BCV aplicada: {tasa_bcv:.2f} VES/USD.
    Registros con excepciones o partidas pendientes detectadas en la conciliación:
    {json.dumps(muestra_registros, default=str, ensure_ascii=False)}

    Genera un informe ejecutivo estructurado que incluya:
    1. **Resumen de Excepciones:** Clasificación rápida por tipo de discrepancia.
    2. **Análisis de Causa Raíz por Partida:** Evalúa si la diferencia corresponde a retenciones SENIAT (IVA 75%/100% o ISLR), IGTF (3%), comisiones bancarias no registradas, cheques/transferencias en tránsito o diferencial por redondeo cambiario BCV.
    3. **Acción Recomendada al Analista:** Asiento o verificación específica sugerida en el ERP.
    """

    # Respaldo local automático si aún no se ha ingresado la clave de OpenRouter
    if not api_key or api_key.strip() == "":
        lineas = [
            "#### 📋 Dictamen de Auditoría Asistida TESORIQ (Motor de Reglas Local Activo)",
            f"*Tasa BCV aplicada: **{tasa_bcv:.2f} VES/USD** | Total de partidas auditadas: **{len(df_trabajo)}** | Integridad de datos: **100% Preservada***
"
        ]
        for idx, row in df_trabajo.head(15).iterrows():
            ref_val = row.get(col_ref, f"Partida-{idx+1}")
            dif_ves = row.get("Diferencia_Absoluta_VES", "N/A")
            dif_usd = row.get("Equivalente_USD_BCV", "N/A")
            lineas.append(
                f"- **Referencia `{ref_val}`**: Diferencia Absoluta **{dif_ves} VES** (*Eq. {dif_usd} USD a tasa BCV*).  
"
                f"  ↳ *Recomendación:* Verificar soporte de retención IVA/ISLR, comisión bancaria/IGTF o confirmar estatus en tránsito antes de asentar en el ERP."
            )
        lineas.append("
> 💡 *Para activar la redacción generativa profunda con LLMs (GPT-4o / Gemini Flash / Claude), ingresa tu API Key de OpenRouter en la configuración.*")
        return "
".join(lineas)

    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key.strip()}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://tesoriq.app",
        "X-Title": "TESORIQ Enterprise Reconciliation"
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
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            return data["choices"][0]["message"]["content"]
    except urllib.error.HTTPError as e:
        return f"⚠️ Error HTTP {e.code} al conectar con OpenRouter. Verifica tu API Key."
    except Exception as e:
        return f"⚠️ No fue posible conectar con el motor de IA ({str(e)})."


def renderizar_seccion_partidas_pendientes(df_pendientes: pd.DataFrame, tasa_bcv: float = 45.50, api_key_default: str = ""):
    """
    Componente visual completo de Partidas Pendientes que puedes invocar en una sola línea
    dentro de tu archivo aplicacion.py de 2.036 líneas:
        renderizar_seccion_partidas_pendientes(df_no_conciliados, tasa_bcv)
    """
    st.markdown("### ⚠️ Auditoría Inteligente de Partidas Pendientes y Excepciones")
    st.caption("Gobernanza TESORIQ: La IA clasifica discrepancias en valor absoluto y sugiere acciones contables; el analista humano conserva el control total de aprobación.")

    if df_pendientes is None or df_pendientes.empty:
        st.success("🎉 ¡Excelente! No hay partidas pendientes ni diferencias abiertas en esta conciliación.")
        return

    df_vista = df_pendientes.copy()
    col_banco = next((c for c in df_vista.columns if "banco" in c.lower() and ("monto" in c.lower() or "saldo" in c.lower() or "importe" in c.lower())), None)
    col_erp = next((c for c in df_vista.columns if ("erp" in c.lower() or "libro" in c.lower() or "mayor" in c.lower()) and ("monto" in c.lower() or "saldo" in c.lower() or "importe" in c.lower())), None)

    if col_banco and col_erp:
        m_b = pd.to_numeric(df_vista[col_banco], errors="coerce").fillna(0.0)
        m_e = pd.to_numeric(df_vista[col_erp], errors="coerce").fillna(0.0)
        df_vista["Diferencia_Absoluta_VES"] = (m_b - m_e).abs().round(2)
        df_vista["Equivalente_USD_BCV"] = (df_vista["Diferencia_Absoluta_VES"] / max(float(tasa_bcv), 0.01)).round(2)

    # Resumen rápido en tarjetas KPI
    k1, k2, k3 = st.columns(3)
    total_excepciones = len(df_vista)
    suma_dif_ves = df_vista["Diferencia_Absoluta_VES"].sum() if "Diferencia_Absoluta_VES" in df_vista.columns else 0.0
    suma_dif_usd = (suma_dif_ves / max(float(tasa_bcv), 0.01)) if suma_dif_ves else 0.0

    with k1:
        st.markdown(f"""
        <div class="tesoriq-kpi-card">
            <div class="tesoriq-kpi-title">Partidas Abiertas Detectadas</div>
            <p class="tesoriq-kpi-value">{total_excepciones}</p>
            <span class="tesoriq-kpi-tag">Requieren revisión humana</span>
        </div>
        """, unsafe_allow_html=True)
    with k2:
        st.markdown(f"""
        <div class="tesoriq-kpi-card">
            <div class="tesoriq-kpi-title">Discrepancia Total (Valor Absoluto)</div>
            <p class="tesoriq-kpi-value">{suma_dif_ves:,.2f} VES</p>
            <span class="tesoriq-kpi-tag">Datos originales intactos</span>
        </div>
        """, unsafe_allow_html=True)
    with k3:
        st.markdown(f"""
        <div class="tesoriq-kpi-card">
            <div class="tesoriq-kpi-title">Impacto Equivalente en Divisas</div>
            <p class="tesoriq-kpi-value">${suma_dif_usd:,.2f} USD</p>
            <span class="tesoriq-kpi-tag">Tasa BCV: {tasa_bcv:.2f}</span>
        </div>
        """, unsafe_allow_html=True)

    st.dataframe(df_vista, use_container_width=True, hide_index=True)

    with st.expander("⚙️ Configurar Motor IA (OpenRouter / Modelo)", expanded=False):
        c_api, c_mod = st.columns([2, 1])
        with c_api:
            api_key_input = st.text_input(
                "Clave API de OpenRouter",
                value=api_key_default,
                type="password",
                key="tesoriq_or_key"
            )
        with c_mod:
            modelo_sel = st.selectbox(
                "Modelo Analítico",
                options=[
                    "google/gemini-2.0-flash-001",
                    "openai/gpt-4o-mini",
                    "anthropic/claude-3.5-haiku"
                ],
                key="tesoriq_or_model"
            )

    col_b1, col_b2 = st.columns([1.2, 1.8])
    with col_b1:
        if st.button("🤖 Auditar Partidas Pendientes con IA", use_container_width=True):
            with st.spinner("Ejecutando análisis contable de excepciones con reglas TESORIQ..."):
                st.session_state["tesoriq_ultimo_dictamen"] = analizar_partidas_pendientes_ia(
                    df_vista,
                    tasa_bcv=tasa_bcv,
                    api_key=api_key_input,
                    modelo=modelo_sel
                )

    with col_b2:
        buffer_exc = io.BytesIO()
        with pd.ExcelWriter(buffer_exc, engine="openpyxl") as writer:
            df_vista.to_excel(writer, index=False, sheet_name="Partidas_Pendientes")
        st.download_button(
            label="📥 Exportar Partidas Pendientes (.xlsx)",
            data=buffer_exc.getvalue(),
            file_name=f"Partidas_Pendientes_TESORIQ_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    if st.session_state.get("tesoriq_ultimo_dictamen"):
        st.markdown("<div class='tesoriq-ia-panel'>", unsafe_allow_html=True)
        st.markdown(st.session_state["tesoriq_ultimo_dictamen"])
        st.markdown("</div>", unsafe_allow_html=True)


# ==============================================================================
# 4. EJECUCIÓN AUTÓNOMA DE DEMOSTRACIÓN (SI EJECUTAS ESTE ARCHIVO DIRECTAMENTE)
# ==============================================================================
if __name__ == "__main__":
    st.set_page_config(
        page_title="TESORIQ | Inteligencia y Control Financiero",
        page_icon="☘️",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # 1. Aplicamos la identidad visual TESORIQ
    aplicar_identidad_tesoriq()

    # 2. Barra lateral con el nuevo logo Opción 3 (Vectorial Mate Suave)
    with st.sidebar:
        renderizar_logo_sidebar(modo="oscuro")
        st.markdown("---")
        vista_activa = st.radio(
            "Módulos TESORIQ",
            ["📊 Dashboard y Partidas Pendientes", "📘 Guía de Integración (2.036 líneas)"]
        )
        tasa_actual = st.number_input("Tasa Oficial BCV (VES/USD)", min_value=1.0, value=45.50, step=0.10)

    # 3. Contenido Principal
    renderizar_banner_tesoriq()

    if vista_activa == "📊 Dashboard y Partidas Pendientes":
        df_ejemplo_pendientes = pd.DataFrame({
            "Referencia": ["REF-1002", "REF-1003", "REF-1006", "REF-1009"],
            "Fecha": ["2026-11-11", "2026-11-12", "2026-11-15", "2026-11-16"],
            "Concepto": [
                "Cobro Distribuidora (Dif. Redondeo BCV)",
                "Comisión Mantenimiento + IGTF 3%",
                "Pago Proveedor en Tránsito",
                "Factura con Retención IVA 75% SENIAT"
            ],
            "Monto_Banco": [3200.50, 45.00, 0.00, 1160.00],
            "Monto_ERP": [3200.00, 0.00, 740.00, 1000.00],
            "Estado": [
                "Diferencia en Monto",
                "Pendiente en ERP (Solo Banco)",
                "Pendiente en Banco (Solo ERP)",
                "Diferencia por Retención"
            ]
        })
        renderizar_seccion_partidas_pendientes(df_ejemplo_pendientes, tasa_bcv=tasa_actual)
    else:
        st.markdown("""
        ### Cómo conectar este archivo a tu código de 2.036 líneas sin perder nada
        1. Guarda este archivo como **`tesoriq_modulos.py`** en la misma carpeta donde está tu `aplicacion.py` (o `conciliar.py`).
        2. Al principio de tu archivo de 2.036 líneas (justo después de `st.set_page_config(...)`), agrega estas **3 líneas**:
        ```python
        from tesoriq_modulos import aplicar_identidad_tesoriq, renderizar_logo_sidebar, renderizar_seccion_partidas_pendientes
        aplicar_identidad_tesoriq()
        ```
        3. Dentro de tu `with st.sidebar:` actual, coloca `renderizar_logo_sidebar()` para mostrar el logo del trébol en **T** (Opción 3: Vectorial Mate Suave).
        4. En la sección o pestaña donde muestras las diferencias o registros no conciliados, llama a:
        ```python
        renderizar_seccion_partidas_pendientes(tu_dataframe_no_conciliados, tasa_bcv=tu_variable_tasa)
        ```
        """)
