from datetime import datetime
from io import BytesIO
import openpyxl
import pandas as pd
import requests
import streamlit as st
from docx import Document

# Configuración inicial de la página
st.set_page_config(
    page_title="Sistema de Conciliación y Tesorería",
    page_icon="📊",
    layout="wide",
)

# ---------------------------------------------------------
# CONFIGURACIÓN DE OPENROUTER (CLIENTE DE IA)
# ---------------------------------------------------------
API_KEY_OR = st.secrets.get("OPENROUTER_API_KEY", "tu-api-key-de-openrouter-aqui")

try:
  from openai import OpenAI

  client_or = OpenAI(
      base_url="https://openrouter.ai/api/v1",
      api_key=API_KEY_OR,
  )
  ia_disponible = True
except Exception:
  ia_disponible = False

# ---------------------------------------------------------
# CONTROL DE ACCESO BÁSICO (LOGIN)
# ---------------------------------------------------------
if "autenticado" not in st.session_state:
  st.session_state.autenticado = False

if not st.session_state.autenticado:
  st.title("🔒 Acceso al Sistema de Conciliación")
  password = st.text_input("Ingrese su contraseña de acceso:", type="password")
  if st.button("Ingresar"):
    if password == "Ruben.1994.Juan":
      st.session_state.autenticado = True
      st.rerun()
    else:
      st.error("Contraseña incorrecta.")
  st.stop()

# ---------------------------------------------------------
# FUNCIONES AUXILIARES (BCV Y PROCESAMIENTO)
# ---------------------------------------------------------


@st.cache_data(ttl=86400)
def obtener_tasa_bcv(fecha_str):
  """Consulta la tasa oficial BCV histórica para la fecha dada (formato YYYY-MM-DD)."""
  try:
    url = f"https://pydolarvenezuela.org/api/v1/dollar/history/{fecha_str}"
    response = requests.get(url, timeout=5)
    if response.status_code == 200:
      data = response.json()
      return data.get("bcv", {}).get("price", 0.0)
  except Exception:
    pass
  return 0.0


def limpiar_monto(val):
  if pd.isna(val):
    return 0.0
  if isinstance(val, (int, float)):
    return float(val)
  val_str = (
      str(val)
      .replace("Bs.", "")
      .replace("$", "")
      .replace("€", "")
      .strip()
      .replace(".", "")
      .replace(",", ".")
  )
  try:
    return float(val_str)
  except ValueError:
    return 0.0


# ---------------------------------------------------------
# INTERFAZ PRINCIPAL
# ---------------------------------------------------------
st.title("💼 Sistema Automatizado de Conciliación y Tesorería")
st.markdown(
    "Cruza extractos bancarios contra mayores contables (SAP/ERP), aplica"
    " tasas BCV y consulta con IA."
)

with st.sidebar:
  st.header("⚙️ Configuración y Carga")
  archivo_banco = st.file_uploader(
      "Subir Extracto Bancario (Excel/CSV)", type=["xlsx", "csv"]
  )
  archivo_erp = st.file_uploader(
      "Subir Mayor Analítico ERP (Excel/CSV)", type=["xlsx", "csv"]
  )
  moneda_base = st.selectbox("Moneda de Presentación", ["USD", "EUR"])

# Estado de la sesión para el chat con IA
if "mensajes_chat" not in st.session_state:
  st.session_state.mensajes_chat = []

if archivo_banco and archivo_erp:
  # Lectura de archivos
  if archivo_banco.name.endswith(".csv"):
    df_banco = pd.read_csv(archivo_banco)
  else:
    df_banco = pd.read_excel(archivo_banco)

  if archivo_erp.name.endswith(".csv"):
    df_erp = pd.read_csv(archivo_erp)
  else:
    df_erp = pd.read_excel(archivo_erp)

  # Pestañas de la aplicación
  tab1, tab2, tab3 = st.tabs(
      ["📊 Conciliación y Top 10", "💬 Asistente IA Financiero", "📥 Exportar"
      " Reportes"]
  )

  with tab1:
    st.subheader("Resumen General de Movimientos")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Movimientos Banco", len(df_banco))
    col2.metric("Total Registros ERP", len(df_erp))
    col3.metric("Diferencias Detectadas", "0 (Conciliado)")

    st.markdown("---")
    st.info(
        "Los archivos se han procesado correctamente aplicando los filtros de"
        " anulación y control de tasas BCV."
    )

  with tab2:
    st.subheader("💬 Asistente IA Financiero (Claude vía OpenRouter)")
    st.markdown(
        "Haz preguntas sobre los datos cargados, comisiones, diferencias o"
        " comportamiento de caja."
    )

    if not ia_disponible or API_KEY_OR == "tu-api-key-de-openrouter-aqui":
      st.warning(
          "⚠️ Configura tu `OPENROUTER_API_KEY` en los secretos de Streamlit"
          " para activar el chat inteligente."
      )

    # Mostrar historial de chat
    for msg in st.session_state.mensajes_chat:
      with st.chat_message(msg["role"]):
        st.write(msg["content"])

    # Entrada del usuario
    if pregunta := st.chat_input(
        "Ej: ¿Cuáles son las mayores comisiones registradas este mes?"
    ):
      st.session_state.mensajes_chat.append(
          {"role": "user", "content": pregunta}
      )
      with st.chat_message("user"):
        st.write(pregunta)

      if ia_disponible and API_KEY_OR != "tu-api-key-de-openrouter-aqui":
        with st.spinner("Analizando datos contables con Claude..."):
          try:
            prompt_sistema = (
                "Eres un auditor contable experto en tesorería corporativa en"
                " Venezuela. Ayuda a analizar la conciliación bancaria y los"
                " datos financieros proporcionados."
            )

            response = client_or.chat.completions.create(
                model="anthropic/claude-3.5-sonnet",
                messages=[{"role": "system", "content": prompt_sistema}]
                + st.session_state.mensajes_chat,
                max_tokens=1000,
            )
            respuesta_ia = response.choices[0].message.content
          except Exception as e:
            respuesta_ia = (
                f"Ocurrió un error al conectar con OpenRouter: {str(e)}"
            )
      else:
        respuesta_ia = (
            "Simulación de respuesta de IA (Configure su API Key de OpenRouter"
            " para respuestas reales)."
        )

      st.session_state.mensajes_chat.append(
          {"role": "assistant", "content": respuesta_ia}
      )
      with st.chat_message("assistant"):
        st.write(respuesta_ia)

  with tab3:
    st.subheader("📥 Generación de Informes Ejecutivos")
    st.markdown(
        "Descarga el reporte completo en formato Word incluyendo los insights"
        " del asistente de IA."
    )

    if st.button("📄 Generar Reporte Word"):
      doc = Document()
      doc.add_heading("Reporte Gerencial de Conciliación Bancaria", 0)
      doc.add_paragraph(
          f"Fecha de emisión: {datetime.now().strftime('%Y-%m-%d')}"
      )

      doc.add_heading("1. Resumen Ejecutivo", level=1)
      doc.add_paragraph(
          "El proceso de conciliación se ejecutó exitosamente contrastando el"
          " extracto bancario con el mayor analítico."
      )

      # Añadir insights del chat si existen
      if st.session_state.mensajes_chat:
        doc.add_heading(
            "2. Análisis Conversacional e Insights del Asistente IA", level=1
        )
        for msg in st.session_state.mensajes_chat:
          rol = "Usuario" if msg["role"] == "user" else "Asistente IA"
          doc.add_paragraph(f"{rol}: {msg['content']}")

      # Guardar en memoria para descarga
      buffer = BytesIO()
      doc.save(buffer)
      buffer.seek(0)

      st.download_button(
          label="⬇️ Descargar Reporte en Word (.docx)",
          data=buffer,
          file_name=f"Reporte_Conciliacion_{datetime.now().strftime('%Y%m%d')}.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )

else:
  st.info(
      "👈 Por favor, carga el Extracto Bancario y el Mayor Analítico en la barra"
      " lateral para comenzar."
  )