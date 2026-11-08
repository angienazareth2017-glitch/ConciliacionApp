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

  st.markdown("---")
  st.header("💱 Tasas y Conversión BCV")
  modo_tasa = st.radio(
      "Fuente de Tasa de Cambio:",
      ["Automática (API BCV)", "Ingreso Manual"],
  )

  tasa_manual = 1.0
  if modo_tasa == "Ingreso Manual":
    tasa_manual = st.number_input(
        "Ingrese la Tasa BCV del día:", min_value=0.01, value=36.00, step=0.1
    )
  else:
    fecha_consulta = st.date_input(
        "Fecha para consultar tasa BCV:", datetime.today()
    )
    fecha_str = fecha_consulta.strftime("%Y-%m-%d")
    tasa_api = obtener_tasa_bcv(fecha_str)
    if tasa_api > 0:
      st.success(f"Tasa BCV ({fecha_str}): Bs. {tasa_api:,.2f}")
      tasa_manual = tasa_api
    else:
      st.warning(
          "No se pudo conectar con la API para esa fecha. Ingrese tasa"
          " manual:"
      )
      tasa_manual = st.number_input("Tasa de respaldo:", value=36.00, step=0.1)

# Estado de la sesión para el chat con IA
if "mensajes_chat" not in st.session_state:
  st.session_state.mensajes_chat = []

if archivo_banco and archivo_erp:
  # Lectura robusta de archivos
  try:
    if archivo_banco.name.endswith(".csv"):
      df_banco = pd.read_csv(archivo_banco)
    else:
      df_banco = pd.read_excel(archivo_banco)

    if archivo_erp.name.endswith(".csv"):
      df_erp = pd.read_csv(archivo_erp)
    else:
      df_erp = pd.read_excel(archivo_erp)
  except Exception as e:
    st.error(f"Error al leer los archivos subidos: {e}")
    st.stop()

  # Procesamiento de montos y detección de columnas clave
  col_monto_banco = next(
      (
          c
          for c in df_banco.columns
          if any(
              k in c.lower() for k in ["monto", "importe", "debito", "credito"]
          )
      ),
      df_banco.columns[-1],
  )
  col_desc_banco = next(
      (
          c
          for c in df_banco.columns
          if any(k in c.lower() for k in ["desc", "concepto", "referencia"])
      ),
      df_banco.columns[0],
  )

  df_banco["Monto_LimPIO"] = df_banco[col_monto_banco].apply(limpiar_monto)
  df_banco[f"Monto_{moneda_base}"] = df_banco["Monto_LimPIO"] / (
      tasa_manual if tasa_manual > 0 else 1
  )

  # Cálculos de métricas gerenciales
  total_movimientos = len(df_banco)
  total_ingresos = df_banco[df_banco["Monto_LimPIO"] > 0][
      "Monto_LimPIO"
  ].sum()
  total_egresos = df_banco[df_banco["Monto_LimPIO"] < 0][
      "Monto_LimPIO"
  ].sum()
  total_general = df_banco["Monto_LimPIO"].sum()

  # Filtros de comisiones y reversos
  mask_comisiones = df_banco[col_desc_banco].astype(str).str.contains(
      "comis|comisión|com", case=False, na=False
  )
  total_comisiones = df_banco[mask_comisiones]["Monto_LimPIO"].sum()

  mask_reversos = (
      df_banco[col_desc_banco]
      .astype(str)
      .str.contains("anul|cancel|revers", case=False, na=False)
  )
  total_reversos = df_banco[mask_reversos]["Monto_LimPIO"].sum()
  cant_reversos = mask_reversos.sum()

  # Pestañas de la aplicación
  tab1, tab2, tab3 = st.tabs(
      [
          "📊 Conciliación, Métricas y Consolidado",
          "💬 Asistente IA Financiero",
          "📥 Exportar Reportes",
      ]
  )

  with tab1:
    st.subheader("Resumen General de Movimientos y Tesorería")

    # Métricas detalladas
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total Movimientos", total_movimientos)
    col2.metric("Total General", f"{total_general:,.2f}")
    col3.metric("Total Ingresos", f"{total_ingresos:,.2f}")
    col4.metric("Total Egresos", f"{abs(total_egresos):,.2f}")
    col5.metric(
        "Partidas Reversadas", f"{cant_reversos} ({total_reversos:,.2f})"
    )

    st.markdown("---")
    st.subheader("Detalle de Comisiones Detectadas")
    st.metric("Monto Total en Comisiones", f"{abs(total_comisiones):,.2f}")

    st.markdown("---")
    st.subheader(
        "📋 Consolidado de Conciliación Procesado (Tasa Aplicada:"
        f" {tasa_manual:,.2f})"
    )
    st.dataframe(df_banco, use_container_width=True)

    # Botón para descargar el consolidado en Excel directamente desde la pestaña 1
    buffer_excel = BytesIO()
    with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
      df_banco.to_excel(writer, index=False, sheet_name="Consolidado_Conciliacion")
    buffer_excel.seek(0)

    st.download_button(
        label="⬇️ Descargar Consolidado de Conciliación en Excel (.xlsx)",
        data=buffer_excel,
        file_name=f"Consolidado_Conciliacion_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  with tab2:
    st.subheader("💬 Asistente IA Financiero (Claude vía OpenRouter)")
    st.markdown(
        "Haz preguntas directas sobre los datos cargados. La IA ya tiene"
        " acceso automático al resumen de tu tesorería y extractos."
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
        "Ej: ¿Cuáles son las mayores comisiones y qué partidas fueron"
        " reversadas?"
    ):
      st.session_state.mensajes_chat.append(
          {"role": "user", "content": pregunta}
      )
      with st.chat_message("user"):
        st.write(pregunta)

      if ia_disponible and API_KEY_OR != "tu-api-key-de-openrouter-aqui":
        with st.spinner("Analizando datos contables con Claude..."):
          try:
            resumen_contextual = (
                f"DATOS FINANCIEROS ACTUALES:\n- Tasa BCV Aplicada: {tasa_manual}\n"
                f"- Total Movimientos: {total_movimientos}\n- Total General: {total_general}\n"
                f"- Total Ingresos: {total_ingresos}\n- Total Egresos: {total_egresos}\n"
                f"- Total Comisiones: {total_comisiones}\n"
                f"- Partidas Reversadas / Anuladas: {cant_reversos} (Suma: {total_reversos})\n"
                f"Muestra de transacciones recientes:\n{df_banco[[col_desc_banco, col_monto_banco]].head(15).to_string()}"
            )

            prompt_sistema = (
                "Eres un auditor contable experto en tesorería corporativa en"
                " Venezuela. Responde de forma directa, analítica y profesional"
                " a las preguntas del usuario utilizando los datos"
                " proporcionados a continuación:\n\n"
                + resumen_contextual
            )

            response = client_or.chat.completions.create(
                model="anthropic/claude-haiku-4.5",
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
        "Descarga el reporte completo en formato Word incluyendo las métricas,"
        " la tasa BCV aplicada y los insights del asistente de IA."
    )

    if st.button("📄 Generar Reporte Word"):
      doc = Document()
      doc.add_heading("Reporte Gerencial de Conciliación Bancaria", 0)
      doc.add_paragraph(
          f"Fecha de emisión: {datetime.now().strftime('%Y-%m-%d')}"
      )
      doc.add_paragraph(f"Tasa BCV Aplicada: Bs. {tasa_manual:,.2f}")

      doc.add_heading("1. Resumen Ejecutivo y Métricas", level=1)
      doc.add_paragraph(f"• Total de Movimientos: {total_movimientos}")
      doc.add_paragraph(f"• Total General: {total_general:,.2f}")
      doc.add_paragraph(f"• Total Ingresos: {total_ingresos:,.2f}")
      doc.add_paragraph(f"• Total Egresos: {abs(total_egresos):,.2f}")
      doc.add_paragraph(f"• Total Comisiones: {abs(total_comisiones):,.2f}")
      doc.add_paragraph(
          f"• Partidas Reversadas: {cant_reversos} (Monto:"
          f" {total_reversos:,.2f})"
      )

      if st.session_state.mensajes_chat:
        doc.add_heading(
            "2. Análisis Conversacional e Insights del Asistente IA", level=1
        )
        for msg in st.session_state.mensajes_chat:
          rol = "Usuario" if msg["role"] == "user" else "Asistente IA"
          doc.add_paragraph(f"{rol}: {msg['content']}")

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