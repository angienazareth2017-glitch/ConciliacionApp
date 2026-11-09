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
# ESTILOS VISUALES AVANZADOS (INTERFAZ TIPO ERP)
# ---------------------------------------------------------
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }

        .main {
            background-color: #0b0f19;
            color: #f3f4f6;
        }

        [data-testid="stSidebar"] {
            background-color: #111827;
            border-right: 1px solid #1f2937;
        }

        div[data-testid="metric-container"] {
            background: linear-gradient(135deg, #1f2937 0%, #111827 100%);
            border: 1px solid #374151;
            padding: 16px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
            transition: transform 0.2s ease;
        }
        div[data-testid="metric-container"]:hover {
            transform: translateY(-2px);
            border-color: #4b5563;
        }
        div[data-testid="metric-container"] label {
            color: #9ca3af !important;
            font-size: 0.85rem !important;
            font-weight: 500;
        }
        div[data-testid="metric-container"] div[data-testid="stMetricValue"] {
            color: #ffffff !important;
            font-size: 1.6rem !important;
            font-weight: 700;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
            background-color: #111827;
            padding: 6px;
            border-radius: 8px;
            border: 1px solid #1f2937;
        }
        .stTabs [data-baseweb="tab"] {
            background-color: transparent;
            color: #9ca3af;
            border-radius: 6px;
            padding: 8px 16px;
            font-weight: 500;
        }
        .stTabs [aria-selected="true"] {
            background-color: #2563eb !important;
            color: #ffffff !important;
        }

        [data-testid="stDataFrame"] {
            border-radius: 8px;
            border: 1px solid #1f2937;
        }

        .stButton button {
            background-color: #2563eb;
            color: white;
            border-radius: 6px;
            font-weight: 600;
            border: none;
            padding: 0.5rem 1rem;
            transition: background-color 0.2s;
        }
        .stButton button:hover {
            background-color: #1d4ed8;
        }

        .stAlert {
            border-radius: 8px;
            border: none;
        }
    </style>
""",
    unsafe_allow_html=True,
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
def obtener_tasa_bcv_api(fecha_str):
  """Consulta la tasa oficial BCV histórica para la fecha dada mediante la API."""
  try:
    url = f"https://pydolarvenezuela.org/api/v1/dollar/history/{fecha_str}"
    response = requests.get(url, timeout=4)
    if response.status_code == 200:
      data = response.json()
      return data.get("bcv", {}).get("price", 0.0)
  except Exception:
    pass
  return 0.0


def cargar_diccionario_tasas_excel(archivo_tasas):
  """Carga el archivo Excel de respaldo y crea un diccionario de fechas para búsqueda rápida."""
  dict_tasas = {}
  ultima_tasa = 37.03
  try:
    df_t = (
        pd.read_excel(archivo_tasas)
        if not archivo_tasas.name.endswith(".csv")
        else pd.read_csv(archivo_tasas)
    )
    col_f = next(
        (c for c in df_t.columns if "fecha" in str(c).lower()), df_t.columns[0]
    )
    col_v = next(
        (
            c
            for c in df_t.columns
            if any(
                k in str(c).lower()
                for k in ["tasa", "precio", "bcv", "valor", "cambio"]
            )
        ),
        df_t.columns[1],
    )

    for _, row in df_t.iterrows():
      f_val = pd.to_datetime(row[col_f], errors="coerce")
      v_val = float(row[col_v])
      if not pd.isna(f_val) and v_val > 0:
        f_str = f_val.strftime("%Y-%m-%d")
        dict_tasas[f_str] = v_val
        ultima_tasa = v_val
  except Exception:
    pass
  return dict_tasas, ultima_tasa


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


def limpiar_tabla_financiera(archivo):
  """Limpia metadatos corporativos superiores y detecta la tabla real de transacciones."""
  if archivo.name.endswith(".csv"):
    df_raw = pd.read_csv(archivo, header=None)
  else:
    df_raw = pd.read_excel(archivo, header=None)

  header_row = 0
  for idx, row in df_raw.iterrows():
    row_str = " ".join([str(val) for val in row.values]).lower()
    if "fecha" in row_str and (
        "monto" in row_str
        or "importe" in row_str
        or "cargo" in row_str
        or "abono" in row_str
        or "referencia" in row_str
    ):
      header_row = idx
      break

  if archivo.name.endswith(".csv"):
    df = pd.read_csv(archivo, skiprows=header_row)
  else:
    df = pd.read_excel(archivo, skiprows=header_row)

  df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
  df = df.dropna(how="all")

  if len(df) > 0:
    col_fecha = next(
        (c for c in df.columns if "fecha" in str(c).lower()), df.columns[0]
    )
    df = df[df[col_fecha].notna()]
    df = df[
        ~df[col_fecha]
        .astype(str)
        .str.lower()
        .str.contains(
            "nombre|cuenta|mes|total|banco|rif|empresa|número", na=False
        )
    ]

  return df.reset_index(drop=True)


# ---------------------------------------------------------
# INTERFAZ PRINCIPAL
# ---------------------------------------------------------
st.title("💼 Sistema Automatizado de Conciliación y Tesorería")
st.markdown(
    "Cruza extractos bancarios contra mayores contables (SAP/ERP) fila por"
    " fila para el mes completo, aplicando tasas BCV automáticas (API con"
    " respaldo en Excel), evaluando Cargo y Abono, y generando reporte de 4"
    " pestañas."
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
  archivo_tasas = st.file_uploader(
      "Subir Archivo Excel de Tasas (Respaldo)", type=["xlsx", "csv"]
  )
  st.info(
      "💡 El sistema intentará consultar la API BCV por cada fecha. Si no hay"
      " conexión, tomará la tasa del Excel de respaldo adjunto."
  )

if "mensajes_chat" not in st.session_state:
  st.session_state.mensajes_chat = []

if archivo_banco and archivo_erp:
  try:
    df_banco = limpiar_tabla_financiera(archivo_banco)
    df_erp = limpiar_tabla_financiera(archivo_erp)
  except Exception as e:
    st.error(f"Error al leer y limpiar los archivos subidos: {e}")
    st.stop()

  # Cargar diccionario de tasas del archivo si existe
  dict_tasas_respaldo = {}
  tasa_default = 37.03
  if archivo_tasas is not None:
    dict_tasas_respaldo, tasa_default = cargar_diccionario_tasas_excel(
        archivo_tasas
    )

  # Detección de columnas de Fecha y Monto en Banco
  col_fecha_banco = next(
      (c for c in df_banco.columns if "fecha" in str(c).lower()),
      df_banco.columns[0],
  )
  candidatos_monto = [
      "monto",
      "importe",
      "debito",
      "crédito",
      "credito",
      "valor",
      "haber",
      "debe",
  ]
  col_monto_banco = next(
      (
          c
          for c in df_banco.columns
          if any(k in str(c).lower() for k in candidatos_monto)
      ),
      df_banco.columns[-1],
  )
  col_desc_banco = next(
      (
          c
          for c in df_banco.columns
          if any(
              k in str(c).lower()
              for k in ["desc", "concepto", "referencia", "detalle"]
          )
      ),
      df_banco.columns[0],
  )

  # Detección precisa de columnas Cargo y Abono en el ERP
  col_cargo_erp = next(
      (
          c
          for c in df_erp.columns
          if any(k in str(c).lower() for k in ["cargo", "debe", "débito"])
      ),
      None,
  )
  col_abono_erp = next(
      (
          c
          for c in df_erp.columns
          if any(k in str(c).lower() for k in ["abono", "haber", "crédito"])
      ),
      None,
  )

  if col_cargo_erp and col_abono_erp:
    df_erp["Monto_Clean"] = (
        df_erp[col_cargo_erp].apply(limpiar_monto)
        - df_erp[col_abono_erp].apply(limpiar_monto)
    ).round(2)
  else:
    col_monto_erp = next(
        (
            c
            for c in df_erp.columns
            if any(k in str(c).lower() for k in candidatos_monto)
        ),
        df_erp.columns[-1],
    )
    df_erp["Monto_Clean"] = (
        df_erp[col_monto_erp].apply(limpiar_monto).round(2)
    )

  # Limpieza de montos del banco
  df_banco["Monto_Clean"] = (
      df_banco[col_monto_banco].apply(limpiar_monto).round(2)
  )


  # Función Dual Automática: API BCV con respaldo de Excel
  def obtener_tasa_dual(fecha_val):
    f_dt = pd.to_datetime(fecha_val, errors="coerce")
    if pd.isna(f_dt):
      return tasa_default
    f_str = f_dt.strftime("%Y-%m-%d")

    # 1. Intentar API BCV en línea
    t_api = obtener_tasa_bcv_api(f_str)
    if t_api > 0:
      return t_api

    # 2. Si falla la API, buscar en el Excel de respaldo
    if f_str in dict_tasas_respaldo:
      return dict_tasas_respaldo[f_str]

    # 3. Fallback final si no está en ninguno de los dos
    return tasa_default


  # Asignación de tasa y conversión fila por fila según la fecha del movimiento
  tasas_banco = [obtener_tasa_dual(f) for f in df_banco[col_fecha_banco]]
  df_banco["Tasa_BCV"] = tasas_banco
  df_banco[f"Monto_{moneda_base}"] = (
      df_banco["Monto_Clean"] / df_banco["Tasa_BCV"]
  ).round(2)

  col_fecha_erp = next(
      (c for c in df_erp.columns if "fecha" in str(c).lower()), None
  )
  if col_fecha_erp:
    tasas_erp = [obtener_tasa_dual(f) for f in df_erp[col_fecha_erp]]
    df_erp["Tasa_BCV"] = tasas_erp
  else:
    df_erp["Tasa_BCV"] = tasas_banco[0] if len(tasas_banco) > 0 else tasa_default

  df_erp[f"Monto_{moneda_base}"] = (
      df_erp["Monto_Clean"] / df_erp["Tasa_BCV"]
  ).round(2)

  # MOTOR DE CONCILIACIÓN FILA POR FILA
  erp_usados = set()
  estatus_banco_lista = []
  match_erp_indices = []

  for idx_b, row_b in df_banco.iterrows():
    monto_b = row_b["Monto_Clean"]
    desc_b = str(row_b[col_desc_banco]).lower()

    if any(k in desc_b for k in ["anul", "cancel", "revers"]):
      estatus_banco_lista.append("Pendiente por contabilizar (Reversado)")
      match_erp_indices.append(None)
    else:
      encontrado = None
      for idx_e, row_e in df_erp.iterrows():
        if idx_e in erp_usados:
          continue
        monto_e = row_e["Monto_Clean"]
        if monto_b == monto_e or monto_b == -monto_e:
          encontrado = idx_e
          break

      if encontrado is not None:
        erp_usados.add(encontrado)
        estatus_banco_lista.append("Conciliado")
        match_erp_indices.append(encontrado)
      else:
        estatus_banco_lista.append("Pendiente por conciliar")
        match_erp_indices.append(None)

  df_banco["Estatus_Conciliacion"] = estatus_banco_lista

  # Construcción del consolidado lado a lado (Side-by-Side)
  df_erp_temp = df_erp.copy()
  df_erp_temp.columns = [f"ERP_{c}" for c in df_erp_temp.columns]

  erp_data_row_list = []
  for idx_e in match_erp_indices:
    if idx_e is not None and idx_e < len(df_erp_temp):
      erp_data_row_list.append(df_erp_temp.iloc[idx_e].to_dict())
    else:
      empty_dict = {c: "" for c in df_erp_temp.columns}
      erp_data_row_list.append(empty_dict)

  df_erp_matched = pd.DataFrame(erp_data_row_list)
  df_consolidado = pd.concat(
      [df_banco.reset_index(drop=True), df_erp_matched.reset_index(drop=True)],
      axis=1,
  )

  # Métricas generales
  total_movimientos = len(df_banco)
  total_ingresos = df_banco[df_banco["Monto_Clean"] > 0][
      "Monto_Clean"
  ].sum()
  total_egresos = df_banco[df_banco["Monto_Clean"] < 0][
      "Monto_Clean"
  ].sum()
  total_general = df_banco["Monto_Clean"].sum()

  cant_conciliados = (
      df_banco["Estatus_Conciliacion"] == "Conciliado"
  ).sum()
  cant_pendientes = (
      df_banco["Estatus_Conciliacion"] == "Pendiente por conciliar"
  ).sum()
  mask_reversos = df_banco["Estatus_Conciliacion"].str.contains("Reversado")
  cant_reversos = mask_reversos.sum()

  mask_comisiones = df_banco[col_desc_banco].astype(str).str.contains(
      "comis|comisión|com", case=False, na=False
  )
  total_comisiones = df_banco[mask_comisiones]["Monto_Clean"].sum()

  tab1, tab2, tab3 = st.tabs(
      [
          "📊 Conciliación Fila por Fila y Consolidado",
          "💬 Asistente IA Financiero",
          "📥 Exportar Reportes",
      ]
  )

  with tab1:
    st.subheader("Resumen General de Tesorería y Cruce Fila por Fila")

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total Movimientos", total_movimientos)
    col2.metric("Conciliados", cant_conciliados)
    col3.metric("Pendientes Conciliar", cant_pendientes)
    col4.metric("Partidas Reversadas", cant_reversos)
    col5.metric("Comisiones Totales", f"{abs(total_comisiones):,.2f}")

    st.markdown("---")
    st.subheader(
        "📋 Consolidado de Conciliación Bancaria (Mes Completo - Tasas API /"
        " Respaldo Excel)"
    )
    st.markdown(
        "Detalle side-by-side: Extracto Bancario, Tasa BCV Dinámica, Conversión"
        " en Divisas y Mayor Analítico ERP."
    )

    st.dataframe(df_consolidado, use_container_width=True)

    # CREACIÓN DEL EXCEL CON LAS 4 PESTAÑAS EXACTAS SOLICITADAS
    buffer_excel = BytesIO()
    with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
      df_consolidado.to_excel(
          writer, index=False, sheet_name="Consolidado_Conciliacion"
      )
      df_banco.to_excel(writer, index=False, sheet_name="Extracto_Banco")
      df_erp.to_excel(writer, index=False, sheet_name="Mayor_ERP")

      df_resumen = pd.DataFrame({
          "Metrica": [
              "Total Movimientos Banco",
              "Total General Banco",
              "Total Ingresos Banco",
              "Total Egresos Banco",
              "Partidas Conciliadas",
              "Partidas Pendientes por Conciliar",
              "Partidas Reversadas",
              "Total Comisiones",
          ],
          "Valor": [
              total_movimientos,
              total_general,
              total_ingresos,
              total_egresos,
              cant_conciliados,
              cant_pendientes,
              cant_reversos,
              total_comisiones,
          ],
      })
      df_resumen.to_excel(writer, index=False, sheet_name="Resumen_Gerencial")

    buffer_excel.seek(0)

    st.download_button(
        label="⬇️ Descargar Consolidado Completo (4 Pestañas en Excel .xlsx)",
        data=buffer_excel,
        file_name=f"Consolidado_Conciliacion_4_Pestanas_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  with tab2:
    st.subheader("💬 Asistente IA Financiero (Claude vía OpenRouter)")
    st.markdown(
        "Haz preguntas sobre las diferencias, comisiones o el estatus de las"
        " partidas."
    )

    if not ia_disponible or API_KEY_OR == "tu-api-key-de-openrouter-aqui":
      st.warning(
          "⚠️ Configura tu `OPENROUTER_API_KEY` en los secretos de Streamlit"
          " para activar el chat inteligente."
      )

    for msg in st.session_state.mensajes_chat:
      with st.chat_message(msg["role"]):
        st.write(msg["content"])

    if pregunta := st.chat_input(
        "Ej: ¿Cuáles partidas quedaron pendientes por conciliar?"
    ):
      st.session_state.mensajes_chat.append(
          {"role": "user", "content": pregunta}
      )
      with st.chat_message("user"):
        st.write(pregunta)

      if ia_disponible and API_KEY_OR != "tu-api-key-de-openrouter-aqui":
        with st.spinner("Analizando conciliación con Claude..."):
          try:
            resumen_contextual = (
                f"DATOS DE CONCILIACIÓN BANCARIA (MES COMPLETO):\n- Total"
                f" Movimientos: {total_movimientos}\n- Conciliados:"
                f" {cant_conciliados}\n- Pendientes por conciliar:"
                f" {cant_pendientes}\n- Partidas Reversadas:"
                f" {cant_reversos}\n- Total Comisiones: {total_comisiones}\n\nDetalle"
                f" de filas:\n{df_consolidado[[col_desc_banco, col_monto_banco, 'Estatus_Conciliacion']].head(20).to_string()}"
            )

            prompt_sistema = (
                "Eres un auditor contable experto en tesorería corporativa en"
                " Venezuela. Responde de forma directa, analítica y profesional"
                " utilizando estrictamente los datos proporcionados:\n\n"
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
        respuesta_ia = "Simulación de respuesta de IA (Configure su API Key)."

      st.session_state.mensajes_chat.append(
          {"role": "assistant", "content": respuesta_ia}
      )
      with st.chat_message("assistant"):
        st.write(respuesta_ia)

  with tab3:
    st.subheader("📥 Exportar Reportes")
    st.markdown("Descarga el reporte gerencial completo en formato Word.")

    if st.button("📄 Generar Reporte Word"):
      doc = Document()
      doc.add_heading("Reporte Gerencial de Conciliación Bancaria", 0)
      doc.add_paragraph(
          f"Fecha de emisión: {datetime.now().strftime('%Y-%m-%d')}"
      )
      doc.add_paragraph(
          "Periodo analizado: Mes completo según transacciones cargadas"
      )

      doc.add_heading("1. Resumen de Estatus de Conciliación", level=1)
      doc.add_paragraph(f"• Total de Movimientos Analizados: {total_movimientos}")
      doc.add_paragraph(f"• Partidas en Estatus 'Conciliado': {cant_conciliados}")
      doc.add_paragraph(
          f"• Partidas en Estatus 'Pendiente por conciliar': {cant_pendientes}"
      )
      doc.add_paragraph(
          "• Partidas en Estatus 'Pendiente por contabilizar (Reversado)':"
          f" {cant_reversos}"
      )
      doc.add_paragraph(
          f"• Total Comisiones Bancarias: {abs(total_comisiones):,.2f}"
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
      "👈 Por favor, carga el Extracto Bancario, el Mayor Analítico y el Archivo"
      " de Tasas de Respaldo en la barra lateral para iniciar la conciliación"
      " del mes completo."
  )