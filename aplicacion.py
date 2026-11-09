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
        or "debito" in row_str
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
    " fila, aplicando tasas BCV y generando el consolidado side-by-side de 4"
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
  modo_tasa = st.radio(
      "Fuente de Tasa de Cambio:",
      ["Automática (API BCV)", "Ingreso Manual"],
  )

  tasa_aplicada = 36.00
  archivo_tasas = st.file_uploader(
      "Subir Histórico de Tasas (Excel de Respaldo)", type=["xlsx", "csv"]
  )

  fecha_consulta = st.date_input(
      "Fecha para consultar tasa BCV:", datetime.today()
  )
  fecha_str = fecha_consulta.strftime("%Y-%m-%d")

  tasa_api = 0.0
  if modo_tasa == "Automática (API BCV)":
    with st.spinner("Consultando API BCV..."):
      tasa_api = obtener_tasa_bcv(fecha_str)

  if tasa_api > 0:
    st.success(f"Tasa BCV vía API ({fecha_str}): Bs. {tasa_api:,.2f}")
    tasa_aplicada = tasa_api
  else:
    tasa_encontrada = False
    if archivo_tasas is not None:
      try:
        df_t = (
            pd.read_excel(archivo_tasas)
            if not archivo_tasas.name.endswith(".csv")
            else pd.read_csv(archivo_tasas)
        )
        col_f = next(
            (c for c in df_t.columns if "fecha" in str(c).lower()),
            df_t.columns[0],
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

        df_t["Fecha_Filtro"] = pd.to_datetime(
            df_t[col_f], errors="coerce"
        ).dt.strftime("%Y-%m-%d")
        match = df_t[df_t["Fecha_Filtro"] == fecha_str]

        if not match.empty:
          tasa_aplicada = float(match.iloc[0][col_v])
          st.success(
              f"Tasa exacta del respaldo ({fecha_str}): Bs."
              f" {tasa_aplicada:,.2f}"
          )
          tasa_encontrada = True
        else:
          tasa_aplicada = float(df_t.iloc[-1][col_v])
          st.success(
              f"Tasa tomada del registro final del respaldo: Bs."
              f" {tasa_aplicada:,.2f}"
          )
          tasa_encontrada = True
      except Exception as e:
        pass

    if not tasa_encontrada:
      tasa_aplicada = st.number_input(
          "Tasa de respaldo / manual:", min_value=0.01, value=37.03, step=0.01
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

  # Detección de columnas clave
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

  col_monto_erp = next(
      (
          c
          for c in df_erp.columns
          if any(k in str(c).lower() for k in candidatos_monto)
      ),
      df_erp.columns[-1],
  )

  # Limpieza de montos
  df_banco["Monto_Clean"] = (
      df_banco[col_monto_banco].apply(limpiar_monto).round(2)
  )
  df_erp["Monto_Clean"] = (
      df_erp[col_monto_erp].apply(limpiar_monto).round(2)
  )

  # Asignación de tasa y conversión en divisa
  df_banco["Tasa_BCV"] = tasa_aplicada
  df_banco[f"Monto_{moneda_base}"] = (
      df_banco["Monto_Clean"] / df_banco["Tasa_BCV"]
  ).round(2)

  df_erp["Tasa_BCV"] = tasa_aplicada
  df_erp[f"Monto_{moneda_base}"] = (
      df_erp["Monto_Clean"] / df_erp["Tasa_BCV"]
  ).round(2)

  # MOTOR DE CONCILIACIÓN FILA POR FILA (Cruce exacto Bancos vs ERP)
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

  # Construcción del consolidado lado a lado (Side-by-Side) sin duplicar nombres
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
  total_reversos = df_banco[mask_reversos]["Monto_Clean"].sum()

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
        "📋 Consolidado de Conciliación Bancaria (Tasa Aplicada:"
        f" {tasa_aplicada:,.2f})"
    )
    st.markdown(
        "Detalle side-by-side: Extracto Bancario, Tasa, Conversión en Divisas y"
        " su respectivo Mayor Analítico ERP fila por fila."
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
              "Tasa BCV Aplicada",
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
              tasa_aplicada,
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
                f"DATOS DE CONCILIACIÓN BANCARIA:\n- Tasa BCV Aplicada: {tasa_aplicada}\n"
                f"- Total Movimientos: {total_movimientos}\n- Conciliados: {cant_conciliados}\n"
                f"- Pendientes por conciliar: {cant_pendientes}\n- Partidas Reversadas: {cant_reversos}\n"
                f"- Total Comisiones: {total_comisiones}\n\nDetalle de filas:\n{df_consolidado[[col_desc_banco, col_monto_banco, 'Estatus_Conciliacion']].head(20).to_string()}"
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
      doc.add_paragraph(f"Tasa BCV Aplicada: Bs. {tasa_aplicada:,.2f}")

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
      "👈 Por favor, carga el Extracto Bancario y el Mayor Analítico en la barra"
      " lateral para iniciar el cruce fila por fila."
  )