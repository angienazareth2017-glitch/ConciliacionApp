from datetime import datetime
from io import BytesIO
import openpyxl
import pandas as pd
import requests
import streamlit as st
from docx import Document

st.set_page_config(
    page_title="Sistema de Conciliación y Tesorería",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
        .main { background-color: #0b0f19; color: #f3f4f6; }
        [data-testid="stSidebar"] { background-color: #111827; border-right: 1px solid #1f2937; }
        div[data-testid="metric-container"] {
            background: linear-gradient(135deg, #1f2937 0%, #111827 100%);
            border: 1px solid #374151; padding: 16px; border-radius: 10px;
        }
        div[data-testid="metric-container"] label { color: #9ca3af !important; font-size: 0.85rem !important; }
        div[data-testid="metric-container"] div[data-testid="stMetricValue"] { color: #ffffff !important; font-size: 1.6rem !important; }
    </style>
""",
    unsafe_allow_html=True,
)

API_KEY_OR = st.secrets.get("OPENROUTER_API_KEY", "")

try:
  from openai import OpenAI

  client_or = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=API_KEY_OR)
  ia_disponible = True
except Exception:
  ia_disponible = False

password_correcta = st.secrets.get("PASSWORD_ACCESO", "Ruben.1994.Juan")

if "autenticado" not in st.session_state:
  st.session_state.autenticado = False

if not st.session_state.autenticado:
  st.title("🔒 Acceso al Sistema de Conciliación")
  password = st.text_input("Ingrese su contraseña de acceso:", type="password")
  if st.button("Ingresar"):
    if password == password_correcta:
      st.session_state.autenticado = True
      st.rerun()
    else:
      st.error("Contraseña incorrecta.")
  st.stop()

if "mensajes_chat" not in st.session_state:
  st.session_state.mensajes_chat = []


@st.cache_data(ttl=86400)
def obtener_tasa_bcv_api(fecha_str):
  try:
    url = f"https://pydolarvenezuela.org/api/v1/dollar/history/{fecha_str}"
    response = requests.get(url, timeout=4)
    if response.status_code == 200:
      return response.json().get("bcv", {}).get("price", 0.0)
  except Exception:
    pass
  return 0.0


def cargar_diccionario_tasas_excel(archivo_tasas):
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
        dict_tasas[f_val.strftime("%Y-%m-%d")] = v_val
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


def limpiar_tabla_financiera(archivo, banco):
  df_raw = (
      pd.read_csv(archivo, header=None)
      if archivo.name.endswith(".csv")
      else pd.read_excel(archivo, header=None)
  )
  header_row = 0
  for idx, row in df_raw.iterrows():
    row_str = " ".join([str(val) for val in row.values]).lower()
    if "fecha" in row_str:
      header_row = idx
      break
  df = (
      pd.read_csv(archivo, skiprows=header_row)
      if archivo.name.endswith(".csv")
      else pd.read_excel(archivo, skiprows=header_row)
  )
  df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
  df = df.dropna(how="all")
  if len(df) > 0:
    col_fecha = next(
        (c for c in df.columns if "fecha" in str(c).lower()), df.columns[0]
    )
    df = df[df[col_fecha].notna()]
  return df.reset_index(drop=True)


st.title("💼 Sistema Automatizado de Conciliación y Tesorería")
st.markdown(
    "Cruza extractos bancarios contra mayores contables (SAP/ERP) en una sola"
    " fila horizontal por transacción, aplicando tasas BCV automáticas."
)

with st.sidebar:
  st.header("⚙️ Configuración y Carga")
  banco_seleccionado = st.selectbox(
      "Banco a Conciliar",
      [
          "Banco Mercantil",
          "Banesco",
          "BBVA Provincial",
          "Banco de Venezuela",
          "Otro / Genérico",
      ],
  )
  archivo_banco = st.file_uploader(
      "Subir Extracto Bancario (Excel/CSV)", type=["xlsx", "csv"]
  )
  archivo_erp = st.file_uploader(
      "Subir Mayor Analítico ERP (Excel/CSV)", type=["xlsx", "csv"]
  )
  moneda_base = st.selectbox("Moneda de Presentación", ["USD", "EUR"])
  archivo_tasas = st.file_uploader(
      "Subir Archivo Excel de Tasas (Respaldo)", type=["xlsx", "csv"]
  )

if archivo_banco and archivo_erp:
  try:
    df_banco = limpiar_tabla_financiera(archivo_banco, banco_seleccionado)
    df_erp = limpiar_tabla_financiera(archivo_erp, "Genérico")
  except Exception as e:
    st.error(f"Error al leer y limpiar los archivos: {e}")
    st.stop()

  dict_tasas_respaldo = {}
  tasa_default = 37.03
  if archivo_tasas is not None:
    dict_tasas_respaldo, tasa_default = cargar_diccionario_tasas_excel(
        archivo_tasas
    )

  col_fecha_banco = next(
      (c for c in df_banco.columns if "fecha" in str(c).lower()),
      df_banco.columns[0],
  )
  col_monto_banco = next(
      (
          c
          for c in df_banco.columns
          if any(
              k in str(c).lower()
              for k in ["monto", "importe", "debito", "haber", "crédito"]
          )
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
  col_cta_contra = next(
      (
          c
          for c in df_erp.columns
          if "contrapartida" in str(c).lower()
          and "nombre" not in str(c).lower()
      ),
      None,
  )
  col_nom_contra = next(
      (
          c
          for c in df_erp.columns
          if "contrapartida" in str(c).lower() and "nombre" in str(c).lower()
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
            if any(
                k in str(c).lower()
                for k in ["monto", "importe", "debito", "haber"]
            )
        ),
        df_erp.columns[-1],
    )
    df_erp["Monto_Clean"] = (
        df_erp[col_monto_erp].apply(limpiar_monto).round(2)
    )

  patron_anulacion = "anul|cancelar entrada|cancel|revers|nula|reverso"
  candidatos_desc_erp_full = [
      c
      for c in df_erp.columns
      if any(
          k in str(c).lower() for k in ["desc", "concepto", "detalle", "nombre"]
      )
  ]
  cant_reversos = 0
  if candidatos_desc_erp_full:
    col_eval_erp = candidatos_desc_erp_full[0]
    mask_reversos_erp = (
        df_erp[col_eval_erp]
        .astype(str)
        .str.lower()
        .str.contains(patron_anulacion, na=False)
    )
    cant_reversos = mask_reversos_erp.sum()

  df_erp = df_erp[df_erp["Monto_Clean"] != 0.0].copy()
  if candidatos_desc_erp_full:
    df_erp = df_erp[~mask_reversos_erp].copy()

  df_banco["Monto_Clean"] = (
      df_banco[col_monto_banco].apply(limpiar_monto).round(2)
  )


  def obtener_tasa_dual(fecha_val):
    f_dt = pd.to_datetime(fecha_val, errors="coerce")
    if pd.isna(f_dt):
      return tasa_default
    f_str = f_dt.strftime("%Y-%m-%d")
    t_api = obtener_tasa_bcv_api(f_str)
    if t_api > 0:
      return t_api
    return dict_tasas_respaldo.get(f_str, tasa_default)


  df_banco["Tasa_BCV"] = [
      obtener_tasa_dual(f) for f in df_banco[col_fecha_banco]
  ]
  df_banco[f"Monto_{moneda_base}"] = (
      df_banco["Monto_Clean"] / df_banco["Tasa_BCV"]
  ).round(2)

  col_fecha_erp = next(
      (c for c in df_erp.columns if "fecha" in str(c).lower()), None
  )
  df_erp["Tasa_BCV"] = (
      [obtener_tasa_dual(f) for f in df_erp[col_fecha_erp]]
      if col_fecha_erp
      else df_banco["Tasa_BCV"].iloc[0]
  )
  df_erp[f"Monto_{moneda_base}"] = (
      df_erp["Monto_Clean"] / df_erp["Tasa_BCV"]
  ).round(2)

  # --- CRUCE SIDE-BY-SIDE USANDO VALORES ABSOLUTOS ---
  erp_usados = set()
  match_erp_indices = []
  estatus_banco_lista = []

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
        # Cruce robusto por valor absoluto para ignorar diferencias de signo banco/ERP
        if abs(monto_b) == abs(row_e["Monto_Clean"]):
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

  total_movimientos = len(df_banco)
  total_ingresos = df_banco[df_banco["Monto_Clean"] > 0][
      "Monto_Clean"
  ].sum()
  total_egresos = df_banco[df_banco["Monto_Clean"] < 0][
      "Monto_Clean"
  ].sum()
  total_general = df_banco["Monto_Clean"].sum()
  cant_conciliados = len(erp_usados)
  cant_pendientes = total_movimientos - cant_conciliados

  mask_comisiones = df_banco[col_desc_banco].astype(str).str.contains(
      "comis|comisión|com|igtf|serv", case=False, na=False
  )
  total_comisiones = df_banco[mask_comisiones]["Monto_Clean"].sum()

  top_10_egresos = (
      df_erp[df_erp["Monto_Clean"] < 0]
      .sort_values(by="Monto_Clean", ascending=True)
      .head(10)
  )
  top_10_ingresos = (
      df_erp[df_erp["Monto_Clean"] > 0]
      .sort_values(by="Monto_Clean", ascending=False)
      .head(10)
  )

  tab1, tab2, tab3 = st.tabs(
      [
          "📊 Conciliación Fila por Fila y Consolidado",
          "💬 Asistente IA Financiero",
          "📥 Exportar Reportes",
      ]
  )

  with tab1:
    st.subheader(
        "Resumen General de Tesorería y Consolidado Horizontal (Side-by-Side)"
    )
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Movimientos", total_movimientos)
    c2.metric("Conciliados", cant_conciliados)
    c3.metric("Pendientes", cant_pendientes)
    c4.metric("Reversadas (ERP)", cant_reversos)
    c5.metric("Comisiones (Banco)", f"{abs(total_comisiones):,.2f}")

    st.markdown("---")
    st.subheader("🏆 Top 10 Egresos e Ingresos")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
      st.markdown("#### 🔴 Top 10 Egresos")
      if not top_10_egresos.empty:
        cols_e = [
            c
            for c in [col_cta_contra, col_nom_contra, "Monto_Clean", "Tasa_BCV"]
            if c and c in top_10_egresos.columns
        ]
        st.dataframe(top_10_egresos[cols_e], use_container_width=True)
      else:
        st.info("Sin egresos.")
    with col_t2:
      st.markdown("#### 🟢 Top 10 Ingresos")
      if not top_10_ingresos.empty:
        cols_i = [
            c
            for c in [col_cta_contra, col_nom_contra, "Monto_Clean", "Tasa_BCV"]
            if c and c in top_10_ingresos.columns
        ]
        st.dataframe(top_10_ingresos[cols_i], use_container_width=True)
      else:
        st.info("Sin ingresos.")

    st.markdown("---")
    st.subheader("📋 Consolidado en Una Sola Fila (Side-by-Side)")
    st.dataframe(df_consolidado, use_container_width=True)

    buffer_excel = BytesIO()
    with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
      df_consolidado.to_excel(
          writer, index=False, sheet_name="Consolidado_Conciliacion"
      )
      df_banco.to_excel(writer, index=False, sheet_name="Extracto_Banco")
      df_erp.to_excel(writer, index=False, sheet_name="Mayor_ERP")
      pd.DataFrame({
          "Metrica": [
              "Total Movimientos",
              "Total General",
              "Ingresos",
              "Egresos",
              "Conciliados",
              "Pendientes",
              "Reversados",
              "Comisiones",
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
      }).to_excel(writer, index=False, sheet_name="Resumen_Gerencial")
    buffer_excel.seek(0)

    st.download_button(
        label="⬇️ Descargar Excel Consolidado (4 Pestañas)",
        data=buffer_excel,
        file_name=f"Consolidado_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        key="btn_excel",
    )

  with tab2:
    st.subheader("💬 Asistente IA Financiero (Claude)")
    for msg in st.session_state.mensajes_chat:
      with st.chat_message(msg["role"]):
        st.write(msg["content"])

    if pregunta := st.chat_input("Escribe tu pregunta financiera..."):
      st.session_state.mensajes_chat.append(
          {"role": "user", "content": pregunta}
      )
      with st.chat_message("user"):
        st.write(pregunta)

      if ia_disponible and API_KEY_OR != "":
        with st.spinner("Analizando con Claude..."):
          try:
            resumen_ctx = (
                f"Total Movimientos: {total_movimientos}\nIngresos:"
                f" {total_ingresos:,.2f}\nEgresos: {total_egresos:,.2f}\nReversados:"
                f" {cant_reversos}\nComisiones: {abs(total_comisiones):,.2f}"
            )
            mensajes_api = [
                {
                    "role": "system",
                    "content": (
                        "Eres un auditor contable experto en Venezuela.\n"
                        + resumen_ctx
                    ),
                }
            ]
            for m in st.session_state.mensajes_chat:
              mensajes_api.append({"role": m["role"], "content": m["content"]})
            resp = client_or.chat.completions.create(
                model="anthropic/claude-haiku-4.5",
                messages=mensajes_api,
                max_tokens=1000,
            )
            respuesta_ia = resp.choices[0].message.content
          except Exception as e:
            respuesta_ia = f"Error en la IA: {e}"
      else:
        respuesta_ia = "IA no configurada (Configure OpenRouter API Key)."

      st.session_state.mensajes_chat.append(
          {"role": "assistant", "content": respuesta_ia}
      )
      with st.chat_message("assistant"):
        st.write(respuesta_ia)

  with tab3:
    st.subheader("📥 Exportar Reporte Word")
    if st.button("📄 Generar Reporte Word", key="btn_word"):
      doc = Document()
      doc.add_heading("Reporte Gerencial de Conciliación", 0)
      doc.add_paragraph(
          f"Fecha: {datetime.now().strftime('%Y-%m-%d')}"
          f"\nMovimientos: {total_movimientos} | Conciliados:"
          f" {cant_conciliados} | Reversados: {cant_reversos}"
          f"\nComisiones: {abs(total_comisiones):,.2f}"
      )
      buffer_w = BytesIO()
      doc.save(buffer_w)
      buffer_w.seek(0)
      st.download_button(
          label="⬇️ Descargar Word (.docx)",
          data=buffer_w,
          file_name=f"Reporte_{datetime.now().strftime('%Y%m%d')}.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
          key="btn_desc_word",
      )
else:
  st.info("👈 Sube tus archivos en la barra lateral para iniciar.")