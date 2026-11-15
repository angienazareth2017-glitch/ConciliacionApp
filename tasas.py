import pandas as pd
import streamlit as st

# Suponiendo que ya estás dentro de la app principal autenticada...
st.subheader("📁 Carga Masiva de Tasas BCV desde Excel")

st.markdown("""
Sube un archivo Excel (`.xlsx` o `.xls`) que contenga al menos una columna con la **fecha** y columnas con las tasas de cambio para **USD** y/o **EUR**.
""")

# 1. Componente para subir el archivo
archivo_subido = st.file_uploader(
    "Selecciona tu archivo Excel de tasas", type=["xlsx", "xls"]
)

if archivo_subido is not None:
  try:
    # Leemos el Excel usando pandas
    df_tasas = pd.read_excel(archivo_subido)

    st.write("🔍 **Vista previa de los datos cargados:**")
    st.dataframe(df_tasas.head())

    # 2. Configuración de columnas para mapear el Excel con tu clase TablaTasas
    st.markdown("---")
    st.write("⚙️ **Mapeo de Columnas**")

    columnas_disponibles = list(df_tasas.columns)

    col1, col2, col3 = st.columns(3)
    with col1:
      col_fecha = st.selectbox(
          "Columna de Fecha", columnas_disponibles, index=0
      )
    with col2:
      col_usd = st.selectbox(
          "Columna Tasa USD (Opcional)", ["Ninguna"] + columnas_disponibles
      )
    with col3:
      col_eur = st.selectbox(
          "Columna Tasa EUR (Opcional)", ["Ninguna"] + columnas_disponibles
      )

    # Diccionario dinámico para indicar qué columnas mapear a la clase TablaTasas
    columnas_moneda_map = {}
    if col_usd != "Ninguna":
      columnas_moneda_map["USD"] = col_usd
    if col_eur != "Ninguna":
      columnas_moneda_map["EUR"] = col_eur

    # 3. Botón para ejecutar la carga masiva
    if st.button("Procesar y Cargar Tasas al Sistema"):
      if not columnas_moneda_map:
        st.warning("⚠️ Debes seleccionar al menos una columna de moneda (USD o EUR).")
      else:
        # Aseguramos que la tabla esté inicializada en el session_state
        if "tabla_tasas" not in st.session_state:
          st.session_state.tabla_tasas = TablaTasas(max_antiguedad_dias=7)

        # Convertimos las filas del DataFrame a un iterable de diccionarios
        filas = df_tasas.to_dict(orient="records")

        # Llamamos al método nativo de tu archivo tasas.py
        errores = st.session_state.tabla_tasas.cargar_filas(
            filas=filas,
            columna_fecha=col_fecha,
            columnas_moneda=columnas_moneda_map,
            tipo="BCV",
            fuente=archivo_subido.name,
        )

        st.success("¡Proceso de carga masiva finalizado!")

        # Reportar resultados y posibles errores de formato si los hay
        if errores:
          st.warning(
              f"Se encontraron {len(errores)} filas con advertencias o errores"
              " que fueron omitidas:"
          )
          df_errores = pd.DataFrame([vars(e) for e in errores])
          st.dataframe(df_errores)
        else:
          st.success("🎉 ¡Todas las tasas se cargaron sin ningún error!")

  except Exception as e:
    st.error(f"Error al leer el archivo Excel: {e}")