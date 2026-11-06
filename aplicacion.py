import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
import io

st.set_page_config(page_title="Agente de Conciliación Bancaria", layout="wide")

# --- SISTEMA DE CONTRASEÑA PRIVADA ---
password_ingresada = st.sidebar.text_input("Ingrese su contraseña", type="password")
CONTRASENA_CORRECTA = "Ruben.1994.Juan"

if password_ingresada != CONTRASENA_CORRECTA:
    st.warning("🔒 Por favor, introduzca la contraseña correcta en la barra lateral para acceder al agente de conciliación.")
    st.stop()

# --- INTERFAZ PRINCIPAL ---
st.title("🏦 Agente de Conciliación Bancaria Inteligente")
st.write("Sube tus archivos de extracto bancario y libro mayor para emparejar transacciones automáticamente.")

# --- BARRA LATERAL: Configuración de Reglas ---
st.sidebar.header("⚙️ Configuración del Motor")
tolerancia_dias = st.sidebar.slider("Tolerancia de fechas (días)", 0, 7, 3)
umbral_similitud = st.sidebar.slider("Umbral de similitud de texto (%)", 50, 100, 80)
tolerancia_monto = st.sidebar.number_input("Tolerancia en monto", min_value=0.0, value=0.0, step=0.01)

# --- CARGA DE ARCHIVOS ---
col_up1, col_up2 = st.columns(2)

with col_up1:
    archivo_banco = st.file_uploader("Sube tu Extracto Bancario", type=["xlsx", "csv"])

with col_up2:
    archivo_libro = st.file_uploader("Sube tu Libro Mayor", type=["xlsx", "csv"])

# --- PROCESAMIENTO Y CONCILIACIÓN ---
if archivo_banco and archivo_libro:
    try:
        # Lectura de archivos según su extensión
        if archivo_banco.name.endswith('.csv'):
            df_banco = pd.read_csv(archivo_banco)
        else:
            df_banco = pd.read_excel(archivo_banco)

        if archivo_libro.name.endswith('.csv'):
            df_libro = pd.read_csv(archivo_libro)
        else:
            df_libro = pd.read_excel(archivo_libro)

        # Limpiar espacios en los nombres de las columnas
        df_banco.columns = [str(c).strip() for c in df_banco.columns]
        df_libro.columns = [str(c).strip() for c in df_libro.columns]

        st.success("¡Archivos cargados exitosamente!")
        
        # Vista previa de los datos
        st.subheader("Vista previa: Extracto Bancario")
        st.dataframe(df_banco.head(3))

        st.subheader("Vista previa: Libro Mayor")
        st.dataframe(df_libro.head(3))

        # Botón para ejecutar el motor
        if st.button("🚀 Ejecutar Conciliación Automática"):
            with st.spinner("Procesando y cruzando transacciones..."):
                
                banco = df_banco.copy()
                libro = df_libro.copy()
                
                # Inicializar estatus
                banco['Estatus'] = 'Pendiente Banco'
                libro['Estatus'] = 'Pendiente Libros'
                
                # Detectar automáticamente columnas clave por palabras clave
                col_monto_banco = next((col for col in banco.columns if any(k in col.lower() for k in ['monto', 'importe', 'debe', 'haber'])), banco.columns[-1])
                col_monto_libro = next((col for col in libro.columns if any(k in col.lower() for k in ['monto', 'importe', 'debe', 'haber'])), libro.columns[-1])
                
                col_desc_banco = next((col for col in banco.columns if any(k in col.lower() for k in ['desc', 'concepto', 'detalle'])), banco.columns[1])
                col_desc_libro = next((col for col in libro.columns if any(k in col.lower() for k in ['desc', 'concepto', 'detalle'])), libro.columns[1])

                # Limpieza numérica
                banco[col_monto_banco] = pd.to_numeric(banco[col_monto_banco].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                libro[col_monto_libro] = pd.to_numeric(libro[col_monto_libro].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

                # Motor de cruce inteligente
                for idx_b, row_b in banco.iterrows():
                    monto_b = row_b[col_monto_banco]
                    desc_b = str(row_b[col_desc_banco])
                    
                    for idx_l, row_l in libro.iterrows():
                        if libro.loc[idx_l, 'Estatus'] == 'Conciliado':
                            continue
                        monto_l = row_l[col_monto_libro]
                        desc_l = str(row_l[col_desc_libro])
                        
                        # Validar tolerancia de monto
                        if abs(monto_b - monto_l) <= tolerancia_monto:
                            similitud = fuzz.ratio(desc_b.lower(), desc_l.lower())
                            if similitud >= umbral_similitud or monto_b == monto_l:
                                banco.loc[idx_b, 'Estatus'] = 'Conciliado'
                                libro.loc[idx_l, 'Estatus'] = 'Conciliado'
                                break

                st.success("¡Conciliación finalizada con éxito!")
                
                # Métricas visuales
                total_b = len(banco)
                conciliados_b = len(banco[banco['Estatus'] == 'Conciliado'])
                pendientes_b = total_b - conciliados_b
                
                m1, m2, m3 = st.columns(3)
                m1.metric("Total Registros Banco", total_b)
                m2.metric("Conciliados", conciliados_b)
                m3.metric("Por Conciliar", pendientes_b)

                # Tablas de resultados
                st.subheader("📊 Resultados en Extracto Bancario")
                st.dataframe(banco)

                st.subheader("📊 Resultados en Libro Mayor")
                st.dataframe(libro)

                # Generar archivo Excel en memoria para descarga
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    banco.to_excel(writer, sheet_name='Extracto_Conciliado', index=False)
                    libro.to_excel(writer, sheet_name='Mayor_Conciliado', index=False)
                excel_data = output.getvalue()

                st.download_button(
                    label="📥 Descargar Informe de Conciliación en Excel",
                    data=excel_data,
                    file_name="Reporte_Conciliacion_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos (Extracto Bancario y Libro Mayor) para habilitar el motor de conciliación.")