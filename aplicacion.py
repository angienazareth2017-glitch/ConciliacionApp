import streamlit as st
import pandas as pd
from rapidfuzz import fuzz

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
tolerancia_monto = st.sidebar.number_input("Tolerancia en monto ($ / local)", min_value=0.0, value=0.0, step=0.01)

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

        st.success("¡Archivos cargados exitosamente!")
        
        # Vista previa de los datos
        st.subheader("Vista previa: Extracto Bancario")
        st.dataframe(df_banco.head())

        st.subheader("Vista previa: Libro Mayor")
        st.dataframe(df_libro.head())

        # Botón para ejecutar la conciliación
        if st.button("Ejecutar Conciliación Automática"):
            st.info("Procesando el emparejamiento de transacciones...")
            
            # Aquí puedes integrar la lógica principal de cruce con tus columnas específicas
            # Ejemplo básico de estructura de resultados
            st.success("¡Conciliación finalizada con éxito!")
            
    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos (Extracto Bancario y Libro Mayor) para habilitar el motor de conciliación.")