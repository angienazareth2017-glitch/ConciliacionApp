import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
import io

st.set_page_config(page_title="Agente de Conciliación Bancaria - Quimio-Vet", layout="wide")

# --- SISTEMA DE CONTRASEÑA PRIVADA ---
password_ingresada = st.sidebar.text_input("Ingrese su contraseña", type="password")
CONTRASENA_CORRECTA = "Ruben.1994.Juan"

if password_ingresada != CONTRASENA_CORRECTA:
    st.warning("🔒 Por favor, introduzca la contraseña correcta en la barra lateral para acceder al agente de conciliación.")
    st.stop()

# --- INTERFAZ PRINCIPAL ---
st.title("🏦 Agente de Conciliación Bancaria Inteligente")
st.write("Laboratorios Quimio-Vet, C.A. - Conciliación Automatizada de Banco Mercantil")

# --- BARRA LATERAL: Configuración de Reglas ---
st.sidebar.header("⚙️ Configuración del Motor")
tolerancia_dias = st.sidebar.slider("Tolerancia de fechas (días)", 0, 7, 3)
umbral_similitud = st.sidebar.slider("Umbral de similitud de texto (%)", 50, 100, 80)
tolerancia_monto = st.sidebar.number_input("Tolerancia en monto", min_value=0.0, value=0.0, step=0.01)

# --- CARGA DE ARCHIVOS ---
col_up1, col_up2 = st.columns(2)

with col_up1:
    archivo_banco = st.file_uploader("Sube tu Extracto Bancario (.xlsx)", type=["xlsx", "csv"])

with col_up2:
    archivo_libro = st.file_uploader("Sube tu Mayor Analítico Contable (.xlsx)", type=["xlsx", "csv"])

# --- PROCESAMIENTO Y CONCILIACIÓN ---
if archivo_banco and archivo_libro:
    try:
        # Lectura adaptada a las estructuras reales de tus archivos
        if archivo_banco.name.endswith('.csv'):
            df_banco = pd.read_csv(archivo_banco)
        else:
            # El banco tiene los encabezados en la fila 6 (skiprows=5)
            df_banco = pd.read_excel(archivo_banco, sheet_name=0, header=5)

        if archivo_libro.name.endswith('.csv'):
            df_libro = pd.read_csv(archivo_libro)
        else:
            # El libro mayor tiene los encabezados en la fila 1 (header=0) y salimos de la fila resumen de saldos (row 0)
            df_libro = pd.read_excel(archivo_libro, sheet_name=0, header=0)
            # Omitir la primera fila si es saldo inicial / texto de cabecera
            if len(df_libro) > 0 and 'Activos' in str(df_libro.iloc[0, 0]):
                df_libro = df_libro.iloc[1:].reset_index(drop=True)

        # Limpieza de espacios en los nombres de columnas
        df_banco.columns = [str(c).strip() for c in df_banco.columns]
        df_libro.columns = [str(c).strip() for c in df_libro.columns]

        st.success("¡Archivos cargados y leídos exitosamente!")
        
        # Vistas previas
        st.subheader("Vista previa: Extracto Bancario")
        st.dataframe(df_banco.head(3))

        st.subheader("Vista previa: Libro Mayor")
        st.dataframe(df_libro.head(3))

        # Botón para ejecutar la conciliación
        if st.button("🚀 Ejecutar Conciliación y Hallazgos"):
            with st.spinner("Procesando limpieza, cruce de datos y hallazgos..."):
                
                banco = df_banco.copy()
                libro = df_libro.copy()
                
                # Inicializar estatus
                banco['Estatus'] = 'Por Conciliar (Banco)'
                libro['Estatus'] = 'Por Conciliar (Libros)'
                
                # Identificar columnas en Banco
                col_fecha_banco = next((c for c in banco.columns if 'fecha' in c.lower()), 'Fecha')
                col_desc_banco = next((c for c in banco.columns if any(k in c.lower() for k in ['desc', 'concepto', 'detalle'])), 'Descripción')
                col_monto_banco = next((c for c in banco.columns if any(k in c.lower() for k in ['monto', 'importe'])), 'Monto')

                # Identificar columnas en Libro Mayor (Cargo = Debe / Egreso, Abono = Haber / Ingreso)
                col_fecha_libro = next((c for c in libro.columns if 'fecha' in c.lower() and 'contabilización' in c.lower()), libro.columns[0])
                col_desc_libro = next((c for c in libro.columns if any(k in c.lower() for k in ['comentario', 'concepto', 'detalle'])), libro.columns[6])
                col_cargo_libro = next((c for c in libro.columns if 'cargo' in c.lower()), 'Cargo (ML)')
                col_abono_libro = next((c for c in libro.columns if 'abono' in c.lower()), 'Abono (ML)')

                # Limpieza de montos en banco
                banco[col_monto_banco] = pd.to_numeric(
                    banco[col_monto_banco].astype(str).str.replace(r'[^0-9\.\-]', '', regex=True), 
                    errors='coerce'
                ).fillna(0)

                # Limpieza de montos en libro (combinando cargo y abono en un monto neto con signo)
                def limpiar_monto_libro(row):
                    cargo = str(row.get(col_cargo_libro, '0'))
                    abono = str(row.get(col_abono_libro, '0'))
                    
                    val_c = float(pd.to_numeric(pd.Series([cargo]).str.replace(r'[^0-9\.\-]', '', regex=True), errors='coerce').fillna(0).iloc[0])
                    val_a = float(pd.to_numeric(pd.Series([abono]).str.replace(r'[^0-9\.\-]', '', regex=True), errors='coerce').fillna(0).iloc[0])
                    
                    # Si hay abono es ingreso (+), si hay cargo es egreso (-)
                    if val_a > 0:
                        return val_a
                    elif val_c > 0:
                        return -val_c
                    return 0.0

                libro['Monto_Neto'] = libro.apply(limpiar_monto_libro, axis=1)

                # Motor de cruce inteligente (Monto exacto/tolerancia + Similitud de texto)
                for idx_b, row_b in banco.iterrows():
                    monto_b = row_b[col_monto_banco]
                    desc_b = str(row_b[col_desc_banco])
                    
                    # Omitir saldo inicial en banco si aplica
                    if 'SALDO INICIAL' in desc_b.upper():
                        banco.loc[idx_b, 'Estatus'] = 'Saldo Inicial'
                        continue

                    for idx_l, row_l in libro.iterrows():
                        if libro.loc[idx_l, 'Estatus'] == 'Conciliado':
                            continue
                        monto_l = row_l['Monto_Neto']
                        desc_l = str(row_l[col_desc_libro])
                        
                        # Validar tolerancia de monto
                        if abs(abs(monto_b) - abs(monto_l)) <= tolerancia_monto:
                            similitud = fuzz.ratio(desc_b.lower(), desc_l.lower())
                            if similitud >= umbral_similitud or monto_b == monto_l:
                                banco.loc[idx_b, 'Estatus'] = 'Conciliado'
                                libro.loc[idx_l, 'Estatus'] = 'Conciliado'
                                break

                st.success("¡Proceso de conciliación completado con éxito!")
                
                # Cálculo de Métricas y Porcentajes
                total_b = len(banco[banco['Estatus'] != 'Saldo Inicial'])
                conciliados_b = len(banco[banco['Estatus'] == 'Conciliado'])
                porcentaje_conciliado = (conciliados_b / total_b * 100) if total_b > 0 else 0
                
                st.subheader("📈 Resumen y Porcentaje de Avance")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total Movimientos Banco", total_b)
                m2.metric("Partidas Conciliadas", conciliados_b)
                m3.metric("Por Conciliar", total_b - conciliados_b)
                m4.metric("% Conciliado", f"{porcentaje_conciliado:.1f}%")

                # Visualización de Tablas con Resultados
                st.subheader("📊 Detalle del Extracto Bancario con Estatus")
                st.dataframe(banco)

                st.subheader("📊 Detalle del Mayor Contable con Estatus")
                st.dataframe(libro)

                # Generar archivo Excel descargable con los hallazgos
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    banco.to_excel(writer, sheet_name='Banco_Conciliado', index=False)
                    libro.to_excel(writer, sheet_name='Mayor_Conciliado', index=False)
                excel_data = output.getvalue()

                st.download_button(
                    label="📥 Descargar Reporte de Hallazgos en Excel",
                    data=excel_data,
                    file_name="Reporte_Conciliacion_QuimioVet.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"Ocurrió un error al procesar o limpiar los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos para activar el motor de conciliación y hallar las diferencias automáticamente.")