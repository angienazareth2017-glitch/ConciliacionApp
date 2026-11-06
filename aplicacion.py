import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
import io

st.set_page_config(page_title="Agente de Conciliación Bancaria - Quimio-Vet", layout="wide")

# --- SISTEMA DE CONTRASEÑA PRIVADA ---
password_ingresada = st.sidebar.text_input("Ingrese su contraseña", type="password")
CONTRASENA_CORRECTA = "Ruben.1994.Juan"

if password_ingresada != CONTRASENA_CORRECTA:
    st.warning("🔒 Por favor, introduzca la contraseña correcta en la barra lateral para acceder al agente.")
    st.stop()

# --- INTERFAZ PRINCIPAL ---
st.title("🏦 Agente de Conciliación Bancaria Inteligente")
st.write("Laboratorios Quimio-Vet, C.A. - Conciliación Automatizada y Análisis de Anulaciones")

# --- BARRA LATERAL: Configuración ---
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
        # Lectura de archivos
        df_banco = pd.read_excel(archivo_banco, sheet_name=0, header=5)
        df_libro = pd.read_excel(archivo_libro, sheet_name=0, header=0)
        
        if len(df_libro) > 0 and 'Activos' in str(df_libro.iloc[0, 0]):
            df_libro = df_libro.iloc[1:].reset_index(drop=True)

        df_banco.columns = [str(c).strip() for c in df_banco.columns]
        df_libro.columns = [str(c).strip() for c in df_libro.columns]

        st.success("¡Archivos cargados y estructurados correctamente!")

        if st.button("🚀 Ejecutar Conciliación y Análisis de Anulados"):
            with st.spinner("Procesando montos venezolanos, cruce y análisis de asientos..."):
                banco = df_banco.copy()
                libro = df_libro.copy()
                
                # Función para limpiar formato de moneda venezolano (ej: "BsS 2.500.000,00" -> 2500000.00)
                def limpiar_monto_ve(val):
                    if pd.isna(val):
                        return 0.0
                    s = str(val).replace('BsS', '').replace('Bs.', '').strip()
                    # Quitar puntos de miles y cambiar coma decimal por punto
                    s = s.replace('.', '').replace(',', '.')
                    try:
                        return float(s)
                    except:
                        return 0.0

                # Normalizar Banco
                col_monto_b = next((c for c in banco.columns if 'monto' in c.lower()), 'Monto')
                col_ref_b = next((c for c in banco.columns if 'referencia' in c.lower()), 'Referencia')
                col_desc_b = next((c for c in banco.columns if 'descripción' in c.lower()), 'Descripción')

                banco['Monto_Clean'] = pd.to_numeric(banco[col_monto_b], errors='coerce').fillna(0)
                banco['Ref_Clean'] = banco[col_ref_b].astype(str).str.strip()
                banco['Estatus'] = 'Por Conciliar (Banco)'

                # Normalizar Libro Mayor
                col_doc_l = next((c for c in libro.columns if 'documento' in c.lower()), libro.columns[3])
                col_trans_l = next((c for c in libro.columns if 'transacción' in c.lower()), libro.columns[4])
                col_comm_l = next((c for c in libro.columns if 'comentario' in c.lower()), libro.columns[6])
                col_cargo_l = next((c for c in libro.columns if 'cargo' in c.lower()), 'Cargo (ML)')
                col_abono_l = next((c for c in libro.columns if 'abono' in c.lower()), 'Abono (ML)')

                libro['Cargo_Num'] = libro[col_cargo_l].apply(limpiar_monto_ve)
                libro['Abono_Num'] = libro[col_abono_l].apply(limpiar_monto_ve)
                
                # En libro: Abono = Ingreso (+), Cargo = Egreso (-)
                libro['Monto_Neto'] = libro['Abono_Num'] - libro['Cargo_Num']
                libro['Doc_Clean'] = libro[col_doc_l].astype(str).str.strip()
                libro['Trans_Clean'] = pd.to_numeric(libro[col_trans_l], errors='coerce').fillna(0).astype(int).astype(str)
                libro['Estatus'] = 'Por Conciliar (Libros)'

                # 1. ANÁLISIS DE ASIENTOS ANULADOS / REVERSADOS INTERNOS
                # Detectar si en el libro hay pares de cargos y abonos con el mismo monto exacto y referencia que se anulan
                for idx, row in libro.iterrows():
                    monto = row['Monto_Neto']
                    if monto == 0:
                        libro.loc[idx, 'Estatus'] = 'Anulado / Cero'
                    else:
                        # Buscar contrapartida opuesta en el mismo libro
                        opuesto = libro[(libro.index != idx) & (libro['Monto_Neto'] == -monto) & (libro['Estatus'] == 'Por Conciliar (Libros)')]
                        if len(opuesto) > 0:
                            libro.loc[idx, 'Estatus'] = 'Asiento Reversado / Anulado'
                            libro.loc[opuesto.index[0], 'Estatus'] = 'Asiento Reversado / Anulado'

                # 2. MOTOR DE CRUCE INTELIGENTE BANCO VS LIBROS
                for idx_b, row_b in banco.iterrows():
                    m_b = row_b['Monto_Clean']
                    ref_b = row_b['Ref_Clean']
                    desc_b = str(row_b[col_desc_b])

                    if 'SALDO INICIAL' in desc_b.upper():
                        banco.loc[idx_b, 'Estatus'] = 'Saldo Inicial'
                        continue

                    for idx_l, row_l in libro.iterrows():
                        if libro.loc[idx_l, 'Estatus'] in ['Conciliado', 'Asiento Reversado / Anulado', 'Anulado / Cero']:
                            continue
                        m_l = row_l['Monto_Neto']
                        doc_l = row_l['Doc_Clean']
                        trans_l = row_l['Trans_Clean']
                        desc_l = str(row_l[col_comm_l])

                        # Coincidencia por monto y referencia o similitud de texto
                        monto_coincide = (abs(m_b - m_l) <= tolerancia_monto) or (abs(abs(m_b) - abs(m_l)) <= tolerancia_monto)
                        ref_coincide = (ref_b != '' and (ref_b == doc_l or ref_b == trans_l))
                        
                        similitud = fuzz.ratio(desc_b.lower(), desc_l.lower())

                        if monto_coincide and (ref_coincide or similitud >= umbral_similitud or m_b == m_l):
                            banco.loc[idx_b, 'Estatus'] = 'Conciliado'
                            libro.loc[idx_l, 'Estatus'] = 'Conciliado'
                            break

                # Métricas y Porcentajes
                total_b = len(banco[banco['Estatus'] != 'Saldo Inicial'])
                conciliados_b = len(banco[banco['Estatus'] == 'Conciliado'])
                porcentaje = (conciliados_b / total_b * 100) if total_b > 0 else 0

                anulados_l = len(libro[libro['Estatus'] == 'Asiento Reversado / Anulado'])

                st.subheader("📈 Resumen y Porcentaje de Conciliación")
                m1, m2, m3, m4, m5 = st.columns(5)
                m1.metric("Movimientos Banco", total_b)
                m2.metric("Conciliados", conciliados_b)
                m3.metric("Pendientes Banco", total_b - conciliados_b)
                m4.metric("% Conciliado", f"{porcentaje:.1f}%")
                m5.metric("Asientos Anulados/Reversados", anulados_l)

                st.subheader("📊 Detalle del Extracto Bancario")
                st.dataframe(banco)

                st.subheader("📊 Detalle del Mayor Contable (con Hallazgos de Anulados)")
                st.dataframe(libro)

                # Exportar a Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    banco.to_excel(writer, sheet_name='Banco_Analizado', index=False)
                    libro.to_excel(writer, sheet_name='Mayor_Analizado', index=False)
                excel_data = output.getvalue()

                st.download_button(
                    label="📥 Descargar Reporte Completo en Excel",
                    data=excel_data,
                    file_name="Reporte_Conciliacion_QuimioVet_Avanzado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Error procesando los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos para activar el motor avanzado de conciliación y detección de anulados.")