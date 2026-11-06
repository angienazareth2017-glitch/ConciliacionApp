import streamlit as st
import pandas as pd
from datetime import datetime
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
st.write("Laboratorios Quimio-Vet, C.A. - Conciliación Automatizada y Consolidación")

# --- BARRA LATERAL: Configuración ---
st.sidebar.header("⚙️ Configuración del Motor")
tolerancia_dias = st.sidebar.slider("Tolerancia de fechas (días)", 0, 7, 3)
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

        st.success("¡Archivos cargados y leídos correctamente!")

        if st.button("🚀 Ejecutar Conciliación Masiva"):
            with st.spinner("Procesando cruce masivo de partidas..."):
                banco = df_banco.copy()
                libro = df_libro.copy()
                
                # Función para limpiar montos venezolanos (ej: "BsS 2.500.000,00" -> 2500000.00)
                def limpiar_monto_ve(val):
                    if pd.isna(val):
                        return 0.0
                    s = str(val).replace('BsS', '').replace('Bs.', '').strip()
                    s = s.replace('.', '').replace(',', '.')
                    try:
                        return float(s)
                    except:
                        return 0.0

                # Normalizar Banco
                col_monto_b = next((c for c in banco.columns if 'monto' in c.lower()), 'Monto')
                col_fecha_b = next((c for c in banco.columns if 'fecha' in c.lower()), 'Fecha')
                col_desc_b = next((c for c in banco.columns if 'descripción' in c.lower()), 'Descripción')

                banco['Monto_Clean'] = pd.to_numeric(banco[col_monto_b], errors='coerce').fillna(0)
                banco['Fecha_Dt'] = pd.to_datetime(banco[col_fecha_b], format='%d/%m/%Y', errors='coerce')
                banco['Estatus'] = 'Pendiente Banco'
                banco['Libro_Match_Idx'] = None

                # Normalizar Libro Mayor
                col_fecha_l = next((c for c in libro.columns if 'fecha' in c.lower() and 'contabilización' in c.lower()), libro.columns[0])
                col_comm_l = next((c for c in libro.columns if 'comentario' in c.lower()), libro.columns[6])
                col_cargo_l = next((c for c in libro.columns if 'cargo' in c.lower()), 'Cargo (ML)')
                col_abono_l = next((c for c in libro.columns if 'abono' in c.lower()), 'Abono (ML)')

                libro['Cargo_Num'] = libro[col_cargo_l].apply(limpiar_monto_ve)
                libro['Abono_Num'] = libro[col_abono_l].apply(limpiar_monto_ve)
                
                # Neto contable: Abono (ingreso +) - Cargo (egreso -)
                libro['Monto_Neto'] = libro['Abono_Num'] - libro['Cargo_Num']
                libro['Fecha_Dt'] = pd.to_datetime(libro[col_fecha_l], errors='coerce')
                libro['Estatus'] = 'Pendiente Libros'

                # 1. Marcar Saldo Inicial y Anulados
                for idx_b, row_b in banco.iterrows():
                    if 'SALDO INICIAL' in str(row_b[col_desc_b]).upper():
                        banco.loc[idx_b, 'Estatus'] = 'Saldo Inicial (Banco)'
                        break

                if len(libro) > 0:
                    libro.loc[0, 'Estatus'] = 'Saldo Inicial (Libros)'

                for idx_l, row_l in libro.iterrows():
                    if libro.loc[idx_l, 'Estatus'] == 'Saldo Inicial (Libros)':
                        continue
                    carg = row_l['Cargo_Num']
                    abon = row_l['Abono_Num']
                    comm = str(row_l[col_comm_l]).upper()
                    if (carg == 0 and abon == 0) or ('ANUL' in comm) or (carg > 0 and abon > 0 and carg == abon):
                        libro.loc[idx_l, 'Estatus'] = 'Anulado / Reversado'

                # 2. MOTOR DE CRUCE MASIVO (Optimizado)
                libro_usados = set()
                for idx_b, row_b in banco.iterrows():
                    if banco.loc[idx_b, 'Estatus'] != 'Pendiente Banco':
                        continue
                    
                    m_b = row_b['Monto_Clean']
                    f_b = row_b['Fecha_Dt']
                    if pd.isna(f_b) or m_b == 0:
                        continue

                    for idx_l, row_l in libro.iterrows():
                        if idx_l in libro_usados or libro.loc[idx_l, 'Estatus'] not in ['Pendiente Libros']:
                            continue
                        
                        f_l = row_l['Fecha_Dt']
                        if pd.isna(f_l):
                            continue

                        # Validar tolerancia de fechas (hasta 3 días)
                        if abs((f_l - f_b).days) > tolerancia_dias:
                            continue

                        m_l = row_l['Monto_Neto']

                        # Cruce exacto o por valor absoluto magnitud
                        if abs(m_b - m_l) <= tolerancia_monto or abs(abs(m_b) - abs(m_l)) <= tolerancia_monto:
                            banco.loc[idx_b, 'Estatus'] = 'Conciliado'
                            banco.loc[idx_b, 'Libro_Match_Idx'] = idx_l
                            libro.loc[idx_l, 'Estatus'] = 'Conciliado'
                            libro_usados.add(idx_l)
                            break

                # Métricas y Porcentajes
                total_b = len(banco[~banco['Estatus'].str.contains('Saldo Inicial')])
                conciliados_b = len(banco[banco['Estatus'] == 'Conciliado'])
                porcentaje = (conciliados_b / total_b * 100) if total_b > 0 else 0
                anulados_l = len(libro[libro['Estatus'] == 'Anulado / Reversado'])

                st.subheader("📈 Resumen y Porcentaje de Conciliación")
                m1, m2, m3, m4, m5 = st.columns(5)
                m1.metric("Movimientos Banco", total_b)
                m2.metric("Conciliados", conciliados_b)
                m3.metric("Pendientes Banco", total_b - conciliados_b)
                m4.metric("% Conciliado", f"{porcentaje:.1f}%")
                m5.metric("Anulados / Reversados", anulados_l)

                # 3. CONSOLIDADO (Predomina el Banco)
                consolidado_rows = []
                for idx_b, row_b in banco.iterrows():
                    fila_cons = {
                        'Banco_Fecha': row_b.get('Fecha'),
                        'Banco_Referencia': row_b.get('Referencia'),
                        'Banco_Descripcion': row_b.get('Descripción'),
                        'Banco_Monto': row_b.get('Monto_Clean'),
                        'Estatus_Conciliacion': row_b.get('Estatus')
                    }
                    
                    match_idx = row_b.get('Libro_Match_Idx')
                    if match_idx is not None and not pd.isna(match_idx):
                        r_libro = libro.loc[int(match_idx)]
                        fila_cons.update({
                            'Libro_Fecha_Contabilizacion': r_libro.get('Fecha de contabilización'),
                            'Libro_Nro_Documento': r_libro.get('Nº documento'),
                            'Libro_Nro_Transaccion': r_libro.get('Nº de transacción'),
                            'Libro_Comentarios': r_libro.get('Comentarios'),
                            'Libro_Cargo_Egreso': r_libro.get('Cargo (ML)'),
                            'Libro_Abono_Ingreso': r_libro.get('Abono (ML)')
                        })
                    else:
                        fila_cons.update({
                            'Libro_Fecha_Contabilizacion': '',
                            'Libro_Nro_Documento': '',
                            'Libro_Nro_Transaccion': '',
                            'Libro_Comentarios': '',
                            'Libro_Cargo_Egreso': '',
                            'Libro_Abono_Ingreso': ''
                        })
                    consolidado_rows.append(fila_cons)

                df_consolidado = pd.DataFrame(consolidado_rows)

                st.subheader("📊 Vista Previa: Consolidado (Predomina el Banco)")
                st.dataframe(df_consolidado.head(10))

                # Exportar a Excel con 3 Hojas
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_consolidado.to_excel(writer, sheet_name='Banco_vs_Libro_Consolidado', index=False)
                    banco.to_excel(writer, sheet_name='Extracto_Banco', index=False)
                    libro.to_excel(writer, sheet_name='Mayor_Contable', index=False)
                excel_data = output.getvalue()

                st.download_button(
                    label="📥 Descargar Reporte Completo en Excel (3 Hojas)",
                    data=excel_data,
                    file_name="Reporte_Conciliacion_QuimioVet_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Error procesando los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos para activar el motor de conciliación.")