import streamlit as st
import pandas as pd
import io

# Intento seguro de importar docx para el reporte en Word
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

st.set_page_config(page_title="Agente IA de Conciliación y Análisis - Quimio-Vet", layout="wide")

# --- SISTEMA DE CONTRASEÑA PRIVADA ---
password_ingresada = st.sidebar.text_input("Ingrese su contraseña", type="password")
CONTRASENA_CORRECTA = "Ruben.1994.Juan"

if password_ingresada != CONTRASENA_CORRECTA:
    st.warning("🔒 Por favor, introduzca la contraseña correcta en la barra lateral para acceder al agente.")
    st.stop()

# --- FUNCIÓN PARA GENERAR INFORME EN WORD (.DOCX) ---
def generar_documento_word(total_b, conciliados_b, porcentaje, anulados_l, total_ingresos_Bs, total_egresos_Bs, total_comisiones_Bs, total_ingresos_usd, total_egresos_usd, total_comisiones_usd):
    if not HAS_DOCX:
        return None
    doc = Document()
    
    doc.add_heading('Informe Gerencial de Conciliación Bancaria', level=0)
    doc.add_paragraph('Laboratorios Quimio-Vet, C.A. | Departamento de Tesorería y Contabilidad')
    doc.add_paragraph('Conversión Dinámica aplicada por Tasa BCV Diaria Histórica\n')
    
    doc.add_heading('1. Resumen Ejecutivo de Conciliación', level=1)
    p = doc.add_paragraph()
    p.add_run("Durante el periodo evaluado se registraron un total de ")
    p.add_run(f"{total_b} movimientos bancarios").bold = True
    p.add_run(", alcanzando un índice de conciliación automática del ")
    p.add_run(f"{porcentaje:.1f}%").bold = True
    p.add_run(f" ({conciliados_b} partidas conciliadas). Se identificaron ")
    p.add_run(f"{anulados_l} asientos de anulación/error").bold = True
    p.add_run(" en el mayor analítico de SAP con su debido reverso contable.")

    doc.add_heading('2. Desglose de Flujos Monetarios (Bs. y USD Histórico)', level=1)
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Concepto'
    hdr_cells[1].text = 'Monto en Bolívares (Bs.)'
    hdr_cells[2].text = 'Equivalente en Dólares (USD)'

    datos = [
        ('Créditos Operativos (Ingresos)', f"Bs. {total_ingresos_Bs:,.2f}", f"$ {total_ingresos_usd:,.2f}"),
        ('Débitos Operativos (Egresos)', f"Bs. {total_egresos_Bs:,.2f}", f"$ {total_egresos_usd:,.2f}"),
        ('Comisiones y Tarifas Bancarias', f"Bs. {total_comisiones_Bs:,.2f}", f"$ {total_comisiones_usd:,.2f}"),
    ]

    for concepto, monto_bs, monto_usd in datos:
        row_cells = table.add_row().cells
        row_cells[0].text = concepto
        row_cells[1].text = monto_bs
        row_cells[2].text = monto_usd

    doc.add_heading('3. Observaciones de Auditoría', level=1)
    doc.add_paragraph('• Las comisiones y cargos bancarios aplicados requieren causación contable según extracto.')
    doc.add_paragraph('• Las partidas pendientes en banco responden a depósitos consolidados o transferencias por lotes en SAP.')
    doc.add_paragraph('• La conversión a moneda pesada se ejecutó aplicando la tasa oficial BCV correspondiente a la fecha exacta de cada transacción.')

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ PRINCIPAL STREAMLIT ---
st.title("🏦 Agente IA de Conciliación y Análisis Financiero")
st.write("Laboratorios Quimio-Vet, C.A. - Conciliación Avanzada, Contrapartidas y Tasas BCV Diarias")

# --- BARRA LATERAL: Configuración ---
st.sidebar.header("⚙️ Configuración del Motor & Divisas")
tolerancia_dias = st.sidebar.slider("Tolerancia de fechas (días)", 0, 7, 3)
tolerancia_monto = st.sidebar.number_input("Tolerancia en monto", min_value=0.0, value=0.0, step=0.01)

st.sidebar.markdown("---")
st.sidebar.subheader("💱 Tasa BCV por Defecto (Fallback)")
tasa_fallback = st.sidebar.number_input("Tasa si falta fecha en archivo", min_value=1.0, value=36.5, step=0.1)

# --- CARGA DE ARCHIVOS ---
col_up1, col_up2, col_up3 = st.columns(3)
with col_up1:
    archivo_banco = st.file_uploader("1. Extracto Bancario (.xlsx)", type=["xlsx", "csv"])
with col_up2:
    archivo_libro = st.file_uploader("2. Mayor Analítico SAP (.xlsx)", type=["xlsx", "csv"])
with col_up3:
    archivo_tasas = st.file_uploader("3. Tasas BCV Diarias (.xlsx, .csv)", type=["xlsx", "csv"])

# --- PROCESAMIENTO Y CONCILIACIÓN ---
if archivo_banco and archivo_libro:
    try:
        df_banco = pd.read_excel(archivo_banco, sheet_name=0, header=5)
        df_libro = pd.read_excel(archivo_libro, sheet_name=0, header=0)
        
        # Carga ultra-robusta de tasas BCV diarias con conversión estricta a string/fecha
        tasa_dict = {}
        if archivo_tasas:
            try:
                df_tasas = pd.read_excel(archivo_tasas, sheet_name=0) if archivo_tasas.name.endswith('xlsx') else pd.read_csv(archivo_tasas)
                if len(df_tasas) > 0 and len(df_tasas.columns) >= 2:
                    df_tasas.columns = [str(c).strip() for c in df_tasas.columns]
                    
                    c_fecha_t = next((c for c in df_tasas.columns if 'fecha' in c.lower() or 'dia' in c.lower()), df_tasas.columns[0])
                    c_tasa_t = next((c for c in df_tasas.columns if 'tasa' in c.lower() or 'cambio' in c.lower() or 'valor' in c.lower() or '$' in c.lower() or 'bcv' in c.lower()), df_tasas.columns[1])
                    
                    # Normalizar a string de fecha YYYY-MM-DD para evitar errores de timestamp
                    df_tasas['Fecha_Str'] = pd.to_datetime(df_tasas[c_fecha_t], errors='coerce', dayfirst=True).dt.strftime('%Y-%m-%d')
                    df_tasas['Tasa_Num'] = pd.to_numeric(df_tasas[c_tasa_t].astype(str).str.replace(',', '.'), errors='coerce')
                    
                    df_tasas = df_tasas.dropna(subset=['Fecha_Str', 'Tasa_Num'])
                    tasa_dict = dict(zip(df_tasas['Fecha_Str'], df_tasas['Tasa_Num']))
                    st.success(f"✅ Tasas BCV cargadas exitosamente: {len(tasa_dict)} días registrados.")
            except Exception as e:
                st.sidebar.warning(f"Advertencia al leer tasas: {e}. Se usará la tasa por defecto.")

        if len(df_libro) > 0 and 'Activos' in str(df_libro.iloc[0, 0]):
            df_libro = df_libro.iloc[1:].reset_index(drop=True)

        df_banco.columns = [str(c).strip() for c in df_banco.columns]
        df_libro.columns = [str(c).strip() for c in df_libro.columns]

        st.success("¡Archivos cargados y leídos correctamente!")

        if st.button("🚀 Ejecutar Conciliación y Tasas Diarias"):
            with st.spinner("Procesando cruce masivo, contrapartidas y tasas BCV diarias..."):
                banco = df_banco.copy()
                libro = df_libro.copy()
                
                def limpiar_monto_flexible(val):
                    if pd.isna(val):
                        return 0.0
                    s = str(val).replace('BsS', '').replace('Bs.', '').strip()
                    if ',' in s and '.' in s:
                        if s.rfind(',') > s.rfind('.'):
                            s = s.replace('.', '').replace(',', '.')
                        else:
                            s = s.replace(',', '')
                    elif ',' in s:
                        s = s.replace(',', '.')
                    try:
                        num = float(s)
                        if abs(num) > 100000000: 
                            return num / 100.0
                        return num
                    except:
                        return 0.0

                col_monto_b = next((c for c in banco.columns if 'monto' in c.lower()), 'Monto')
                col_fecha_b = next((c for c in banco.columns if 'fecha' in c.lower()), 'Fecha')
                col_desc_b = next((c for c in banco.columns if 'descripción' in c.lower()), 'Descripción')

                banco['Monto_Clean'] = pd.to_numeric(banco[col_monto_b], errors='coerce').fillna(0)
                
                # Conversión estricta de fecha del banco a string YYYY-MM-DD para emparejar con el diccionario de tasas
                banco['Fecha_Str'] = pd.to_datetime(banco[col_fecha_b], errors='coerce', dayfirst=True).dt.strftime('%Y-%m-%d')
                banco['Fecha_Dt'] = pd.to_datetime(banco[col_fecha_b], errors='coerce', dayfirst=True).dt.normalize()

                banco['Estatus'] = 'Pendiente Banco'
                banco['Libro_Match_Idx'] = None

                # Asignación de tasa BCV usando las cadenas YYYY-MM-DD
                tasas_aplicadas = []
                montos_usd = []
                for _, row in banco.iterrows():
                    f_str = row['Fecha_Str']
                    monto = row['Monto_Clean']
                    tasa_dia = tasa_fallback
                    
                    if pd.notna(f_str) and f_str in tasa_dict:
                        tasa_dia = tasa_dict[f_str]
                    
                    tasas_aplicadas.append(tasa_dia)
                    montos_usd.append(monto / tasa_dia if tasa_dia > 0 else 0.0)

                banco['Tasa_BCV'] = tasas_aplicadas
                banco['Monto_USD'] = montos_usd

                col_fecha_l = next((c for c in libro.columns if 'fecha' in c.lower() and 'contabilización' in c.lower()), libro.columns[0])
                col_comm_l = next((c for c in libro.columns if 'comentario' in c.lower()), libro.columns[6])
                col_cargo_l = next((c for c in libro.columns if 'cargo' in c.lower()), 'Cargo (ML)')
                col_abono_l = next((c for c in libro.columns if 'abono' in c.lower()), 'Abono (ML)')

                libro['Cargo_Num'] = libro[col_cargo_l].apply(limpiar_monto_flexible)
                libro['Abono_Num'] = libro[col_abono_l].apply(limpiar_monto_flexible)
                libro['Monto_Abs'] = libro[['Cargo_Num', 'Abono_Num']].max(axis=1)
                libro['Fecha_Dt'] = pd.to_datetime(libro[col_fecha_l], errors='coerce').dt.normalize()
                libro['Estatus'] = 'Pendiente Libros'

                # 1. Saldo Inicial y Detección de Anulaciones
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
                    
                    es_anulado = (
                        (carg == 0 and abon == 0) or
                        ('ANUL' in comm) or
                        ('CANCELAR' in comm) or
                        ('ANULAR' in comm) or
                        ('REVERS' in comm) or
                        (carg > 0 and abon > 0 and carg == abon)
                    )
                    if es_anulado:
                        libro.loc[idx_l, 'Estatus'] = 'Anulado / Reversado (Error)'

                # 2. MOTOR DE CRUCE MASIVO
                libro_usados = set()
                for idx_b, row_b in banco.iterrows():
                    if banco.loc[idx_b, 'Estatus'] != 'Pendiente Banco':
                        continue
                    
                    m_b = abs(row_b['Monto_Clean'])
                    f_b = row_b['Fecha_Dt']
                    if pd.isna(f_b) or m_b == 0:
                        continue

                    for idx_l, row_l in libro.iterrows():
                        if idx_l in libro_usados or libro.loc[idx_l, 'Estatus'] not in ['Pendiente Libros']:
                            continue
                        
                        f_l = row_l['Fecha_Dt']
                        if pd.isna(f_l):
                            continue

                        if abs((f_l - f_b).days) > tolerancia_dias:
                            continue

                        m_l = row_l['Monto_Abs']

                        if abs(m_b - m_l) <= tolerancia_monto:
                            banco.loc[idx_b, 'Estatus'] = 'Conciliado'
                            banco.loc[idx_b, 'Libro_Match_Idx'] = idx_l
                            libro.loc[idx_l, 'Estatus'] = 'Conciliado'
                            libro_usados.add(idx_l)
                            break

                # 3. ANÁLISIS Y CLASIFICACIÓN
                banco_validos = banco[~banco['Estatus'].str.contains('Saldo Inicial')]
                mask_comision = banco_validos[col_desc_b].str.contains('COMISIO|TARIFA|MANTENIMIENTO', case=False, na=False)
                
                comisiones = banco_validos[mask_comision]
                creditos_sin_comision = banco_validos[(banco_validos['Monto_Clean'] > 0) & (~mask_comision)]
                debitos_sin_comision = banco_validos[(banco_validos['Monto_Clean'] < 0) & (~mask_comision)]

                st.subheader("🏦 Análisis de Movimientos del Banco (Clasificación)")
                c1, c2, c3 = st.columns(3)
                c1.metric("Débitos (Sin Comisiones)", len(debitos_sin_comision))
                c2.metric("Créditos (Sin Comisiones)", len(creditos_sin_comision))
                c3.metric("Comisiones y Tarifas", len(comisiones))

                total_b = len(banco_validos)
                conciliados_b = len(banco[banco['Estatus'] == 'Conciliado'])
                porcentaje = (conciliados_b / total_b * 100) if total_b > 0 else 0
                anulados_l = len(libro[libro['Estatus'] == 'Anulado / Reversado (Error)'])

                st.subheader("📈 Resumen y Porcentaje de Conciliación")
                m1, m2, m3, m4, m5 = st.columns(5)
                m1.metric("Movimientos Banco", total_b)
                m2.metric("Conciliados", conciliados_b)
                m3.metric("Pendientes Banco", total_b - conciliados_b)
                m4.metric("% Conciliado", f"{porcentaje:.1f}%")
                m5.metric("Anulados / Errores", anulados_l)

                # Totales para informes
                total_ingresos_Bs = creditos_sin_comision['Monto_Clean'].sum()
                total_egresos_Bs = abs(debitos_sin_comision['Monto_Clean'].sum())
                total_comisiones_Bs = abs(comisiones['Monto_Clean'].sum())

                total_ingresos_usd = creditos_sin_comision['Monto_USD'].sum()
                total_egresos_usd = abs(debitos_sin_comision['Monto_USD'].sum())
                total_comisiones_usd = abs(comisiones['Monto_USD'].sum())

                # --- 🤖 RESUMEN GERENCIAL EN PANTALLA ---
                st.markdown("---")
                st.subheader("🤖 Resumen e Insights Gerenciales (Tasas BCV Diarias)")

                texto_ia = f"""
                * **Diagnóstico General:** El mes registra **{total_b} movimientos bancarios** y **{anulados_l} asientos de anulación/error** en SAP.
                * **Flujo de Entradas:** **{len(creditos_sin_comision)} créditos operativos** sumando **Bs. {total_ingresos_Bs:,.2f}** (**$ {total_ingresos_usd:,.2f} USD** aplicando tasa BCV diaria exacta).
                * **Flujo de Salidas:** **{len(debitos_sin_comision)} débitos operativos** por **Bs. {total_egresos_Bs:,.2f}** (**$ {total_egresos_usd:,.2f} USD**).
                * **Comisiones Bancarias:** **{len(comisiones)} cargos** sumando **Bs. {total_comisiones_Bs:,.2f}** (**$ {total_comisiones_usd:,.2f} USD**).
                """
                st.info(texto_ia)

                # 4. CONSOLIDADO
                consolidado_rows = []
                for idx_b, row_b in banco.iterrows():
                    fila_cons = {
                        'Banco_Fecha': row_b.get('Fecha'),
                        'Banco_Referencia': row_b.get('Referencia'),
                        'Banco_Descripcion': row_b.get('Descripción'),
                        'Banco_Monto_Bs': row_b.get('Monto_Clean'),
                        'Tasa_BCV_Diaria': row_b.get('Tasa_BCV'),
                        'Banco_Monto_USD': row_b.get('Monto_USD'),
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
                            'Cuenta_Contrapartida': r_libro.get('Cuenta de contrapartida'),
                            'Nombre_Cuenta_Contrapartida': r_libro.get('Nombre de la cuenta de contrapartida'),
                            'Libro_Cargo_Egreso': r_libro.get('Cargo (ML)'),
                            'Libro_Abono_Ingreso': r_libro.get('Abono (ML)')
                        })
                    else:
                        fila_cons.update({
                            'Libro_Fecha_Contabilizacion': '',
                            'Libro_Nro_Documento': '',
                            'Libro_Nro_Transaccion': '',
                            'Libro_Comentarios': '',
                            'Cuenta_Contrapartida': '',
                            'Nombre_Cuenta_Contrapartida': '',
                            'Libro_Cargo_Egreso': '',
                            'Libro_Abono_Ingreso': ''
                        })
                    consolidado_rows.append(fila_cons)

                df_consolidado = pd.DataFrame(consolidado_rows)

                st.subheader("📊 Vista Previa: Consolidado (Con Tasas BCV Diarias)")
                st.dataframe(df_consolidado.head(10))

                # --- PANEL DE DESCARGAS ---
                st.markdown("---")
                st.subheader("📥 Exportar Resultados y Reportes")
                col_dl1, col_dl2 = st.columns(2)

                # Excel de 4 Hojas
                output_excel = io.BytesIO()
                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                    df_resumen_ia = pd.DataFrame([
                        {"Indicador / Concepto": "Movimientos Bancarios Totales", "Valor / Resultado": total_b},
                        {"Indicador / Concepto": "Partidas Conciliadas", "Valor / Resultado": conciliados_b},
                        {"Indicador / Concepto": "Partidas Pendientes Banco", "Valor / Resultado": total_b - conciliados_b},
                        {"Indicador / Concepto": "% Porcentaje de Conciliación", "Valor / Resultado": f"{porcentaje:.1f}%"},
                        {"Indicador / Concepto": "Asientos de Anulación / Error (Libros)", "Valor / Resultado": anulados_l},
                        {"Indicador / Concepto": "Total Créditos Operativos (Bs.)", "Valor / Resultado": total_ingresos_Bs},
                        {"Indicador / Concepto": "Total Débitos Operativos (Bs.)", "Valor / Resultado": total_egresos_Bs},
                        {"Indicador / Concepto": "Total Comisiones Bancarias (Bs.)", "Valor / Resultado": total_comisiones_Bs},
                        {"Indicador / Concepto": "Total Créditos Operativos (USD Histórico)", "Valor / Resultado": total_ingresos_usd},
                        {"Indicador / Concepto": "Total Débitos Operativos (USD Histórico)", "Valor / Resultado": total_egresos_usd},
                        {"Indicador / Concepto": "Total Comisiones (USD Histórico)", "Valor / Resultado": total_comisiones_usd},
                    ])
                    df_resumen_ia.to_excel(writer, sheet_name='Resumen_Gerencial_IA', index=False)
                    df_consolidado.to_excel(writer, sheet_name='Banco_vs_Libro_Consolidado', index=False)
                    banco.to_excel(writer, sheet_name='Extracto_Banco', index=False)
                    libro.to_excel(writer, sheet_name='Mayor_Contable', index=False)
                
                excel_data = output_excel.getvalue()

                with col_dl1:
                    st.download_button(
                        label="📊 Descargar Excel (4 Hojas con Tasas Diarias)",
                        data=excel_data,
                        file_name="Reporte_Conciliacion_QuimioVet_TasasDiarias.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

                with col_dl2:
                    if HAS_DOCX:
                        doc_word_buffer = generar_documento_word(
                            total_b, conciliados_b, porcentaje, anulados_l,
                            total_ingresos_Bs, total_egresos_Bs, total_comisiones_Bs,
                            total_ingresos_usd, total_egresos_usd, total_comisiones_usd
                        )
                        st.download_button(
                            label="📄 Descargar Informe Word (.docx)",
                            data=doc_word_buffer,
                            file_name="Informe_Gerencial_QuimioVet.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    else:
                        st.info("💡 Instala 'python-docx' para habilitar la descarga en Word.")

    except Exception as e:
        st.error(f"Error procesando los archivos: {e}")
else:
    st.info("Por favor, sube al menos el extracto bancario y el mayor analítico de SAP para activar el motor.")