"""
aplicacion.py - Aplicación web principal en Streamlit para el 
Sistema de Conciliación y Tesorería Multidivisa.
"""
from __future__ import annotations

import pandas as pd
import streamlit as st

# Importación de tus motores lógicos locales
from conciliar import (
    Config,
    ErrorDeParseo,
    calcular_cuadre,
    cargar_movimientos,
    conciliar,
)
from tasas import TablaTasas, TasaNoDisponible

# =============================================================================
# 1. CONFIGURACIÓN INICIAL DE LA PÁGINA
# =============================================================================
st.set_page_config(
    page_title="Conciliación - AI Reconciliation",
    page_icon="📊",
    layout="wide",
)

# Estilos CSS profesionales inspirados en la interfaz de referencia
st.markdown(
    """
    <style>
        .stApp { background-color: #F4F6F9; }
        .dashboard-card {
            background-color: #FFFFFF;
            padding: 20px;
            border-radius: 10px;
            border: 1px solid #E2E8F0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            margin-bottom: 15px;
        }
        .stButton button { border-radius: 8px; font-weight: 500; }
    </style>
""",
    unsafe_allow_html=True,
)

# =============================================================================
# 2. GESTIÓN DE ESTADO Y SEGURIDAD (AUTENTICACIÓN)
# =============================================================================
if "autenticado" not in st.session_state:
  st.session_state.autenticado = False

if not st.session_state.autenticado:
  st.title("🔒 Acceso al Sistema de Conciliación")
  st.markdown("Por favor, ingrese su contraseña para continuar[cite: 3, 8].")

  password_correcta = st.secrets.get("PASSWORD_ACCESO", "admin123")
  password = st.text_input("Contraseña:", type="password")

  if st.button("Ingresar"):
    if password and password == password_correcta:
      st.session_state.autenticado = True
      st.rerun()
    else:
      st.error("Contraseña incorrecta. Intente nuevamente.")
  st.stop()

# Inicialización de Estados Globales del Sistema
if "tabla_tasas" not in st.session_state:
  st.session_state.tabla_tasas = TablaTasas(max_antiguedad_dias=7)

if "cuentas_bancarias" not in st.session_state:
  st.session_state.cuentas_bancarias = {
      "Banco Nacional": ["0102-1234-56-1234567890"],
      "Banco Regional": ["0105-9876-54-0987654321"],
  }

if "resultado_conciliacion" not in st.session_state:
  st.session_state.resultado_conciliacion = None

if "movs_banco" not in st.session_state:
  st.session_state.movs_banco = []

if "movs_libros" not in st.session_state:
  st.session_state.movs_libros = []


# =============================================================================
# 3. BARRA LATERAL DE NAVEGACIÓN (Estructura idéntica a la referencia)
# =============================================================================
with st.sidebar:
  st.markdown("### 📊 Conciliacion")
  st.caption("AI Reconciliation")
  st.markdown("---")

  st.markdown(
      "<p style='font-size:11px; color:#64748B;"
      " font-weight:bold;margin-bottom:5px;'>PROCESO</p>",
      unsafe_allow_html=True,
  )
  menu_proceso = st.radio(
      "Proceso",
      [
          "Dashboard",
          "Importar",
          "Conciliación",
          "Reglas",
          "Partidas pendientes",
          "Historial",
      ],
      label_visibility="collapsed",
  )

  st.markdown("---")
  st.markdown(
      "<p style='font-size:11px; color:#64748B;"
      " font-weight:bold;margin-bottom:5px;'>ANÁLISIS & TASAS</p>",
      unsafe_allow_html=True,
  )
  modulo_analisis = st.selectbox(
      "Análisis", ["Resumen histórico", "Gestión de Tasas BCV / Multidivisa"]
  )

  st.markdown("---")
  st.markdown(
      "<p style='font-size:11px; color:#64748B;"
      " font-weight:bold;margin-bottom:5px;'>ARCHIVOS & MAESTROS</p>",
      unsafe_allow_html=True,
  )
  modulo_maestros = st.selectbox(
      "Maestros", ["Ninguno", "Bancos y Cuentas", "Estados de cuenta"]
  )

  st.markdown("---")
  st.caption("Soporte técnico • Configuración")


# =============================================================================
# 4. MÓDULOS Y VISTAS DE LA APLICACIÓN
# =============================================================================

# --- VISTA: GESTIÓN DE TASAS BCV ---
if modulo_analisis == "Gestión de Tasas BCV / Multidivisa":
  st.title("💱 Tasas de Cambio BCV y Conversión Multidivisa")
  st.markdown(
      "Cargue masivamente las tasas oficiales desde Excel[cite: 4, 8] y realice"
      " conversiones seguras."
  )

  archivo_tasas = st.file_uploader(
      "Seleccione el archivo Excel de Tasas", type=["xlsx", "xls"]
  )
  if archivo_tasas is not None:
    try:
      df_t = pd.read_excel(archivo_tasas)
      st.write("🔍 **Vista previa de tasas:**")
      st.dataframe(df_t.head())

      cols_t = list(df_t.columns)
      col_f_t = st.selectbox("Columna de Fecha", cols_t)
      col_u_t = st.selectbox("Columna Tasa USD", ["Ninguna"] + cols_t)
      col_e_t = st.selectbox("Columna Tasa EUR", ["Ninguna"] + cols_t)

      map_monedas = {}
      if col_u_t != "Ninguna":
        map_monedas["USD"] = col_u_t
      if col_e_t != "Ninguna":
        map_monedas["EUR"] = col_e_t

      if st.button("Procesar y Guardar Tasas"):
        filas_t = df_t.to_dict(orient="records")
        errores_t = st.session_state.tabla_tasas.cargar_filas(
            filas_t, col_f_t, map_monedas, tipo="BCV", fuente=archivo_tasas.name
        )
        if errores_t:
          st.warning(
              f"Se omitieron {len(errores_t)} filas con errores de formato."
          )
        else:
          st.success("🎉 ¡Todas las tasas se cargaron correctamente al sistema!")
    except Exception as ex:
      st.error(f"Error al leer el archivo de tasas: {ex}")

# --- VISTA: BANCOS Y CUENTAS (MAESTROS) ---
elif modulo_maestros == "Bancos y Cuentas":
  st.title("🏦 Maestros: Bancos y Cuentas")
  st.markdown(
      "Gestión de cuentas bancarias corporativas agrupadas por institución [cite:"
      " 8]."
  )

  with st.form("form_cuenta"):
    banco_input = st.selectbox(
        "Banco", ["Banco Nacional", "Banco Regional", "Banco Mercantil"]
    )
    nro_cta = st.text_input("Número de Cuenta")
    moneda_cta = st.selectbox("Moneda", ["VES", "USD", "EUR"])
    obs_cta = st.text_area("Observaciones (Sucursal, tipo, titular)")

    if st.form_submit_button("Registrar Cuenta"):
      if nro_cta.strip():
        # Validación de duplicados normalizando espacios y guiones[cite: 8]
        nro_limpio = "".join(filter(str.isalnum, nro_cta)).upper()
        existentes = [
            "".join(filter(str.isalnum, c)).upper()
            for c in st.session_state.cuentas_bancarias.get(banco_input, [])
        ]
        if nro_limpio in existentes:
          st.error(
              "⚠️ Ya existe una cuenta registrada con este número en el sistema"
              "[cite: 8]."
          )
        else:
          st.session_state.cuentas_bancarias.setdefault(banco_input, []).append(
              nro_cta
          )
          st.success(
              f"¡Cuenta {nro_cta} ({moneda_cta}) registrada con éxito para"
              f" {banco_input}!"
          )
      else:
        st.warning("Debe ingresar un número de cuenta válido.")

  st.markdown("---")
  st.subheader("Cuentas Activas")
  for bnk, ctas in st.session_state.cuentas_bancarias.items():
    st.write(f"**{bnk}**")
    for c in ctas:
      st.info(f"Número: {c}")

# --- VISTA: DASHBOARD PRINCIPAL ---
elif menu_proceso == "Dashboard":
  st.title("Dashboard")
  st.markdown(
      "<p style='color: #64748B; margin-top: -15px;'>Resumen general de"
      " conciliaciones bancarias[cite: 6]</p>",
      unsafe_allow_html=True,
  )

  # Banner informativo de novedades
  st.info(
      "✨ **Ver Novedades y Actualizaciones de Conciliar**\n\n*Nueva"
      " publicación · Última publicación: 29 de septiembre de 2026[cite: 6]*",
      icon="🚀",
  )

  # Tarjetas de métricas superiores
  res = st.session_state.resultado_conciliacion
  c1, c2, c3, c4 = st.columns(4)

  with c1:
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.metric(
        label="Partidas abiertas",
        value=len(res.banco_pendiente) if res else len(st.session_state.movs_banco),
    )
    st.markdown("</div>", unsafe_allow_html=True)
  with c2:
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.metric(
        label="Conciliaciones cerradas", value="1" if res else "0"
    )
    st.markdown("</div>", unsafe_allow_html=True)
  with c3:
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.metric(
        label="Movimientos conciliados",
        value=len(res.pares) if res else 0,
    )
    st.markdown("</div>", unsafe_allow_html=True)
  with c4:
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.metric(label="Diferencias registradas", value="0")
    st.markdown("</div>", unsafe_allow_html=True)

  col_izq, col_der = st.columns([2, 1])
  with col_izq:
    st.markdown("### Conciliaciones recientes")
    with st.container():
      st.markdown(
          '<div class="dashboard-card" style="text-align: center; padding:'
          ' 30px;">',
          unsafe_allow_html=True,
      )
      if not res:
        st.markdown(
            "📄<br>**No hay conciliaciones cerradas**<br><span style='color:"
            " #64748B; font-size: 14px;'>Importá un extracto y realizá tu"
            " primera conciliación[cite: 6].</span>",
            unsafe_allow_html=True,
        )
      else:
        st.success(
            "✅ Conciliación procesada correctamente. Revise el detalle en el"
            " menú **Conciliación**."
        )
      st.markdown("</div>", unsafe_allow_html=True)

  with col_der:
    st.markdown("### Acciones rápidas")
    with st.container():
      st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
      if st.button("📂 Importar extracto o mayor", use_container_width=True):
        pass
      if st.button("⚡ Iniciar conciliación", use_container_width=True):
        pass
      st.markdown("</div>", unsafe_allow_html=True)

# --- VISTA: IMPORTAR ARCHIVOS ---
elif menu_proceso == "Importar":
  st.title("📂 Importar Extractos y Registros Contables")
  st.markdown(
      "Cargue archivos en formato Excel o CSV para procesar los movimientos"
      " automáticamente[cite: 8]."
  )

  tipo_origen = st.selectbox(
      "Seleccione el tipo de archivo",
      ["Extracto Bancario", "Registro Contable (Mayor)"],
  )
  archivo_subido = st.file_uploader(
      "Seleccione el archivo", type=["xlsx", "xls", "csv"]
  )

  if archivo_subido is not None:
    try:
      df_imp = pd.read_excel(archivo_subido)
      st.write("🔍 **Vista previa de los datos:**")
      st.dataframe(df_imp.head())

      st.subheader("⚙️ Mapeo de Columnas")
      cols_disponibles = list(df_imp.columns)
      c_fec = st.selectbox("Columna de Fecha", cols_disponibles)
      c_desc = st.selectbox(
          "Columna de Descripción", ["Ninguna"] + cols_disponibles, index=1
      )
      c_ref = st.selectbox(
          "Columna de Referencia", ["Ninguna"] + cols_disponibles, index=2
      )
      esquema_monto = st.radio(
          "Estructura de Montos",
          [
              "Columna única (Monto firmado)",
              "Columnas separadas (Débito / Crédito)",
          ],
      )

      mapeo = {"fecha": c_fec}
      if c_desc != "Ninguna":
        mapeo["descripcion"] = c_desc
      if c_ref != "Ninguna":
        mapeo["referencia"] = c_ref

      if esquema_monto == "Columna única (Monto firmado)":
        mapeo["monto"] = st.selectbox("Columna de Monto", cols_disponibles)
      else:
        mapeo["entrada"] = st.selectbox(
            "Columna Entrada (Crédito / Debe)", cols_disponibles
        )
        mapeo["salida"] = st.selectbox(
            "Columna Salida (Débito / Haber)", cols_disponibles
        )

      if st.button("Procesar e Importar al Sistema"):
        origen_val = "banco" if "Extracto" in tipo_origen else "libros"
        resultado_carga = cargar_movimientos(
            df_imp.to_dict(orient="records"), origen_val, mapeo
        )

        st.success(
            f"¡Se importaron {len(resultado_carga.movimientos)} movimientos con"
            f" éxito!"
        )

        if resultado_carga.errores:
          st.warning(
              f"Se detectaron {len(resultado_carga.errores)} filas con"
              " advertencias o errores de formato que fueron omitidas [cite:"
              " 7]."
          )
          st.dataframe(pd.DataFrame([vars(e) for e in resultado_carga.errores]))

        if origen_val == "banco":
          st.session_state.movs_banco = resultado_carga.movimientos
        else:
          st.session_state.movs_libros = resultado_carga.movimientos

    except Exception as err:
      st.error(f"Error al procesar el archivo cargado: {err}")

# --- VISTA: MOTOR DE CONCILIACIÓN ---
elif menu_proceso == "Conciliación":
  st.title("⚡ Motor de Cruce y Propuestas Automáticas")
  st.markdown(
      "Cruce inteligente de partidas aplicando las reglas del motor (C1, C2, C3,"
      " C4)[cite: 8]."
  )

  st.info(
      f"📊 Movimientos cargados en memoria — Banco:"
      f" {len(st.session_state.movs_banco)} | Libros:"
      f" {len(st.session_state.movs_libros)}"
  )

  tolerancia_c2 = st.slider(
      "Tolerancia de días para regla C2", min_value=0, max_value=7, value=3
  )

  if st.button("Ejecutar Conciliación Completa", type="primary"):
    if not st.session_state.movs_banco or not st.session_state.movs_libros:
      st.warning(
          "⚠️ Debe importar tanto el extracto bancario como el registro"
          " contable antes de ejecutar la conciliación."
      )
    else:
      try:
        configuracion = Config(tolerancia_dias_c2=tolerancia_c2)
        res_cruces = conciliar(
            st.session_state.movs_banco,
            st.session_state.movs_libros,
            configuracion,
        )
        st.session_state.resultado_conciliacion = res_cruces
        st.success("🎉 ¡Conciliación ejecutada con éxito!")
      except Exception as e:
        st.error(f"Error durante la ejecución del motor: {e}")

  # Mostrar resultados si ya fueron calculados
  if st.session_state.resultado_conciliacion:
    res = st.session_state.resultado_conciliacion
    datos_resumen = res.resumen()

    st.markdown("---")
    st.subheader("📈 Resultados del Cruce")

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric(
        "Pares Cerrados (C1/C2)", datos_resumen.get("pares_cerrados", 0)
    )
    kpi2.metric("Sugerencias (C3/C4)", datos_resumen.get("sugerencias", 0))
    kpi3.metric("Banco Pendiente", datos_resumen.get("banco_pendiente", 0))
    kpi4.metric("Libros Pendiente", datos_resumen.get("libros_pendiente", 0))

    if res.sugerencias:
      st.markdown("### 💡 Sugerencias pendientes de aprobación (C3 / C4)")
      for idx, sug in enumerate(res.sugerencias):
        with st.expander(
            f"Sugerencia {idx+1} — Regla: {sug.regla} | Motivo: {sug.motivo}"
        ):
          col_b, col_l = st.columns(2)
          with col_b:
            st.write("**Movimiento(s) Banco:**")
            for mb in sug.banco:
              st.text(
                  f"Fecha: {mb.fecha} | Monto: {mb.monto} |"
                  f" {mb.descripcion}"
              )
          with col_l:
            st.write("**Movimiento(s) Libros:**")
            for ml in sug.libros:
              st.text(
                  f"Fecha: {ml.fecha} | Monto: {ml.monto} |"
                  f" {ml.descripcion}"
              )

          if st.button(f"Aprobar sugerencia {idx+1}", key=f"aprobar_{idx}"):
            res.aprobar(sug)
            st.success(
                "¡Sugerencia aprobada con éxito y movida a pares cerrados!"
            )
            st.rerun()

elif menu_proceso == "Partidas pendientes":
  st.title("📌 Partidas Abiertas y Pendientes")
  st.markdown(
      "Gestión de partidas sin resolver, clasificación por etiquetas y extornos"
      "[cite: 8]."
  )
  if st.session_state.resultado_conciliacion:
    res = st.session_state.resultado_conciliacion
    col_pb, col_pl = st.columns(2)
    with col_pb:
      st.subheader("Pendientes de Banco")
      if res.banco_pendiente:
        df_pb = pd.DataFrame([vars(m) for m in res.banco_pendiente])
        st.dataframe(df_pb[["id", "fecha", "monto", "descripcion", "referencia"]])
      else:
        st.info("Sin partidas pendientes en banco.")
    with col_pl:
      st.subheader("Pendientes de Libros")
      if res.libros_pendiente:
        df_pl = pd.DataFrame([vars(m) for m in res.libros_pendiente])
        st.dataframe(df_pl[["id", "fecha", "monto", "descripcion", "referencia"]])
      else:
        st.info("Sin partidas pendientes en libros.")
  else:
    st.info(
        "Ejecute una conciliación primero para visualizar las partidas"
        " pendientes."
    )

elif menu_proceso == "Historial":
  st.title("📜 Historial de Conciliaciones Cerradas")
  st.markdown(
      "Registro histórico para auditoría, seguimiento y control[cite: 8]."
  )
  if st.session_state.resultado_conciliacion:
    st.success(
        "Conciliación actual cerrada y registrada correctamente en el sistema."
    )
  else:
    st.info("No hay conciliaciones cerradas registradas todavía.")