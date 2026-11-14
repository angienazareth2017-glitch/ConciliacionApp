import streamlit as st

st.set_page_config(
    page_title="Sistema de Conciliación y Tesorería",
    page_icon="🔒",
    layout="wide",
)

# Inicializamos el estado de autenticación
if "autenticado" not in st.session_state:
  st.session_state.autenticado = False

# Pantalla de Login (si no está autenticado)
if not st.session_state.autenticado:
  st.title("🔒 Acceso al Sistema de Conciliación")
  st.markdown("Por favor, ingrese su contraseña para continuar.")

  # Obtenemos la contraseña exclusivamente desde los secrets de Streamlit Cloud
  password_correcta = st.secrets.get("PASSWORD_ACCESO", "")

  password = st.text_input("Contraseña:", type="password")

  if st.button("Ingresar"):
    if password and password == password_correcta:
      st.session_state.autenticado = True
      st.rerun()
    else:
      st.error("Contraseña incorrecta. Intente nuevamente.")

  # Detenemos la ejecución aquí para que no muestre el resto de la app
  st.stop()

# =============================================================================
# A PARTIR DE AQUÍ COMIENZA TU APLICACIÓN PRINCIPAL (Solo visible al autenticarse)
# =============================================================================
st.success("¡Bienvenido al sistema!")

# Coloca aquí el resto de tu código de conciliación y procesamiento de archivos...