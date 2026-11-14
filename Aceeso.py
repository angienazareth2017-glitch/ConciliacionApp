import streamlit as st

st.set_page_config(page_title="Acceso al Sistema", page_icon="🔒", layout="centered")

# Recupera la contraseña desde los secretos de Streamlit (o usa una por defecto)
password_correcta = st.secrets.get("PASSWORD_ACCESO",)

if "autenticado" not in st.session_state:
  st.session_state.autenticado = False

if not st.session_state.autenticado:
  st.title("🔒 Acceso al Sistema de Conciliación")
  st.markdown("Por favor, ingrese su contraseña para acceder al sistema.")

  password = st.text_input("Contraseña:", type="password")

  if st.button("Ingresar"):
    if password == password_correcta:
      st.session_state.autenticado = True
      st.rerun()
    else:
      st.error("Contraseña incorrecta. Intente nuevamente.")
  st.stop()

# Si ya está autenticado, aquí continúa tu aplicación principal
st.success("¡Acceso concedido!")