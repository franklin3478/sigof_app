import streamlit as st
from modulo_selfie import ejecutar_selfie
from modulo_fise import ejecutar_fise

st.set_page_config(page_title="Sistema SIGOF", layout="wide")

# ----------------------------
# USUARIOS DEL SISTEMA
# ----------------------------
USUARIOS = {
    "huancayo": "2026"
 }

# ----------------------------
# ESTADO DE SESIÓN
# ----------------------------
if "logueado" not in st.session_state:
    st.session_state.logueado = False

# ----------------------------
# LOGIN
# ----------------------------
if not st.session_state.logueado:
    st.title("🔐 Login del Sistema")

    user = st.text_input("Usuario")
    password = st.text_input("Contraseña", type="password")

    if st.button("Ingresar"):
        if user in USUARIOS and USUARIOS[user] == password:
            st.session_state.logueado = True
            st.success("Bienvenido al sistema")
            st.rerun()  # 🔥 CLAVE
        else:
            st.error("Credenciales incorrectas")

# ----------------------------
# SISTEMA PRINCIPAL
# ----------------------------
else:
    
    # BOTÓN CERRAR SESIÓN
    if st.sidebar.button("Cerrar sesión"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]

        st.rerun()

    # MENÚ LATERAL
    opcion = st.sidebar.selectbox(
        "Selecciona módulo",
        ["Selfie Lectura", "Validación FISE"]
    )

    # NAVEGACIÓN
    if opcion == "Selfie Lectura":
        ejecutar_selfie()

    elif opcion == "Validación FISE":
        ejecutar_fise()