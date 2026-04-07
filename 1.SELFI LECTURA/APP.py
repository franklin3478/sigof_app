import streamlit as st
from modulo_selfie import ejecutar_selfie
from modulo_fise import ejecutar_fise

st.set_page_config(page_title="Sistema SIGOF", layout="wide")

# ----------------------------
# USUARIOS DEL SISTEMA
# ----------------------------
# "rol": "admin" -> tiene privilegios
# "rol": "usuario" -> solo puede interactuar con la app
USUARIOS = {
    "huancayo": {"password": "2026", "rol": "admin"},  # tu usuario principal
    "aria": {"password": "abcd", "rol": "usuario"},    # usuario normal
    "juan": {"password": "5678", "rol": "usuario"}     # otro usuario
}

# ----------------------------
# ESTADO DE SESIÓN
# ----------------------------
if "logueado" not in st.session_state:
    st.session_state.logueado = False
    st.session_state.usuario_actual = None
    st.session_state.rol = None

# ----------------------------
# LOGIN
# ----------------------------
if not st.session_state.logueado:
    st.title("🔐 Login del Sistema")

    user = st.text_input("Usuario")
    password = st.text_input("Contraseña", type="password")

    if st.button("Ingresar"):
        if user in USUARIOS and USUARIOS[user]["password"] == password:
            st.session_state.logueado = True
            st.session_state.usuario_actual = user
            st.session_state.rol = USUARIOS[user]["rol"]
            st.success(f"Bienvenido {user}")
            st.rerun()  # 🔥 CLAVE para refrescar la app
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

    # SOLO ADMIN PUEDE VER BOTONES DE CONFIGURACIÓN (opcional)
    if st.session_state.rol == "admin":
        st.sidebar.markdown("⚙️ **Admin Settings**")
        st.sidebar.button("Manage app")  # solo visible para admin

    # NAVEGACIÓN
    if opcion == "Selfie Lectura":
        ejecutar_selfie()

    elif opcion == "Validación FISE":
        ejecutar_fise()