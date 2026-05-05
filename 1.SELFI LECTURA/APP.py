import streamlit as st
from modulo_selfie import ejecutar_selfie
from modulo_fise import ejecutar_fise
from modulo_galeria_reparto import ejecutar_galeria_reparto
from modulo_galeria_lectura import ejecutar_galeria_lectura
from modulo_seguimiento_reparto import ejecutar_seguimiento_reparto 

st.set_page_config(page_title="Sistema SIGOF", layout="wide")

# ----------------------------
# USUARIOS DEL SISTEMA
# ----------------------------
USUARIOS = {
    "huancayo": {"password": "2026", "rol": "admin"},
    "aria": {"password": "abcd", "rol": "usuario"},
    "juan": {"password": "5678", "rol": "usuario"}
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
            st.rerun()
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
        [
            "Selfie Lectura",
            "Validación FISE",
            "Galería Fotos Reparto",
            "Galería Fotos Lectura",
            "Seguimiento Reparto", 
        ]
    )

    # NAVEGACIÓN
    if opcion == "Selfie Lectura":
        ejecutar_selfie()

    elif opcion == "Validación FISE":
        ejecutar_fise()

    elif opcion == "Galería Fotos Reparto":
        ejecutar_galeria_reparto()

    elif opcion == "Galería Fotos Lectura":
        ejecutar_galeria_lectura()

    elif opcion == "Seguimiento Reparto": 
        ejecutar_seguimiento_reparto()