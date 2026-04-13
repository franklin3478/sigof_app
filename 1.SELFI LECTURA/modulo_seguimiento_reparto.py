import streamlit as st
import requests
import re
from datetime import datetime
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import pandas as pd
import math
from io import BytesIO

# ----------------------------
# UNIDADES
# ----------------------------
mapa_unidades = {
    "Valle Mantaro": {"idempresa": "4", "iduunn": "83"},
    "Tarma": {"idempresa": "4", "iduunn": "79"},
    "Huancavelica": {"idempresa": "4", "iduunn": "78"},
    "Ayacucho": {"idempresa": "4", "iduunn": "76"},
    "Selva Central": {"idempresa": "4", "iduunn": "80"},
    "Huánuco": {"idempresa": "4", "iduunn": "82"},
    "Pasco": {"idempresa": "4", "iduunn": "81"},
    "Huancayo": {"idempresa": "4", "iduunn": "77"},
    "Tingo María": {"idempresa": "4", "iduunn": "84"},
}

# ----------------------------
# LIMPIEZA GLOBAL
# ----------------------------
def limpiar_datos_reparto():
    for k in [
        "df_reparto",
        "resumen_reparto",
        "imagenes_reparto",
        "pagina_reparto",
        "filtros_reparto",
        "firma_actual"
    ]:
        st.session_state.pop(k, None)

# ----------------------------
# LOGIN
# ----------------------------
def login_sigof():

    st.subheader("🔐 Login SIGOF")

    usuario = st.text_input("Usuario")
    password = st.text_input("Contraseña", type="password")

    if st.button("Ingresar"):

        if not usuario or not password:
            st.warning("Ingresa credenciales")
            return

        with st.spinner("Conectando a SIGOF..."):
            session = requests.Session()
            r = session.post(
                "http://sigof.distriluz.com.pe/plus/usuario/login",
                data={
                    "data[Usuario][usuario]": usuario,
                    "data[Usuario][pass]": password
                }
            )

        if "Salir" in r.text:
            st.session_state.session_sigof_reparto = session
            st.session_state.logueado_sigof_reparto = True
            st.session_state.usuario_sigof = usuario
            st.rerun()
        else:
            st.error("Credenciales incorrectas")

# ----------------------------
# CICLOS
# ----------------------------
def obtener_ciclos(session, unidad):

    hoy = datetime.today().strftime("%Y-%m-%d")
    u = mapa_unidades[unidad]

    session.post(
        "http://sigof.distriluz.com.pe/plus/usuario/ajax_cambiar_sesion",
        data=u
    )

    url = f"http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_listar_tabla_repartos_historico/U/0/0/0/0/{hoy}/{hoy}/0/"

    r = session.get(url, headers={"X-Requested-With": "XMLHttpRequest"})

    return {
        idc: f"{idc}-{desc}".strip()
        for idc, desc in re.findall(r'(\d+)-(Ciclo[^<"]+)', r.text)
    }

# ----------------------------
# HELPERS
# ----------------------------
def obtener_id_ciclo(ciclo_texto):
    return ciclo_texto.split("-")[0]

def descargar_excel_ciclo(session, unidad, idc):

    u = mapa_unidades[unidad]

    session.post(
        "http://sigof.distriluz.com.pe/plus/usuario/ajax_cambiar_sesion",
        data=u
    )

    url = f"http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_reporte_excel_ordenes_historico/U/0/{idc}/0/0/2026-01-14/2026-01-14/0/"

    r = session.get(url)

    if r.status_code == 200 and r.content:
        return load_workbook(BytesIO(r.content))

    return None

def descargar_ciclos_excel(session, unidad, ciclos_seleccionados):

    resultados = []

    with ThreadPoolExecutor(max_workers=4) as executor:
        tareas = []

        for ciclo in ciclos_seleccionados:
            idc = obtener_id_ciclo(ciclo)
            tareas.append(executor.submit(descargar_excel_ciclo, session, unidad, idc))

        for f in as_completed(tareas):
            wb = f.result()
            if wb:
                resultados.append(wb)

    return resultados

# ----------------------------
# EXCEL → DATAFRAME
# ----------------------------
def excel_a_dataframe(wb):

    ws = wb.active

    foto_col_idx = None

    for idx, cell in enumerate(ws[1], start=1):
        if str(cell.value).strip().lower() == "foto":
            foto_col_idx = idx
            break

    foto_links = {}

    if foto_col_idx:
        foto_col_letter = get_column_letter(foto_col_idx)

        for cell in ws[foto_col_letter]:
            if cell.hyperlink:
                foto_links[cell.row] = cell.hyperlink.target

    data = list(ws.values)
    df = pd.DataFrame(data)

    df.columns = df.iloc[0]
    df = df[1:]
    df.columns = df.columns.str.strip()

    if "foto" in df.columns:
        df["foto"] = df.apply(
            lambda row: foto_links.get(row.name + 1, row["foto"]),
            axis=1
        )
        df["foto"] = df["foto"].astype(str)

    return df

# ----------------------------
# PROCESAR
# ----------------------------
def procesar_dataframe(df):

    df.columns = df.columns.str.strip()

    df["fise"] = df["fise"].astype(str).str.strip()
    df["foto"] = df["foto"].astype(str)

    df["tiene_foto"] = df["foto"].astype(str).str.startswith("http")

    df["fise_es_1"] = df["fise"].isin(["1", "1.0"])

    resumen = (
        df.groupby("ruta")
        .agg(
            lecturista=("lecturista", "first"),
            ciclo=("idciclo", "first"),
            total_suministros=("ruta", "count"),
            fise_obligatorio=("fise_es_1", "sum"),
            fotos_tomadas=("tiene_foto", "sum")
        )
        .reset_index()
    )

    def calc(total):
        return math.ceil(total * 0.10) if total < 10 else math.floor(total * 0.10)

    resumen["cant_min_foto"] = resumen["total_suministros"].apply(calc)

    resumen["consignado"] = resumen.apply(
        lambda r: "NO CUENTA CON FOTOS" 
        if r["fotos_tomadas"] == 0
        else ("CUMPLIÓ EL 10%" if r["fotos_tomadas"] >= r["cant_min_foto"] else "NO CUMPLIÓ EL 10%"),
        axis=1
    )

    df["fise_con_foto_flag"] = df["fise_es_1"] & df["tiene_foto"]

    resumen["fise_con_foto"] = (
        df.groupby("ruta")["fise_con_foto_flag"].sum().values
    )

    resumen["fise_sin_foto"] = (
        resumen["fise_obligatorio"] - resumen["fise_con_foto"]
    )
       
    resumen["cumplimiento_fise"] = resumen.apply(
       lambda r: "CUMPLIÓ" if r["fise_con_foto"] == r["fise_obligatorio"]
       else "NO CUMPLIÓ",
       axis=1
    )

    return resumen

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ----------------------------
# MAIN
# ----------------------------
def ejecutar_seguimiento_reparto():

    for k, v in {
        "logueado_sigof_reparto": False,
        "session_sigof_reparto": None,
        "ciclos_reparto": {},
        "unidad_actual_reparto": None,
        "ultima_actualizacion": None,
        "mensaje_ciclos": None,
        "tipo_mensaje": None,
        "ciclos_previos": []
    }.items():
        st.session_state.setdefault(k, v)

    if not st.session_state.logueado_sigof_reparto:
        login_sigof()
        return

    session = st.session_state.session_sigof_reparto

    st.title("📦 Seguimiento de Reparto")
    st.caption(f"Usuario: {st.session_state.usuario_sigof}")

    col1, col2 = st.columns(2)

    with col1:
        unidad = st.selectbox("Unidad", list(mapa_unidades.keys()))

    if unidad != st.session_state.unidad_actual_reparto:
        st.session_state.unidad_actual_reparto = unidad
        st.session_state.ciclos_reparto = {}
        st.session_state.ultima_actualizacion = None
        limpiar_datos_reparto()

    with col2:
        ciclos_dict = st.session_state.ciclos_reparto

        ciclos_seleccionados = (
            st.multiselect("Ciclos", list(ciclos_dict.values()))
            if ciclos_dict else []
        )

        if not ciclos_dict:
            st.selectbox("Ciclos", ["-- obtener primero --"], disabled=True)

    if set(ciclos_seleccionados) != set(st.session_state.ciclos_previos):
        limpiar_datos_reparto()
        st.session_state.ciclos_previos = ciclos_seleccionados

    if st.session_state.ultima_actualizacion:
        st.caption(f"Última actualización: {st.session_state.ultima_actualizacion}")

    if st.session_state.mensaje_ciclos:
        getattr(st, st.session_state.tipo_mensaje)(st.session_state.mensaje_ciclos)
        st.session_state.mensaje_ciclos = None

    colb1, colb2, colb3 = st.columns(3)

    with colb1:
        if st.button("📡 Obtener ciclos"):
            if st.session_state.ciclos_reparto:
                st.session_state.mensaje_ciclos = "Ya cargados (usa actualizar)"
                st.session_state.tipo_mensaje = "info"
            else:
                with st.spinner("Consultando..."):
                    ciclos = obtener_ciclos(session, unidad)

                if ciclos:
                    limpiar_datos_reparto()
                    st.session_state.ciclos_reparto = ciclos
                    st.session_state.ultima_actualizacion = datetime.now().strftime("%H:%M:%S")
                    st.session_state.mensaje_ciclos = f"{len(ciclos)} ciclos encontrados"
                    st.session_state.tipo_mensaje = "success"

                else:
                    st.session_state.mensaje_ciclos = "Sin resultados"
                    st.session_state.tipo_mensaje = "warning"

            st.rerun()

    with colb2:
        if st.button("🔄 Actualizar"):
            with st.spinner("Verificando cambios..."):
                nuevos = obtener_ciclos(session, unidad)

            if nuevos != st.session_state.ciclos_reparto:
                limpiar_datos_reparto()
                st.session_state.ciclos_reparto = nuevos
                st.session_state.ultima_actualizacion = datetime.now().strftime("%H:%M:%S")
                st.session_state.mensaje_ciclos = f"Actualizado: {len(nuevos)} ciclos"
                st.session_state.tipo_mensaje = "success"
            else:
                st.session_state.mensaje_ciclos = "No hay nuevos datos"
                st.session_state.tipo_mensaje = "info"

            st.rerun()

    with colb3:
        if st.button("🔓 Cerrar sesión"):
            st.session_state.clear()
            st.rerun()

    if ciclos_seleccionados:

        st.markdown("---")

        if st.button("📊 Procesar ciclos"):
            with st.spinner("Descargando y procesando..."):
                wbs = descargar_ciclos_excel(session, unidad, ciclos_seleccionados)
                dfs = [excel_a_dataframe(wb) for wb in wbs]
                df_total = pd.concat(dfs, ignore_index=True)
                resumen = procesar_dataframe(df_total)
                st.session_state.resumen_reparto = resumen

            st.success("✅ Procesamiento completo")

    # ----------------------------
    # FILTROS (RUTA DEPENDE DE LECTURISTA + CICLO)
    # ----------------------------
    if "resumen_reparto" in st.session_state:

        st.markdown("---")
        df = st.session_state.resumen_reparto.copy()

        st.subheader("🔎 Filtros")

        # ----------------------------
        # LECTURISTA OPTIONS (depende de ciclo + ruta)
        # ----------------------------
        df_lect = df.copy()

        # CICLO FILTRO TEMPORAL
        ciclos_tmp = st.session_state.get("filtro_ciclo", [])
        rutas_tmp = st.session_state.get("filtro_ruta", [])

        if ciclos_tmp:
            df_lect = df_lect[df_lect["ciclo"].isin(ciclos_tmp)]

        if rutas_tmp:
            df_lect = df_lect[df_lect["ruta"].isin(rutas_tmp)]

        lecturistas_options = sorted(df_lect["lecturista"].dropna().unique())

        # ----------------------------
        # CICLO OPTIONS
        # ----------------------------
        df_ciclo = df.copy()

        lect_tmp = st.session_state.get("filtro_lecturista", [])
        rutas_tmp = st.session_state.get("filtro_ruta", [])

        if lect_tmp:
            df_ciclo = df_ciclo[df_ciclo["lecturista"].isin(lect_tmp)]

        if rutas_tmp:
            df_ciclo = df_ciclo[df_ciclo["ruta"].isin(rutas_tmp)]

        ciclos_options = sorted(df_ciclo["ciclo"].dropna().unique())

        # ----------------------------
        # RUTA OPTIONS
        # ----------------------------
        df_ruta = df.copy()

        lect_tmp = st.session_state.get("filtro_lecturista", [])
        ciclo_tmp = st.session_state.get("filtro_ciclo", [])

        if lect_tmp:
            df_ruta = df_ruta[df_ruta["lecturista"].isin(lect_tmp)]

        if ciclo_tmp:
            df_ruta = df_ruta[df_ruta["ciclo"].isin(ciclo_tmp)]

        rutas_options = sorted(df_ruta["ruta"].dropna().unique())

        # ----------------------------
        # WIDGETS (UNA SOLA VEZ)
        # ----------------------------
        colf1, colf2, colf3 = st.columns(3)

        with colf1:
            lecturistas = st.multiselect(
                "Lecturista",
                lecturistas_options,
                key="filtro_lecturista"
            )

        with colf2:
            ciclos = st.multiselect(
                "Ciclo",
                ciclos_options,
                key="filtro_ciclo"
            )

        with colf3:
            rutas = st.multiselect(
                "Ruta",
                rutas_options,
                key="filtro_ruta"
            )
        
        # 🔥 BASE PARA DEPENDENCIA DE RUTA
        df_ruta_base = df.copy()

        if lecturistas:
            df_ruta_base = df_ruta_base[df_ruta_base["lecturista"].isin(lecturistas)]

        if ciclos:
            df_ruta_base = df_ruta_base[df_ruta_base["ciclo"].isin(ciclos)]

                # FILTRADO FINAL
        if lecturistas:
            df = df[df["lecturista"].isin(lecturistas)]

        if ciclos:
            df = df[df["ciclo"].isin(ciclos)]

        if rutas:
            df = df[df["ruta"].isin(rutas)]

        st.subheader("📊 Cumplimiento FISE")

        df_fise = df[
            ["lecturista","ciclo","ruta","total_suministros",
             "fise_obligatorio","fise_con_foto","fise_sin_foto","cumplimiento_fise"]
        ]

        st.dataframe(df_fise, use_container_width=True)

        st.download_button(
            "⬇ Descargar FISE",
            data=to_excel(df_fise),
            file_name="fise.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("📊 Cumplimiento 10% Fotos")

        df_10 = df[
            ["lecturista","ciclo","ruta","total_suministros",
             "fotos_tomadas","cant_min_foto","consignado"]
        ]

        st.dataframe(df_10, use_container_width=True)

        st.download_button(
            "⬇ Descargar 10%",
            data=to_excel(df_10),
            file_name="cumplimiento_10.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )