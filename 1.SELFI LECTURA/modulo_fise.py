import streamlit as st
import requests
from datetime import datetime
from bs4 import BeautifulSoup
import pandas as pd
import re
from concurrent.futures import ThreadPoolExecutor, as_completed

# ----------------------------
# UNIDADES (COMPLETO)
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
# CAMBIAR UNIDAD
# ----------------------------
def cambiar_unidad(session, unidad):
    session.post(
        "http://sigof.distriluz.com.pe/plus/usuario/ajax_cambiar_sesion",
        data=mapa_unidades[unidad],
        timeout=20
    )

# ----------------------------
# LOGIN
# ----------------------------
def login_sigof():
    st.subheader("🔐 Login SIGOF")

    usuario = st.text_input("Usuario")
    password = st.text_input("Contraseña", type="password")

    if st.button("Ingresar SIGOF"):
        session = requests.Session()

        try:
            r = session.post(
                "http://sigof.distriluz.com.pe/plus/usuario/login",
                data={
                    "data[Usuario][usuario]": usuario,
                    "data[Usuario][pass]": password
                },
                timeout=30
            )

            if "Salir" in r.text:
                st.session_state.session_sigof = session
                st.session_state.logueado_sigof = True
                st.success("Login correcto")
                st.rerun()
            else:
                st.error("Credenciales incorrectas")

        except requests.exceptions.RequestException as e:
            st.error(f"Error de conexión: {e}")

# ----------------------------
# CICLOS
# ----------------------------
def obtener_ciclos(session, unidad, fecha):
    cambiar_unidad(session, unidad)

    url = f"http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_listar_tabla_repartos_historico/U/0/0/0/0/{fecha}/{fecha}/0/"

    r = session.get(url, headers={"X-Requested-With": "XMLHttpRequest"}, timeout=30)
    matches = re.findall(r'(\d+)-(Ciclo[^"<]+)', r.text)

    return {idc: f"{idc} - {desc.strip()}" for idc, desc in matches}

# ----------------------------
# MONITOREO
# ----------------------------
def obtener_monitoreo(session, unidad, fecha, idciclo):
    cambiar_unidad(session, unidad)

    url = f"http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_monitorear_reparto_reload/U/{fecha}/{fecha}/{idciclo}/0/0"

    r = session.get(url, headers={"X-Requested-With": "XMLHttpRequest"}, timeout=30)
    return r.text

# ----------------------------
# PARSEAR TABLA
# ----------------------------
def parsear_monitoreo(html):
    soup = BeautifulSoup(html, "html.parser")
    tabla = soup.find("table", {"id": "list-monitorear-lecturistas"})

    if tabla is None:
        return pd.DataFrame()

    filas = tabla.find("tbody").find_all("tr")
    data = []

    for fila in filas:
        cols = fila.find_all("td")
        if len(cols) < 18:
            continue

        # ----------------------------
        # 🔥 TIEMPO LIMPIO (HI / HF)
        # ----------------------------
        tiempo_raw = cols[11].get_text(" ", strip=True)

        match = re.search(
            r"(HI:\s*\d{2}:\d{2}:\d{2}).?(HF:\s\d{2}:\d{2}:\d{2})",
            tiempo_raw
        )

        if match:
            tiempo_limpio = f"{match.group(1)} / {match.group(2)}"
        else:
            tiempo_limpio = tiempo_raw

        # ----------------------------
        # 🔥 AVANCE LIMPIO
        # ----------------------------
        avance_raw = cols[14].get_text(strip=True)
        avance_limpio = re.sub(r"\s+", "", avance_raw)

        # ----------------------------
        # 🔥 DATOS BASE
        # ----------------------------
        asi = cols[9].get_text(strip=True)
        des = cols[10].get_text(strip=True)

        # ----------------------------
        # 🔥 CONDICIONES (CLAVE)
        # ----------------------------

        # NO INICIÓ → cuando tiempo está vacío o TT = 0
        if "TT: 00:00:00" in tiempo_raw or tiempo_limpio.strip() == "":
            tiempo_limpio = "NO INICIÓ"

        # NO DESCARGÓ → cuando des = 0
        if des == "0":
            des = "NO DESCARGÓ"

        # ----------------------------
        # DATA FINAL
        # ----------------------------
        td_repartidor = cols[8]

        repartidor = td_repartidor.get_text(strip=True)

        id_repartidor = (
            td_repartidor.get("data-empleado-id")
            or td_repartidor.get("empleado_id")
            or td_repartidor.get("id")
        )

        data.append({
            "id_repartidor": id_repartidor, 
            "repartidor": repartidor,
            "asi": asi,
            "des": des,
            "tiempo": tiempo_limpio,
            "fin": cols[12].get_text(strip=True),
            "p": cols[13].get_text(strip=True),
            "% avance": avance_limpio,
            "validados": cols[15].get_text(strip=True),
            "entregado": cols[16].get_text(strip=True),
            "paso_ruta": cols[17].get_text(strip=True),
        })

    return pd.DataFrame(data)
# ----------------------------
# FUNCIÓN PRINCIPAL
# ----------------------------
def obtener_monitoreo_sigof():
    st.session_state.setdefault("logueado_sigof", False)

    if not st.session_state.logueado_sigof:
        login_sigof()
        return None

    session = st.session_state.session_sigof
    fecha = datetime.today().strftime("%Y-%m-%d")

    # 🔥 SELECTOR DE UNIDAD (CLAVE)
    col1, col2 = st.columns(2)

    with col1:
        unidad = st.selectbox("Unidad", list(mapa_unidades.keys()))
    
    # 🔥 detectar cambio de unidad
    if st.session_state.get("unidad_actual") != unidad:
        st.session_state["unidad_actual"] = unidad
        st.session_state.pop("ciclos", None)
        st.rerun()

    # ----------------------------
    # CICLOS
    # ----------------------------
    if st.button("📡 Obtener ciclos"):
        st.session_state.ciclos = obtener_ciclos(session, unidad, fecha)

    ciclos_dict = st.session_state.get("ciclos", {})

    with col2:
        if ciclos_dict:
            ciclos_sel = st.multiselect("Ciclo",list(ciclos_dict.values()))
        else:
            ciclos_sel = None

    if not ciclos_dict or not ciclos_sel:
        st.info("Primero obtén ciclos")
        return None
    
    ids_ciclos = [list(ciclos_dict.keys())[list(ciclos_dict.values()).index(c)]for c in ciclos_sel]

    # ----------------------------
    # MONITOREO
    # ----------------------------
    mostrar_monitoreo = st.button("📊 Obtener Monitoreo")

    if mostrar_monitoreo:
        st.session_state["cargar_monitoreo"] = True

    if st.session_state.get("cargar_monitoreo", False):

        dfs = []

        for idciclo in ids_ciclos:
            html = obtener_monitoreo(session, unidad, fecha, idciclo)
            df_tmp = parsear_monitoreo(html)

            if not df_tmp.empty:
                df_tmp["ciclo"] = ciclos_dict[idciclo]
                cols = df_tmp.columns.tolist()

                if "ciclo" in cols and "repartidor" in cols:
                    cols.remove("ciclo")
                    idx = cols.index("repartidor")
                    cols.insert(idx, "ciclo")

                df_tmp = df_tmp[cols]
                dfs.append(df_tmp)

        if not dfs:
            st.warning("No hay datos")
            return None

        df_final = pd.concat(dfs, ignore_index=True)

        # Guardar completo para cálculos internos
        st.session_state.df_final = df_final

        # ----------------------------
        # FILTRO POR LECTURISTA
        # ----------------------------
        lecturistas = sorted(
            df_final["repartidor"]
            .dropna()
            .astype(str)
            .unique()
        )

        lecturistas = sorted(
            df_final["repartidor"]
            .dropna()
            .astype(str)
            .unique()
        )

        lecturistas_sel = st.multiselect(
            "👤 Lecturista",
            lecturistas,
            placeholder="Seleccione uno o varios lecturistas"
        )

        if lecturistas_sel:

            df_filtrado = df_final[
                df_final["repartidor"].astype(str).isin(lecturistas_sel)
            ].copy()

        else:

            # Si no selecciona ninguno → TODOS
            df_filtrado = df_final.copy()

        # ----------------------------
        # MOSTRAR TABLA SEGÚN FILTRO
        # ----------------------------

        df_mostrar = df_filtrado.drop(
            columns=["id_repartidor"],
            errors="ignore"
        )

        st.dataframe(
            df_mostrar,
            use_container_width=True,
            hide_index=True
        )

        # 🔥 AQUÍ VA TODO LO NUEVO
        
        resultado = {}
        resultado_fise = {}

        iduunn = mapa_unidades[unidad]["iduunn"]

        for _, row in df_filtrado.iterrows():

            nombre = row["repartidor"]
            id_rep = row["id_repartidor"]
            ciclo_txt = row["ciclo"]

            clave = f"{id_rep} - {nombre}"

            idciclo = [
                k for k, v in ciclos_dict.items()
                if v == ciclo_txt
            ][0]

            suministros, fise_suministros = obtener_suministros_rojos(
                session,
                unidad,
                fecha,
                iduunn,
                idciclo,
                id_rep
            )

            # =====================================
            # TODOS LOS PUNTOS ROJOS
            # =====================================

            if clave not in resultado:
                resultado[clave] = set()

            resultado[clave].update(suministros)

            # =====================================
            # SOLO FISE ROJOS
            # =====================================

            if clave not in resultado_fise:
                resultado_fise[clave] = set()

            resultado_fise[clave].update(fise_suministros)

        # ===================================
        # RESÚMENES
        # ===================================

        total_rojos = sum(
            len(suministros)
            for suministros in resultado.values()
        )

        total_fise = sum(
            len(suministros)
            for suministros in resultado_fise.values()
        )

        col_fise, col_general = st.columns(2)


        # ===================================
        # RESUMEN FISE
        # ===================================

        with col_fise:

            # ===================================
            # RESUMEN FISE
            # ===================================

            st.subheader("🔥 FISE pendientes")

            if total_fise == 0:

                st.warning(
                    "No se encontraron FISE pendientes de entrega"
                )

            else:

                # Resumen general siempre visible
                st.info(
                    f"🔥 Se encontraron {total_fise} FISE pendientes de entrega"
                )

                # Lista desplegable
                with st.expander(
                    f"➕ Ver detalle de FISE pendientes ({total_fise})"
                ):

                    df_fise_debug = pd.DataFrame(
                        [
                            (repartidor, len(suministros))
                            for repartidor, suministros in resultado_fise.items()
                            if len(suministros) > 0
                        ],
                        columns=[
                            "Repartidor",
                            "Cantidad FISE pendientes"
                        ]
                    )

                    st.dataframe(
                        df_fise_debug,
                        use_container_width=True,
                        hide_index=True
                    )

        # ===================================
        # RESUMEN GENERAL
        # ===================================

        with col_general:

            st.subheader("🔴 Resumen general")

            st.info(
                f"Se encontraron {total_rojos} "
                f"suministros pendientes de validación"
            )     
                        
        # ===================================
        # VALIDACIÓN MASIVA
        # ===================================

        st.subheader("⚙️ Validación automática")

        opcion_validacion = st.radio(
            "¿Qué deseas validar?",
            [
                "🔴 Todos los pendientes",
                "🔥 Solo FISE pendientes",
                "🔵 Pendientes excepto FISE"
            ],
            horizontal=True
        )


        # ===================================
        # DETERMINAR QUÉ SUMINISTROS VALIDAR
        # ===================================

        if opcion_validacion == "🔴 Todos los pendientes":

            # Todos los puntos rojos
            # Incluye FISE y NO FISE
            datos_validar = resultado


        elif opcion_validacion == "🔥 Solo FISE pendientes":

            # Solamente FISE que están rojos
            datos_validar = resultado_fise


        else:

            # ===================================
            # ROJOS EXCEPTO FISE
            # ===================================

            datos_validar = {}

            for repartidor, suministros in resultado.items():

                fises = resultado_fise.get(
                    repartidor,
                    set()
                )

                # Todos los rojos menos los FISE
                no_fise = suministros - fises

                if no_fise:
                    datos_validar[repartidor] = no_fise


        # ===================================
        # TOTAL A VALIDAR
        # ===================================

        total_validar = sum(
            len(suministros)
            for suministros in datos_validar.values()
        )


        # ===================================
        # MOSTRAR CANTIDAD
        # ===================================

        if total_validar == 0:

            if opcion_validacion == "🔴 Todos los pendientes":

                st.warning(
                    "No existen pendientes para validar."
                )

            elif opcion_validacion == "🔥 Solo FISE pendientes":

                st.warning(
                    "No existen FISES pendientes para validar."
                )

            else:

                st.warning(
                    "No existen pendientes que no sean FISE."
                )

        else:

            st.info(
                f"Se encontraron {total_validar} suministros "
                f"para validar."
            )


            # ===================================
            # BOTÓN EJECUTAR
            # ===================================

            if st.button(
                "✅ Validar seleccionados",
                key="btn_validacion_masiva"
            ):

                # ---------------------------------
                # ASEGURAR UNIDAD CORRECTA
                # ---------------------------------

                cambiar_unidad(session, unidad)

                # ---------------------------------
                # CREAR LISTA DE TRABAJO
                # ---------------------------------

                tareas = []

                for repartidor, suministros in datos_validar.items():

                    for suministro in suministros:

                        tareas.append(
                            (repartidor, suministro)
                        )

                total_validar = len(tareas)

                progreso = st.progress(0)

                resultados_validacion = []

                # ---------------------------------
                # VALIDACIÓN EN PARALELO
                # ---------------------------------

                with ThreadPoolExecutor(
                    max_workers=5
                ) as executor:

                    futuros = {
                        executor.submit(
                            validar_suministro,
                            session,
                            suministro
                        ): (
                            repartidor,
                            suministro
                        )
                        for repartidor, suministro in tareas
                    }

                    # ---------------------------------
                    # RECIBIR RESULTADOS
                    # ---------------------------------

                    for contador, futuro in enumerate(
                        as_completed(futuros),
                        start=1
                    ):

                        repartidor, suministro = futuros[futuro]

                        try:

                            respuesta = futuro.result()

                            # ---------------------------------
                            # RESPUESTA SIGOF
                            # ---------------------------------

                            if isinstance(respuesta, dict):

                                status_code = respuesta.get(
                                    "status_code"
                                )

                                if status_code == 200:
                                    estado = "VALIDADO"
                                else:
                                    estado = "ERROR"

                            else:

                                estado = (
                                    "VALIDADO"
                                    if respuesta
                                    else "ERROR"
                                )

                        except Exception:

                            estado = "ERROR"

                        # ---------------------------------
                        # TIPO FISE
                        # ---------------------------------

                        es_fise = (
                            suministro
                            in resultado_fise.get(
                                repartidor,
                                set()
                            )
                        )

                        # ---------------------------------
                        # GUARDAR RESULTADO
                        # ---------------------------------

                        resultados_validacion.append({

                            "Repartidor": repartidor,

                            "Suministro": suministro,

                            "Tipo": (
                                "FISE"
                                if es_fise
                                else "NO FISE"
                            ),

                            "Estado": estado
                        })

                        # ---------------------------------
                        # PROGRESO
                        # ---------------------------------

                        progreso.progress(
                            contador / total_validar
                        )

                # ===================================
                # RESULTADO FINAL
                # ===================================

                df_val = pd.DataFrame(
                    resultados_validacion
                )

                st.subheader(
                    "📋 Resultado de la validación"
                )

                st.dataframe(
                    df_val,
                    use_container_width=True
                )


                # ===================================
                # RESUMEN
                # ===================================

                cantidad_validos = (
                    df_val["Estado"]
                    .eq("VALIDADO")
                    .sum()
                )

                cantidad_error = (
                    df_val["Estado"]
                    .eq("ERROR")
                    .sum()
                )

                st.success(
                    f"Proceso terminado. "
                    f"Validados: {cantidad_validos} | "
                    f"Errores: {cantidad_error}"
                )

# ----------------------------
# WRAPPER PARA APP PRINCIPAL
# ----------------------------
def ejecutar_fise():
    st.header("📊 Validación FISE")

    df = obtener_monitoreo_sigof()

    if df is not None:
        st.success("Datos obtenidos correctamente")
        st.dataframe(df, use_container_width=True)
      

def obtener_suministros_rojos(
    session,
    unidad,
    fecha,
    iduunn,
    idciclo,
    id_repartidor
):

    cambiar_unidad(session, unidad)

    url = (
        f"http://sigof.distriluz.com.pe/plus/"
        f"ComrepOrdenrepartos/ajax_evaluar_repartointro/"
        f"U,L/{fecha}/{fecha}/{iduunn}/{idciclo}/0/0/"
        f"{id_repartidor}/0/0/0/0/0"
    )

    rojos = []
    fise_rojos = []

    start = 0
    length = 1000

    while True:

        params = {
            "sEcho": 1,
            "iColumns": 8,
            "iDisplayStart": start,
            "iDisplayLength": length,
        }

        r = session.get(
            url,
            params=params,
            headers={
                "X-Requested-With": "XMLHttpRequest",
                "Accept": "application/json, text/javascript, */*; q=0.01"
            },
            timeout=30
        )

        try:
            data = r.json()
        except Exception:
            return rojos, fise_rojos

        filas = data.get("aaData", [])

        if not filas:
            break

        for fila in filas:

            if len(fila) < 8:
                continue

            # ==========================================
            # COLUMNA R = RESULTADO EVALUACIÓN
            # ==========================================

            col_r = str(fila[5]).lower()

            es_rojo = (
                "color-danger" in col_r
                or "text-danger" in col_r
                or "fa-times" in col_r
                or "no entregado" in col_r
            )

            # ==========================================
            # COLUMNA T = TIPO DE REPARTO
            # ==========================================

            col_t = str(fila[6])

            soup_t = BeautifulSoup(
                col_t,
                "html.parser"
            )

            tipo_reparto = soup_t.get_text(
                " ",
                strip=True
            ).upper()

            # Detectamos FISE de dos maneras:
            # 1. El contenido sea F
            # 2. El span tenga title="Fise"

            es_fise = (
                tipo_reparto == "F"
                or 'title="Fise"' in col_t
                or "title='Fise'" in col_t
            )

            # ==========================================
            # SOLO PROCESAMOS LOS ROJOS
            # ==========================================

            if es_rojo:

                match = re.search(
                    r'val_suministro\s*=\s*"(\d+)"',
                    str(fila[7])
                )

                if match:

                    suministro = match.group(1)

                    # TODOS LOS ROJOS
                    rojos.append(suministro)

                    # ROJO + FISE
                    if es_fise:
                        fise_rojos.append(suministro)

        start += length

        total_registros = int(
            data.get("iTotalRecords", 0)
        )

        if start >= total_registros:
            break

    return rojos, fise_rojos

# ----------------------------
# VALIDAR SUMINISTRO (ROJO → VERDE)
# ----------------------------
def validar_suministro(
    session,
    suministro
):

    url = (
        "http://sigof.distriluz.com.pe"
        "/plus/ComrepOrdenrepartos/"
        f"ajax_save_entrega_reparto_si/{suministro}"
    )

    try:

        r = session.post(
            url,
            headers={
                "X-Requested-With": "XMLHttpRequest",
                "Accept": "application/json, text/javascript, */*; q=0.01"
            },
            timeout=30
        )

        return {
            "status_code": r.status_code,
            "respuesta": r.text,
            "url": url
        }

    except Exception as e:

        return {
            "status_code": "ERROR",
            "respuesta": str(e),
            "url": url
        }