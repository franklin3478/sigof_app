import streamlit as st
import requests
from datetime import datetime
from bs4 import BeautifulSoup
import pandas as pd
import re

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

        # Mostrar sin ID
        df_mostrar = df_final.drop(
            columns=["id_repartidor"],
            errors="ignore"
        )

        st.dataframe(
            df_mostrar,
            use_container_width=True
        )

        # 🔥 AQUÍ VA TODO LO NUEVO
        
        resultado = {}

        iduunn = mapa_unidades[unidad]["iduunn"]

        for _, row in st.session_state.df_final.iterrows():
            nombre = row["repartidor"]
            id_rep = row["id_repartidor"]
            ciclo_txt = row["ciclo"]

            # 🔑 clave única (evita problemas de nombres repetidos)
            clave = f"{id_rep} - {nombre}"

            idciclo = [k for k, v in ciclos_dict.items() if v == ciclo_txt][0]

            suministros = obtener_suministros_rojos(
                session,
                unidad,
                fecha,
                iduunn,
                idciclo,
                id_rep
            )

            # 🔥 usar SET para evitar duplicados entre ciclos
            if clave not in resultado:
                resultado[clave] = set()

            resultado[clave].update(suministros)

        # 🔥 convertir a conteo real (sin duplicados)
        st.subheader("🔴 Resumen de puntos rojos por repartidor")

        df_debug = pd.DataFrame(
            [(k, len(v)) for k, v in resultado.items()],
            columns=["Repartidor", "Cantidad puntos rojos"]
        )
        
        st.dataframe(df_debug, use_container_width=True)

        total_rojos = sum(len(v) for v in resultado.values())

        st.subheader("🔴 Resumen general")

        st.info(f"Se encontraron {total_rojos} suministros pendientes de validación")

        st.subheader("🧪 Prueba de validación")

        todos_rojos = []

        for repartidor, suministros in resultado.items():

            for suministro in suministros:

                todos_rojos.append(
                    f"{suministro} | {repartidor}"
                )

        if todos_rojos:

            suministro_sel = st.selectbox(
                "Seleccione un suministro rojo",
                todos_rojos
            )

            if st.button("Validar suministro seleccionado"):

                suministro = suministro_sel.split("|")[0].strip()

                resultado_prueba = validar_suministro(
                    session,
                    unidad,
                    suministro
                )

                st.json(resultado_prueba)

        else:

            st.warning("No se encontraron puntos rojos")
        
        # ===================================
        # BOTÓN VALIDACIÓN MASIVA
        # ===================================

        if st.button(
                "✅ Validar automáticamente puntos rojos"
        ):

            progreso = st.progress(0)

            resultados = []

            total = sum(
                len(v)
                for v in resultado.values()
            )

            contador = 0

            for repartidor, suministros in resultado.items():

                # PRUEBA SOLO 1
                for suministro in list(suministros):

                    ok = validar_suministro(
                        session,
                        unidad,
                        suministro
                    )

                    resultados.append({

                        "repartidor":
                            repartidor,

                        "suministro":
                            suministro,

                        "estado":
                            "VALIDADO"
                            if ok
                            else "ERROR"

                    })

                    contador += 1

                    progreso.progress(
                        contador / total
                    )

            df_val = pd.DataFrame(
                resultados
            )

            st.dataframe(df_val)

            st.success(
                "Proceso terminado"
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
      

def obtener_suministros_rojos(session, unidad, fecha, iduunn, idciclo, id_repartidor):

    cambiar_unidad(session, unidad)

    url = f"http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_evaluar_repartointro/U,L/{fecha}/{fecha}/{iduunn}/{idciclo}/0/0/{id_repartidor}/0/0/0/0/0"

    rojos = []
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
                "Accept": "application/json, text/javascript, /; q=0.01"
            },
            timeout=30
        )

        try:
            data = r.json()
        except:
            return rojos

        filas = data.get("aaData", [])
        
        if not filas:
            break

        # 🔴 PROCESAR FILAS
        for fila in filas:
           
            if len(fila) < 6:
                continue
            
            col_r = fila[5].lower()

            if (
                "color-danger" in col_r or
                "text-danger" in col_r or
                "fa-times" in col_r or
                "no entregado" in col_r
            ):
                match = re.search(
                    r'val_suministro\s*=\s*"(\d+)"',
                    fila[7]
                )

                if match:
                    suministro = match.group(1)
                    rojos.append(suministro)
        
        start += length

        # 🔥 cortar si ya no hay más
        if start >= int(data.get("iTotalRecords", 0)):
            break
    return rojos

# ----------------------------
# VALIDAR SUMINISTRO (ROJO → VERDE)
# ----------------------------
def validar_suministro(
        session,
        unidad,
        suministro
):

    cambiar_unidad(session, unidad)

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
                "Accept": "application/json, text/javascript, /; q=0.01"
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