import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
import json
import html
import base64


# ============================================================
# LIMPIAR SUMINISTRO
# ============================================================

def limpiar_suministro(valor):

    if pd.isna(valor):
        return ""

    valor = str(valor).strip()

    # Excel puede convertir 67179860 en 67179860.0
    if valor.endswith(".0"):
        valor = valor[:-2]

    return valor


# ============================================================
# CREAR MAPA GOOGLE
# ============================================================

def mostrar_mapa_google(df_puntos, api_key):

    # --------------------------------------------------------
    # Preparar puntos para JavaScript
    # --------------------------------------------------------

    puntos = []

    for numero, (_, fila) in enumerate(
        df_puntos.iterrows(),
        start=1
    ):

        suministro = str(
            fila["suministro"]
        )

        latitud = float(
            fila["latitud"]
        )

        longitud = float(
            fila["longitud"]
        )

        puntos.append({
            "numero": numero,
            "suministro": suministro,
            "lat": latitud,
            "lng": longitud
        })


    if not puntos:

        st.error(
            "No existen puntos con coordenadas válidas."
        )

        return


    # --------------------------------------------------------
    # Centro inicial
    # --------------------------------------------------------

    lat_centro = sum(
        p["lat"] for p in puntos
    ) / len(puntos)

    lng_centro = sum(
        p["lng"] for p in puntos
    ) / len(puntos)


    # --------------------------------------------------------
    # Convertir datos a JSON
    # --------------------------------------------------------

    puntos_json = json.dumps(
        puntos,
        ensure_ascii=False
    )


    # --------------------------------------------------------
    # HTML + JAVASCRIPT
    # --------------------------------------------------------

    mapa_html = f"""
<!DOCTYPE html>

<html>

<head>

<meta charset="UTF-8">

<meta
    name="viewport"
    content="width=device-width,
             initial-scale=1.0,
             maximum-scale=1.0,
             user-scalable=no"
>


<style>

html,
body {{

    margin: 0;
    padding: 0;

    width: 100%;
    height: 100%;

    overflow: hidden;

    font-family: Arial, sans-serif;
}}


#map {{

    width: 100%;
    height: 100vh;
}}


/* ---------------------------------------------------------
   MARCADOR NUMERADO
--------------------------------------------------------- */

.numero-marcador {{

    width: 34px;
    height: 34px;

    background: #1683d8;

    border: 3px solid white;

    border-radius: 50%;

    box-shadow:
        0 2px 6px rgba(0,0,0,0.45);

    display: flex;

    align-items: center;

    justify-content: center;

    color: white;

    font-size: 14px;

    font-weight: bold;

    box-sizing: border-box;

    cursor: pointer;
}}


/* ---------------------------------------------------------
   VENTANA DE INFORMACIÓN
--------------------------------------------------------- */

.info-suministro {{

    font-family: Arial, sans-serif;

    min-width: 180px;

    padding: 4px;
}}


.info-suministro .titulo {{

    font-size: 16px;

    font-weight: bold;

    margin-bottom: 8px;
}}


.info-suministro .dato {{

    font-size: 14px;

    margin-bottom: 6px;
}}


.boton-google {{

    display: inline-block;

    margin-top: 8px;

    padding: 9px 12px;

    background: #1a73e8;

    color: white;

    text-decoration: none;

    border-radius: 6px;

    font-weight: bold;

    font-size: 13px;
}}

</style>

</head>


<body>


<div id="map"></div>


<script>

const puntos = {puntos_json};


/* ==========================================================
   INICIALIZAR MAPA
========================================================== */

async function initMap() {{

    const {{ Map }} =
        await google.maps.importLibrary("maps");

    const {{ AdvancedMarkerElement }} =
        await google.maps.importLibrary("marker");


    /* ------------------------------------------------------
       CREAR MAPA
    ------------------------------------------------------ */

    const map = new Map(
        document.getElementById("map"),
        {{

            center: {{
                lat: {lat_centro},
                lng: {lng_centro}
            }},

            zoom: 16,

            mapTypeId: "satellite",

            mapTypeId: "hybrid",

            mapId: "DEMO_MAP_ID",

            streetViewControl: true,

            mapTypeControl: true,

            fullscreenControl: true,

            zoomControl: true,

            gestureHandling: "greedy"

        }}
    );


    /* ======================================================
       BOUNDS
    ====================================================== */

    const bounds =
        new google.maps.LatLngBounds();


    /* ======================================================
       MARCADORES
    ====================================================== */

    puntos.forEach((punto) => {{

        const posicion = {{

            lat: punto.lat,

            lng: punto.lng

        }};


        bounds.extend(posicion);


        /* --------------------------------------------------
           CREAR MARCADOR NUMERADO
        -------------------------------------------------- */

        const contenido =
            document.createElement("div");


        contenido.className =
            "numero-marcador";


        contenido.textContent =
            punto.numero;


        contenido.title =
            "Suministro " +
            punto.suministro;


        /* --------------------------------------------------
           MARCADOR AVANZADO
        -------------------------------------------------- */

        const marker =
            new AdvancedMarkerElement({{

                map: map,

                position: posicion,

                content: contenido,

                title:
                    "Suministro " +
                    punto.suministro,

                gmpClickable: true

            }});


        /* --------------------------------------------------
           INFO WINDOW
        -------------------------------------------------- */

        const infoWindow =
            new google.maps.InfoWindow();


        contenido.addEventListener(
            "click",
            () => {{

                const urlGoogle =
                    "https://www.google.com/maps/dir/?api=1"
                    + "&destination="
                    + punto.lat
                    + ","
                    + punto.lng;


                const contenidoInfo = `

                    <div class="info-suministro">

                        <div class="titulo">

                            Suministro
                            ${{punto.suministro}}

                        </div>


                        <div class="dato">

                            <b>Orden:</b>
                            ${{punto.numero}}

                        </div>


                        <div class="dato">

                            <b>Latitud:</b>
                            ${{punto.lat}}

                        </div>


                        <div class="dato">

                            <b>Longitud:</b>
                            ${{punto.lng}}

                        </div>


                        <a
                            class="boton-google"
                            href="${{urlGoogle}}"
                            target="_blank"
                        >

                            🧭 Ir con Google Maps

                        </a>

                    </div>

                `;


                infoWindow.setContent(
                    contenidoInfo
                );


                infoWindow.open({{

                    map: map,

                    anchor: marker

                }});

            }}
        );

    }});


    /* ======================================================
       AJUSTAR MAPA A TODOS LOS PUNTOS
    ====================================================== */

    if (puntos.length > 1) {{

        map.fitBounds(bounds);

    }}

}}


/* ==========================================================
   CARGAR GOOGLE MAPS
========================================================== */

function cargarGoogleMaps() {{

    const script =
        document.createElement("script");


    script.src =
        "https://maps.googleapis.com/maps/api/js"
        + "?key={html.escape(api_key)}"
        + "&v=weekly"
        + "&callback=initMap";


    script.async = true;

    script.defer = true;


    document.head.appendChild(script);

}}


cargarGoogleMaps();

</script>


</body>

</html>
"""


    # --------------------------------------------------------
    # Mostrar mapa
    # --------------------------------------------------------

    components.html(
        mapa_html,
        height=700,
        scrolling=False
    )

    st.download_button(
        "📄 Guardar mapa HTML",
        data=mapa_html,
        file_name="mapa_suministros.html",
        mime="text/html",
        use_container_width=True
    )


# ============================================================
# FUNCIÓN PRINCIPAL
# ============================================================

def ejecutar_generacion_mapa():

    st.title("🗺️ Generación de Mapa")

    st.write(
        "Genera un mapa de Google Maps a partir "
        "del Excel maestro y la lista de pendientes."
    )


    # ========================================================
    # API KEY
    # ========================================================

    try:

        api_key = st.secrets["GOOGLE_MAPS_API_KEY"]

    except Exception:

        st.error(
            "No se encontró GOOGLE_MAPS_API_KEY "
            "en los secrets de Streamlit."
        )

        st.info(
            "Crea el archivo .streamlit/secrets.toml "
            "y agrega tu API Key."
        )

        return


    st.divider()


    # ========================================================
    # EXCEL MAESTRO
    # ========================================================

    st.subheader(
        "1️⃣ Excel maestro de coordenadas"
    )


    # ========================================================
    # SI YA EXISTE UN EXCEL CARGADO
    # ========================================================

    if "df_maestro" in st.session_state:

        st.success(
            f"✅ Excel maestro cargado: "
            f"{st.session_state['nombre_maestro']}"
        )

        col_excel1, col_excel2 = st.columns([4, 1])

        with col_excel1:

            st.caption(
                f"Registros cargados: "
                f"{len(st.session_state['df_maestro']):,}"
            )

        with col_excel2:

            if st.button(
                "🔄 Cambiar Excel",
                use_container_width=True
            ):

                # Eliminar información anterior

                st.session_state.pop(
                    "df_maestro",
                    None
                )

                st.session_state.pop(
                    "nombre_maestro",
                    None
                )

                st.session_state.pop(
                    "columnas_maestro",
                    None
                )

                st.session_state.pop(
                    "mapa_google_generado",
                    None
                )

                st.rerun()


        # Recuperar Excel desde memoria

        df = st.session_state[
            "df_maestro"
        ]

        columnas = st.session_state[
            "columnas_maestro"
        ]


    # ========================================================
    # SI TODAVÍA NO EXISTE EXCEL
    # ========================================================

    else:

        archivo_excel = st.file_uploader(
            "Selecciona el Excel maestro",
            type=["xlsx", "xls"],
            key="excel_maestro_google"
        )


        if archivo_excel is None:

            st.info(
                "Carga el Excel maestro de coordenadas."
            )

            return


        # ----------------------------------------------------
        # LEER EXCEL SOLO UNA VEZ
        # ----------------------------------------------------

        try:

            df = pd.read_excel(
                archivo_excel
            )

        except Exception as e:

            st.error(
                f"No se pudo leer el Excel: {e}"
            )

            return


        columnas = list(
            df.columns
        )


        if len(columnas) < 3:

            st.error(
                "El Excel debe tener al menos "
                "tres columnas."
            )

            return


        # ----------------------------------------------------
        # GUARDAR EN MEMORIA DE LA SESIÓN
        # ----------------------------------------------------

        st.session_state[
            "df_maestro"
        ] = df.copy()


        st.session_state[
            "nombre_maestro"
        ] = archivo_excel.name


        st.session_state[
            "columnas_maestro"
        ] = columnas


        st.success(
            f"✅ Excel maestro cargado: "
            f"{archivo_excel.name}"
        )


    # ========================================================
    # SELECCIONAR COLUMNAS
    # ========================================================

    st.subheader(
        "2️⃣ Configurar columnas"
    )


    col1, col2, col3 = st.columns(3)


    # --------------------------------------------------------
    # Detectar columnas automáticamente
    # --------------------------------------------------------

    suministro_default = 0
    latitud_default = 0
    longitud_default = 0


    for i, columna in enumerate(columnas):

        nombre = str(
            columna
        ).lower()


        if (
            "suministro" in nombre
            or "sumin" in nombre
            or "nis" in nombre
            or "codigo" in nombre
            or "código" in nombre
        ):

            suministro_default = i


        if (
            "latitud" in nombre
            or nombre == "lat"
            or "latitude" in nombre
        ):

            latitud_default = i


        if (
            "longitud" in nombre
            or nombre == "lon"
            or nombre == "lng"
            or "longitude" in nombre
        ):

            longitud_default = i


    with col1:

        columna_suministro = st.selectbox(
            "Suministro",
            columnas,
            index=suministro_default
        )


    with col2:

        columna_latitud = st.selectbox(
            "Latitud",
            columnas,
            index=latitud_default
        )


    with col3:

        columna_longitud = st.selectbox(
            "Longitud",
            columnas,
            index=longitud_default
        )


    # ========================================================
    # PREPARAR DATOS
    # ========================================================

    df_mapa = df[
        [
            columna_suministro,
            columna_latitud,
            columna_longitud
        ]
    ].copy()


    df_mapa.columns = [
        "suministro",
        "latitud",
        "longitud"
    ]


    df_mapa["suministro"] = (
        df_mapa["suministro"]
        .apply(limpiar_suministro)
    )


    df_mapa["latitud"] = pd.to_numeric(
        df_mapa["latitud"],
        errors="coerce"
    )


    df_mapa["longitud"] = pd.to_numeric(
        df_mapa["longitud"],
        errors="coerce"
    )


    df_mapa = df_mapa[
        df_mapa["suministro"] != ""
    ].copy()


    # ========================================================
    # PENDIENTES
    # ========================================================

    st.subheader(
        "3️⃣ Suministros pendientes"
    )

    with st.expander(
        "📋 Pegar / ver suministros pendientes",
        expanded=False
    ):

        pendientes_texto = st.text_area(
            "Lista de suministros",
            height=200,
            placeholder=(
                "67179860\n"
                "67179861\n"
                "67179862\n"
                "67179863"
            ),
            key="pendientes_google"
        )


    if not pendientes_texto.strip():

        st.info(
            "Pega la lista de suministros pendientes."
        )

        return


    # ========================================================
    # PROCESAR PENDIENTES
    # ========================================================

    pendientes = []


    for linea in pendientes_texto.splitlines():

        suministro = limpiar_suministro(
            linea
        )


        if suministro:

            pendientes.append(
                suministro
            )


    # Eliminar duplicados
    pendientes = list(
        dict.fromkeys(
            pendientes
        )
    )


    # ========================================================
    # COMPARAR
    # ========================================================

    df_resultado = df_mapa[
        df_mapa["suministro"].isin(
            pendientes
        )
    ].copy()


    encontrados = set(
        df_resultado["suministro"]
    )


    no_encontrados = [
        suministro
        for suministro in pendientes
        if suministro not in encontrados
    ]


    # ========================================================
    # COORDENADAS
    # ========================================================

    df_puntos = df_resultado[
        df_resultado["latitud"].notna()
        &
        df_resultado["longitud"].notna()
    ].copy()


    sin_coordenadas = df_resultado[
        df_resultado["latitud"].isna()
        |
        df_resultado["longitud"].isna()
    ].copy()


    # ========================================================
    # RESUMEN
    # ========================================================

    st.divider()

    st.subheader(
        "📊 Resultado"
    )


    c1, c2, c3, c4 = st.columns(4)


    with c1:

        st.metric(
            "Pendientes",
            len(pendientes)
        )


    with c2:

        st.metric(
            "Encontrados",
            len(df_resultado)
        )


    with c3:

        st.metric(
            "Puntos en mapa",
            len(df_puntos)
        )


    with c4:

        st.metric(
            "No encontrados",
            len(no_encontrados)
        )


    # ========================================================
    # NO ENCONTRADOS
    # ========================================================

    if no_encontrados:

        with st.expander(
            "⚠️ Ver suministros no encontrados"
        ):

            st.dataframe(
                pd.DataFrame({
                    "suministro":
                        no_encontrados
                }),
                use_container_width=True,
                hide_index=True
            )


    # ========================================================
    # SIN COORDENADAS
    # ========================================================

    if len(sin_coordenadas) > 0:

        with st.expander(
            "⚠️ Ver suministros sin coordenadas"
        ):

            st.dataframe(
                sin_coordenadas,
                use_container_width=True,
                hide_index=True
            )


    # ========================================================
    # GENERAR
    # ========================================================

    st.divider()


    if len(df_puntos) == 0:

        st.error(
            "No hay suministros con coordenadas "
            "válidas para mostrar."
        )

        return


    if st.button(
        "🗺️ GENERAR MAPA GOOGLE",
        type="primary",
        use_container_width=True
    ):

        st.session_state[
            "mapa_google_generado"
        ] = True


    # ========================================================
    # MOSTRAR
    # ========================================================

    if st.session_state.get(
        "mapa_google_generado",
        False
    ):

        st.success(
            f"Mapa generado con "
            f"{len(df_puntos)} suministros."
        )


        mostrar_mapa_google(
            df_puntos,
            api_key
        )