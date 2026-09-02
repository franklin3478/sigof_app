import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path
import json
import html


# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Mapa de Personal",
    page_icon="🗺️",
    layout="wide"
)


# ============================================================
# USUARIOS
# ============================================================

USUARIOS = {

    "enlace1": {
        "password": "1234",
        "nombre": "Mapa Personal 1",
        "archivo": "mapa_1.json"
    },

    "enlace2": {
        "password": "5678",
        "nombre": "Mapa Personal 2",
        "archivo": "mapa_2.json"
    }

}


# ============================================================
# CARPETA DE MAPAS
# ============================================================

CARPETA_MAPAS = (
    Path(__file__).parent / "mapas"
)


# ============================================================
# LOGIN
# ============================================================

def pantalla_login():

    st.title("🗺️ MAPA DE PERSONAL")

    st.write(
        "Ingrese sus credenciales para acceder."
    )

    st.divider()

    col1, col2, col3 = st.columns(
        [1, 2, 1]
    )

    with col2:

        usuario = st.text_input(
            "Usuario"
        )

        password = st.text_input(
            "Contraseña",
            type="password"
        )

        if st.button(
            "🔐 INGRESAR",
            type="primary",
            use_container_width=True
        ):

            if (
                usuario in USUARIOS
                and
                USUARIOS[usuario]["password"]
                == password
            ):

                st.session_state[
                    "logueado"
                ] = True

                st.session_state[
                    "usuario"
                ] = usuario

                st.rerun()

            else:

                st.error(
                    "❌ Usuario o contraseña incorrectos."
                )


# ============================================================
# GENERAR HTML DEL MAPA
# ============================================================

def generar_mapa_html(
    puntos,
    api_key
):

    if not puntos:

        return ""


    lat_centro = sum(
        p["lat"] for p in puntos
    ) / len(puntos)

    lng_centro = sum(
        p["lng"] for p in puntos
    ) / len(puntos)


    puntos_json = json.dumps(
        puntos,
        ensure_ascii=False
    )


    mapa_html = f"""
<!DOCTYPE html>

<html>

<head>

<meta charset="UTF-8">

<meta name="viewport"
content="width=device-width,
initial-scale=1.0">

<style>

html,
body {{

    margin: 0;
    padding: 0;

    width: 100%;
    height: 100%;

    overflow: hidden;

    font-family: Arial;
}}


#map {{

    width: 100%;
    height: 100vh;

}}


.numero-marcador {{

    width: 34px;
    height: 34px;

    background: #1683d8;

    border: 3px solid white;

    border-radius: 50%;

    box-shadow:
        0 2px 6px rgba(0,0,0,.45);

    display: flex;

    align-items: center;

    justify-content: center;

    color: white;

    font-size: 14px;

    font-weight: bold;

    cursor: pointer;

}}


.info-suministro {{

    font-family: Arial;

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

}}

</style>

</head>


<body>

<div id="map"></div>


<script>

const puntos = {puntos_json};


async function initMap() {{

    const {{ Map }} =
        await google.maps.importLibrary(
            "maps"
        );


    const {{ AdvancedMarkerElement }} =
        await google.maps.importLibrary(
            "marker"
        );


    const map = new Map(
        document.getElementById("map"),
        {{

            center: {{

                lat: {lat_centro},

                lng: {lng_centro}

            }},

            zoom: 16,

            mapTypeId: "hybrid",

            mapId: "DEMO_MAP_ID",

            streetViewControl: true,

            mapTypeControl: true,

            fullscreenControl: true,

            zoomControl: true,

            gestureHandling: "greedy"

        }}

    );


    const bounds =
        new google.maps.LatLngBounds();


    puntos.forEach((punto) => {{

        const posicion = {{

            lat: punto.lat,

            lng: punto.lng

        }};


        bounds.extend(
            posicion
        );


        const contenido =
            document.createElement(
                "div"
            );


        contenido.className =
            "numero-marcador";


        contenido.textContent =
            punto.numero;


        contenido.title =
            "Suministro "
            + punto.suministro;


        const marker =
            new AdvancedMarkerElement({{

                map: map,

                position: posicion,

                content: contenido,

                title:
                    "Suministro "
                    + punto.suministro,

                gmpClickable: true

            }});


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


    if (puntos.length > 1) {{

        map.fitBounds(
            bounds
        );

    }}

}}


function cargarGoogleMaps() {{

    const script =
        document.createElement(
            "script"
        );


    script.src =
        "https://maps.googleapis.com/maps/api/js"
        + "?key={html.escape(api_key)}"
        + "&v=weekly"
        + "&callback=initMap";


    script.async = true;

    script.defer = true;


    document.head.appendChild(
        script
    );

}}


cargarGoogleMaps();

</script>

</body>

</html>
"""

    return mapa_html


# ============================================================
# MOSTRAR MAPA
# ============================================================

def mostrar_mapa(usuario):

    datos = USUARIOS[usuario]

    archivo = (
        CARPETA_MAPAS /
        datos["archivo"]
    )


    st.title("🗺️ " + datos["nombre"])


    col1, col2 = st.columns(
        [5, 1]
    )


    with col1:

        st.success(
            f"👤 {datos['nombre']}"
        )


    with col2:

        if st.button(
            "🚪 Salir",
            use_container_width=True
        ):

            st.session_state.clear()

            st.rerun()


    st.divider()


    # --------------------------------------------------------
    # COMPROBAR ARCHIVO
    # --------------------------------------------------------

    if not archivo.exists():

        st.warning(
            "⚠️ Este mapa todavía no ha sido publicado."
        )

        return


    # --------------------------------------------------------
    # LEER JSON
    # --------------------------------------------------------

    try:

        with open(
            archivo,
            "r",
            encoding="utf-8"
        ) as f:

            datos_mapa = json.load(f)


    except Exception as e:

        st.error(
            f"Error leyendo el mapa: {e}"
        )

        return


    puntos = datos_mapa.get(
        "puntos",
        []
    )


    nombre_mapa = datos_mapa.get(
        "nombre",
        datos["nombre"]
    )


    fecha = datos_mapa.get(
        "actualizado",
        "Sin fecha"
    )


    # --------------------------------------------------------
    # INFORMACIÓN
    # --------------------------------------------------------

    col1, col2, col3 = st.columns(
        3
    )


    with col1:

        st.metric(
            "Suministros",
            len(puntos)
        )


    with col2:

        st.metric(
            "Última actualización",
            fecha
        )


    with col3:

        st.metric(
            "Estado",
            "ACTUALIZADO"
        )


    st.divider()


    # --------------------------------------------------------
    # BOTÓN ACTUALIZAR
    # --------------------------------------------------------

    if st.button(
        "🔄 ACTUALIZAR MAPA",
        type="primary",
        use_container_width=True
    ):

        st.rerun()


    # --------------------------------------------------------
    # API KEY
    # --------------------------------------------------------

    try:

        api_key = st.secrets[
            "GOOGLE_MAPS_API_KEY"
        ]

    except Exception:

        st.error(
            "No se encontró GOOGLE_MAPS_API_KEY."
        )

        return


    # --------------------------------------------------------
    # MOSTRAR MAPA
    # --------------------------------------------------------

    mapa_html = generar_mapa_html(
        puntos,
        api_key
    )


    components.html(
        mapa_html,
        height=750,
        scrolling=False
    )


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

if not st.session_state.get(
    "logueado",
    False
):

    pantalla_login()

else:

    usuario = st.session_state[
        "usuario"
    ]

    mostrar_mapa(
        usuario
    )

