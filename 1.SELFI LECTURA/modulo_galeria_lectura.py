import streamlit as st
import pandas as pd
import requests
from openpyxl import load_workbook
from io import BytesIO
import re
import json
import html as html_lib
import streamlit.components.v1 as components
from concurrent.futures import ThreadPoolExecutor
import threading

# ============================================================
# CONFIGURACIÓN
# ============================================================

TIMEOUT = 20

URL_API_FIELSERVICE = (
    "https://servicios.distriluz.com.pe:51000/"
    "OptimusNGC_FieldService/api/reportes-publicos/fotos/"
)

USER_AGENT = (
    "Mozilla/5.0 "
    "(Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 "
    "(KHTML, like Gecko) "
    "Chrome/139.0.0.0 "
    "Safari/537.36"
)


# ============================================================
# SESIÓN HTTP POR HILO
# ============================================================

_sesiones_hilos = threading.local()


def obtener_sesion_http():

    sesion = getattr(
        _sesiones_hilos,
        "session",
        None
    )

    if sesion is None:

        sesion = requests.Session()

        sesion.headers.update({
            "Accept": "application/json",
            "User-Agent": USER_AGENT
        })

        _sesiones_hilos.session = sesion

    return sesion


# ============================================================
# LEER EXCEL Y OBTENER HIPERVÍNCULOS
# ============================================================

@st.cache_data(show_spinner=False)
def leer_excel_con_links(archivo):

    contenido = archivo.getvalue()

    df = pd.read_excel(
        BytesIO(contenido)
    )

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    columna_foto = None

    for col in df.columns:

        if str(col).strip().lower() == "foto":

            columna_foto = col
            break

    if columna_foto is None:

        raise ValueError(
            "El Excel debe contener una columna llamada 'foto'."
        )

    wb = load_workbook(
        BytesIO(contenido),
        data_only=False
    )

    ws = wb.active

    numero_columna_foto = None

    for celda in ws[1]:

        if (
            str(celda.value)
            .strip()
            .lower()
            == "foto"
        ):

            numero_columna_foto = celda.column
            break

    if numero_columna_foto is None:

        raise ValueError(
            "No se encontró la columna 'foto'."
        )

    enlaces = []

    for fila in range(
        2,
        ws.max_row + 1
    ):

        celda = ws.cell(
            row=fila,
            column=numero_columna_foto
        )

        url = None

        # ====================================================
        # HIPERVÍNCULO REAL
        # ====================================================

        if celda.hyperlink:

            url = celda.hyperlink.target

        # ====================================================
        # FÓRMULA HYPERLINK
        # ====================================================

        elif isinstance(celda.value, str):

            valor = celda.value.strip()

            coincidencia = re.search(
                r'HYPERLINK\s*\(\s*"([^"]+)"',
                valor,
                flags=re.IGNORECASE
            )

            if coincidencia:

                url = coincidencia.group(1)

        enlaces.append(url)

    if len(enlaces) < len(df):

        enlaces.extend(
            [None] * (
                len(df) - len(enlaces)
            )
        )

    df["__url_foto"] = enlaces[:len(df)]

    return df, columna_foto


# ============================================================
# EXTRAER UUID DEL HIPERVÍNCULO
# MISMA LÓGICA DE LA MACRO
# ============================================================

def extraer_uuid_fieldservice(url):

    if not url:
        return ""

    url = str(url).strip()

    if not url:
        return ""

    try:

        uuid = url.rstrip("/").rsplit(
            "/",
            1
        )[-1]

        return uuid.strip()

    except Exception:

        return ""


# ============================================================
# EXTRAER LAS DOS URLS DEL JSON
# LECTURA + MEDIDOR
# ============================================================

def extraer_urls_json_fieldservice(respuesta):

    url_lectura = ""
    url_medidor = ""

    try:

        datos = respuesta.json()

        # ====================================================
        # CASO NORMAL
        # ====================================================

        elementos = datos

        if isinstance(datos, dict):

            if isinstance(
                datos.get("data"),
                list
            ):

                elementos = datos["data"]

            elif isinstance(
                datos.get("result"),
                list
            ):

                elementos = datos["result"]

            elif isinstance(
                datos.get("items"),
                list
            ):

                elementos = datos["items"]

            else:

                elementos = [datos]

        if isinstance(elementos, list):

            for elemento in elementos:

                if not isinstance(
                    elemento,
                    dict
                ):
                    continue

                tipo = str(
                    elemento.get(
                        "tipo",
                        ""
                    )
                ).strip().upper()

                url = str(
                    elemento.get(
                        "url",
                        ""
                    )
                ).strip()

                if not url:
                    continue

                if tipo == "LECTURA":

                    url_lectura = url

                elif tipo == "MEDIDOR":

                    url_medidor = url

        # ====================================================
        # RESPALDO
        # ====================================================

        if (
            not url_lectura
            or not url_medidor
        ):

            texto = respuesta.text

            if texto:

                if not url_lectura:

                    patron_lectura = re.search(
                        r'"tipo"\s*:\s*"LECTURA".*?'
                        r'"url"\s*:\s*"([^"]+)"',
                        texto,
                        flags=(
                            re.IGNORECASE |
                            re.DOTALL
                        )
                    )

                    if patron_lectura:

                        url_lectura = (
                            patron_lectura
                            .group(1)
                            .strip()
                        )

                if not url_medidor:

                    patron_medidor = re.search(
                        r'"tipo"\s*:\s*"MEDIDOR".*?'
                        r'"url"\s*:\s*"([^"]+)"',
                        texto,
                        flags=(
                            re.IGNORECASE |
                            re.DOTALL
                        )
                    )

                    if patron_medidor:

                        url_medidor = (
                            patron_medidor
                            .group(1)
                            .strip()
                        )

    except Exception:

        # ====================================================
        # RESPALDO TOTAL
        # ====================================================

        try:

            texto = respuesta.text

            patron_lectura = re.search(
                r'"tipo"\s*:\s*"LECTURA".*?'
                r'"url"\s*:\s*"([^"]+)"',
                texto,
                flags=(
                    re.IGNORECASE |
                    re.DOTALL
                )
            )

            patron_medidor = re.search(
                r'"tipo"\s*:\s*"MEDIDOR".*?'
                r'"url"\s*:\s*"([^"]+)"',
                texto,
                flags=(
                    re.IGNORECASE |
                    re.DOTALL
                )
            )

            if patron_lectura:

                url_lectura = (
                    patron_lectura
                    .group(1)
                    .strip()
                )

            if patron_medidor:

                url_medidor = (
                    patron_medidor
                    .group(1)
                    .strip()
                )

        except Exception:

            pass

    urls_fotos = []

    if url_lectura:

        urls_fotos.append(
            url_lectura
        )

    if url_medidor:

        urls_fotos.append(
            url_medidor
        )

    return urls_fotos[:2]


# ============================================================
# OBTENER URLS DIRECTAS DE FOTOS FIELDSERVICE
# MISMA LÓGICA DE LA MACRO
# ============================================================

def obtener_urls_fotos_fieldservice(url):

    if not url:
        return []

    url = str(url).strip()

    if not url:
        return []

    try:

        # ====================================================
        # EXTRAER UUID
        # ====================================================

        uuid = extraer_uuid_fieldservice(
            url
        )

        if not uuid:
            return []

        # ====================================================
        # CONSTRUIR API
        # ====================================================

        url_api = (
            URL_API_FIELSERVICE
            + uuid
        )

        # ====================================================
        # SESIÓN HTTP
        # ====================================================

        sesion = obtener_sesion_http()

        respuesta = sesion.get(
            url_api,
            timeout=TIMEOUT
        )

        # ====================================================
        # VALIDAR RESPUESTA
        # ====================================================

        if respuesta.status_code != 200:
            return []

        if not respuesta.text:
            return []

        # ====================================================
        # OBTENER LECTURA + MEDIDOR
        # ====================================================

        return extraer_urls_json_fieldservice(
            respuesta
        )

    except requests.RequestException:

        return []

    except Exception:

        return []


# ============================================================
# MOSTRAR INFORMACIÓN
# ============================================================

def mostrar_datos(
    fila,
    columnas_mostrar
):

    for columna in columnas_mostrar:

        valor = fila.get(
            columna,
            ""
        )

        if pd.isna(valor):

            valor = ""

        elif isinstance(
            valor,
            float
        ):

            if valor.is_integer():

                valor = str(
                    int(valor)
                )

            else:

                valor = str(valor)

        else:

            valor = str(valor).strip()

        nombre = (
            str(columna)
            .replace("_", " ")
            .title()
        )

        st.markdown(
            f"**{nombre}:** {valor}"
        )


# ============================================================
# MOSTRAR FOTOS
# ============================================================

def mostrar_fotos(
    url,
    imagenes=None
):

    if not url:

        st.warning(
            "⚠️ Este registro no tiene hipervínculo."
        )

        return

    url = str(url).strip()

    if imagenes is not None:

        urls_fotos = imagenes

    else:

        with st.spinner(
            "📷 Buscando fotografías..."
        ):

            urls_fotos = (
                obtener_urls_fotos_fieldservice(
                    url
                )
            )

    if not urls_fotos:

        st.warning(
            "⚠️ No se pudieron obtener "
            "las fotografías."
        )

        return

    # ========================================================
    # LAS 2 FOTOS DEL MISMO REGISTRO
    # ========================================================

    columnas = st.columns(
        min(
            len(urls_fotos[:2]),
            2
        )
    )

    for indice, url_foto in enumerate(
        urls_fotos[:2]
    ):

        with columnas[indice]:

            st.image(
                url_foto,
                width=300
            )


# ============================================================
# GENERAR HTML PARA PDF DESDE EL NAVEGADOR
# ============================================================

def generar_html_pdf_navegador(df):

    df = df.copy()

    columnas_normalizadas = {}

    for col in df.columns:

        nombre = re.sub(
            r"\s+",
            " ",
            str(col).strip().lower()
        )

        columnas_normalizadas[nombre] = col

    def buscar_columna(*nombres):

        for nombre in nombres:

            nombre_normalizado = re.sub(
                r"\s+",
                " ",
                nombre.strip().lower()
            )

            if (
                nombre_normalizado
                in columnas_normalizadas
            ):

                return columnas_normalizadas[
                    nombre_normalizado
                ]

        return None

    col_suministro = buscar_columna(
        "suministro"
    )

    col_medidor = buscar_columna(
        "serie medidor"
    )

    col_direccion = buscar_columna(
        "direccion",
        "dirección"
    )

    col_obs = buscar_columna(
        "obs"
    )

    col_obs_descripcion = buscar_columna(
        "obs_descripcion"
    )

    col_lectura = buscar_columna(
        "lectura"
    )

    def obtener_valor(
        fila,
        columna
    ):

        if not columna:
            return ""

        valor = fila.get(
            columna,
            ""
        )

        if pd.isna(valor):
            return ""

        if isinstance(
            valor,
            float
        ):

            if valor.is_integer():

                return str(
                    int(valor)
                )

            return str(valor)

        return str(valor).strip()

    def obtener_urls_fotos_pdf(url):

        if not url:
            return []

        url = str(url).strip()

        return (
            st.session_state.get(
                "galeria_imagenes_fieldservice",
                {}
            ).get(
                url,
                []
            )[:2]
        )

    paginas = []

    for indice, (_, fila) in enumerate(
        df.iterrows(),
        start=1
    ):

        suministro = obtener_valor(
            fila,
            col_suministro
        )

        medidor = obtener_valor(
            fila,
            col_medidor
        )

        direccion = obtener_valor(
            fila,
            col_direccion
        )

        obs = obtener_valor(
            fila,
            col_obs
        )

        obs_descripcion = obtener_valor(
            fila,
            col_obs_descripcion
        )

        lectura = obtener_valor(
            fila,
            col_lectura
        )

        url_foto = obtener_valor(
            fila,
            "__url_foto"
        )

        fotos = obtener_urls_fotos_pdf(
            url_foto
        )

        suministro_html = html_lib.escape(
            suministro
        )

        medidor_html = html_lib.escape(
            medidor
        )

        direccion_html = html_lib.escape(
            direccion
        )

        obs_html = html_lib.escape(
            obs
        )

        obs_descripcion_html = html_lib.escape(
            obs_descripcion
        )

        lectura_html = html_lib.escape(
            lectura
        )

        celdas_fotos = []

        for posicion in range(2):

            if posicion < len(fotos):

                url_imagen = html_lib.escape(
                    str(
                        fotos[posicion]
                    ).strip(),
                    quote=True
                )

                celdas_fotos.append(
                    f"""
                    <td class="celda-foto">
                        <img
                            src="{url_imagen}"
                            alt="Foto {posicion + 1}"
                        >
                    </td>
                    """
                )

            else:

                celdas_fotos.append(
                    """
                    <td class="celda-foto">
                        <div class="sin-foto">
                            Sin fotografía
                        </div>
                    </td>
                    """
                )

        pagina = f"""
        <section class="registro">

            <div class="titulo-registro">
                📷 Registro {indice}
            </div>

            <table class="tabla">

                <thead>

                    <tr>
                        <th>SUMINISTRO</th>
                        <th>SERIE MEDIDOR</th>
                        <th>DIRECCIÓN</th>
                        <th>OBS</th>
                        <th>OBS_DESCRIPCION</th>
                        <th>LECTURA</th>
                        <th>FOTO 1</th>
                        <th>FOTO 2</th>
                    </tr>

                </thead>

                <tbody>

                    <tr>

                        <td>
                            {suministro_html}
                        </td>

                        <td>
                            {medidor_html}
                        </td>

                        <td>
                            {direccion_html}
                        </td>

                        <td>
                            {obs_html}
                        </td>

                        <td>
                            {obs_descripcion_html}
                        </td>

                        <td>
                            {lectura_html}
                        </td>

                        {celdas_fotos[0]}

                        {celdas_fotos[1]}

                    </tr>

                </tbody>

            </table>

        </section>
        """

        paginas.append(
            pagina
        )

    contenido = "\n".join(
        paginas
    )

    contenido_json = json.dumps(
        contenido,
        ensure_ascii=False
    )

    html_documento = f"""
<!DOCTYPE html>

<html lang="es">

<head>

<meta charset="UTF-8">

<title>Galería de Fotos Lectura</title>

<style>

@page {{
    size: A3 landscape;
    margin: 8mm;
}}

* {{
    box-sizing: border-box;
}}

html,
body {{
    margin: 0;
    padding: 0;
    font-family: Arial, Helvetica, sans-serif;
}}

body {{
    background: white;
}}

#barra {{
    width: 100%;
    padding: 12px;
    text-align: center;
}}

#btn_pdf {{
    border: none;
    border-radius: 6px;
    padding: 12px 24px;
    font-size: 16px;
    font-weight: bold;
    cursor: pointer;
    background: #ff4b4b;
    color: white;
}}

#btn_pdf:hover {{
    opacity: 0.9;
}}

#estado {{
    margin-top: 8px;
    font-size: 13px;
}}

#contenido_pdf {{
    display: none;
}}

.registro {{
    width: 100%;
    min-height: 281mm;
    page-break-after: always;
    break-after: page;
    overflow: visible;
    padding: 0;
}}

.registro:last-child {{
    page-break-after: auto;
    break-after: auto;
}}

.titulo-registro {{
    font-size: 18px;
    font-weight: bold;
    text-align: left;
    margin-bottom: 5mm;
}}

.tabla {{
    width: 100%;
    min-width: 100%;
    border-collapse: collapse;
    table-layout: fixed;
}}

.tabla th,
.tabla td {{
    border: 0.7px solid #000;
    padding: 2mm;
    text-align: center;
    vertical-align: middle;
    word-break: break-word;
}}

.tabla th {{
    background: #D9E1F2;
    font-size: 10px;
    font-weight: bold;
    height: 12mm;
}}

.tabla td {{
    font-size: 10px;
}}

.tabla th:nth-child(7),
.tabla td:nth-child(7),
.tabla th:nth-child(8),
.tabla td:nth-child(8) {{
    width: 124mm !important;
    min-width: 124mm !important;
    max-width: 124mm !important;
    padding: 1mm !important;
    margin: 0 !important;
}}

.celda-foto {{
    padding: 1mm !important;
    vertical-align: middle;
    text-align: center;
    overflow: hidden;
    white-space: nowrap;
}}

.celda-foto img {{
    display: block;
    width: 120mm;
    max-width: 100%;
    height: auto;
    margin: 0;
    padding: 0;
}}

.sin-foto {{
    color: #777;
    font-size: 10px;
}}

@media print {{

    #barra {{
        display: none !important;
    }}

    #contenido_pdf {{
        display: block !important;
    }}

    body {{
        margin: 0;
        padding: 0;
    }}

}}

</style>

</head>

<body>

<div id="barra">

    <button
        id="btn_pdf"
        onclick="generarPDF()"
    >
        📥 Generar PDF con Fotos
    </button>

    <div id="estado">
        El PDF se abrirá mediante el navegador.
        Seleccione <b>Guardar como PDF</b>.
    </div>

</div>

<div id="contenido_pdf"></div>

<script>

const contenido = {contenido_json};

async function esperarImagenes() {{

    const imagenes = Array.from(
        document.querySelectorAll(
            "#contenido_pdf img"
        )
    );

    await Promise.all(
        imagenes.map(
            imagen => new Promise(
                resolve => {{

                    if (imagen.complete) {{
                        resolve();
                        return;
                    }}

                    imagen.onload = resolve;
                    imagen.onerror = resolve;

                }}
            )
        )
    );
}}

async function generarPDF() {{

    const boton = document.getElementById(
        "btn_pdf"
    );

    const estado = document.getElementById(
        "estado"
    );

    const contenidoPDF =
        document.getElementById(
            "contenido_pdf"
        );

    boton.disabled = true;

    boton.innerText =
        "⏳ Preparando PDF...";

    estado.innerText =
        "Cargando las fotografías ya obtenidas...";

    contenidoPDF.innerHTML = contenido;

    await esperarImagenes();

    estado.innerText =
        "Abriendo el diálogo de impresión...";

    await new Promise(
        resolve => setTimeout(
            resolve,
            300
        )
    );

    window.print();

    boton.disabled = false;

    boton.innerText =
        "📥 Generar PDF con Fotos";

    estado.innerHTML =
        "Seleccione <b>Guardar como PDF</b> para descargarlo.";
}}

window.onafterprint = function() {{

    const contenidoPDF =
        document.getElementById(
            "contenido_pdf"
        );

    contenidoPDF.innerHTML = "";

}};

</script>

</body>

</html>
"""

    return html_documento


# ============================================================
# FUNCIÓN PRINCIPAL
# ============================================================

def ejecutar_galeria_lectura():

    st.title(
        "📷 Galería de Fotos Lectura"
    )

    st.caption(
        "Las fotografías se obtienen automáticamente "
        "desde los enlaces de FieldService de la columna 'foto'."
    )

    # ========================================================
    # CARGAR EXCEL
    # ========================================================

    archivo = st.file_uploader(
        "📂 Seleccionar archivo Excel",
        type=["xlsx", "xls"],
        key="galeria_lectura_excel"
    )

    if archivo is None:

        st.info(
            "Seleccione un archivo Excel para comenzar."
        )

        return

    archivo_id = (
        archivo.name,
        len(archivo.getvalue())
    )

    if (
        st.session_state.get(
            "galeria_archivo_id"
        )
        != archivo_id
    ):

        try:

            with st.spinner(
                "📂 Procesando archivo Excel..."
            ):

                df, columna_foto = (
                    leer_excel_con_links(
                        archivo
                    )
                )

            st.session_state.galeria_archivo_id = (
                archivo_id
            )

            st.session_state.galeria_df = df

            st.session_state.galeria_columna_foto = (
                columna_foto
            )

            st.session_state.galeria_fotos_hasta = 0

            st.session_state.galeria_resultado_filtro_id = None

            st.session_state.pop(
                "galeria_columnas_mostrar",
                None
            )

            st.session_state.pop(
                "galeria_filtros_habilitados",
                None
            )

            # =================================================
            # RESULTADOS POR URL
            # =================================================

            st.session_state.galeria_imagenes_fieldservice = {}

            # =================================================
            # RESULTADOS POR UUID
            # =================================================

            st.session_state.galeria_resultados_uuid = {}

        except Exception as e:

            st.error(
                f"❌ Error al leer el Excel:\n\n{e}"
            )

            return

    df = st.session_state.get(
        "galeria_df"
    )

    columna_foto = st.session_state.get(
        "galeria_columna_foto"
    )

    if df is None:

        st.error(
            "❌ No se pudo cargar el archivo."
        )

        return

    if df.empty:

        st.warning(
            "El Excel no contiene registros."
        )

        return

    total = len(df)

    con_link = (
        df["__url_foto"]
        .notna()
        .sum()
    )

    col1, col2 = st.columns(2)

    with col1:

        st.metric(
            "📋 Registros",
            total
        )

    with col2:

        st.metric(
            "🔗 Registros con foto",
            con_link
        )

    st.divider()

    # ========================================================
    # SIDEBAR
    # ========================================================

    st.sidebar.header(
        "⚙️ Configuración"
    )

    columnas_disponibles = [
        columna
        for columna in df.columns
        if columna not in [
            columna_foto,
            "__url_foto"
        ]
    ]

    st.sidebar.subheader(
        "👁️ Información"
    )

    columnas_mostrar = st.sidebar.multiselect(
        "Seleccionar columnas",
        options=columnas_disponibles,
        default=[],
        key="galeria_columnas_mostrar"
    )

    st.sidebar.subheader(
        "🔎 Filtros"
    )

    filtros_habilitados = st.sidebar.multiselect(
        "Seleccionar filtros",
        options=columnas_disponibles,
        default=[],
        help=(
            "Seleccione las columnas que desea "
            "utilizar como filtros."
        ),
        key="galeria_filtros_habilitados"
    )

    # ========================================================
    # FILTRAR REGISTROS CON FOTO
    # ========================================================

    df_filtrado = df.copy()

    df_filtrado = df_filtrado[
        df_filtrado["__url_foto"].notna()
        &
        (
            df_filtrado["__url_foto"]
            .astype(str)
            .str.strip()
            .ne("")
        )
    ].copy()

    # ========================================================
    # APLICAR FILTROS
    # ========================================================

    for columna in filtros_habilitados:

        valores_filtro = df_filtrado[
            columna
        ].copy()

        valores_filtro = valores_filtro.apply(
            lambda x: (
                str(int(x))
                if pd.notna(x)
                and isinstance(
                    x,
                    (int, float)
                )
                and float(x).is_integer()

                else str(x).strip()
                if pd.notna(x)

                else "[VACÍO]"
            )
        )

        valores = sorted(
            valores_filtro.unique()
        )

        seleccion = st.sidebar.multiselect(
            f"🔎 {columna}",
            options=valores,
            key=f"filtro_valor_{columna}"
        )

        if seleccion:

            df_filtrado = df_filtrado[
                valores_filtro.isin(
                    seleccion
                )
            ].copy()

    # ========================================================
    # RESULTADOS
    # ========================================================

    total_filtrado = len(
        df_filtrado
    )

    st.subheader(
        f"📸 Registros filtrados: {total_filtrado}"
    )

    if df_filtrado.empty:

        st.warning(
            "No existen registros con "
            "los filtros seleccionados."
        )

        return

    # ========================================================
    # DETECTAR CAMBIO DE RESULTADO DEL FILTRO
    # ========================================================

    resultado_filtro_id = tuple(
        df_filtrado.index.tolist()
    )

    if st.session_state.get(
        "galeria_resultado_filtro_id"
    ) != resultado_filtro_id:

        st.session_state.galeria_resultado_filtro_id = (
            resultado_filtro_id
        )

        st.session_state.galeria_fotos_hasta = 0

        # No eliminamos las URLs ya obtenidas.
        # Esto permite reutilizar el caché.


    # ========================================================
    # TAMAÑO DEL BLOQUE
    # ========================================================

    TAMANO_BLOQUE = 200

    fotos_hasta = st.session_state.get(
        "galeria_fotos_hasta",
        0
    )

    # ========================================================
    # TODAVÍA NO CARGAMOS FOTOGRAFÍAS
    # ========================================================

    if fotos_hasta == 0:

        st.info(
            f"📋 Hay {total_filtrado:,} registros "
            "que cumplen los filtros.\n\n"
            f"Se mostrarán en bloques de "
            f"{TAMANO_BLOQUE} registros."
        )

        if st.button(
            "🚀 Cargar fotografías",
            type="primary",
            use_container_width=True,
            key="btn_cargar_fotografias"
        ):

            siguiente = min(
                TAMANO_BLOQUE,
                total_filtrado
            )

            st.session_state.galeria_fotos_hasta = (
                siguiente
            )

            st.rerun()

        st.stop()

    # ========================================================
    # TOMAR REGISTROS QUE SE DEBEN MOSTRAR
    # ========================================================

    registros_a_mostrar = df_filtrado.iloc[
        :fotos_hasta
    ]

    cantidad_mostrada = len(
        registros_a_mostrar
    )

    st.success(
        f"📷 Mostrando {cantidad_mostrada:,} "
        f"de {total_filtrado:,} registros"
    )

    progreso = (
        cantidad_mostrada /
        total_filtrado
    )

    st.progress(
        progreso,
        text=(
            f"{cantidad_mostrada:,} "
            f"de {total_filtrado:,} registros"
        )
    )

    st.divider()

    # ========================================================
    # INICIALIZAR ESTADO DE FOTOS
    # ========================================================

    if (
        "galeria_imagenes_fieldservice"
        not in st.session_state
    ):

        st.session_state.galeria_imagenes_fieldservice = {}

    # ========================================================
    # INICIALIZAR RESULTADOS POR UUID
    # ========================================================

    if (
        "galeria_resultados_uuid"
        not in st.session_state
    ):

        st.session_state.galeria_resultados_uuid = {}

    # ========================================================
    # REGISTROS DEL ÚLTIMO BLOQUE
    #
    # SOLO ESTOS SE CONSULTAN EN ESTA EJECUCIÓN.
    # ========================================================

    inicio_bloque = max(
        0,
        fotos_hasta - TAMANO_BLOQUE
    )

    registros_consulta = list(
        df_filtrado.iloc[
            inicio_bloque:fotos_hasta
        ].iterrows()
    )

    # ========================================================
    # PREPARAR URLS FIELDSERVICE
    # ========================================================

    urls_fieldservice = []

    for _, fila in registros_consulta:

        url = fila.get(
            "__url_foto"
        )

        if isinstance(
            url,
            pd.Series
        ):

            url = (
                url.iloc[0]
                if not url.empty
                else ""
            )

        if pd.isna(url):

            continue

        url = str(url).strip()

        if not url:

            continue

        urls_fieldservice.append(
            url
        )

    # ========================================================
    # QUITAR DUPLICADOS
    # ========================================================

    urls_fieldservice = list(
        dict.fromkeys(
            urls_fieldservice
        )
    )

    # ========================================================
    # OBTENER LAS URLs DEL BLOQUE
    # ========================================================

    urls_pendientes = []

    for url in urls_fieldservice:

        uuid = extraer_uuid_fieldservice(
            url
        )

        if not uuid:

            continue

        # ====================================================
        # SI YA EXISTE RESULTADO POR UUID
        # ====================================================

        if (
            uuid
            in st.session_state.galeria_resultados_uuid
        ):

            resultado_uuid = (
                st.session_state
                .galeria_resultados_uuid
                .get(
                    uuid,
                    []
                )
            )

            st.session_state.galeria_imagenes_fieldservice[
                url
            ] = resultado_uuid

            continue

        # ====================================================
        # SI YA EXISTE RESULTADO POR URL
        # ====================================================

        if (
            url
            in st.session_state.galeria_imagenes_fieldservice
        ):

            continue

        urls_pendientes.append(
            url
        )

    # ========================================================
    # CONSULTAR URLS PENDIENTES
    #
    # 10 CONSULTAS SIMULTÁNEAS
    # PERO SIN MOSTRAR NADA HASTA TERMINAR EL BLOQUE.
    # ========================================================

    if urls_pendientes:

        total_consultas = len(
            urls_pendientes
        )

        with st.spinner(
            f"📷 Obteniendo fotografías... "
            f"0 de {total_consultas:,}"
        ):

            with ThreadPoolExecutor(
                max_workers=10
            ) as executor:

                resultados = executor.map(
                    obtener_urls_fotos_fieldservice,
                    urls_pendientes
                )

                for url, resultado in zip(
                    urls_pendientes,
                    resultados
                ):

                    uuid = extraer_uuid_fieldservice(
                        url
                    )

                    # ========================================
                    # GUARDAR POR UUID
                    # ========================================

                    if uuid:

                        st.session_state.galeria_resultados_uuid[
                            uuid
                        ] = resultado

                    # ========================================
                    # GUARDAR POR URL
                    # ========================================

                    st.session_state.galeria_imagenes_fieldservice[
                        url
                    ] = resultado

    # ========================================================
    # AQUÍ YA TERMINARON TODAS LAS CONSULTAS.
    #
    # AHORA SÍ MOSTRAMOS LAS FOTOS.
    # ========================================================

    st.success(
        "✅ Enlaces de fotografías obtenidos. "
        "Mostrando imágenes..."
    )

    # ========================================================
    # DOS REGISTROS POR FILA
    # ========================================================

    registros = list(
        registros_a_mostrar.iterrows()
    )

    for inicio in range(
        0,
        len(registros),
        2
    ):

        registros_fila = registros[
            inicio:inicio + 2
        ]

        columnas = st.columns(2)

        for posicion_columna, (
            indice_real,
            fila
        ) in enumerate(
            registros_fila
        ):

            url = fila.get(
                "__url_foto"
            )

            if isinstance(
                url,
                pd.Series
            ):

                url = (
                    url.iloc[0]
                    if not url.empty
                    else ""
                )

            if pd.isna(url):

                continue

            url = str(url).strip()

            if not url:

                continue

            with columnas[posicion_columna]:

                with st.container(
                    border=True
                ):

                    # ========================================
                    # NÚMERO DEL REGISTRO
                    # ========================================

                    st.markdown(
                        f"### 📷 Registro {inicio + posicion_columna + 1}"
                    )

                    st.divider()

                    # ========================================
                    # OBTENER LAS DOS FOTOS
                    # ========================================

                    imagenes = (
                        st.session_state
                        .galeria_imagenes_fieldservice
                        .get(
                            url,
                            []
                        )
                    )

                    # ========================================
                    # MOSTRAR LAS DOS FOTOS
                    # ========================================

                    mostrar_fotos(
                        url,
                        imagenes=imagenes
                    )

                    # ========================================
                    # INFORMACIÓN ADICIONAL
                    # ========================================

                    if columnas_mostrar:

                        st.divider()

                        mostrar_datos(
                            fila,
                            columnas_mostrar
                        )

                    # ========================================
                    # OBSERVACIONES
                    # ========================================

                    observaciones = [
                        "Desenfocada",
                        "Sin observación",
                        "Foto borroso",
                        "Foto de lejos",
                        "Celular a Celular"
                    ]

                    st.selectbox(
                        "",
                        observaciones,
                        index=None,
                        placeholder=(
                            "Seleccionar observación..."
                        ),
                        key=(
                            f"observacion_foto_"
                            f"{inicio + posicion_columna + 1}"
                        )
                    )

    # ========================================================
    # SIGUIENTE BLOQUE
    # ========================================================

    if fotos_hasta < total_filtrado:

        restantes = (
            total_filtrado -
            fotos_hasta
        )

        siguiente = min(
            TAMANO_BLOQUE,
            restantes
        )

        st.divider()

        st.info(
            f"📦 Ya se cargaron "
            f"{fotos_hasta:,} registros. "
            f"Quedan {restantes:,}."
        )

        if st.button(
            f"🚀 Cargar siguientes "
            f"{siguiente} registros",
            type="primary",
            use_container_width=True,
            key=(
                f"btn_siguiente_"
                f"{fotos_hasta}"
            )
        ):

            st.session_state.galeria_fotos_hasta = (
                fotos_hasta +
                siguiente
            )

            st.rerun()

    else:

        st.divider()

        st.success(
            f"✅ Se han mostrado los "
            f"{total_filtrado:,} "
            f"registros filtrados."
        )

    # ========================================================
    # EXPORTAR PDF CON FOTOS
    # ========================================================

    st.subheader(
        "📄 Exportar registros"
    )

    st.info(
        "El PDF se generará en formato A3 horizontal, "
        "con un registro por página y hasta dos fotografías "
        "por registro."
    )

    col_pdf, col_excel = st.columns(2)

    # ========================================================
    # PDF
    # ========================================================

    with col_pdf:

        html_pdf = generar_html_pdf_navegador(
            registros_a_mostrar
        )

        components.html(
            html_pdf,
            height=80,
            scrolling=False
        )

    # ========================================================
    # EXCEL
    # ========================================================

    with col_excel:

        df_excel = df_filtrado.copy()

        observaciones_excel = []

        for posicion, (
            indice_real,
            fila
        ) in enumerate(
            df_excel.iterrows(),
            start=1
        ):

            observacion = st.session_state.get(
                f"observacion_foto_{posicion}",
                ""
            )

            if (
                observacion
                == "Seleccionar observación..."
            ):

                observacion = ""

            observaciones_excel.append(
                observacion
            )

        df_excel["Observacion_foto"] = (
            observaciones_excel
        )

        if "__url_foto" in df_excel.columns:

            df_excel = df_excel.drop(
                columns=["__url_foto"]
            )

        excel_salida = BytesIO()

        with pd.ExcelWriter(
            excel_salida,
            engine="openpyxl"
        ) as writer:

            df_excel.to_excel(
                writer,
                index=False,
                sheet_name="Galeria Lectura"
            )

        excel_salida.seek(0)

        st.download_button(
            label="📊 Descargar Excel",
            data=excel_salida,
            file_name="galeria_lectura.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
            key="btn_descargar_excel"
        )