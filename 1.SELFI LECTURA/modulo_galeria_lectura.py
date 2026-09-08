import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from io import BytesIO
from urllib.parse import urljoin
import re
import asyncio
import sys
import subprocess
import html as html_lib
import json
import streamlit.components.v1 as components
from concurrent.futures import ThreadPoolExecutor, as_completed


# ============================================================
# CONFIGURACIÓN
# ============================================================
TIMEOUT = 20
EXTENSIONES_IMAGEN = (
    ".jpg",
    ".jpeg",
    ".png",
    ".webp",
    ".gif"
)


# ============================================================
# CONFIGURACIÓN PLAYWRIGHT PARA WINDOWS + STREAMLIT
# ============================================================
if sys.platform == "win32":
    try:
        asyncio.set_event_loop_policy(
            asyncio.WindowsProactorEventLoopPolicy()
        )
    except Exception:
        pass


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

        if celda.hyperlink:

            url = celda.hyperlink.target

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
                len(df) -
                len(enlaces)
            )
        )

    df["__url_foto"] = enlaces[:len(df)]

    return df, columna_foto


# ============================================================
# EXTRAER ENLACES DE IMÁGENES
# ============================================================
def extraer_imagenes(url):

    """
    Obtiene las URLs de las fotografías desde SIGOF.
    Optimizada para responder más rápido.
    NO se utiliza para FieldService.
    """

    if not url:
        return []

    url = str(url).strip()

    if "servicios.distriluz.com.pe/FieldService" in url:
        return []

    try:

        respuesta = requests.get(
            url,
            timeout=TIMEOUT,
            verify=False,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/131.0.0.0 Safari/537.36"
                ),
                "Accept": "text/html,application/xhtml+xml"
            }
        )

        if respuesta.status_code != 200:
            return []

        content_type = respuesta.headers.get(
            "Content-Type",
            ""
        ).lower()

        if "image/" in content_type:
            return [url]

        soup = BeautifulSoup(
            respuesta.content,
            "html.parser"
        )

        imagenes = []
        vistas = set()

        for img in soup.find_all("img"):

            for atributo in (
                "src",
                "data-src",
                "data-original",
                "data-lazy-src"
            ):

                valor = img.get(atributo)

                if not valor:
                    continue

                valor = str(valor).strip()

                if not valor:
                    continue

                imagen_url = urljoin(
                    url,
                    valor
                )

                if imagen_url in vistas:
                    continue

                extension = (
                    imagen_url
                    .lower()
                    .split("?")[0]
                )

                if extension.endswith(
                    EXTENSIONES_IMAGEN
                ):

                    vistas.add(imagen_url)
                    imagenes.append(imagen_url)

        for enlace in soup.find_all(
            "a",
            href=True
        ):

            href = str(
                enlace.get("href")
            ).strip()

            if not href:
                continue

            imagen_url = urljoin(
                url,
                href
            )

            if imagen_url in vistas:
                continue

            extension = (
                imagen_url
                .lower()
                .split("?")[0]
            )

            if extension.endswith(
                EXTENSIONES_IMAGEN
            ):

                vistas.add(imagen_url)
                imagenes.append(imagen_url)

        return imagenes

    except requests.RequestException:
        return []

    except Exception:
        return []


# ============================================================
# EXTRAER IMÁGENES SIGOF EN PARALELO
# ============================================================
def extraer_imagenes_sigof_paralelo(
    urls,
    trabajadores=10
):

    resultados = {}

    urls_validas = []

    for url in urls:

        if not url:
            continue

        url = str(url).strip()

        if not url:
            continue

        if (
            "servicios.distriluz.com.pe/FieldService"
            in url
        ):
            continue

        urls_validas.append(url)

    if not urls_validas:
        return resultados

    urls_validas = list(
        dict.fromkeys(urls_validas)
    )

    with ThreadPoolExecutor(
        max_workers=trabajadores
    ) as executor:

        futuros = {
            executor.submit(
                extraer_imagenes,
                url
            ): url
            for url in urls_validas
        }

        for futuro in as_completed(futuros):

            url = futuros[futuro]

            try:

                resultados[url] = (
                    futuro.result()
                )

            except Exception:

                resultados[url] = []

    return resultados


# ============================================================
# OBTENER FOTOS FIELDSERVICE EN PARALELO
# ============================================================
def obtener_fotos_fieldservice_paralelo(
    urls,
    trabajadores=10
):

    resultados = {}

    urls_validas = []

    for url in urls:

        if not url:
            continue

        url = str(url).strip()

        if not url:
            continue

        if "servicios.distriluz.com.pe/FieldService" not in url:
            continue

        urls_validas.append(url)

    urls_validas = list(
        dict.fromkeys(urls_validas)
    )

    if not urls_validas:
        return resultados

    with ThreadPoolExecutor(
        max_workers=trabajadores
    ) as executor:

        futuros = {
            executor.submit(
                obtener_urls_fotos_fieldservice,
                url
            ): url
            for url in urls_validas
        }

        for futuro in as_completed(futuros):

            url = futuros[futuro]

            try:

                resultados[url] = (
                    futuro.result()
                )

            except Exception:

                resultados[url] = []

    return resultados


# ============================================================
# OBTENER ENLACES DIRECTOS DE FOTOS FIELDSERVICE
# MEDIANTE LA API
# ============================================================
def obtener_urls_fotos_fieldservice(url):

    if not url:
        return []

    url = str(url).strip()

    if not url:
        return []

    if "servicios.distriluz.com.pe/FieldService" not in url:
        return []

    try:

        uuid = url.rstrip("/").split("/")[-1]

        if not uuid:
            return []

        url_api = (
            "https://servicios.distriluz.com.pe:51000/"
            "OptimusNGC_FieldService/api/reportes-publicos/fotos/"
            + uuid
        )

        respuesta = requests.get(
            url_api,
            headers={
                "Accept": "application/json",
                "User-Agent": (
                    "Mozilla/5.0 "
                    "(Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 "
                    "(KHTML, like Gecko) "
                    "Chrome/139.0.0.0 "
                    "Safari/537.36"
                )
            },
            timeout=TIMEOUT
        )

        if respuesta.status_code != 200:
            return []

        json_texto = respuesta.text

        if not json_texto:
            return []

        url_lectura = ""

        patron_lectura = re.search(
            r'"tipo"\s*:\s*"LECTURA".*?"url"\s*:\s*"([^"]+)"',
            json_texto,
            flags=re.IGNORECASE | re.DOTALL
        )

        if patron_lectura:

            url_lectura = (
                patron_lectura
                .group(1)
                .strip()
            )

        url_medidor = ""

        patron_medidor = re.search(
            r'"tipo"\s*:\s*"MEDIDOR".*?"url"\s*:\s*"([^"]+)"',
            json_texto,
            flags=re.IGNORECASE | re.DOTALL
        )

        if patron_medidor:

            url_medidor = (
                patron_medidor
                .group(1)
                .strip()
            )

        urls_fotos = []

        if url_lectura:
            urls_fotos.append(url_lectura)

        if url_medidor:
            urls_fotos.append(url_medidor)

        return urls_fotos[:2]

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

        elif isinstance(valor, float):

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
    imagenes_sigof=None
):

    if not url:

        st.warning(
            "⚠️ Este registro no tiene hipervínculo."
        )

        return

    url = str(url).strip()

    # ========================================================
    # FIELDSERVICE
    # ========================================================
    if "servicios.distriluz.com.pe/FieldService" in url:

        if imagenes_sigof is not None:

            urls_fotos = imagenes_sigof

        else:

            urls_fotos = (
                obtener_urls_fotos_fieldservice(
                    url
                )
            )

        if urls_fotos:

            columnas_fieldservice = st.columns(
                min(
                    len(urls_fotos[:2]),
                    2
                )
            )

            for indice, url_foto in enumerate(
                urls_fotos[:2]
            ):

                with columnas_fieldservice[indice]:

                    st.image(
                        url_foto,
                        width=300
                    )

        else:

            st.warning(
                "⚠️ No se pudieron obtener "
                "las fotografías."
            )

        return

    # ========================================================
    # SIGOF
    # ========================================================
    if imagenes_sigof is not None:

        imagenes = imagenes_sigof

    else:

        with st.spinner(
            "📷 Buscando fotografías..."
        ):

            imagenes = extraer_imagenes(
                url
            )

    if not imagenes:

        st.warning(
            "⚠️ No se encontraron imágenes "
            "en este enlace."
        )

        st.link_button(
            "🔗 Abrir enlace original",
            url
        )

        return

    st.caption(
        f"📷 {len(imagenes)} fotografía(s)"
    )

    if len(imagenes) == 1:

        st.image(
            imagenes[0],
            use_container_width=True
        )

    else:

        columnas = st.columns(
            min(
                len(imagenes),
                2
            )
        )

        for posicion, imagen in enumerate(
            imagenes
        ):

            with columnas[
                posicion % 2
            ]:

                st.image(
                    imagen,
                    use_container_width=True
                )


# ============================================================
# MOSTRAR UN REGISTRO
# ============================================================
def mostrar_registro(
    indice,
    fila,
    columnas_mostrar,
    imagenes_sigof=None
):

    url = fila.get(
        "__url_foto"
    )

    if pd.isna(url) or not str(url).strip():
        return

    with st.container(
        border=True
    ):
        
        st.divider()

        mostrar_fotos(
            url,
            imagenes_sigof=imagenes_sigof
        )

        if columnas_mostrar:

            st.divider()

            mostrar_datos(
                fila,
                columnas_mostrar
            )

        
# ============================================================
# MOSTRAR REGISTRO PROGRESIVO
# ============================================================
def mostrar_registro_progresivo(
    indice,
    fila,
    columnas_mostrar
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
        url = ""

    url = str(url).strip()

    placeholder_foto_resultado = None

    es_fieldservice = (
        "servicios.distriluz.com.pe/FieldService"
        in url
    )

    with st.container(
        border=True
    ):

        st.markdown(
            f"### 📷 Registro {indice}"
        )

        st.divider()

        placeholder_foto = st.empty()

        if es_fieldservice:

            imagenes = (
                st.session_state.get(
                    "galeria_imagenes_fieldservice",
                    {}
                ).get(url)
            )

        else:

            imagenes = (
                st.session_state.get(
                    "galeria_imagenes_sigof",
                    {}
                ).get(url)
            )

        if imagenes is not None:

            with placeholder_foto.container():

                mostrar_fotos(
                    url,
                    imagenes_sigof=imagenes
                )

        else:

            placeholder_foto.info(
                "⏳ Cargando fotografía..."
            )

            placeholder_foto_resultado = (
                placeholder_foto
            )

        for columna in columnas_mostrar:

            valor = fila.get(
                columna,
                ""
            )

            if pd.isna(valor):
                valor = ""

            st.markdown(
                f"**{columna}:** {valor}"
            )

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
            placeholder="Seleccionar observación...",
            key=f"observacion_foto_{indice}"
        )

    return placeholder_foto_resultado


# ============================================================
# DETECTAR SI LA URL ES FIELDSERVICE
# ============================================================
def es_fieldservice(url):

    if not url:
        return False

    return (
        "servicios.distriluz.com.pe/FieldService"
        in str(url)
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
        "suministro",
        "nro suministro",
        "n° suministro",
        "numero suministro",
        "número suministro"
    )

    col_medidor = buscar_columna(
        "medidor",
        "nro medidor",
        "n° medidor",
        "numero medidor",
        "número medidor"
    )

    col_direccion = buscar_columna(
        "direccion",
        "dirección"
    )

    col_obs = buscar_columna(
        "obs"
    )

    col_obs_descripcion = buscar_columna(
        "obs_descripcion",
        "obs descripcion",
        "obs descripción",
        "observacion descripcion",
        "observación descripción"
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

        if isinstance(valor, float):

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

        if es_fieldservice(url):

            return (
                st.session_state.get(
                    "galeria_imagenes_fieldservice",
                    {}
                ).get(
                    url,
                    []
                )[:2]
            )

        return (
            st.session_state.get(
                "galeria_imagenes_sigof",
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
                        <th>MEDIDOR</th>
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

<div
    id="contenido_pdf"
>
</div>

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

    const contenidoPDF = document.getElementById(
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
        "desde el hipervínculo de la columna 'foto'."
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

            st.session_state.pop(
                "galeria_job_id",
                None
            )

            st.session_state.pop(
                "galeria_proceso_terminado",
                None
            )

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
                and isinstance(x, (int, float))
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

        st.session_state.pop(
            "galeria_job_id",
            None
        )

        st.session_state.pop(
            "galeria_proceso_terminado",
            None
        )

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

            st.session_state.pop(
                "galeria_job_id",
                None
            )

            st.session_state.pop(
                "galeria_proceso_terminado",
                None
            )

            st.rerun()

        st.stop()

    # ========================================================
    # TOMAR SOLO LOS REGISTROS DEL BLOQUE
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

    registros = list(
        registros_a_mostrar.iterrows()
    )

    # ============================================================
    # PREPARAR CONSULTAS SIGOF Y FIELDSERVICE
    # ============================================================
    urls_sigof = []
    urls_fieldservice = []

    for _, fila in registros:

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

        if "servicios.distriluz.com.pe/FieldService" in url:

            urls_fieldservice.append(
                url
            )

        else:

            urls_sigof.append(
                url
            )

    urls_sigof = list(
        dict.fromkeys(
            urls_sigof
        )
    )

    urls_fieldservice = list(
        dict.fromkeys(
            urls_fieldservice
        )
    )

    # ============================================================
    # INICIALIZAR CACHÉ DE FOTOS
    # ============================================================
    if "galeria_imagenes_sigof" not in st.session_state:

        st.session_state.galeria_imagenes_sigof = {}

    if "galeria_imagenes_fieldservice" not in st.session_state:

        st.session_state.galeria_imagenes_fieldservice = {}

    # ============================================================
    # CONFIGURACIÓN
    # ============================================================
    TAMANO_CONSULTAS = 20

    # ============================================================
    # FUNCIÓN PARA DIBUJAR LOS REGISTROS YA LISTOS
    # ============================================================
    def renderizar_registros_listos():

        registros_listos = (
            st.session_state.get(
                "galeria_registros_listos",
                []
            )
        )

        for contador, (
            posicion,
            indice_real,
            fila
        ) in enumerate(
            registros_listos
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

            if contador % 2 == 0:

                columnas = st.columns(2)

            columna = columnas[
                contador % 2
            ]

            with columna:

                with st.container(
                    border=True
                ):

                    st.markdown(
                        f"### 📷 Registro {posicion + 1}"
                    )

                    st.divider()

                    if es_fieldservice(url):

                        imagenes = (
                            st.session_state
                            .galeria_imagenes_fieldservice
                            .get(
                                url,
                                []
                            )
                        )

                    else:

                        imagenes = (
                            st.session_state
                            .galeria_imagenes_sigof
                            .get(
                                url,
                                []
                            )
                        )

                    mostrar_fotos(
                        url,
                        imagenes_sigof=imagenes
                    )

                    if columnas_mostrar:

                        st.divider()

                        mostrar_datos(
                            fila,
                            columnas_mostrar
                        )
                    
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
                            f"{posicion + 1}"
                        )
                    )

    # ============================================================
    # IDENTIFICADOR DEL PROCESO ACTUAL
    # ============================================================
    job_id = (
        resultado_filtro_id,
        fotos_hasta
    )

    # ============================================================
    # CREAR EL PROCESO SOLO UNA VEZ
    # ============================================================
    if st.session_state.get(
        "galeria_job_id"
    ) != job_id:

        executor_anterior = st.session_state.get(
            "galeria_executor"
        )

        if executor_anterior is not None:

            try:

                executor_anterior.shutdown(
                    wait=False,
                    cancel_futures=True
                )

            except Exception:

                pass

        urls_pendientes = []

        # --------------------------------------------------------
        # FIELDSERVICE
        # --------------------------------------------------------
        for url in urls_fieldservice:

            if (
                url
                not in st.session_state.galeria_imagenes_fieldservice
            ):

                urls_pendientes.append(
                    (
                        "fieldservice",
                        url
                    )
                )

        # --------------------------------------------------------
        # SIGOF
        # --------------------------------------------------------
        for url in urls_sigof:

            if (
                url
                not in st.session_state.galeria_imagenes_sigof
            ):

                urls_pendientes.append(
                    (
                        "sigof",
                        url
                    )
                )

        # ========================================================
        # REGISTROS YA CARGADOS
        # ========================================================
        registros_listos = []

        for posicion, (
            indice_real,
            fila
        ) in enumerate(registros):

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

            if es_fieldservice(url):

                ya_cargada = (
                    url
                    in st.session_state.galeria_imagenes_fieldservice
                )

            else:

                ya_cargada = (
                    url
                    in st.session_state.galeria_imagenes_sigof
                )

            if ya_cargada:

                registros_listos.append(
                    (
                        posicion,
                        indice_real,
                        fila
                    )
                )

        # ========================================================
        # GUARDAR ESTADO
        # ========================================================
        st.session_state.galeria_job_id = (
            job_id
        )

        st.session_state.galeria_urls_pendientes = (
            urls_pendientes
        )

        st.session_state.galeria_futuros = {}

        st.session_state.galeria_registros_listos = (
            registros_listos
        )

        st.session_state.galeria_consultas_terminadas = 0

        st.session_state.galeria_total_consultas = (
            len(urls_pendientes)
        )

        # ========================================================
        # SI NO HAY CONSULTAS PENDIENTES
        # ========================================================
        if not urls_pendientes:

            st.session_state.galeria_executor = None

            st.session_state.galeria_proceso_terminado = (
                True
            )

        else:

            st.session_state.galeria_proceso_terminado = (
                False
            )

            # ====================================================
            # CREAR EJECUTOR CON MÁXIMO 20 CONSULTAS
            # ====================================================
            executor = ThreadPoolExecutor(
                max_workers=TAMANO_CONSULTAS
            )

            st.session_state.galeria_executor = (
                executor
            )

            # ====================================================
            # LANZAR SOLO LAS PRIMERAS 20
            # ====================================================
            lote_inicial = (
                urls_pendientes[
                    :TAMANO_CONSULTAS
                ]
            )

            st.session_state.galeria_urls_pendientes = (
                urls_pendientes[
                    TAMANO_CONSULTAS:
                ]
            )

            for tipo, url in lote_inicial:

                if tipo == "fieldservice":

                    futuro = executor.submit(
                        obtener_urls_fotos_fieldservice,
                        url
                    )

                else:

                    futuro = executor.submit(
                        extraer_imagenes,
                        url
                    )

                st.session_state.galeria_futuros[
                    futuro
                ] = (
                    tipo,
                    url
                )

    # ============================================================
    # PROCESO PROGRESIVO
    # ============================================================
    if not st.session_state.get(
        "galeria_proceso_terminado",
        False
    ):

        @st.fragment(
            run_every=0.3
        )
        def actualizar_galeria():

            futuros_pendientes = (
                st.session_state.get(
                    "galeria_futuros",
                    {}
                )
            )

            # ====================================================
            # DETECTAR CONSULTAS TERMINADAS
            # ====================================================
            terminados = []

            for futuro in list(
                futuros_pendientes.keys()
            ):

                if futuro.done():

                    terminados.append(
                        futuro
                    )

            # ====================================================
            # PROCESAR TODAS LAS TERMINADAS
            # ====================================================
            for futuro in terminados:

                tipo, url = (
                    futuros_pendientes.pop(
                        futuro
                    )
                )

                try:

                    resultado = (
                        futuro.result()
                    )

                except Exception:

                    resultado = []

                # ------------------------------------------------
                # GUARDAR RESULTADO
                # ------------------------------------------------
                if tipo == "fieldservice":

                    st.session_state.galeria_imagenes_fieldservice[
                        url
                    ] = resultado

                else:

                    st.session_state.galeria_imagenes_sigof[
                        url
                    ] = resultado

                # ------------------------------------------------
                # CONTADOR
                # ------------------------------------------------
                st.session_state.galeria_consultas_terminadas += 1

                # =================================================
                # AGREGAR TODOS LOS REGISTROS QUE USAN ESA URL
                # =================================================
                registros_listos = (
                    st.session_state.get(
                        "galeria_registros_listos",
                        []
                    )
                )

                posiciones_existentes = {
                    registro[0]
                    for registro in registros_listos
                }

                for posicion, (
                    indice_real,
                    fila
                ) in enumerate(registros):

                    url_registro = fila.get(
                        "__url_foto"
                    )

                    if isinstance(
                        url_registro,
                        pd.Series
                    ):

                        url_registro = (
                            url_registro.iloc[0]
                            if not url_registro.empty
                            else ""
                        )

                    if pd.isna(url_registro):
                        continue

                    url_registro = str(
                        url_registro
                    ).strip()

                    if (
                        url_registro == url
                        and posicion
                        not in posiciones_existentes
                    ):

                        registros_listos.append(
                            (
                                posicion,
                                indice_real,
                                fila
                            )
                        )

                        posiciones_existentes.add(
                            posicion
                        )

                st.session_state.galeria_registros_listos = (
                    registros_listos
                )

            # ====================================================
            # MANTENER SIEMPRE 20 CONSULTAS ACTIVAS
            # ====================================================
            executor_actual = (
                st.session_state.get(
                    "galeria_executor"
                )
            )

            urls_pendientes = (
                st.session_state.get(
                    "galeria_urls_pendientes",
                    []
                )
            )

            futuros_actuales = (
                st.session_state.get(
                    "galeria_futuros",
                    {}
                )
            )

            while (
                executor_actual is not None
                and len(futuros_actuales)
                < TAMANO_CONSULTAS
                and urls_pendientes
            ):

                siguiente_tipo, siguiente_url = (
                    urls_pendientes.pop(0)
                )

                if siguiente_tipo == "fieldservice":

                    nuevo_futuro = executor_actual.submit(
                        obtener_urls_fotos_fieldservice,
                        siguiente_url
                    )

                else:

                    nuevo_futuro = executor_actual.submit(
                        extraer_imagenes,
                        siguiente_url
                    )

                futuros_actuales[
                    nuevo_futuro
                ] = (
                    siguiente_tipo,
                    siguiente_url
                )

            st.session_state.galeria_urls_pendientes = (
                urls_pendientes
            )

            st.session_state.galeria_futuros = (
                futuros_actuales
            )

            # ====================================================
            # ESTADO
            # ====================================================
            total_consultas = (
                st.session_state.get(
                    "galeria_total_consultas",
                    0
                )
            )

            consultas_terminadas = (
                st.session_state.get(
                    "galeria_consultas_terminadas",
                    0
                )
            )

            consultas_activas = (
                len(
                    futuros_actuales
                )
            )

            consultas_en_espera = (
                len(
                    urls_pendientes
                )
            )

            if total_consultas > 0:

                st.info(
                    f"📷 Fotografías listas: "
                    f"{consultas_terminadas:,} "
                    f"de {total_consultas:,}"
                    f"  |  🔄 Procesando: "
                    f"{consultas_activas}"
                    f"  |  📦 En espera: "
                    f"{consultas_en_espera:,}"
                )

            # ====================================================
            # MOSTRAR REGISTROS QUE YA ESTÁN LISTOS
            # ====================================================
            renderizar_registros_listos()

            # ====================================================
            # VERIFICAR FIN DEL PROCESO
            # ====================================================
            if (
                not futuros_actuales
                and not urls_pendientes
            ):

                st.session_state.galeria_proceso_terminado = (
                    True
                )

                executor_final = (
                    st.session_state.get(
                        "galeria_executor"
                    )
                )

                if executor_final is not None:

                    try:

                        executor_final.shutdown(
                            wait=False
                        )

                    except Exception:

                        pass

                # =================================================
                # IMPORTANTE:
                # HACER RERUN COMPLETO PARA QUE LA GALERÍA
                # PERMANEZCA Y LUEGO APAREZCAN PDF/EXCEL
                # =================================================
                st.rerun()

        actualizar_galeria()

        # ========================================================
        # DETENER EL RESTO MIENTRAS SE CARGAN FOTOS
        # ========================================================
        st.stop()

    # ============================================================
    # PROCESO TERMINADO
    # ============================================================
    # AQUÍ ESTÁ EL CAMBIO IMPORTANTE:
    #
    # Después del st.rerun() anterior, Streamlit vuelve a ejecutar
    # la aplicación completa. Como el proceso ya está terminado,
    # el fragmento no se ejecuta nuevamente.
    #
    # Por eso debemos dibujar nuevamente los registros aquí.
    # ============================================================

    renderizar_registros_listos()

    st.success(
        "✅ Proceso terminado. "
        "Todas las fotografías del bloque están listas."
    )

    # ============================================================
    # SIGUIENTE BLOQUE
    # ============================================================
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

            st.session_state.pop(
                "galeria_job_id",
                None
            )

            st.session_state.pop(
                "galeria_proceso_terminado",
                None
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

    # ========================================================
    # BOTONES DE EXPORTACIÓN
    # ========================================================
    col_pdf, col_excel = st.columns(2)

    # ========================================================
    # GENERAR PDF DESDE EL NAVEGADOR
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
    # DESCARGAR EXCEL
    # ========================================================
    with col_excel:

        df_excel = df_filtrado.copy()

        # ====================================================
        # AGREGAR OBSERVACIÓN
        # ====================================================
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

            if observacion == "Seleccionar observación...":

                observacion = ""

            observaciones_excel.append(
                observacion
            )

        df_excel["Observacion_foto"] = (
            observaciones_excel
        )

        # ====================================================
        # ELIMINAR COLUMNA INTERNA
        # ====================================================
        if "__url_foto" in df_excel.columns:

            df_excel = df_excel.drop(
                columns=["__url_foto"]
            )

        # ====================================================
        # CREAR EXCEL EN MEMORIA
        # ====================================================
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

        # ====================================================
        # BOTÓN
        # ====================================================
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