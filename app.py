import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode
from io import BytesIO
import uuid  # Asegúrate de importar uuid al inicio del archivo
import requests
from bs4 import BeautifulSoup
import msal
from urllib.parse import urlencode, urlparse, parse_qs
import openpyxl
from openpyxl import load_workbook

###############################################################################
# 1. CONFIGURACIÓN INICIAL STREAMLIT
###############################################################################
# Título: "AppWebSGA"
# Favicon: URL de la imagen solicitada
st.set_page_config(
    page_title="AppWebSGA",
    page_icon="https://blogger.googleusercontent.com/img/a/AVvXsEgqcaKJ1VLBjTRUn-Jz8DNxGx2xuonGQitE2rZjDm_y_uLKe1_6oi5qMiinWMB91JLtS5IvR4Tj-RU08GEfx7h8FdXAEI5HuNoV9YumyfwyXL5qFQ6MJmZw2sKWqR6LWWT8OuEGEkIRRnS2cqP86TgHOoBVkqPPSIRgnHGa4uSEu4O4gM0iNBb7a8Dunfw1",
    layout="wide"
)

# Configuración de la aplicación (usando valores desde st.secrets)
CLIENT_ID = st.secrets["ms_graph"]["client_id"]
CLIENT_SECRET = st.secrets["ms_graph"]["client_secret"]
TENANT_ID = st.secrets["ms_graph"]["tenant_id"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Files.ReadWrite", "User.Read"] # Permisos para leer y escribir en OneDrive


###############################################################################
# 2. CSS PARA PERSONALIZAR LA BARRA LATERAL Y ESTILOS GENERALES
###############################################################################
st.markdown("""
<style>
/* Mover la barra lateral a la derecha */
.css-18e3th9 {
    flex-direction: row-reverse;
}
/* Ajustar ancho (un poco más angosta) */
section[data-testid="stSidebar"] {
    min-width: 250px !important;
    max-width: 250px !important;
    background-color: #001F3F !important;
    color: white !important;
}
/* Forzar color blanco en la sidebar (incluyendo emojis) */
section[data-testid="stSidebar"] * {
    color: #ffffff !important;
}

/* Ajuste general para no superponer */
.css-1laz8t5 {
    flex: 1 1 0%;
}

/* Contenedor principal */
.main-container {
  background-color: #F3F3F3;
  padding: 20px 30px;
  border-radius: 10px;
}

/* Tarjetas (subdivisiones) */
.card {
  background: #fff;
  border: 1px solid #ddd;
  border-radius: 8px;
  padding:16px;
  margin-bottom:16px;
  box-shadow:0 2px 5px rgba(0,0,0,.1);
}

/* Títulos */
h2.title {
  font-size:26px;
  color:#333;
  margin-bottom: 0;
}
p.credit {
  margin: 0;
  font-size: 14px;
  color: #888;
}
p.desc {
  margin-top: 8px;
  font-size:15px;
  color:#444;
}
</style>
""", unsafe_allow_html=True)
# Función para obtener el correo electrónico del usuario desde Microsoft Graph
def get_user_email(access_token):
    url = "https://graph.microsoft.com/v1.0/me"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    data = response.json()
    if "mail" in data:
        return data["mail"]
    elif "userPrincipalName" in data:  # Si "mail" no está disponible, se utiliza el nombre principal del usuario
        return data["userPrincipalName"]
    return None


# Función para listar archivos de OneDrive
def list_onedrive_files(access_token):
    """Obtiene la lista de archivos desde OneDrive usando el token de acceso."""
    url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    data = response.json()
    if "value" in data:
        return data["value"]  # Lista de archivos
    return None

redirect_uri = "https://localhost:8501/signout-oidc/"
# redirect_uri = st.secrets["ms_graph"]["redirect_uri"]

# Función para generar el enlace de autorización con la redirección correcta
def get_authorization_url():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    auth_url = app.get_authorization_request_url(SCOPES, redirect_uri=redirect_uri)
    print(f"auth_url$:${auth_url}")
    return auth_url
# Función para intercambiar el código de autorización por un token de acceso
def get_access_token_from_code(auth_code):
    print(f"${auth_code}")
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    result = app.acquire_token_by_authorization_code(auth_code, scopes=SCOPES, redirect_uri=redirect_uri)
    print(f"${result}")
    # Verificar si se obtuvo el token
    if result is None:
        st.error("El resultado de la autenticación es None. Verifica los parámetros de la solicitud.")
        return None
    
    if "access_token" in result:
        return result["access_token"]
    else:
        st.error(f"Error al obtener el token de acceso: {result.get('error_description')}")
        return None

def login_button():
    auth_url = get_authorization_url()
    

# Función para obtener el código de autorización desde la URL
def get_auth_code_from_url():
    url = st.query_params  # Obtener los parámetros de la URL
    if 'code' in url:
        return url['code'][0]  # Extraer el código de autorización
    return None

###############################################################################
# 3. TABLA INTERACTIVA SIN AUTO-ACTUALIZACIÓN (NO_UPDATE)
###############################################################################

# Función para mostrar y exportar el DataFrame
def interactive_table_no_autoupdate(df: pd.DataFrame, key: str = None) -> pd.DataFrame:
    """
    Muestra un DataFrame con st_aggrid usando update_mode=NO_UPDATE:
      - La tabla es interactiva: los usuarios pueden filtrar y ordenar los datos
      - Exportación a Excel para descargar los datos.
    """
    # Configurar la tabla interactiva con AgGrid
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(filter=True)  # Habilitar filtros sin permitir edición
    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=20)
    gb.configure_grid_options(
        paginationPageSize=20,
        paginationPageSizeOptions=[20, 50, 100, 200, 500]
    )
    gb.configure_side_bar()

    grid_options = gb.build()
    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        theme="blue",
        key=key
    )

    # Botón para exportar la tabla a Excel
    if st.button("Exportar a Excel", key=f"exportar_excel_{key}"):
        # Crear el archivo Excel en memoria (sin guardarlo en el disco)
        output = BytesIO()

        # Usamos el motor 'openpyxl' para crear el archivo Excel
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Datos")
            # No es necesario llamar a writer.save(), ya que openpyxl lo maneja automáticamente

        output.seek(0)  # Volver al inicio del archivo

        # Crear el botón de descarga usando `st.download_button`
        st.download_button(
            label="Descargar archivo Excel",
            data=output,
            file_name=f"{key}_editado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    return df


###############################################################################
# 4. FUNCIONES AUXILIARES DE NEGOCIO
###############################################################################
def set_directories():
    """Crea la carpeta local 'DATA_MAUI_PJLT' si no existe."""
    os.makedirs('DATA_MAUI_PJLT', exist_ok=True)

def extraer_descripcion(descripcion):
    valores_bano = [
        "T. Baño Entero", "T. Baño Stretch", "T. Baño Corto",
        "T. Baño Microfibra", "Traje de Baño", "T.BAÑO"
    ]
    if isinstance(descripcion, str):
        if descripcion[:2] == " *":
            match = re.search(r'Pack\s([A-Za-z\s]+)\s\d', descripcion)
            resultado = match.group(1).strip() if match else None
        else:
            match = re.search(r'\*\s([A-Za-z\s]+)\s\d', descripcion)
            resultado = match.group(1).strip() if match else None

        if resultado in valores_bano:
            return "Traje de Baño"
        return resultado
    else:
        return None

def extraer_codigo_marca(descripcion, subfamilia):
    if not isinstance(descripcion, str) or not isinstance(subfamilia, str):
        return None
    patron = rf"{re.escape(subfamilia)}\s+([\w\-]+)"
    match = re.search(patron, descripcion, flags=re.IGNORECASE)
    if match:
        code = match.group(1)
        if "-" in code:
            code = code.split("-")[0]
        return code.strip()
    return None

def calcular_marca(codigo_marca):
    if not isinstance(codigo_marca, str) or not codigo_marca.strip():
        return None
    match = re.search(r'(\d)', codigo_marca)
    if match:
        d = match.group(1)
        if d == '5':
            return "Maui"
        elif d == '6':
            return "Rip Curl"
        elif d in ['4', '7']:
            return "Rusty y Otros"
    return None

def calcular_zona(marca):
    if marca == "Maui":
        return "B2.ME02 - B2.ME03 - B2.ME04.C09 a B2.ME04.C15"
    elif marca == "Rip Curl":
        return "B1.ME03 - B1.M0E04"
    elif marca == "Rusty y Otros":
        return "B2.ME04.C01 a B2.ME04.C05"
    return None
#Función auxiliar para calcular Tipo_Pack
def calcular_tipo_pack(descripcion: str) -> str:
    """
    Retorna "Inner" si los dos primeros caracteres son " *",
    de lo contrario retorna "Unidad".
    """
    if isinstance(descripcion, str) and len(descripcion) >= 2 and descripcion[:2] == " *":
        return "Inner"
    return "Unidad"
def calcular_factor_por_caja(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula la columna 'Factor por Caja' en el dataframe `df` basado en:
    - Agrupación por la columna 'Prtnum Padre'
    - Suma de la columna 'Cantidad Empleada'
    """
    # Verificamos que existan las columnas requeridas
    required_columns = ["Prtnum Padre", "Cantidad Empleada"]
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Falta la columna obligatoria '{col}' en df_curva_articulo")

    # Usamos 'transform' para sumar en cada grupo el valor de 'Cantidad Empleada'
    df["Factor por Caja"] = df.groupby("Prtnum Padre")["Cantidad Empleada"].transform("sum")

    return df
def calcular_factor_caja(df_consolidado: pd.DataFrame, 
                         df_curva_articulo: pd.DataFrame) -> pd.DataFrame:
    """
    1. Toma el valor de 'Num Producto' en df_consolidado,
    2. Busca coincidencia en df_curva_articulo['Prtnum Padre']
    3. Retorna el 'Factor por Caja' si hay coincidencia; si no, 1.
    """
    # Verificar que existan las columnas necesarias
    if "Num Producto" not in df_consolidado.columns:
        raise ValueError("No se encontró la columna 'Num Producto' en df_consolidado.")
    if "Prtnum Padre" not in df_curva_articulo.columns:
        raise ValueError("No se encontró la columna 'Prtnum Padre' en df_curva_articulo.")
    if "Factor por Caja" not in df_curva_articulo.columns:
        raise ValueError("No se encontró la columna 'Factor por Caja' en df_curva_articulo.")

    # Creamos un diccionario para mapear: { Prtnum Padre -> Factor por Caja }
    mapping = dict(zip(df_curva_articulo["Prtnum Padre"], df_curva_articulo["Factor por Caja"]))

    # Para cada fila de df_consolidado, busca 'Num Producto' en mapping; si no existe, asigna 1
    df_consolidado["Factor_Caja"] = df_consolidado["Num Producto"].map(mapping).fillna(1)
    return df_consolidado


def calcular_qty_inners(df_consolidado: pd.DataFrame) -> pd.DataFrame:
    """
    1. Si Tipo_Pack == "Inner" => Qty_Inners = Cantidad Esperada
    2. Si Tipo_Pack == "Unidad" => Qty_Inners = Cant Cajas
    """
    # Verificar que existan las columnas necesarias
    required_cols = ["Tipo_Pack", "Cantidad Esperada", "Cant Cajas"]
    for col in required_cols:
        if col not in df_consolidado.columns:
            raise ValueError(f"No se encontró la columna '{col}' en df_consolidado.")

    # Aplicar la lógica
    df_consolidado["Qty_Inners"] = df_consolidado.apply(
        lambda row: row["Cantidad Esperada"] if row["Tipo_Pack"] == "Inner" 
                    else row["Cant Cajas"],
        axis=1
    )
    return df_consolidado


def calcular_qty_unidades(df_consolidado: pd.DataFrame) -> pd.DataFrame:
    """
    1. Si Tipo_Pack == "Inner" => Qty_Unidades = Qty_Inners * Factor_Caja
    2. Si Tipo_Pack == "Unidad" => Qty_Unidades = Cantidad Esperada
    """
    # Verificar que existan las columnas necesarias
    required_cols = ["Tipo_Pack", "Factor_Caja", "Qty_Inners", "Cantidad Esperada"]
    for col in required_cols:
        if col not in df_consolidado.columns:
            raise ValueError(f"No se encontró la columna '{col}' en df_consolidado.")

    # Aplicar la lógica
    df_consolidado["Qty_Unidades"] = df_consolidado.apply(
        lambda row: row["Qty_Inners"] * row["Factor_Caja"] if row["Tipo_Pack"] == "Inner"
                    else row["Cantidad Esperada"],
        axis=1
    )
    return df_consolidado
def generar_df_f_expl_unid(df_consolidado: pd.DataFrame) -> pd.DataFrame:
    """
    Genera el DataFrame df_f_expl_unid tomando solo las filas con 'Tipo_Pack' == 'Unidad'
    y ajustando las columnas al orden y nombre requeridos.
    """
    # Filtrar los registros con Tipo_Pack == 'Unidad'
    df_unid = df_consolidado[df_consolidado["Tipo_Pack"] == "Unidad"].copy()

    # Verificar columnas base
    columnas_requeridas = [
        "No Factura",
        "Num Producto",
        "Descripcion",          # o "Descripción", según tu CSV
        "Qty_Unidades",
        "Familia De Producto",
        "Zona"
    ]
    for col in columnas_requeridas:
        if col not in df_unid.columns:
            raise ValueError(f"Falta la columna '{col}' en df_consolidado para generar df_f_expl_unid.")

    # Crear nueva columna Observaciones (vacía)
    df_unid["Observaciones"] = ""

    # Renombrar y reordenar columnas
    df_unid.rename(
        columns={
            "No Factura": "No Factura",
            "Num Producto": "Código Hijo",
            "Descripcion": "Descripción",
            "Qty_Unidades": "Cantidad Und",
            "Familia De Producto": "Familia",
            "Zona": "Piso",
        },
        inplace=True
    )

    # Orden final
    df_unid = df_unid[
        ["No Factura",
         "Código Hijo",
         "Descripción",
         "Cantidad Und",
         "Familia",
         "Observaciones",
         "Piso"]
    ]

    return df_unid

def mostrar_y_descargar_dataframe(df: pd.DataFrame, nombre: str):
    # Mostrar con AgGrid (o st.dataframe)
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_pagination(paginationAutoPageSize=True)
    gb.configure_side_bar()
    grid_options = gb.build()

    AgGrid(
        df,
        gridOptions=grid_options,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.NO_UPDATE,
        theme="blue"
    )

    # Botón para exportar a Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nombre)
    output.seek(0)

    st.download_button(
        label="Descargar en Excel",
        data=output,
        file_name=f"{nombre}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# ------------------------------------------------------------------------------
# 2. Generar DataFrames de Salida (ejemplos)
# ------------------------------------------------------------------------------
def generar_df_f_expl_unid(df_consolidado: pd.DataFrame) -> pd.DataFrame:
    """
    Filtra Tipo_Pack == 'Unidad' y reordena/renombra columnas:
      1) No Factura -> (misma)
      2) Código Hijo -> Num Producto
      3) Descripción -> Descripcion
      4) Cantidad Und -> Qty_Unidades
      5) Familia -> Familia De Producto
      6) Observaciones -> (columna vacía)
      7) Piso -> Zona
    """
    df_unid = df_consolidado[df_consolidado["Tipo_Pack"] == "Unidad"].copy()
    required_cols = [
        "No Factura",
        "Num Producto",
        "Descripcion",
        "Qty_Unidades",
        "Familia De Producto",
        "Zona"
    ]
    for col in required_cols:
        if col not in df_unid.columns:
            raise ValueError(f"Falta la columna '{col}' en df_consolidado.")

    df_unid["Observaciones"] = ""  # Columna vacía
    df_unid.rename(columns={
        "Num Producto": "Código Hijo",
        "Descripcion": "Descripción",
        "Qty_Unidades": "Cantidad Und",
        "Familia De Producto": "Familia",
        "Zona": "Piso"
    }, inplace=True)

    df_unid = df_unid[[
        "No Factura",
        "Código Hijo",
        "Descripción",
        "Cantidad Und",
        "Familia",
        "Observaciones",
        "Piso"
    ]]
    return df_unid

def generar_df_f_recepcion(df_consolidado: pd.DataFrame) -> pd.DataFrame:
    """
    Placeholder: aquí implementa la lógica real para df_f_recepción.
    """
    # Ejemplo: por ahora devolvemos el mismo df_consolidado
    return df_consolidado.copy()

def generar_df_expl_inner(df_consolidado: pd.DataFrame) -> pd.DataFrame:
    """
    Placeholder: aquí implementa la lógica real para df_expl_inner.
    """
    # Ejemplo: por ahora devolvemos el mismo df_consolidado
    return df_consolidado.copy()

###############################################################################
# 5. MENÚ LATERAL (A LA DERECHA) CON ICONOS BLANCOS
###############################################################################
MENU_OPCIONES = {
    "Inicio": "🏠",
    "Realizar Análisis": "🔍",
    "Registro de OC´s": "📝",
    "Consultar BD": "📂",
    "Salir": "🚪"
}

def set_menu_selection():
    """Se asegura de que 'menu_selected' exista en st.session_state."""
    if "menu_selected" not in st.session_state:
        st.session_state["menu_selected"] = "Inicio"

def radio_menu_con_iconos():
    """
    Crea en la barra lateral:
      - Logo en la parte superior
      - Radio con opciones (iconos + texto)
      - Texto 'Developed by: PJLT' al final
    """
    with st.sidebar:
        # Logo superior
        st.image(
            "https://blogger.googleusercontent.com/img/a/AVvXsEgG46LCtcs4m21eiV-0iDqPHZpdfuEEQrJAqwKNY2WPZWdaC1eoAokveaOPXpitT2a_vsKB7zCnxhRfadp0Edz0q5CcfERwYVzrTZSIeeay_o31XrYlqRxocgNau6kWPjAA61uD42zK--pQlZ6wsyIp97mKU53kHZO-yZXjp_wMNv6Coo_CMiitELregplf=w320-h320",
            use_container_width=True
        )

        # Radio con iconos
        set_menu_selection()
        labels = [f"{MENU_OPCIONES[k]} {k}" for k in MENU_OPCIONES.keys()]
        opciones_keys = list(MENU_OPCIONES.keys())
        seleccion_actual = st.session_state["menu_selected"]
        idx_seleccion_actual = opciones_keys.index(seleccion_actual)

        chosen_label = st.radio(
            "Menú Principal",
            labels,
            index=idx_seleccion_actual,
            key="radio_menu_key"
        )

        # Crear espacio para mover el botón abajo, sin afectar el menú ni el logo
        st.sidebar.markdown("<br><br><br><br>", unsafe_allow_html=True)  # Añadir espacio para el botón

        # Estilo del botón más pequeño y en la barra lateral
        auth_url = get_authorization_url()  # Obtener el auth_url para el botón
        st.sidebar.markdown(
            f'<a href="{auth_url}" target="_blank">'
            f'<button style="background-color:#0078d4; color:white; padding:5px 10px; font-size:12px; border-radius:8px; width: 100%;">'
            'Iniciar sesión con OneDrive'
            '</button>'
            '</a>', unsafe_allow_html=True
        )

        # Texto final con tamaño de fuente configurado
        st.sidebar.markdown(
            "<p style='text-align:center; font-size:12px;'>Developed by: PJLT</p>",  # Cambiar el tamaño de la fuente aquí
            unsafe_allow_html=True
        )

    # Determinar la opción elegida
    for k, icono in MENU_OPCIONES.items():
        if chosen_label.startswith(icono):
            st.session_state["menu_selected"] = k
            return k

    return seleccion_actual

###############################################################################
# 6. LECTURA DE token
###############################################################################
def get_access_token():
    """
    Obtiene un token de Microsoft Graph usando Client Credentials,
    leyendo las credenciales desde st.secrets.
    """
    # Lee la sección [ms_graph] definida en secrets.toml
    tenant_id = st.secrets["ms_graph"]["tenant_id"]
    client_id = st.secrets["ms_graph"]["client_id"]
    client_secret = st.secrets["ms_graph"]["client_secret"]

    authority_url = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["https://graph.microsoft.com/.default"]

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority_url
    )
    result = app.acquire_token_for_client(scopes=scopes)

    if not result:
        st.error("No se obtuvo ninguna respuesta al solicitar el token.")
    else:
        # Si 'result' no es None, verifica si tiene 'access_token' o si contiene 'error'
        if 'access_token' in result:
            # Maneja tu lógica con el token
            access_token = result['access_token']
            st.success("Token obtenido correctamente.")
        else:
            # Muestra el error completo o un mensaje amigable
            st.error(f"No se obtuvo el token. Respuesta devuelta: {result}")
    return result["access_token"]

###############################################################################
# 6. PÁGINAS / SECCIONES
###############################################################################
def page_home():
    col1, col2 = st.columns([1,3])
    with col1:
        # Logotipo en la sección de inicio
        st.image("https://www.dinet.com.pe/img/logo-dinet.png", width=120)
    with col2:
        st.markdown("<h2 class='title'>Sistema de Gestión de Abastecimiento - MAUI</h2>", unsafe_allow_html=True)
        st.markdown("<p class='credit'>Developed by: <b>PJLT</b></p>", unsafe_allow_html=True)
        st.markdown("<p class='desc'>Este sistema realiza análisis y registro de datos de abastecimiento.</p>", unsafe_allow_html=True)

    st.markdown("---")
    st.info("Selecciona una opción en la barra lateral para comenzar.")


def page_consultar_bd():
    st.markdown("## Conectar a OneDrive con Microsoft Graph (Protegido con st.secrets)")

    token = get_access_token()
    if not token:
        return  # Error en obtención de token
    
    headers = {"Authorization": f"Bearer {token}"}
    # Ejemplo: listar archivos en la raíz de OneDrive del usuario
    resp = requests.get("https://graph.microsoft.com/v1.0/me/drive/root/children", headers=headers)
    data = resp.json()

    if "value" not in data:
        st.error(f"No se encontraron archivos: {data}")
        return
    
    archivos = data["value"]
    if not archivos:
        st.warning("No hay archivos en la carpeta raíz de OneDrive.")
        return

    nombres = [item["name"] for item in archivos]
    seleccionado = st.selectbox("Seleccionar archivo", nombres)

    if st.button("Cargar archivo"):
        # Buscamos el item
        item = next((i for i in archivos if i["name"] == seleccionado), None)
        if not item:
            st.warning("No se encontró el archivo en la respuesta.")
            return
        
        download_url = item.get("@microsoft.graph.downloadUrl")
        if not download_url:
            st.warning("No hay enlace de descarga.")
            return
        
        r_file = requests.get(download_url)
        r_file.raise_for_status()

        # Asumimos un Excel .xlsx de ejemplo
        try:
            df = pd.read_excel(BytesIO(r_file.content), engine="openpyxl")
            st.success("Archivo leído con éxito. Vista previa:")
            st.dataframe(df.head(20))
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")


def page_realizar_analisis():
    icon = MENU_OPCIONES["Realizar Análisis"]
    st.markdown(f"## {icon} Realizar Análisis")

    set_directories()
    anio = st.number_input("Año para filtrar", min_value=1990, max_value=2100, value=2024, step=1)
    uploaded_file = st.file_uploader("Subir archivo XLSX/XLSB", type=["xlsx","xlsb"])

    if st.button("Procesar Análisis"):
        if not uploaded_file:
            st.error("No se ha subido ningún archivo.")
            return

        filename = uploaded_file.name
        st.info(f"Archivo recibido: {filename}")

        try:
            import io
            content = uploaded_file.read()
            if filename.endswith(".xlsb"):
                df = pd.read_excel(io.BytesIO(content), engine='pyxlsb')
            else:
                df = pd.read_excel(io.BytesIO(content))

            if 'fecha_despacho' in df.columns:
                df['fecha_despacho'] = pd.to_datetime(
                    df['fecha_despacho'],
                    unit='d',
                    origin='1899-12-30',
                    errors='coerce'
                )
                df['Mes'] = df['fecha_despacho'].dt.month
                df['Año'] = df['fecha_despacho'].dt.year

            if 'Año' in df.columns:
                df_anio = df[df['Año'] == anio].copy()
            else:
                df_anio = df.copy()

            if df_anio.empty:
                st.warning(f"No hay datos para el año {anio}.")
                return

            if 'Cant_Unidad' not in df_anio.columns:
                st.error("No se encontró la columna 'Cant_Unidad'.")
                return
            if 'Sub Familia' not in df_anio.columns:
                st.error("No se encontró la columna 'Sub Familia'.")
                return
            if 'Mes' not in df_anio.columns:
                st.error("No se encontró la columna 'Mes'.")
                return

            agrupado = df_anio.groupby(['Mes', 'Sub Familia'], as_index=False)['Cant_Unidad'].sum()
            agrupado['Total_Mes'] = agrupado.groupby('Mes')['Cant_Unidad'].transform('sum')
            agrupado['Porcentaje'] = (agrupado['Cant_Unidad'] / agrupado['Total_Mes'] * 100).round(2)
            agrupado.sort_values(['Mes', 'Sub Familia'], inplace=True)

            output_file = f"porcentaje_subfamilias_{anio}.xlsx"
            full_path = os.path.join('DATA_MAUI_PJLT', output_file)
            agrupado.to_excel(full_path, index=False)

            st.success(f"Proceso finalizado. Archivo guardado en: {full_path}")

            st.markdown("### Resultado (Editar sin re-run)")
            df_table = interactive_table_no_autoupdate(agrupado, key="analisis")

            if st.button("Aplicar Cambios (Análisis)"):
                st.session_state["df_analisis_editado"] = df_table
                st.success("Cambios guardados en session_state. Se recargará la app.")
                st.experimental_rerun()

            if "df_analisis_editado" in st.session_state:
                st.markdown("#### Data en session_state (Análisis Editado):")
                st.dataframe(st.session_state["df_analisis_editado"].head(20))

                if st.button("Exportar a Excel (Análisis Editado)"):
                    out_name = f"{os.path.splitext(output_file)[0]}_editado.xlsx"
                    st.session_state["df_analisis_editado"].to_excel(out_name, index=False)
                    st.success(f"Archivo Excel (editado) guardado localmente: {out_name}")

        except Exception as e:
            st.error(f"Ocurrió un error al procesar el archivo: {e}")


# Función principal para consolidar y procesar los archivos
def page_consolidar_oc():
    icon = "📝"
    st.markdown(f"## {icon} Registro de OC´s")

    # A) Sección para cargar documentos
    st.markdown("### Cargar Documentos")
    uploaded_files = st.file_uploader("Subir uno o más CSV (Consolidado)", 
                                      type=["csv"], 
                                      accept_multiple_files=True)
    curva_articulo_file = st.file_uploader("Cargar Curva Artículo", type=["csv"])
    plantilla_explosion_file = st.file_uploader("Cargar Plantilla Explosión Maui (XLSM)",
                                                type=["xlsm"])
    if plantilla_explosion_file:
        st.session_state["plantilla_explosion_file"] = plantilla_explosion_file

    # B) Variables de entrada
    contenedor = st.text_input("Contenedor:")
    referencia = st.text_input("Referencia:")
    fecha_recepcion = st.date_input("Fecha de Recepción:", datetime.now())

    if uploaded_files and curva_articulo_file:
        # 1. Procesar archivos para df_consolidado
        lista_df = []
        for upf in uploaded_files:
            try:
                df_temp = pd.read_csv(upf)
                lista_df.append(df_temp)
            except Exception as e:
                st.error(f"Error al leer {upf.name}: {e}")

        if lista_df:
            df_consolidado = pd.concat(lista_df, ignore_index=True)
            df_consolidado.insert(0, "Shipment", contenedor)
            df_consolidado.insert(1, "Referencia", referencia)
            df_consolidado.insert(2, "Fecha de Recepción", fecha_recepcion if fecha_recepcion else "")

            st.success("Archivos procesados correctamente (df_consolidado).")

            # Buscar columna 'Descripcion' o 'Descripción'
            desc_cols = [c for c in df_consolidado.columns if c.lower() in ["descripcion", "descripción"]]
            if not desc_cols:
                st.error("No se encontró la columna 'Descripcion' en los CSV.")
                return
            desc_col = desc_cols[0]

            # 2. Aplicar funciones auxiliares iniciales
            df_consolidado["Subfamilias"] = df_consolidado[desc_col].apply(extraer_descripcion)
            df_consolidado["Código Marca"] = df_consolidado.apply(
                lambda row: extraer_codigo_marca(row[desc_col], row["Subfamilias"]),
                axis=1
            )
            df_consolidado["Marca"] = df_consolidado["Código Marca"].apply(calcular_marca)
            df_consolidado["Zona"] = df_consolidado["Marca"].apply(calcular_zona)
            df_consolidado["Tipo_Pack"] = df_consolidado[desc_col].apply(calcular_tipo_pack)

            # 3. Procesar archivo de Curva Artículo
            try:
                df_curva_articulo = pd.read_csv(curva_articulo_file)

                # Crear columna "Factor por Caja" en df_curva_articulo
                df_curva_articulo = calcular_factor_por_caja(df_curva_articulo)

                # Calcular columnas en df_consolidado que dependen del df_curva_articulo
                df_consolidado = calcular_factor_caja(df_consolidado, df_curva_articulo)
                df_consolidado = calcular_qty_inners(df_consolidado)
                df_consolidado = calcular_qty_unidades(df_consolidado)

                st.success("Curva Artículo procesada correctamente.")

                # 4. Crear los dataframes derivados
                df_f_recepcion = generar_df_f_recepcion(df_consolidado)
                df_f_expl_unid = generar_df_f_expl_unid(df_consolidado)
                df_expl_inner  = generar_df_expl_inner(df_consolidado)

                # 5. Guardar en session_state para poder consultarlos
                st.session_state["df_consolidado"]       = df_consolidado
                st.session_state["df_curva_articulo"]    = df_curva_articulo
                st.session_state["df_f_recepción"]       = df_f_recepcion
                st.session_state["df_f_expl_unid"]       = df_f_expl_unid
                st.session_state["df_expl_inner"]        = df_expl_inner

                # 6. Mostrar df_consolidado final
                st.markdown("### Tabla Consolidado OC's (Final)")
                mostrar_y_descargar_dataframe(df_consolidado, "consolidado_oc_final")

                # 7. Sección de "Consultar Formatos generados"
                st.markdown("### Consultar Formatos Generados")
                opciones = [
                    "df_f_expl_unid",
                    "df_curva_articulo",
                    "df_consolidado",
                    "df_f_recepción",
                    "df_expl_inner"
                ]
                seleccion = st.selectbox("Selecciona un DataFrame para visualizar:", opciones)

                if seleccion in st.session_state:
                    df_seleccionado = st.session_state[seleccion]
                    st.info(f"Mostrando: {seleccion}")
                    mostrar_y_descargar_dataframe(df_seleccionado, seleccion)
                else:
                    st.warning("Aún no se ha generado el DataFrame seleccionado.")

                # == BOTÓN PARA EXPORTAR LA PLANTILLA CON LOS 3 DFS (y macros) ==
                if st.button("Exportar Plantilla Explosión"):
                    try:
                        if "plantilla_explosion_file" not in st.session_state:
                            st.error("No se encontró la plantilla XLSM en session_state.")
                            return

                        # Recuperar DataFrames y variables
                        df_f_recep = st.session_state["df_f_recepción"]
                        df_f_unid  = st.session_state["df_f_expl_unid"]
                        df_inner   = st.session_state["df_expl_inner"]
                        # Contenedor, referencia, fecha
                        contenedor_val = df_consolidado["Shipment"].unique()[0]
                        referencia_val = df_consolidado["Referencia"].unique()[0]
                        fecha_val      = df_consolidado["Fecha de Recepción"].unique()[0]

                        # Convertir fecha a string (ajusta formato si deseas)
                        if isinstance(fecha_val, pd.Timestamp):
                            fecha_str = fecha_val.strftime("%Y%m%d")
                        else:
                            fecha_str = str(fecha_val)

                        # Cargar la plantilla con macros
                        plantilla_bytes = st.session_state["plantilla_explosion_file"]
                        in_memory_file = BytesIO(plantilla_bytes.getvalue())
                        wb = load_workbook(in_memory_file, keep_vba=True)

                        # 1) Hoja "df_f_expl_unid"
                        sheet_unid = wb["df_f_expl_unid"]
                        sheet_unid["I2"] = contenedor_val
                        sheet_unid["J2"] = referencia_val
                        sheet_unid["K2"] = str(fecha_val)

                        start_row = 9
                        start_col = 3  # C
                        for i, row_data in df_f_unid.iterrows():
                            for j, value in enumerate(row_data):
                                sheet_unid.cell(row=start_row + i, column=start_col + j, value=value)

                        # 2) Hoja "df_f_recepción"
                        sheet_recep = wb["df_f_recepción"]
                        for i, row_data in df_f_recep.iterrows():
                            for j, value in enumerate(row_data):
                                sheet_recep.cell(row=start_row + i, column=start_col + j, value=value)

                        # 3) Hoja "df_expl_inner"
                        sheet_inner = wb["df_expl_inner"]
                        for i, row_data in df_inner.iterrows():
                            for j, value in enumerate(row_data):
                                sheet_inner.cell(row=start_row + i, column=start_col + j, value=value)

                        # Guardar en memoria
                        out_file = BytesIO()
                        file_name = f"Explosión_Maui_{contenedor_val}_{referencia_val}_{fecha_str}.xlsm"
                        wb.save(out_file)
                        out_file.seek(0)

                        st.download_button(
                            label="Descargar Archivo Explosión",
                            data=out_file,
                            file_name=file_name,
                            mime="application/vnd.ms-excel.sheet.macroEnabled.12"
                        )
                        st.success(f"Plantilla exportada correctamente: {file_name}")

                    except Exception as e:
                        st.error(f"Ocurrió un error al exportar la plantilla: {e}")

            except Exception as e:
                st.error(f"Error al cargar el archivo de Curva Artículo: {e}")
        else:
            st.warning("No se pudo consolidar ningún archivo.")
    else:
        st.warning("Por favor, sube los archivos CSV y el archivo de Curva Artículo para continuar.")

                

###############################################################################
# 7. FUNCIÓN PRINCIPAL (NAVEGACIÓN)
###############################################################################
# Función principal
def main():
    # Si el usuario ya ha iniciado sesión
    auth_code = get_auth_code_from_url()
    if auth_code:
        # Intercambiamos el código por un token de acceso
        result = get_access_token_from_code(auth_code)
        if 'access_token' in result:
            access_token = result['access_token']
            st.success("Autenticación exitosa. Accediendo a OneDrive...")
            
            # Mostrar que está conectado a OneDrive
            email = get_user_email(access_token)
            if email:
                st.write(f"Conectado con OneDrive como: {email}")
            else:
                st.error("No se pudo obtener el correo electrónico del usuario.")
            
            # Listar archivos de OneDrive
            files = list_onedrive_files(access_token)
            if files:
                st.write("Archivos en OneDrive:")
                for file in files:
                    st.write(file["name"])
            else:
                st.error("No se encontraron archivos en OneDrive.")
        else:
            st.error("Error al obtener el token de acceso.")
    else:
        # Si no hay código, mostrar el botón de inicio de sesión
        st.write("Para acceder a los archivos de OneDrive, por favor inicie sesión.")
        login_button()

    # Continuar con la interfaz de la aplicación Streamlit
    opcion = radio_menu_con_iconos()

    with st.container():
        st.markdown("<div class='main-container'>", unsafe_allow_html=True)

        if opcion == "Inicio":
            page_home()
        elif opcion == "Realizar Análisis":
            page_realizar_analisis()
        elif opcion == "Registro de OC´s":
            page_consolidar_oc()
        elif opcion == "Consultar BD":
            page_consultar_bd()
        elif opcion == "Salir":
            icon = MENU_OPCIONES["Sadeflir"]
            st.markdown(f"## {icon} Salir")
            st.warning("Has salido del Sistema de Gestión de Abastecimiento. Cierra la pestaña o selecciona otra opción.")

        st.markdown("</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
