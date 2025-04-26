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

redirect_uri = "http://localhost:8501"

# Función para generar el enlace de autorización con la redirección correcta
def get_authorization_url():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    auth_url = app.get_authorization_request_url(SCOPES, redirect_uri=redirect_uri)
    return auth_url
# Función para intercambiar el código de autorización por un token de acceso
def get_access_token_from_code(auth_code):
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    result = app.acquire_token_by_authorization_code(auth_code, scopes=SCOPES, redirect_uri=redirect_uri)

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
        update_mode=GridUpdateMode.NO_UPDATE,
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
    if "access_token" not in result:
        st.error(f"No se pudo obtener el token: {result.get('error_description')}")
        return None
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

    # Cargar archivo de entrada
    uploaded_files = st.file_uploader("Subir uno o más CSV", type=["csv"], accept_multiple_files=True)
    curva_articulo_file = st.file_uploader("Cargar Curva Artículo", type=["csv"])

    # Variables de entrada
    contenedor = st.text_input("Contenedor:")
    referencia = st.text_input("Referencia:")
    fecha_recepcion = st.date_input("Fecha de Recepción:", datetime.now())

    if uploaded_files and curva_articulo_file:
        # Procesar archivos
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

            # Guardar los datos cargados en session_state para persistencia
            st.session_state["df_consolidado"] = df_consolidado
            st.success("Archivos procesados correctamente.")

            # Procesar las columnas utilizando las funciones auxiliares
            desc_cols = [c for c in df_consolidado.columns if c.lower() in ["descripcion","descripción"]]
            if not desc_cols:
                st.error("No se encontró la columna 'Descripcion' en los CSV.")
                return
            desc_col = desc_cols[0]

            # Aplicar las funciones auxiliares
            df_consolidado["Subfamilias"] = df_consolidado[desc_col].apply(extraer_descripcion)
            df_consolidado["Código Marca"] = df_consolidado.apply(
                lambda row: extraer_codigo_marca(row[desc_col], row["Subfamilias"]),
                axis=1
            )
            df_consolidado["Marca"] = df_consolidado["Código Marca"].apply(calcular_marca)
            df_consolidado["Zona"] = df_consolidado["Marca"].apply(calcular_zona)

            # Guardar en session_state después de aplicar las funciones
            st.session_state["df_consolidado"] = df_consolidado

            # Mostrar tabla cargada
            st.markdown("### Datos Consolidados")
            interactive_table_no_autoupdate(df_consolidado, key="consolidado_oc")

            # Cargar y mostrar archivo de Curva Artículo
            try:
                df_curva_articulo = pd.read_csv(curva_articulo_file)
                st.success("Archivo de Curva Artículo cargado correctamente.")
                st.markdown("### Tabla de Curva Artículo")
                interactive_table_no_autoupdate(df_curva_articulo, key="curva_articulo")
            except Exception as e:
                st.error(f"Error al cargar el archivo de Curva Artículo: {e}")
        else:
            st.warning("No se pudo consolidar ningún archivo.")
    else:
        st.warning("Por favor, sube los archivos CSV.")
                

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
            icon = MENU_OPCIONES["Salir"]
            st.markdown(f"## {icon} Salir")
            st.warning("Has salido del Sistema de Gestión de Abastecimiento. Cierra la pestaña o selecciona otra opción.")

        st.markdown("</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
