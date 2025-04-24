import streamlit as st 
import pandas as pd
import os
import re
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode
from io import BytesIO
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

###############################################################################
# 3. TABLA INTERACTIVA SIN AUTO-ACTUALIZACIÓN (NO_UPDATE)
###############################################################################
def interactive_table_no_autoupdate(df: pd.DataFrame, key: str=None) -> pd.DataFrame:
    """
    Muestra un DataFrame con st_aggrid usando update_mode=NO_UPDATE:
      - No hay re-run automático al editar o filtrar
      - Se requiere un botón manual para "aplicar" los cambios
    """
    from io import BytesIO

    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(editable=True, filter=True)

    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=20)
    gb.configure_grid_options(
        paginationPageSize=20,
        paginationPageSizeOptions=[20, 50, 100, 200, 500]
    )
    gb.configure_side_bar()

    # Evitar re-run al pulsar Enter
    grid_options = gb.build()
    grid_options["stopEnterEventPropagation"] = True
    grid_options["enterMovesDownAfterEdit"] = True

    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.NO_UPDATE,
        theme="blue",
        key=key
    )
    edited_df = pd.DataFrame(grid_response["data"])

    # Botón para exportar la tabla a Excel
    if st.button("Exportar a Excel"):
        # Convertir el DataFrame a un archivo Excel en memoria
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            edited_df.to_excel(writer, index=False, sheet_name="Datos")
            writer.save()
        output.seek(0)

        # Crear el botón de descarga
        st.download_button(
            label="Descargar archivo Excel",
            data=output,
            file_name=f"{key}_editado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    return edited_df

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

        # Espacio para bajar el texto del final
        st.markdown("<br><br><br><br>", unsafe_allow_html=True)
        # Texto final
        st.markdown(
            "<p style='text-align:center;'>Developed by: PJLT</p>",
            unsafe_allow_html=True
        )

    # Determinar la opción elegida
    for k, icono in MENU_OPCIONES.items():
        if chosen_label.startswith(icono):
            st.session_state["menu_selected"] = k
            return k

    return seleccion_actual

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
    icon = MENU_OPCIONES["Consultar BD"]
    st.markdown(f"## {icon} Consultar BD")

    set_directories()
    folder_path = 'DATA_MAUI_PJLT'

    archivos = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xlsb', '.csv'))]
    if not archivos:
        st.error("No se encontraron archivos en la carpeta 'DATA_MAUI_PJLT'.")
        return

    selected_file = st.selectbox("Seleccionar archivo para cargar:", archivos)
    if selected_file:
        file_path = os.path.join(folder_path, selected_file)
        st.write(f"Archivo seleccionado: {selected_file}")

        if st.button("Cargar archivo"):
            try:
                if selected_file.endswith('.xlsx'):
                    df = pd.read_excel(file_path)
                elif selected_file.endswith('.xlsb'):
                    df = pd.read_excel(file_path, engine='pyxlsb')
                else:
                    df = pd.read_csv(file_path)

                st.success(f"Archivo cargado correctamente: {selected_file}")

                # Mostrar tabla NO_UPDATE
                st.markdown("### Datos (Editar sin re-run)")
                df_table = interactive_table_no_autoupdate(df, key="consulta_bd")

                if st.button("Aplicar Cambios (BD)"):
                    st.session_state["df_consultar_bd"] = df_table
                    st.success("Cambios guardados en session_state. Se recargará la app.")
                    st.experimental_rerun()

                if "df_consultar_bd" in st.session_state:
                    st.markdown("#### Data en session_state (BD Editado):")
                    st.dataframe(st.session_state["df_consultar_bd"].head(20))

                    # Exportar a Excel
                    if st.button("Exportar a Excel (BD Editado)"):
                        out_name = f"{os.path.splitext(selected_file)[0]}_editado.xlsx"
                        st.session_state["df_consultar_bd"].to_excel(out_name, index=False)
                        st.success(f"Archivo Excel guardado localmente: {out_name}")

            except Exception as e:
                st.error(f"Error al cargar el archivo: {e}")

    # Nueva funcionalidad para exportar la tabla consultada
    if "df_consultar_bd" in st.session_state:
        if st.button("Exportar tabla consultada a Excel"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                st.session_state["df_consultar_bd"].to_excel(writer, index=False, sheet_name="Datos")
                writer.save()
            output.seek(0)

            st.download_button(
                label="Descargar tabla consultada como Excel",
                data=output,
                file_name="tabla_consultada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


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


def page_consolidar_oc():
    icon = MENU_OPCIONES["Registro de OC´s"]
    st.markdown(f"## {icon} Registro de OC´s")

    set_directories()

    contenedor = st.text_input("Contenedor:")
    referencia = st.text_input("Referencia:")
    fecha_recepcion = st.date_input("Fecha de Recepción:", datetime.now())
    uploaded_files = st.file_uploader("Subir uno o más CSV", type=["csv"], accept_multiple_files=True)

    if st.button("Procesar"):
        if not contenedor or not referencia:
            st.error("Por favor, ingrese Contenedor y Referencia.")
            return
        if not uploaded_files:
            st.error("No se subió ningún CSV.")
            return

        lista_df = []
        nombres_archivos = []
        for upf in uploaded_files:
            try:
                df_temp = pd.read_csv(upf)
                lista_df.append(df_temp)
                nombres_archivos.append(upf.name)
            except Exception as e:
                st.error(f"Error al leer {upf.name}: {e}")

        if not lista_df:
            st.warning("No se pudo consolidar ningún archivo.")
            return

        df_consolidado = pd.concat(lista_df, ignore_index=True)
        df_consolidado.insert(0, "Shipment", contenedor)
        df_consolidado.insert(1, "Referencia", referencia)
        df_consolidado.insert(2, "Fecha de Recepción", fecha_recepcion if fecha_recepcion else "")

        desc_cols = [c for c in df_consolidado.columns if c.lower() in ["descripcion","descripción"]]
        if not desc_cols:
            st.error("No se encontró la columna 'Descripcion' en los CSV.")
            return
        desc_col = desc_cols[0]

        df_consolidado["Subfamilias"] = df_consolidado[desc_col].apply(extraer_descripcion)

        if fecha_recepcion:
            df_consolidado["Mes de Recepción"] = pd.to_datetime(fecha_recepcion).month
        else:
            df_consolidado["Mes de Recepción"] = None

        df_consolidado["Código Marca"] = df_consolidado.apply(
            lambda row: extraer_codigo_marca(row[desc_col], row["Subfamilias"]),
            axis=1
        )
        df_consolidado["Marca"] = df_consolidado["Código Marca"].apply(calcular_marca)
        df_consolidado["Zona"] = df_consolidado["Marca"].apply(calcular_zona)

        cols_final = ["Mes de Recepción", "Subfamilias", "Código Marca", "Marca", "Zona"]
        cols_principales = [c for c in df_consolidado.columns if c not in cols_final]
        df_consolidado = df_consolidado[cols_principales + cols_final]

        output_csv = "ordenes_compra_consolidado.csv"
        full_path_csv = os.path.join('DATA_MAUI_PJLT', output_csv)
        df_consolidado.to_csv(full_path_csv, index=False)

        st.success(f"Consolidación completada. CSV guardado en: {full_path_csv}")

        st.write("#### Archivos cargados:")
        for n in nombres_archivos:
            st.write(f"- {n}")
        st.write("---")

        st.markdown("#### Vista previa (Editar sin re-run)")
        df_table = interactive_table_no_autoupdate(df_consolidado, key="consolidar_oc")

        if st.button("Aplicar Cambios (OC)"):
            st.session_state["df_consolidado_editado"] = df_table
            st.success("Cambios guardados en session_state. Se recargará la app.")
            st.experimental_rerun()

        if "df_consolidado_editado" in st.session_state:
            st.markdown("##### Data en session_state (OC Editado):")
            st.dataframe(st.session_state["df_consolidado_editado"].head(20))

            if st.button("Exportar a Excel (OC Editado)"):
                output_excel = "ordenes_compra_consolidado_editado.xlsx"
                full_path_xlsx = os.path.join('DATA_MAUI_PJLT', output_excel)
                st.session_state["df_consolidado_editado"].to_excel(full_path_xlsx, index=False)
                st.success(f"Archivo Excel consolidado guardado en: {full_path_xlsx}")


###############################################################################
# 7. FUNCIÓN PRINCIPAL (NAVEGACIÓN)
###############################################################################
def main():
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
