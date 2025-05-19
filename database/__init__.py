from sqlmodel import Session
from .models import OCMaui  # tu modelo
from typing import Any
import pandas as pd
from typing import List, Optional
from sqlmodel import Field, Relationship, SQLModel, create_engine, Session, select
from datetime import datetime
import pandas as pd
from configparser import ConfigParser
from .models import OCMaui
from .db_core import get_session

# --------------------------------------------------------------------------------
# 1) Definir la URL de conexión a tu base de datos de Aiven
# --------------------------------------------------------------------------------
# NOTA: Puedes poner la URL directamente o utilizar st.secrets, variables de entorno, etc.
# Ejemplo de URL en texto plano (reemplaza con la tuya):
AIVEN_URL = "postgresql://avnadmin:AVNS_A9dQ9mjpat6wIhkZbrN@appwebsga-paul911000-1cfc.g.aivencloud.com:26193/defaultdb?sslmode=require"
# Creamos el engine con SQLModel
engine = create_engine(AIVEN_URL, echo=True)

def guardar_contenedor_bd(df_renamed: pd.DataFrame):
    """
    Inserta cada fila de df_renamed en la tabla oc_maui.
    Se asume que df_renamed ya tiene las columnas EXACTAS
    que coinciden con los campos de OCMaui.
    """
    from .db_core import get_session  # suponiendo que ahí tienes la función get_session

    # Convertir cada fila a dict
    rows = df_renamed.to_dict(orient="records")

    with get_session() as session:
        for row_data in rows:
            # Creamos instancia del modelo
            record = OCMaui(**row_data)
            session.add(record)
        session.commit()

def get_session():
    return Session(engine)
# --------------------------------------------------------------------------------
# 3) Función para crear la tabla (si no existe) 
# --------------------------------------------------------------------------------
def create_db_and_tables():
    SQLModel.metadata.create_all(engine)


# --------------------------------------------------------------------------------
# 4) Función helper para obtener la sesión
# --------------------------------------------------------------------------------
def get_session():
    return Session(engine)
