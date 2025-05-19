from typing import List, Optional
from sqlmodel import Field, Relationship, SQLModel, create_engine, Session, select
from datetime import datetime
import pandas as pd
from configparser import ConfigParser
from .models import OCMaui
from .db_core import get_session
# database.py






# --------------------------------------------------------------------------------
# 2) Definir el Modelo (tabla) con los campos deseados
# --------------------------------------------------------------------------------
# De acuerdo con tus encabezados en df_consolidado y cómo se llamarán en la BD:
class OCMaui(SQLModel, table=True):
    """
    Representa la tabla "oc_maui" (puedes renombrarla en metadata si gustas).
    """
    # Primera columna (PK):
    id: Optional[int] = Field(default=None, primary_key=True)

    # El resto de columnas, con sus nombres ya en minúsculas, sin espacios,
    # y en el orden que hayas definido (pero el orden no es tan crucial en BD):
    shipment: str
    referencia: str
    fecha_recepcion: str
    cliente: Optional[str] = None
    proveedor: Optional[str] = None
    direccion: Optional[str] = None
    nro_factura: Optional[str] = None
    fecha_limite: Optional[str] = None
    fecha_factura: Optional[str] = None
    familia_producto: Optional[str] = None
    num_producto: Optional[str] = None
    descripcion: Optional[str] = None
    producto_nuevo: Optional[str] = None
    huella: Optional[str] = None
    huella_default: Optional[str] = None
    recibo_habilitado: Optional[str] = None
    cantidad_esperada: Optional[float] = None
    identificada: Optional[str] = None
    cant_cajas: Optional[float] = None
    saldos_un: Optional[float] = None
    vol_m3: Optional[float] = None
    articulo_padre: Optional[str] = None
    recibida: Optional[str] = None
    subfamilia: Optional[str] = None
    codigo_marca: Optional[str] = None
    marca: Optional[str] = None
    zona: Optional[str] = None
    tipo_pack: Optional[str] = None
    factor_caja: Optional[float] = None
    qty_inner: Optional[float] = None
    qty_unidades: Optional[float] = None





