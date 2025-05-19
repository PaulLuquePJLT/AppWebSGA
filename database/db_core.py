# db_core.py
from sqlmodel import create_engine, Session

DATABASE_URL = (
    "postgresql://avnadmin:AVNS_A9dQ9mjpat6wIhkZbrN"
    "@appwebsga-paul911000-1cfc.g.aivencloud.com:26193/defaultdb"
    "?sslmode=require"
)
engine = create_engine(DATABASE_URL, echo=True)


def get_session():
    return Session(engine)
