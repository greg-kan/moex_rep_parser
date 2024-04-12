from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import create_engine
import settings as st
from logger import Logger

SCHEMA_NAME = 'reports'

Base = declarative_base()

params = st.DB_PARAMS
engine = create_engine(
    f"postgresql+psycopg2://{params['user']}:{params['password']}@{params['host']}:5432/{params['database']}")

Base.metadata.create_all(engine)
