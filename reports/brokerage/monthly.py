
from db import Base
from db import engine

SCHEMA_NAME = 'reports'


class BrokerageMonthly(Base):
    __tablename__ = 'brokerage_monthly'
    __table_args__ = {"schema": SCHEMA_NAME}

