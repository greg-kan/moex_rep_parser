import ntpath

from openpyxl import load_workbook
from logger import Logger
import settings as st
from core import *

from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship
from sqlalchemy import Column, Integer, String, DateTime
from sqlalchemy.sql import func
from db import Base

from reports.brokerage.securities import BrokerageMonthlySecurities
from reports.brokerage.money import BrokerageMonthlyMoney

SCHEMA_NAME = 'reports'
STOP_REPORT_STR = 'Владелец: ООО "Компания БКС"'

logger = Logger('brokerage_monthly', st.APPLICATION_LOG, write_to_stdout=st.DEBUG_MODE).get()


class BrokerageMonthly(Base):
    __tablename__ = 'brokerage_monthly'
    __table_args__ = {"schema": SCHEMA_NAME}

    id = Column(Integer, primary_key=True)
    money = relationship("BrokerageMonthlyMoney", uselist=False, backref="report")
    # deals
    securities = relationship("BrokerageMonthlySecurities", uselist=False, backref="report")
    # securities_transactions

    report_path = Column(String, nullable=False)
    year = Column(Integer, nullable=False)
    month = Column(Integer, nullable=False)
    stop_report_str = Column(String, nullable=False)

    inserted = Column(DateTime(), server_default=func.now())  # timezone=True
    updated = Column(DateTime(), onupdate=func.now())  # timezone=True

    def __init__(self, report_path):
        self.year: int = 0
        self.month: int = 0
        self.stop_report_str = STOP_REPORT_STR
        self.class_name = self.__class__.__name__
        self.report_path: str = str(report_path)
        self.workbook = load_workbook(self.report_path)
        self.sheet = self.workbook[self.workbook.sheetnames[0]]

        self._extract_year_and_month()

        self.money = BrokerageMonthlyMoney(self.sheet)
        # self.deals = Deals(self.sheet)
        self.securities = BrokerageMonthlySecurities(self.sheet)
        # self.securities_transactions = SecuritiesTransactions(self.sheet)

    def __repr__(self):
        return f"<BrokerageMonthly({self.year}, {self.month})>"

    def _extract_year_and_month(self):
        bare_file_name = get_file_name(ntpath.basename(self.report_path))
        str_year_month = bare_file_name.split('_ALL_')[-1]
        self.year = 2000 + int(str_year_month.split('-')[0])
        self.month = int(str_year_month.split('-')[-1])
        logger.info(f"Year {self.year} and Month {self.month} were extracted")
