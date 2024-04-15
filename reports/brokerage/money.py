from logger import Logger
import settings as st
from core import *

from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship
from sqlalchemy import Column, Integer, String, DateTime, Numeric, DOUBLE_PRECISION
from sqlalchemy.sql import func
from db import Base

MAX_EXCEL_ROWS_NUM = 20000
EXCEL_START_COLUMN = 2

SCHEMA_NAME = 'reports'

MONEY_START_STR = '1. Движение денежных средств'
MONEY_STOP_STR = '2.1. Сделки:'

MONEY_TOTAL_BEGIN_STR = 'Остаток денежных средств на начало периода (Рубль):'
MONEY_TOTAL_END_STR = 'Остаток денежных средств на конец периода (Рубль):'
MONEY_TOTAL_COLUMN = 8

MONEY_TABLE_START_STR1 = 'Рубль'  # + next str Дата
MONEY_TABLE_START_STR2 = 'Дата'
MONEY_TABLE_STOP_STR = 'Итого по валюте Рубль:'

# MONEY_TABLE_BEGIN_SUMM_NKD_COLUMN = 9
# MONEY_TABLE_BEGIN_SUMM_INCLUDING_NKD_COLUMN = 10
# MONEY_TABLE_END_SUMM_NKD_COLUMN = 13
# MONEY_TABLE_END_SUMM_INCLUDING_NKD_COLUMN = 14

MONEY_FEES_FINES_TABLE_START_STR = 'Вид сбора/штрафа'
MONEY_FEES_FINES_TABLE_STOP_STR = 'Итого по валюте Рубль:'


logger = Logger('brokerage_monthly_money', st.APPLICATION_LOG, write_to_stdout=st.DEBUG_MODE).get()


class BrokerageMonthlyMoneyDetail:  # (Base)
    __tablename__ = 'brokerage_monthly_money_detail'
    __table_args__ = {"schema": SCHEMA_NAME}

    # id = Column(Integer, primary_key=True)
    # securities_id = Column(Integer, ForeignKey(f'{SCHEMA_NAME}.brokerage_monthly_securities.id'))
    # securities = relationship("BrokerageMonthlySecurities", backref="securities")
    #
    # secid = Column(String, nullable=False)


class BrokerageMonthlyMoney:  # (Base)
    __tablename__ = 'brokerage_monthly_money'
    __table_args__ = {"schema": SCHEMA_NAME}

    # id = Column(Integer, primary_key=True)
    # report_id = Column(Integer, ForeignKey(f'{SCHEMA_NAME}.brokerage_monthly.id'), unique=True)
    #
    # portfolio_total_value_begin_rub = Column(Numeric(19, 6))

    def __init__(self, sheet):
        self.class_name = self.__class__.__name__
        self.sheet = sheet
        self.start_column = EXCEL_START_COLUMN
        self.start_row: int | None = None
        self.stop_row: int | None = None

        self._find_boundaries()

    # def __repr__(self):
    #     return (f"<BrokerageMonthlyMoney({self.portfolio_total_value_begin_rub}, "
    #             f"{self.portfolio_total_value_end_rub})>")

    def _find_boundaries(self):
        for i in range(1, MAX_EXCEL_ROWS_NUM+1):
            cell = self.sheet.cell(row=i, column=self.start_column)
            if cell.value == MONEY_START_STR:
                self.start_row = cell.row
                break

        if self.start_row:
            for i in range(self.start_row, MAX_EXCEL_ROWS_NUM+1):
                cell = self.sheet.cell(row=i, column=self.start_column)
                if cell.value == MONEY_STOP_STR:
                    self.stop_row = cell.row - 1
                    break
        else:
            logger.error(f"{self.class_name}._find_boundaries(): "
                         f"Could not find a Money start row index")

            raise Exception('Could not find a Money start row index')

        if not self.stop_row:
            logger.error(f"{self.class_name}._find_boundaries(): "
                         f"Could not find a Money stop row index")

            raise Exception('Could not find a Money stop row index')

        logger.info(f"{self.class_name}._find_boundaries(): "
                    f' boundaries found: {self.start_row}, {self.stop_row}')
