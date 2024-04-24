from logger import Logger
import settings as st
from core import *

from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship
from sqlalchemy import Column, Integer, String, DateTime, Numeric, DOUBLE_PRECISION, Date, Boolean
from sqlalchemy.sql import func
from db import Base
from datetime import datetime

MAX_EXCEL_ROWS_NUM = 20000
EXCEL_START_COLUMN = 2

TRANSACTION_TABLE_START_STR = 'ЦБ'

SCHEMA_NAME = 'reports'

logger = Logger('brokerage_monthly_transactions', st.APPLICATION_LOG, write_to_stdout=st.DEBUG_MODE).get()


class BrokerageMonthlyTransaction(Base):

    __tablename__ = 'brokerage_monthly_transaction'
    __table_args__ = {"schema": SCHEMA_NAME}

    id = Column(Integer, primary_key=True)
    transactions_id = Column(Integer, ForeignKey(f'{SCHEMA_NAME}.brokerage_monthly_transactions.id'))
    transactions = relationship("BrokerageMonthlyTransactions", backref="transactions")

    security = Column(String, nullable=False)
    trans_date = Column(Date, nullable=False)

    balance = Column(DOUBLE_PRECISION)
    credit = Column(DOUBLE_PRECISION)
    debet = Column(DOUBLE_PRECISION)
    saldo = Column(DOUBLE_PRECISION)

    storage = Column(String)
    note = Column(String)

    inserted = Column(DateTime(), server_default=func.now())
    updated = Column(DateTime(), onupdate=func.now())

    def __init__(self, security: str, trans_date: datetime,
                 balance: float | None, credit: float | None, debet: float | None, saldo: float | None,
                 storage: str | None, note: str | None):
        self.class_name = self.__class__.__name__

        self.security = security
        self.trans_date = trans_date
        self.balance = balance
        self.credit = credit
        self.debet = debet
        self.saldo = saldo
        self.storage = storage
        self.note = note

    def __repr__(self):
        return f"<BrokerageMonthlyTransaction({self.security}, {self.trans_date})>"


class BrokerageMonthlyTransactions(Base):

    __tablename__ = 'brokerage_monthly_transactions'
    __table_args__ = {"schema": SCHEMA_NAME}

    id = Column(Integer, primary_key=True)
    report_id = Column(Integer, ForeignKey(f'{SCHEMA_NAME}.brokerage_monthly.id'), unique=True)

    def __init__(self, sheet, start_pos: int, stop_pos: int):
        self.class_name = self.__class__.__name__
        self.sheet = sheet
        self.start_row: int = start_pos
        self.stop_row: int = stop_pos
        self.start_column = EXCEL_START_COLUMN

        self.table_start_row: int | None = None
        self.table_stop_row: int | None = None

        self._find_table_boundaries()
        self._load_table()
        self._make_calculations()

    def __repr__(self):
        return f"<BrokerageMonthlyTransactions()>"  # TODO: {self.secid}

    def _find_table_boundaries(self):
        if self.start_row and self.stop_row:
            for i in range(self.start_row, self.stop_row + 1):
                cell = self.sheet.cell(row=i, column=self.start_column)

                if cell.value == TRANSACTION_TABLE_START_STR:
                    self.table_start_row = cell.row + 1
                    break

            self.table_stop_row = self.stop_row

        else:
            logger.error(f"{self.class_name}._find_table_boundaries(): "
                         f"No transaction start or/and stop row(s) defined")
            raise Exception('No transaction start or/and stop row(s) defined')

        logger.info(f"{self.class_name}._find_table_boundaries(): "
                    f'Transaction table boundaries found: {self.table_start_row}, {self.table_stop_row}')

    def _load_table(self):
        if self.table_start_row and self.table_stop_row:

            for i in range(self.table_start_row, self.table_stop_row + 1):

                cell = self.sheet.cell(row=i, column=2)

                if not cell.value:
                    continue

                security: str = cell.value

                cell = self.sheet.cell(row=i, column=5)
                trans_date: datetime = datetime.strptime(cell.value, "%d.%m.%y")

                cell = self.sheet.cell(row=i, column=7)
                balance: float | None = tofloat(cell.value)

                cell = self.sheet.cell(row=i, column=8)
                credit: float | None = tofloat(cell.value)

                cell = self.sheet.cell(row=i, column=9)
                debet: float | None = tofloat(cell.value)

                cell = self.sheet.cell(row=i, column=10)
                saldo: float | None = tofloat(cell.value)

                cell = self.sheet.cell(row=i, column=11)
                storage: str | None = cell.value

                cell = self.sheet.cell(row=i, column=13)
                note: str | None = cell.value

                transaction = BrokerageMonthlyTransaction(security, trans_date,
                                                          balance, credit, debet, saldo, storage, note)

                self.transactions.append(transaction)

            logger.info(f"{self.class_name}._load_table(): "
                        f"{len(self.transactions)} transactions loaded")

        else:
            logger.error(f"{self.class_name}._load_table(): "
                         f"No transactions table start or/and stop row(s) defined")
            raise Exception('No transactions table start or/and stop row(s) defined')

        logger.info(f"{self.class_name}._load_table(): "
                    f'Transactions table successfully loaded')

    def _make_calculations(self):
        # TODO: any aggregates may be in a future...
        pass
