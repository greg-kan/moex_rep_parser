from logger import Logger
import settings as st
from core import *

from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship
from sqlalchemy import Column, Integer, String, DateTime, Numeric, DOUBLE_PRECISION, Date, Boolean
from sqlalchemy.sql import func
from db import Base
from datetime import date, datetime

MAX_EXCEL_ROWS_NUM = 20000
EXCEL_START_COLUMN = 2

SCHEMA_NAME = 'reports'

MONEY_START_STR = '1. Движение денежных средств'
MONEY_STOP_STR = '2.1. Сделки:'

MONEY_TOTAL_BEGIN_STR = 'Остаток денежных средств на начало периода (Рубль):'
MONEY_TOTAL_END_STR = 'Остаток денежных средств на конец периода (Рубль):'
MONEY_TOTAL_COLUMN = 8

# BROKERAGE_MONTHLY_MONEY_MARKER1 = 'Итого по валюте Рубль:'

MONEY_OPERATIONS_TABLE_START_STR1 = 'Рубль'  # + next str Дата
MONEY_OPERATIONS_TABLE_START_STR2 = 'Дата'
MONEY_OPERATIONS_TABLE_STOP_STR = 'Итого по валюте Рубль:'

MONEY_OPERATIONS_TABLE_CREDIT_SUMM_COLUMN = 7
MONEY_OPERATIONS_TABLE_DEBET_SUMM_COLUMN = 8
MONEY_OPERATIONS_TABLE_NDS_SUMM_COLUMN = 9
MONEY_OPERATIONS_TABLE_SALDO_SUMM_COLUMN = 10

MONEY_FEES_TABLE_START_STR1 = 'Рубль'
MONEY_FEES_TABLE_START_STR2 = 'Вид сбора/штрафа'
MONEY_FEES_TABLE_STOP_STR = 'Итого по валюте Рубль:'

MONEY_FEES_TABLE_FEE_SUMM_COLUMN = 6

logger = Logger('brokerage_monthly_money', st.APPLICATION_LOG, write_to_stdout=st.DEBUG_MODE).get()


class BrokerageMonthlyMoneyFee(Base):
    __tablename__ = 'brokerage_monthly_money_fee'
    __table_args__ = {"schema": SCHEMA_NAME}

    id = Column(Integer, primary_key=True)
    money_id = Column(Integer, ForeignKey(f'{SCHEMA_NAME}.brokerage_monthly_money.id'))
    money = relationship("BrokerageMonthlyMoney", backref="fees")

    fee_name = Column(String, nullable=False)
    amount = Column(Numeric(19, 6))
    nds = Column(Numeric(19, 6))
    platform = Column(String)

    inserted = Column(DateTime(), server_default=func.now())
    updated = Column(DateTime(), onupdate=func.now())

    def __init__(self, fee_name: str, amount: float | None, nds: float | None, platform: str | None):

        self.class_name = self.__class__.__name__

        self.fee_name = fee_name
        self.amount = amount
        self.nds = nds
        self.platform = platform

    def __repr__(self):
        return f"<BrokerageMonthlyMoneyFee({self.fee_name}, {self.amount})>"


class BrokerageMonthlyMoneyOperation(Base):
    __tablename__ = 'brokerage_monthly_money_operation'
    __table_args__ = {"schema": SCHEMA_NAME}

    id = Column(Integer, primary_key=True)
    money_id = Column(Integer, ForeignKey(f'{SCHEMA_NAME}.brokerage_monthly_money.id'))
    money = relationship("BrokerageMonthlyMoney", backref="operations")

    oper_date = Column(Date, nullable=False)
    oper_name = Column(String, nullable=False)
    credit = Column(Numeric(19, 6))
    debet = Column(Numeric(19, 6))
    nds = Column(Numeric(19, 6))
    saldo = Column(Numeric(19, 6))
    warranty = Column(Numeric(19, 6))
    deposit_margin = Column(Numeric(19, 6))

    platform = Column(String)
    note = Column(String)
    intermediate_clearing = Column(String)

    inserted = Column(DateTime(), server_default=func.now())
    updated = Column(DateTime(), onupdate=func.now())

    #
    def __init__(self, oper_date: date, oper_name: str, credit: float | None,
                 debet: float | None, nds: float | None, saldo: float | None,
                 warranty: float | None, deposit_margin: float | None,
                 platform: str | None, note: str | None, intermediate_clearing: str | None):

        self.class_name = self.__class__.__name__

        self.oper_date = oper_date
        self.oper_name = oper_name
        self.credit = credit
        self.debet = debet
        self.dns = nds
        self.saldo = saldo
        self.warranty = warranty
        self.deposit_margin = deposit_margin
        self.platform = platform
        self.note = note
        self.intermediate_clearing = intermediate_clearing

    def __repr__(self):
        return f"<BrokerageMonthlyMoneyOperation({self.oper_date}, {self.oper_name})>"


class BrokerageMonthlyMoney(Base):
    __tablename__ = 'brokerage_monthly_money'
    __table_args__ = {"schema": SCHEMA_NAME}

    id = Column(Integer, primary_key=True)
    report_id = Column(Integer, ForeignKey(f'{SCHEMA_NAME}.brokerage_monthly.id'), unique=True)

    chapter_operations = Column(Boolean, nullable=False)
    chapter_fees = Column(Boolean, nullable=False)

    total_sum_begin_rub = Column(Numeric(19, 6), nullable=False)
    total_sum_end_rub = Column(Numeric(19, 6), nullable=False)
    operations_credit_summ = Column(Numeric(19, 6), nullable=False)
    operations_debet_summ = Column(Numeric(19, 6), nullable=False)
    operations_nds_summ = Column(Numeric(19, 6), nullable=False)
    operations_saldo_summ = Column(Numeric(19, 6), nullable=False)

    fee_summ = Column(Numeric(19, 6), nullable=False)

    inserted = Column(DateTime(), server_default=func.now())
    updated = Column(DateTime(), onupdate=func.now())

    def __init__(self, sheet, start_pos: int, stop_pos: int, chapter_operations: bool, chapter_fees: bool):
        self.class_name = self.__class__.__name__
        self.sheet = sheet
        self.start_row: int = start_pos
        self.stop_row: int = stop_pos

        self.chapter_operations: bool = chapter_operations
        self.chapter_fees: bool = chapter_fees
        self.start_column = EXCEL_START_COLUMN

        self.total_sum_begin_rub = 0
        self.total_sum_end_rub = 0
        self.operations_credit_summ = 0
        self.operations_debet_summ = 0
        self.operations_nds_summ = 0
        self.operations_saldo_summ = 0
        self.fee_summ = 0

        self.total_row_begin: int | None = None
        self.total_row_end: int | None = None

        self.oper_table_start_row: int | None = None
        self.oper_table_stop_row: int | None = None
        self.oper_table_total_row: int | None = None

        self.fees_table_start_row: int | None = None
        self.fees_table_stop_row: int | None = None
        self.fees_table_total_row: int | None = None

        self._check_boundaries()
        self._find_total_rows()
        self._extract_total_sums_rub()

        if self.chapter_operations:
            self._find_oper_table_boundaries()
            self._extract_oper_table_summ_values()
            self._load_oper_table()
            self._check_all_oper_table_summs(6)

        if self.chapter_fees:
            self._find_fees_table_boundaries()
            self._extract_fees_table_summ_value()
            self._load_fees_table()
            self._check_all_fees_table_summs(6)

    def __repr__(self):
        return (f"<BrokerageMonthlyMoney({self.total_sum_begin_rub}, "
                f"{self.total_sum_end_rub})>")

    def _check_boundaries(self):
        if self.start_row and self.stop_row:
            logger.info(f"{self.class_name}._check_boundaries(): "
                        f'Money boundaries found: {self.start_row}, {self.stop_row}')
        else:
            logger.error(f"{self.class_name}._check_boundaries(): "
                         f"Money boundaries not found")
            raise Exception('Money boundaries not found')

    def _find_total_rows(self):
        if self.start_row and self.stop_row:
            for i in range(self.start_row, self.stop_row + 1):
                cell = self.sheet.cell(row=i, column=self.start_column)
                if cell.value == MONEY_TOTAL_BEGIN_STR:
                    self.total_row_begin = cell.row
                    break

            if self.total_row_begin:
                for i in range(self.total_row_begin, self.stop_row + 1):
                    cell = self.sheet.cell(row=i, column=self.start_column)
                    if cell.value == MONEY_TOTAL_END_STR:
                        self.total_row_end = cell.row
                        break

            else:
                logger.error(f"{self.class_name}._find_total_rows(): "
                             f"No self.total_row_begin defined")

                raise Exception('No self.total_row_begin defined')

            if not self.total_row_end:
                logger.error(f"{self.class_name}._find_total_rows(): "
                             f"No self.total_row_end defined")

                raise Exception('No self.total_row_end defined')

        else:
            logger.error(f"{self.class_name}._find_total_rows(): "
                         f"No money boundaries defined")
            raise Exception('No money boundaries defined')

        logger.info(f"{self.class_name}._find_total_rows(): "
                    f'Money total rows found: {self.total_row_begin}, {self.total_row_end}')

    def _extract_total_sums_rub(self):
        if self.total_row_begin and self.total_row_end:
            cell_begin = self.sheet.cell(row=self.total_row_begin,
                                         column=MONEY_TOTAL_COLUMN)
            self.total_sum_begin_rub = float(ifnull(cell_begin.value, 0))
            logger.info(f"{self.class_name}._extract_total_sums_rub(): "
                        f"total_sum_begin_rub = {self.total_sum_begin_rub}")

            cell_end = self.sheet.cell(row=self.total_row_end,
                                       column=MONEY_TOTAL_COLUMN)
            self.total_sum_end_rub = float(ifnull(cell_end.value, 0))
            logger.info(f"{self.class_name}._extract_total_sums_rub(): "
                        f"total_sum_end_rub = {self.total_sum_end_rub}")
        else:
            logger.error(f"{self.class_name}._extract_total_sums_rub(): "
                         f"No money total begin or | and end row(s) defined")
            raise Exception('No money total begin or | and end row(s) defined')

    def _find_oper_table_boundaries(self):
        if self.start_row and self.stop_row:
            for i in range(self.start_row, self.stop_row + 1):
                cell = self.sheet.cell(row=i, column=self.start_column)

                if cell.value == MONEY_OPERATIONS_TABLE_START_STR1:
                    cell_next = self.sheet.cell(row=i+1, column=self.start_column)
                    if cell_next.value == MONEY_OPERATIONS_TABLE_START_STR2:
                        self.oper_table_start_row = cell_next.row + 1
                        break

            for i in range(self.oper_table_start_row, self.stop_row + 1):
                cell = self.sheet.cell(row=i, column=self.start_column)

                if cell.value == MONEY_OPERATIONS_TABLE_STOP_STR:
                    self.oper_table_total_row = cell.row
                    self.oper_table_stop_row = cell.row - 1
                    break

        else:
            logger.error(f"{self.class_name}._find_oper_table_boundaries(): "
                         f"No money start or/and stop row(s) defined")
            raise Exception('No money start or/and stop row(s) defined')

        logger.info(f"{self.class_name}._find_oper_table_boundaries(): "
                    f'Money oper table boundaries found: {self.oper_table_start_row}, {self.oper_table_stop_row}')

    def _extract_oper_table_summ_values(self):
        if self.oper_table_total_row:
            cell_credit_summ = self.sheet.cell(row=self.oper_table_total_row,
                                               column=MONEY_OPERATIONS_TABLE_CREDIT_SUMM_COLUMN)
            cell_debet_summ = self.sheet.cell(row=self.oper_table_total_row,
                                              column=MONEY_OPERATIONS_TABLE_DEBET_SUMM_COLUMN)
            cell_nds_summ = self.sheet.cell(row=self.oper_table_total_row,
                                            column=MONEY_OPERATIONS_TABLE_NDS_SUMM_COLUMN)
            cell_saldo_summ = self.sheet.cell(row=self.oper_table_total_row,
                                              column=MONEY_OPERATIONS_TABLE_SALDO_SUMM_COLUMN)

            self.operations_credit_summ = float(ifnull(cell_credit_summ.value, 0))
            self.operations_debet_summ = float(ifnull(cell_debet_summ.value, 0))
            self.operations_nds_summ = float(ifnull(cell_nds_summ.value, 0))
            self.operations_saldo_summ = float(ifnull(cell_saldo_summ.value, 0))

            logger.info(f"{self.class_name}._extract_oper_table_summ_values(): "
                        f'Money oper table summ values found: {self.operations_credit_summ}, '
                        f'{self.operations_debet_summ}, {self.operations_nds_summ}, '
                        f'{self.operations_saldo_summ}')
        else:
            logger.error(f"{self.class_name}._extract_oper_table_summ_values(): "
                         f"No money oper table total row defined")
            raise Exception("No money oper table total row defined")

    def _load_oper_table(self):

        if self.oper_table_start_row and self.oper_table_stop_row:

            for i in range(self.oper_table_start_row, self.oper_table_stop_row + 1):

                cell = self.sheet.cell(row=i, column=3)
                if not cell.value:
                    raise Exception('Operation name must not be None')

                if cell.value == 'Итого:':
                    continue

                oper_name: str = cell.value

                cell = self.sheet.cell(row=i, column=2)
                if not cell.value:
                    raise Exception('Operation date must not be None')
                oper_date: date = datetime.strptime(cell.value, "%d.%m.%y")

                cell = self.sheet.cell(row=i, column=7)
                credit: float | None = tofloat(cell.value)

                cell = self.sheet.cell(row=i, column=8)
                debet: float | None = tofloat(cell.value)

                cell = self.sheet.cell(row=i, column=9)
                nds: float | None = tofloat(cell.value)

                cell = self.sheet.cell(row=i, column=10)
                saldo: float | None = tofloat(cell.value)

                cell = self.sheet.cell(row=i, column=11)
                warranty: float | None = tofloat(cell.value)

                cell = self.sheet.cell(row=i, column=12)
                deposit_margin: float | None = tofloat(cell.value)

                cell = self.sheet.cell(row=i, column=13)
                platform: str | None = cell.value

                cell = self.sheet.cell(row=i, column=15)
                note: str | None = cell.value

                cell = self.sheet.cell(row=i, column=19)
                intermediate_clearing: str | None = cell.value

                operation = BrokerageMonthlyMoneyOperation(
                    oper_date, oper_name, credit, debet, nds, saldo, warranty, deposit_margin,
                    platform, note, intermediate_clearing
                )

                self.operations.append(operation)

            logger.info(f"{self.class_name}._load_oper_table(): "
                        f"{len(self.operations)} Money operations loaded")

        else:
            logger.error(f"{self.class_name}._load_oper_table(): "
                         f"No oper table start or/and stop row(s) defined")
            raise Exception('No oper table start or/and stop row(s) defined')

        logger.info(f"{self.class_name}._load_oper_table(): "
                    f'Operations table successfully loaded')

    def _check_all_oper_table_summs(self, precision):

        _table_summ_credit: float = 0
        _table_summ_debet: float = 0

        for oper in self.operations:
            if oper.credit:
                _table_summ_credit += oper.credit

            if oper.debet:
                _table_summ_debet += oper.debet

        if round(_table_summ_credit, precision) == round(self.operations_credit_summ, precision) and \
           round(_table_summ_debet, precision) == round(self.operations_debet_summ, precision):
            logger.info(f"{self.class_name}._check_all_oper_table_summs(): "
                        f"Summs in total row correspond summs of all rows")
        else:
            logger.error(f"{self.class_name}._check_all_oper_table_summs(): "
                         f"Summs in total row do not mach summs of all rows")
            raise Exception("Summs in total row do not mach summs of all rows")

        _subtotal = round(self.total_sum_begin_rub + self.operations_credit_summ -
                          self.operations_debet_summ, precision)
        if (self.total_sum_end_rub == self.operations_saldo_summ) and \
           (_subtotal == self.operations_saldo_summ):
            logger.info(f"{self.class_name}._check_all_oper_table_summs(): "
                        f"The portfolio balance has converged")

        else:
            logger.error(f"{self.class_name}._check_all_oper_table_summs(): "
                         f"The portfolio balance has not converged")
            raise Exception('The portfolio balance has not converged')

    def _find_fees_table_boundaries(self):

        if self.oper_table_stop_row and self.stop_row:

            for i in range(self.oper_table_stop_row, self.stop_row + 1):
                cell = self.sheet.cell(row=i, column=self.start_column)

                if cell.value == MONEY_FEES_TABLE_START_STR1:
                    cell_next = self.sheet.cell(row=i+1, column=self.start_column)
                    if cell_next.value == MONEY_FEES_TABLE_START_STR2:
                        self.fees_table_start_row = cell_next.row + 1
                        break

            for i in range(self.fees_table_start_row, self.stop_row + 1):
                cell = self.sheet.cell(row=i, column=self.start_column)

                if cell.value == MONEY_FEES_TABLE_STOP_STR:
                    self.fees_table_total_row = cell.row
                    self.fees_table_stop_row = cell.row - 1
                    break

        else:
            logger.error(f"{self.class_name}._find_fees_table_boundaries(): "
                         f"No money start or/and stop row(s) defined")
            raise Exception('No money start or/and stop row(s) defined')

        logger.info(f"{self.class_name}._find_fees_table_boundaries(): "
                    f'Money fees table boundaries found: {self.fees_table_start_row}, '
                    f'{self.fees_table_stop_row}')

    def _extract_fees_table_summ_value(self):
        if self.fees_table_total_row:
            cell_fee_summ = self.sheet.cell(row=self.fees_table_total_row,
                                            column=MONEY_FEES_TABLE_FEE_SUMM_COLUMN)

            self.fee_summ = float(ifnull(cell_fee_summ.value, 0))

            logger.info(f"{self.class_name}._extract_fees_table_summ_value(): "
                        f"Money fees table summ value found: {self.fee_summ}")
        else:
            logger.error(f"{self.class_name}._extract_fees_table_summ_value(): "
                         f"No money fees table total row defined")
            raise Exception("No money fees table total row defined")

    def _load_fees_table(self):
        if self.fees_table_start_row and self.fees_table_stop_row:

            for i in range(self.fees_table_start_row, self.fees_table_stop_row + 1):

                cell = self.sheet.cell(row=i, column=2)
                if not cell.value:
                    raise Exception('Operation name must not be None')

                fee_name: str = cell.value

                cell = self.sheet.cell(row=i, column=6)
                amount: float | None = tofloat(cell.value)

                cell = self.sheet.cell(row=i, column=8)
                nds: float | None = tofloat(cell.value)

                cell = self.sheet.cell(row=i, column=9)
                platform: str = cell.value

                fee = BrokerageMonthlyMoneyFee(fee_name, amount, nds, platform)

                self.fees.append(fee)

            logger.info(f"{self.class_name}.load_fees_table(): "
                        f"{len(self.fees)} Money fees loaded")

        else:
            logger.error(f"{self.class_name}._load_fees_table(): "
                         f"No fees table start or/and stop row(s) defined")
            raise Exception('No fees table start or/and stop row(s) defined')

        logger.info(f"{self.class_name}._load_fees_table(): "
                    f'Fees table successfully loaded')

    def _check_all_fees_table_summs(self, precision):
        _table_summ_amount: float = 0

        for fee in self.fees:
            if fee.amount:
                _table_summ_amount += fee.amount

        if round(_table_summ_amount, precision) == round(self.fee_summ, precision):
            logger.info(f"{self.class_name}._check_all_fees_table_summs(): "
                        f"Summs in total row correspond summs of all rows")
        else:
            logger.error(f"{self.class_name}._check_all_fees_table_summs(): "
                         f"Summs in total row do not mach summs of all rows")
            raise Exception("Summs in total row do not mach summs of all rows")
