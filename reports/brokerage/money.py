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

MONEY_MAIN_TABLE_START_STR1 = 'Рубль'  # + next str Дата
MONEY_MAIN_TABLE_START_STR2 = 'Дата'
MONEY_MAIN_TABLE_STOP_STR = 'Итого по валюте Рубль:'

MONEY_MAIN_TABLE_CREDIT_SUMM_COLUMN = 7
MONEY_MAIN_TABLE_DEBET_SUMM_COLUMN = 8
MONEY_MAIN_TABLE_NDS_SUMM_COLUMN = 9
MONEY_MAIN_TABLE_SALDO_SUMM_COLUMN = 10

MONEY_FEES_FINES_TABLE_START_STR1 = 'Рубль'
MONEY_FEES_FINES_TABLE_START_STR2 = 'Вид сбора/штрафа'
MONEY_FEES_FINES_TABLE_STOP_STR = 'Итого по валюте Рубль:'

MONEY_FEES_FINES_TABLE_FEE_SUMM_COLUMN = 6

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

        self.total_row_begin: int | None = None
        self.total_row_end: int | None = None

        self.total_sum_begin_rub: float = 0
        self.total_sum_end_rub: float = 0

        self.main_table_start_row: int | None = None
        self.main_table_stop_row: int | None = None
        self.main_table_total_row: int | None = None

        self.main_table_credit_summ: float = 0
        self.main_table_debet_summ: float = 0
        self.main_table_nds_summ: float = 0
        self.main_table_saldo_summ: float = 0

        self.fees_fines_table_start_row: int | None = None
        self.fees_fines_table_stop_row: int | None = None
        self.fees_fines_table_total_row: int | None = None

        self.fees_fines_table_fee_summ: float = 0

        self._find_boundaries()
        self._find_total_rows()
        self._extract_total_sums_rub()
        self._find_main_table_boundaries()
        self._extract_main_table_summ_values()
        self._find_fees_fines_table_boundaries()
        self._extract_fees_fines_table_summ_value()

    def __repr__(self):
        return (f"<BrokerageMonthlyMoney({self.total_sum_begin_rub}, "
                f"{self.total_sum_end_rub})>")

    def _find_boundaries(self):
        for i in range(1, MAX_EXCEL_ROWS_NUM):
            cell = self.sheet.cell(row=i, column=self.start_column)
            if cell.value == MONEY_START_STR:
                self.start_row = cell.row
                break

        if self.start_row:
            for i in range(self.start_row, MAX_EXCEL_ROWS_NUM):
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

    def _find_total_rows(self):
        if self.start_row and self.stop_row:
            for i in range(self.start_row, self.stop_row):
                cell = self.sheet.cell(row=i, column=self.start_column)
                if cell.value == MONEY_TOTAL_BEGIN_STR:
                    self.total_row_begin = cell.row
                    break

            if self.total_row_begin:
                for i in range(self.total_row_begin, self.stop_row):
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

    def _find_main_table_boundaries(self):
        if self.start_row and self.stop_row:
            for i in range(self.start_row, self.stop_row):
                cell = self.sheet.cell(row=i, column=self.start_column)

                if cell.value == MONEY_MAIN_TABLE_START_STR1:
                    cell_next = self.sheet.cell(row=i+1, column=self.start_column)
                    if cell_next.value == MONEY_MAIN_TABLE_START_STR2:
                        self.main_table_start_row = cell_next.row + 1
                        break

            for i in range(self.main_table_start_row, self.stop_row):
                cell = self.sheet.cell(row=i, column=self.start_column)

                if cell.value == MONEY_MAIN_TABLE_STOP_STR:
                    self.main_table_total_row = cell.row
                    self.main_table_stop_row = cell.row - 1
                    break

        else:
            logger.error(f"{self.class_name}._find_main_table_boundaries(): "
                         f"No money start or/and stop row(s) defined")
            raise Exception('No money start or/and stop row(s) defined')

        logger.info(f"{self.class_name}._find_main_table_boundaries(): "
                    f'Money main table boundaries found: {self.main_table_start_row}, {self.main_table_stop_row}')

    def _extract_main_table_summ_values(self):
        if self.main_table_total_row:
            cell_credit_summ = self.sheet.cell(row=self.main_table_total_row,
                                               column=MONEY_MAIN_TABLE_CREDIT_SUMM_COLUMN)
            cell_debet_summ = self.sheet.cell(row=self.main_table_total_row,
                                              column=MONEY_MAIN_TABLE_DEBET_SUMM_COLUMN)
            cell_nds_summ = self.sheet.cell(row=self.main_table_total_row,
                                            column=MONEY_MAIN_TABLE_NDS_SUMM_COLUMN)
            cell_saldo_summ = self.sheet.cell(row=self.main_table_total_row,
                                              column=MONEY_MAIN_TABLE_SALDO_SUMM_COLUMN)

            self.main_table_credit_summ = float(ifnull(cell_credit_summ.value, 0))
            self.main_table_debet_summ = float(ifnull(cell_debet_summ.value, 0))
            self.main_table_nds_summ = float(ifnull(cell_nds_summ.value, 0))
            self.main_table_saldo_summ = float(ifnull(cell_saldo_summ.value, 0))

            logger.info(f"{self.class_name}._extract_main_table_summ_values(): "
                        f'Money main table summ values found: {self.main_table_credit_summ}, '
                        f'{self.main_table_debet_summ}, {self.main_table_nds_summ}, '
                        f'{self.main_table_saldo_summ}')
        else:
            logger.error(f"{self.class_name}._extract_main_table_summ_values(): "
                         f"No money main table total row defined")
            raise Exception("No money main table total row defined")

    def _find_fees_fines_table_boundaries(self):

        if self.main_table_stop_row and self.stop_row:

            for i in range(self.main_table_stop_row, self.stop_row):
                cell = self.sheet.cell(row=i, column=self.start_column)

                if cell.value == MONEY_FEES_FINES_TABLE_START_STR1:
                    cell_next = self.sheet.cell(row=i+1, column=self.start_column)
                    if cell_next.value == MONEY_FEES_FINES_TABLE_START_STR2:
                        self.fees_fines_table_start_row = cell_next.row + 1
                        break

            for i in range(self.fees_fines_table_start_row, self.stop_row):
                cell = self.sheet.cell(row=i, column=self.start_column)

                if cell.value == MONEY_FEES_FINES_TABLE_STOP_STR:
                    self.fees_fines_table_total_row = cell.row
                    self.fees_fines_table_stop_row = cell.row - 1
                    break

        else:
            logger.error(f"{self.class_name}._find_fees_fines_table_boundaries(): "
                         f"No money start or/and stop row(s) defined")
            raise Exception('No money start or/and stop row(s) defined')

        logger.info(f"{self.class_name}._find_fees_fines_table_boundaries(): "
                    f'Money fees and fines table boundaries found: {self.fees_fines_table_start_row}, '
                    f'{self.fees_fines_table_stop_row}')

    def _extract_fees_fines_table_summ_value(self):
        if self.fees_fines_table_total_row:
            cell_fee_summ = self.sheet.cell(row=self.fees_fines_table_total_row,
                                            column=MONEY_FEES_FINES_TABLE_FEE_SUMM_COLUMN)

            self.fees_fines_table_fee_summ = float(ifnull(cell_fee_summ.value, 0))

            logger.info(f"{self.class_name}._extract_fees_fines_table_summ_value(): "
                        f"Money fees and fines table summ value found: {self.fees_fines_table_fee_summ}")
        else:
            logger.error(f"{self.class_name}._extract_fees_fines_table_summ_value(): "
                         f"No money fees and fines table total row defined")
            raise Exception("No money fees and fines table total row defined")
