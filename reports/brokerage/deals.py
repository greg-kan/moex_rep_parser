from logger import Logger
import settings as st
from core import *

from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship
from sqlalchemy import Column, Integer, String, DateTime, Numeric, DOUBLE_PRECISION, Boolean
from sqlalchemy.sql import func
from db import Base

MAX_EXCEL_ROWS_NUM = 20000
EXCEL_START_COLUMN = 2

SCHEMA_NAME = 'reports'

BROKERAGE_MONTHLY_DEALS_ADR_MARKER = 'АДР'
BROKERAGE_MONTHLY_DEALS_SHARES_MARKER = 'Акция'
BROKERAGE_MONTHLY_DEALS_BONDS_MARKER = 'Облигация'

logger = Logger('brokerage_monthly_deals', st.APPLICATION_LOG, write_to_stdout=st.DEBUG_MODE).get()


class BrokerageMonthlyDealsAdr:
    def __init__(self, sheet, start_pos: int, stop_pos: int):
        self.class_name = self.__class__.__name__
        self.sheet = sheet
        self.start_column = EXCEL_START_COLUMN
        self.start_row: int = start_pos
        self.stop_row: int = stop_pos


class BrokerageMonthlyDealsShares:
    def __init__(self, sheet, start_pos: int, stop_pos: int):
        self.class_name = self.__class__.__name__
        self.sheet = sheet
        self.start_column = EXCEL_START_COLUMN
        self.start_row: int = start_pos
        self.stop_row: int = stop_pos


class BrokerageMonthlyDealsBonds:
    def __init__(self, sheet, start_pos: int, stop_pos: int):
        self.class_name = self.__class__.__name__
        self.sheet = sheet
        self.start_column = EXCEL_START_COLUMN
        self.start_row: int = start_pos
        self.stop_row: int = stop_pos


class BrokerageMonthlyDeals(Base):
    __tablename__ = 'brokerage_monthly_deals'
    __table_args__ = {"schema": SCHEMA_NAME}

    id = Column(Integer, primary_key=True)
    report_id = Column(Integer, ForeignKey(f'{SCHEMA_NAME}.brokerage_monthly.id'), unique=True)

    chapter_adr = Column(Boolean, nullable=False)
    chapter_shares = Column(Boolean, nullable=False)
    chapter_bonds = Column(Boolean, nullable=False)

    def __init__(self, sheet, start_pos: int, stop_pos: int):
        self.class_name = self.__class__.__name__
        self.sheet = sheet
        self.start_column = EXCEL_START_COLUMN
        self.start_row: int = start_pos
        self.stop_row: int = stop_pos

        self.chapter_adr = False
        self.chapter_shares = False
        self.chapter_bonds = False

        self.adr_start_row: int | None = None
        self.adr_stop_row: int | None = None
        self.shares_start_row: int | None = None
        self.shares_stop_row: int | None = None
        self.bonds_start_row: int | None = None
        self.bonds_stop_row: int | None = None

        self._detect_chapters()

        if self.chapter_adr:
            self.adr: BrokerageMonthlyDealsAdr = BrokerageMonthlyDealsAdr(
                self.sheet, self.adr_start_row, self.adr_stop_row
            )

        if self.chapter_shares:
            self.shares: BrokerageMonthlyDealsShares = BrokerageMonthlyDealsShares(
                self.sheet, self.shares_start_row, self.shares_stop_row
            )

        if self.chapter_bonds:
            self.bonds: BrokerageMonthlyDealsBonds = BrokerageMonthlyDealsBonds(
                self.sheet, self.bonds_start_row, self.bonds_stop_row
            )

    def __repr__(self):
        return f"<BrokerageMonthlyDeals({self.chapter_adr}, {self.chapter_shares}, {self.chapter_bonds})>"

    def _detect_chapters(self):
        for i in range(self.start_row, self.stop_row + 1):
            cell = self.sheet.cell(row=i, column=self.start_column)

            if cell.value == BROKERAGE_MONTHLY_DEALS_ADR_MARKER:
                self.adr_start_row = cell.row

            if cell.value == BROKERAGE_MONTHLY_DEALS_SHARES_MARKER:
                self.shares_start_row = cell.row

                if self.adr_start_row and not self.adr_stop_row:
                    self.adr_stop_row = cell.row - 1

            if cell.value == BROKERAGE_MONTHLY_DEALS_BONDS_MARKER:
                self.bonds_start_row = cell.row

                if self.adr_start_row and not self.adr_stop_row:
                    self.adr_stop_row = cell.row - 1

                if self.shares_start_row and not self.shares_stop_row:
                    self.shares_stop_row = cell.row - 1

        if self.adr_start_row and not self.adr_stop_row:
            self.adr_stop_row = self.stop_row - 1

        if self.shares_start_row and not self.shares_stop_row:
            self.shares_stop_row = self.stop_row - 1

        if self.bonds_start_row and not self.bonds_stop_row:
            self.bonds_stop_row = self.stop_row - 1

        if self.adr_start_row:
            if self.adr_stop_row:
                self.chapter_adr = True
                logger.info(f"{self.class_name}._detect_chapters(): "
                            f'Adr deals found at: {self.adr_start_row}, {self.adr_stop_row}')
            else:
                logger.error(f"{self.class_name}._detect_chapters(): "
                             f'Error finding ADR stop row')
                raise Exception(f'Error finding ADR stop row')

        if self.shares_start_row:
            if self.shares_stop_row:
                self.chapter_shares = True
                logger.info(f"{self.class_name}._detect_chapters(): "
                            f'Shares deals found at: {self.shares_start_row}, {self.shares_stop_row}')
            else:
                logger.error(f"{self.class_name}._detect_chapters(): "
                             f'Error finding Shares stop row')
                raise Exception(f'Error finding Shares stop row')

        if self.bonds_start_row:
            if self.bonds_stop_row:
                self.chapter_bonds = True
                logger.info(f"{self.class_name}._detect_chapters(): "
                            f'Bonds deals found at: {self.bonds_start_row}, {self.bonds_stop_row}')
            else:
                logger.error(f"{self.class_name}._detect_chapters(): "
                             f'Error finding Bonds stop row')
                raise Exception(f'Error finding Bonds stop row')
