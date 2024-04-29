from logger import Logger
import settings as st
from core import *

from datetime import datetime, date, time
from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship
from sqlalchemy import Column, Integer, String, DateTime, Numeric, DOUBLE_PRECISION, Boolean, Date
from sqlalchemy.sql import func
from db import Base

MAX_EXCEL_ROWS_NUM = 20000
EXCEL_START_COLUMN = 2

SCHEMA_NAME = 'reports'

BROKERAGE_MONTHLY_DEALS_ADR_MARKER = 'АДР'
BROKERAGE_MONTHLY_DEALS_SHARES_MARKER = 'Акция'
BROKERAGE_MONTHLY_DEALS_BONDS_MARKER = 'Облигация'

ADR_TABLE_START_STR1 = 'Валюта цены = Рубль, валюта платежа = Рубль'
ADR_TABLE_START_STR2 = 'Дата'

SHARES_TABLE_START_STR1 = 'Валюта цены = Рубль, валюта платежа = Рубль'
SHARES_TABLE_START_STR2 = 'Дата'

BONDS_TABLE_START_STR1 = 'Валюта цены = Рубль, валюта платежа = Рубль'
BONDS_TABLE_START_STR2 = 'Дата'

logger = Logger('brokerage_monthly_deals', st.APPLICATION_LOG, write_to_stdout=st.DEBUG_MODE).get()


class BrokerageMonthlyDealsAdr(Base):

    __tablename__ = 'brokerage_monthly_deals_adr'
    __table_args__ = {"schema": SCHEMA_NAME}

    id = Column(Integer, primary_key=True)
    deals_id = Column(Integer, ForeignKey(f'{SCHEMA_NAME}.brokerage_monthly_deals.id'))
    deals = relationship("BrokerageMonthlyDeals", backref="adrs")

    secid = Column(String, nullable=False)
    regnumber = Column(String)
    isin = Column(String, nullable=False)
    emitent = Column(String, nullable=False)

    deal_date = Column(Date, nullable=False)
    deal_num = Column(String, nullable=False)
    deal_time = Column(String)
    bought = Column(DOUBLE_PRECISION, nullable=False)
    buy_price = Column(Numeric(19, 6), nullable=False)
    buy_sum = Column(Numeric(19, 6), nullable=False)
    sold = Column(DOUBLE_PRECISION, nullable=False)
    sell_price = Column(Numeric(19, 6), nullable=False)
    sell_sum = Column(Numeric(19, 6), nullable=False)
    currency = Column(String, nullable=False)
    currency_pay = Column(String, nullable=False)
    commission_date = Column(Date, nullable=False)
    commission_time = Column(DateTime, nullable=False)
    deal_type = Column(String)
    pay_date = Column(Date, nullable=False)
    delivery_date = Column(Date, nullable=False)
    platform = Column(String, nullable=False)
    note = Column(String)

    inserted = Column(DateTime(), server_default=func.now())
    updated = Column(DateTime(), onupdate=func.now())

    def __init__(self, secid: str, regnumber: str, isin: str, emitent: str, deal_date: date, deal_num: str,
                 deal_time,
                 bought: float, buy_price: float, buy_sum: float, sold: float, sell_price: float,
                 sell_sum: float, currency: str, currency_pay: str, commission_date: date,
                 commission_time: datetime, deal_type: str, pay_date: date, delivery_date: date,
                 platform: str, note: str):

        self.class_name = self.__class__.__name__

        self.secid = secid
        self.regnumber = regnumber
        self.isin = isin
        self.emitent = emitent
        self.deal_date = deal_date
        self.deal_num = deal_num
        self.deal_time = deal_time
        self.bought = bought
        self.buy_price = buy_price
        self.buy_sum = buy_sum
        self.sold = sold
        self.sell_price = sell_price
        self.sell_sum = sell_sum
        self.currency = currency
        self.currency_pay = currency_pay
        self.commission_date = commission_date
        self.commission_time = commission_time
        self.deal_type = deal_type
        self.pay_date = pay_date
        self.delivery_date = delivery_date
        self.platform = platform
        self.note = note

    def __repr__(self):
        return f"<BrokerageMonthlyDealsArd({self.secid}, {self.isin})>"


class BrokerageMonthlyDealsShare(Base):

    __tablename__ = 'brokerage_monthly_deals_share'
    __table_args__ = {"schema": SCHEMA_NAME}

    id = Column(Integer, primary_key=True)
    deals_id = Column(Integer, ForeignKey(f'{SCHEMA_NAME}.brokerage_monthly_deals.id'))
    deals = relationship("BrokerageMonthlyDeals", backref="shares")

    secid = Column(String, nullable=False)
    regnumber = Column(String)
    isin = Column(String, nullable=False)
    emitent = Column(String, nullable=False)

    deal_date = Column(Date, nullable=False)
    deal_num = Column(String, nullable=False)
    deal_time = Column(String)
    bought = Column(DOUBLE_PRECISION, nullable=False)
    buy_price = Column(Numeric(19, 6), nullable=False)
    buy_sum = Column(Numeric(19, 6), nullable=False)
    sold = Column(DOUBLE_PRECISION, nullable=False)
    sell_price = Column(Numeric(19, 6), nullable=False)
    sell_sum = Column(Numeric(19, 6), nullable=False)
    currency = Column(String, nullable=False)
    currency_pay = Column(String, nullable=False)
    commission_date = Column(Date, nullable=False)
    commission_time = Column(DateTime, nullable=False)
    deal_type = Column(String)
    pay_date = Column(Date, nullable=False)
    delivery_date = Column(Date, nullable=False)
    platform = Column(String, nullable=False)
    note = Column(String)

    inserted = Column(DateTime(), server_default=func.now())
    updated = Column(DateTime(), onupdate=func.now())

    def __init__(self, secid: str, regnumber: str, isin: str, emitent: str, deal_date: date, deal_num: str,
                 deal_time,
                 bought: float, buy_price: float, buy_sum: float, sold: float, sell_price: float,
                 sell_sum: float, currency: str, currency_pay: str, commission_date: date,
                 commission_time: datetime, deal_type: str, pay_date: date, delivery_date: date,
                 platform: str, note: str):

        self.class_name = self.__class__.__name__

        self.secid = secid
        self.regnumber = regnumber
        self.isin = isin
        self.emitent = emitent
        self.deal_date = deal_date
        self.deal_num = deal_num
        self.deal_time = deal_time
        self.bought = bought
        self.buy_price = buy_price
        self.buy_sum = buy_sum
        self.sold = sold
        self.sell_price = sell_price
        self.sell_sum = sell_sum
        self.currency = currency
        self.currency_pay = currency_pay
        self.commission_date = commission_date
        self.commission_time = commission_time
        self.deal_type = deal_type
        self.pay_date = pay_date
        self.delivery_date = delivery_date
        self.platform = platform
        self.note = note

    def __repr__(self):
        return f"<BrokerageMonthlyDealsShare({self.secid}, {self.isin})>"


class BrokerageMonthlyDealsBond(Base):
    __tablename__ = 'brokerage_monthly_deals_bond'
    __table_args__ = {"schema": SCHEMA_NAME}

    id = Column(Integer, primary_key=True)
    deals_id = Column(Integer, ForeignKey(f'{SCHEMA_NAME}.brokerage_monthly_deals.id'))
    deals = relationship("BrokerageMonthlyDeals", backref="bonds")

    secid = Column(String, nullable=False)
    regnumber = Column(String)
    isin = Column(String, nullable=False)
    emitent = Column(String, nullable=False)

    deal_date = Column(Date, nullable=False)
    deal_num = Column(String, nullable=False)
    deal_time = Column(String)
    bought = Column(DOUBLE_PRECISION, nullable=False)
    buy_price = Column(Numeric(19, 6), nullable=False)
    buy_sum = Column(Numeric(19, 6), nullable=False)
    buy_nkd = Column(Numeric(19, 6), nullable=False)
    sold = Column(DOUBLE_PRECISION, nullable=False)
    sell_price = Column(Numeric(19, 6), nullable=False)
    sell_sum = Column(Numeric(19, 6), nullable=False)
    sell_nkd = Column(Numeric(19, 6), nullable=False)
    currency = Column(String, nullable=False)
    currency_pay = Column(String, nullable=False)
    commission_date = Column(Date, nullable=False)
    deal_type = Column(String)
    pay_date = Column(Date, nullable=False)
    delivery_date = Column(Date, nullable=False)
    platform = Column(String, nullable=False)
    note = Column(String)

    inserted = Column(DateTime(), server_default=func.now())
    updated = Column(DateTime(), onupdate=func.now())

    def __init__(self, secid: str, regnumber: str, isin: str, emitent: str, deal_date: date, deal_num: str,
                 deal_time,
                 bought: float, buy_price: float, buy_sum: float, buy_nkd: float, sold: float, sell_price: float,
                 sell_sum: float, sell_nkd: float, currency: str, currency_pay: str, commission_date: date,
                 deal_type: str, pay_date: date, delivery_date: date, platform: str, note: str):

        self.class_name = self.__class__.__name__

        self.secid = secid
        self.regnumber = regnumber
        self.isin = isin
        self.emitent = emitent
        self.deal_date = deal_date
        self.deal_num = deal_num
        self.deal_time = deal_time
        self.bought = bought
        self.buy_price = buy_price
        self.buy_sum = buy_sum
        self.buy_nkd = buy_nkd
        self.sold = sold
        self.sell_price = sell_price
        self.sell_sum = sell_sum
        self.sell_nkd = sell_nkd
        self.currency = currency
        self.currency_pay = currency_pay
        self.commission_date = commission_date
        self.deal_type = deal_type
        self.pay_date = pay_date
        self.delivery_date = delivery_date
        self.platform = platform
        self.note = note

    def __repr__(self):
        return f"<BrokerageMonthlyDealsBond({self.secid}, {self.isin})>"


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
            self._find_adr_table_boundaries()
            self._load_adr_table()

        if self.chapter_shares:
            self._find_shares_table_boundaries()
            self._load_shares_table()

        if self.chapter_bonds:
            self._find_bonds_table_boundaries()
            self._load_bonds_table()

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
            self.adr_stop_row = self.stop_row  # - 1

        if self.shares_start_row and not self.shares_stop_row:
            self.shares_stop_row = self.stop_row  # - 1

        if self.bonds_start_row and not self.bonds_stop_row:
            self.bonds_stop_row = self.stop_row  # - 1

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

    def _find_adr_table_boundaries(self):
        for i in range(self.adr_start_row, self.adr_stop_row + 1):
            cell = self.sheet.cell(row=i, column=self.start_column)
            if cell.value == ADR_TABLE_START_STR1:
                cell_next = self.sheet.cell(row=i+1, column=self.start_column)
                if cell_next.value == ADR_TABLE_START_STR2:
                    self.adr_table_start_row = cell_next.row + 1
                    break

        if self.adr_table_start_row:
            self.adr_table_stop_row = self.adr_stop_row - 1

            logger.info(f"{self.class_name}._find_adr_table_boundaries(): "
                        f"ADR table boundaries found: {self.adr_table_start_row}, {self.adr_table_stop_row}")

        else:
            logger.error(f"{self.class_name}._find_adr_table_boundaries(): "
                         f"Error: ADR table boundaries not found")
            raise Exception(f"Error: ADR table boundaries not found")

    def _load_adr_table(self):
        i = self.adr_table_start_row
        while i < self.adr_table_stop_row:
            cell = self.sheet.cell(row=i, column=self.start_column)
            secid = cell.value
            cell = self.sheet.cell(row=i, column=6)
            regnumber = cell.value
            cell = self.sheet.cell(row=i, column=8)
            isin = cell.value
            cell = self.sheet.cell(row=i, column=9)
            emitent = cell.value

            sec_stop = f'Итого по {secid}:'

            sum_buy = 0
            sum_sell = 0

            while True:
                i += 1
                if i > MAX_EXCEL_ROWS_NUM:
                    break

                cell = self.sheet.cell(row=i, column=self.start_column)
                if cell.value != sec_stop:
                    deal_date: date = datetime.strptime(cell.value, "%d.%m.%y")
                    cell = self.sheet.cell(row=i, column=3)
                    deal_num = cell.value
                    cell = self.sheet.cell(row=i, column=4)
                    deal_time = cell.value

                    cell = self.sheet.cell(row=i, column=5)
                    bought = tofloat(ifnull(cell.value, 0))
                    cell = self.sheet.cell(row=i, column=6)
                    buy_price = tofloat(ifnull(cell.value, 0))
                    cell = self.sheet.cell(row=i, column=7)
                    buy_sum = tofloat(ifnull(cell.value, 0))
                    sum_buy += buy_sum

                    cell = self.sheet.cell(row=i, column=8)
                    sold = tofloat(ifnull(cell.value, 0))
                    cell = self.sheet.cell(row=i, column=9)
                    sell_price = tofloat(ifnull(cell.value, 0))
                    cell = self.sheet.cell(row=i, column=10)
                    sell_sum = tofloat(ifnull(cell.value, 0))
                    sum_sell += sell_sum

                    cell = self.sheet.cell(row=i, column=11)
                    currency = cell.value
                    cell = self.sheet.cell(row=i, column=12)
                    currency_pay = cell.value

                    cell_date = self.sheet.cell(row=i, column=13)
                    commission_date: date = datetime.strptime(cell_date.value, "%d.%m.%y")

                    cell_time = self.sheet.cell(row=i, column=14)
                    str_datetime = f'{cell_date.value} {cell_time.value}'
                    commission_time: datetime = datetime.strptime(str_datetime, "%d.%m.%y %H:%M:%S")

                    cell = self.sheet.cell(row=i, column=15)
                    deal_type = cell.value

                    cell = self.sheet.cell(row=i, column=16)
                    pay_date: date = datetime.strptime(cell.value, "%d.%m.%y")

                    cell = self.sheet.cell(row=i, column=17)
                    delivery_date: date = datetime.strptime(cell.value, "%d.%m.%y")

                    cell = self.sheet.cell(row=i, column=18)
                    platform = cell.value

                    cell = self.sheet.cell(row=i, column=20)
                    note = cell.value

                    adr = BrokerageMonthlyDealsAdr(secid, regnumber, isin, emitent,
                                                   deal_date, deal_num, deal_time,
                                                   bought, buy_price, buy_sum, sold, sell_price, sell_sum,
                                                   currency, currency_pay, commission_date, commission_time,
                                                   deal_type, pay_date, delivery_date, platform, note)
                    self.adrs.append(adr)

                else:
                    cell = self.sheet.cell(row=i, column=7)
                    summary_sum_buy = tofloat(ifnull(cell.value, 0))

                    cell = self.sheet.cell(row=i, column=10)
                    summary_sum_sell = tofloat(ifnull(cell.value, 0))

                    logger.info(f"{self.class_name}._load_shares_table(): "
                                f"Summary: {summary_sum_buy}, {summary_sum_sell}")

                    if round(summary_sum_buy, 6) == round(sum_buy, 6) and \
                       round(summary_sum_sell, 6) == round(sum_sell, 6):
                        logger.info(f"{self.class_name}._load_shares_table(): "
                                    f"Summaries are match calculations. It is good!")

                    else:
                        logger.error(f"{self.class_name}._load_shares_table(): "
                                     f"Summaries are not match calculations. It is bad!")
                        raise Exception("Summaries are not match calculations. It is bad!")

                    break

            i += 1
    def _find_shares_table_boundaries(self):
        for i in range(self.shares_start_row, self.shares_stop_row + 1):
            cell = self.sheet.cell(row=i, column=self.start_column)
            if cell.value == SHARES_TABLE_START_STR1:
                cell_next = self.sheet.cell(row=i+1, column=self.start_column)
                if cell_next.value == SHARES_TABLE_START_STR2:
                    self.shares_table_start_row = cell_next.row + 1
                    break

        if self.shares_table_start_row:
            self.shares_table_stop_row = self.shares_stop_row - 1

            logger.info(f"{self.class_name}._find_shares_table_boundaries(): "
                        f"Shares table boundaries found: {self.shares_table_start_row}, {self.shares_table_stop_row}")

        else:
            logger.error(f"{self.class_name}._find_shares_table_boundaries(): "
                         f"Error: Shares table boundaries not found")
            raise Exception(f"Error: Shares table boundaries not found")

    def _load_shares_table(self):
        i = self.shares_table_start_row
        while i < self.shares_table_stop_row:
            cell = self.sheet.cell(row=i, column=self.start_column)
            secid = cell.value
            cell = self.sheet.cell(row=i, column=6)
            regnumber = cell.value
            cell = self.sheet.cell(row=i, column=8)
            isin = cell.value
            cell = self.sheet.cell(row=i, column=9)
            emitent = cell.value

            sec_stop = f'Итого по {secid}:'

            sum_buy = 0
            sum_sell = 0

            while True:
                i += 1
                if i > MAX_EXCEL_ROWS_NUM:
                    break

                cell = self.sheet.cell(row=i, column=self.start_column)
                if cell.value != sec_stop:
                    deal_date: date = datetime.strptime(cell.value, "%d.%m.%y")
                    cell = self.sheet.cell(row=i, column=3)
                    deal_num = cell.value
                    cell = self.sheet.cell(row=i, column=4)
                    deal_time = cell.value

                    cell = self.sheet.cell(row=i, column=5)
                    bought = tofloat(ifnull(cell.value, 0))
                    cell = self.sheet.cell(row=i, column=6)
                    buy_price = tofloat(ifnull(cell.value, 0))
                    cell = self.sheet.cell(row=i, column=7)
                    buy_sum = tofloat(ifnull(cell.value, 0))
                    sum_buy += buy_sum

                    cell = self.sheet.cell(row=i, column=8)
                    sold = tofloat(ifnull(cell.value, 0))
                    cell = self.sheet.cell(row=i, column=9)
                    sell_price = tofloat(ifnull(cell.value, 0))
                    cell = self.sheet.cell(row=i, column=10)
                    sell_sum = tofloat(ifnull(cell.value, 0))
                    sum_sell += sell_sum

                    cell = self.sheet.cell(row=i, column=11)
                    currency = cell.value
                    cell = self.sheet.cell(row=i, column=12)
                    currency_pay = cell.value

                    cell_date = self.sheet.cell(row=i, column=13)
                    commission_date: date = datetime.strptime(cell_date.value, "%d.%m.%y")

                    cell_time = self.sheet.cell(row=i, column=14)
                    str_datetime = f'{cell_date.value} {cell_time.value}'
                    commission_time: datetime = datetime.strptime(str_datetime, "%d.%m.%y %H:%M:%S")

                    cell = self.sheet.cell(row=i, column=15)
                    deal_type = cell.value

                    cell = self.sheet.cell(row=i, column=16)
                    pay_date: date = datetime.strptime(cell.value, "%d.%m.%y")

                    cell = self.sheet.cell(row=i, column=17)
                    delivery_date: date = datetime.strptime(cell.value, "%d.%m.%y")

                    cell = self.sheet.cell(row=i, column=18)
                    platform = cell.value

                    cell = self.sheet.cell(row=i, column=20)
                    note = cell.value

                    share = BrokerageMonthlyDealsShare(secid, regnumber, isin, emitent,
                                                       deal_date, deal_num, deal_time,
                                                       bought, buy_price, buy_sum, sold, sell_price, sell_sum,
                                                       currency, currency_pay, commission_date, commission_time,
                                                       deal_type, pay_date, delivery_date, platform, note)
                    self.shares.append(share)

                else:
                    cell = self.sheet.cell(row=i, column=7)
                    summary_sum_buy = tofloat(ifnull(cell.value, 0))

                    cell = self.sheet.cell(row=i, column=10)
                    summary_sum_sell = tofloat(ifnull(cell.value, 0))

                    logger.info(f"{self.class_name}._load_shares_table(): "
                                f"Summary: {summary_sum_buy}, {summary_sum_sell}")

                    if round(summary_sum_buy, 6) == round(sum_buy, 6) and \
                       round(summary_sum_sell, 6) == round(sum_sell, 6):
                        logger.info(f"{self.class_name}._load_shares_table(): "
                                    f"Summaries are match calculations. It is good!")

                    else:
                        logger.error(f"{self.class_name}._load_shares_table(): "
                                     f"Summaries are not match calculations. It is bad!")
                        raise Exception("Summaries are not match calculations. It is bad!")

                    break

            i += 1

    def _find_bonds_table_boundaries(self):
        for i in range(self.bonds_start_row, self.bonds_stop_row + 1):
            cell = self.sheet.cell(row=i, column=self.start_column)
            if cell.value == BONDS_TABLE_START_STR1:
                cell_next = self.sheet.cell(row=i + 1, column=self.start_column)
                if cell_next.value == BONDS_TABLE_START_STR2:
                    self.bonds_table_start_row = cell_next.row + 1
                    break

        if self.bonds_table_start_row:
            self.bonds_table_stop_row = self.bonds_stop_row - 1

            logger.info(f"{self.class_name}._find_bonds_table_boundaries(): "
                        f"Bonds table boundaries found: {self.bonds_table_start_row}, {self.bonds_table_stop_row}")

        else:
            logger.error(f"{self.class_name}._find_bonds_table_boundaries(): "
                         f"Error: Bonds table boundaries not found")
            raise Exception(f"Error: Bonds table boundaries not found")

    def _load_bonds_table(self):
        i = self.bonds_table_start_row
        while i < self.bonds_table_stop_row:
            cell = self.sheet.cell(row=i, column=self.start_column)
            secid = cell.value
            cell = self.sheet.cell(row=i, column=6)
            regnumber = cell.value
            cell = self.sheet.cell(row=i, column=8)
            isin = cell.value
            cell = self.sheet.cell(row=i, column=9)
            emitent = cell.value

            sec_stop = f'Итого по {secid}:'

            sum_buy = 0
            nkd_buy = 0
            sum_sell = 0
            nkd_sell = 0

            while True:
                i += 1
                if i > MAX_EXCEL_ROWS_NUM:
                    break

                cell = self.sheet.cell(row=i, column=self.start_column)
                if cell.value != sec_stop:
                    deal_date: date = datetime.strptime(cell.value, "%d.%m.%y")
                    cell = self.sheet.cell(row=i, column=3)
                    deal_num = cell.value
                    cell = self.sheet.cell(row=i, column=4)
                    deal_time = cell.value

                    cell = self.sheet.cell(row=i, column=5)
                    bought = tofloat(ifnull(cell.value, 0))
                    cell = self.sheet.cell(row=i, column=6)
                    buy_price = tofloat(ifnull(cell.value, 0))
                    cell = self.sheet.cell(row=i, column=7)
                    buy_sum = tofloat(ifnull(cell.value, 0))
                    sum_buy += buy_sum
                    cell = self.sheet.cell(row=i, column=8)
                    buy_nkd = tofloat(ifnull(cell.value, 0))
                    nkd_buy += buy_nkd

                    cell = self.sheet.cell(row=i, column=9)
                    sold = tofloat(ifnull(cell.value, 0))
                    cell = self.sheet.cell(row=i, column=10)
                    sell_price = tofloat(ifnull(cell.value, 0))
                    cell = self.sheet.cell(row=i, column=11)
                    sell_sum = tofloat(ifnull(cell.value, 0))
                    sum_sell += sell_sum
                    cell = self.sheet.cell(row=i, column=12)
                    sell_nkd = tofloat(ifnull(cell.value, 0))
                    nkd_sell += sell_nkd

                    cell = self.sheet.cell(row=i, column=13)
                    currency = cell.value
                    cell = self.sheet.cell(row=i, column=14)
                    currency_pay = cell.value

                    cell = self.sheet.cell(row=i, column=15)
                    commission_date: date = datetime.strptime(cell.value, "%d.%m.%y")

                    cell = self.sheet.cell(row=i, column=16)
                    deal_type = cell.value

                    cell = self.sheet.cell(row=i, column=17)
                    pay_date: date = datetime.strptime(cell.value, "%d.%m.%y")

                    cell = self.sheet.cell(row=i, column=18)
                    delivery_date: date = datetime.strptime(cell.value, "%d.%m.%y")

                    cell = self.sheet.cell(row=i, column=19)
                    platform = cell.value

                    cell = self.sheet.cell(row=i, column=21)
                    note = cell.value

                    bond = BrokerageMonthlyDealsBond(secid, regnumber, isin, emitent,
                                                     deal_date, deal_num, deal_time,
                                                     bought, buy_price, buy_sum, buy_nkd,
                                                     sold, sell_price, sell_sum, sell_nkd,
                                                     currency, currency_pay, commission_date, deal_type, pay_date,
                                                     delivery_date, platform, note)
                    self.bonds.append(bond)

                else:
                    cell = self.sheet.cell(row=i, column=7)
                    summary_sum_buy = tofloat(ifnull(cell.value, 0))

                    cell = self.sheet.cell(row=i, column=8)
                    summary_nkd_buy = tofloat(ifnull(cell.value, 0))

                    cell = self.sheet.cell(row=i, column=11)
                    summary_sum_sell = tofloat(ifnull(cell.value, 0))

                    cell = self.sheet.cell(row=i, column=12)
                    summary_nkd_sell = tofloat(ifnull(cell.value, 0))

                    logger.info(f"{self.class_name}._load_bonds_table(): "
                                f"Summary: {summary_sum_buy}, {summary_nkd_buy}, "
                                f"{summary_sum_sell}, {summary_nkd_sell}")

                    if round(summary_sum_buy, 6) == round(sum_buy, 6) and \
                       round(summary_nkd_buy, 6) == round(nkd_buy, 6) and \
                       round(summary_sum_sell, 6) == round(sum_sell, 6) and \
                       round(summary_nkd_sell, 6) == round(nkd_sell, 6):

                        logger.info(f"{self.class_name}._load_bonds_table(): "
                                    f"Summaries are match calculations. It is good!")

                    else:
                        logger.error(f"{self.class_name}._load_bonds_table(): "
                                     f"Summaries are not match calculations. It is bad!")
                        raise Exception("Summaries are not match calculations. It is bad!")

                    break

            i += 1
