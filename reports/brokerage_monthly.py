from openpyxl import load_workbook
from pathlib import Path
from logger import Logger
import settings as st

logger = Logger('brokerage_monthly', st.APPLICATION_LOG, write_to_stdout=st.DEBUG_MODE).get()

MAX_EXCEL_ROWS_NUM = 20000
EXCEL_START_COLUMN = 2

STOP_REPORT_STR = 'Владелец: ООО "Компания БКС"'

# 1. Money
MONEY_START_STR = '1. Движение денежных средств'
MONEY_STOP_STR = '2.1. Сделки:'

# 2. Deals
DEALS_START_STR = '2.1. Сделки:'
DEALS_STOP_STR = '3. Активы:'

# 3. Securities
SECURITIES_START_STR = '3. Активы:'
SECURITIES_STOP_STR = '4. Движение Ценных бумаг'
SECURITIES_PORTFOLIO_TOTAL_STR = 'Стоимость портфеля (руб.):'
SECURITIES_PORTFOLIO_TOTAL_BEGIN_COLUMN = 8
SECURITIES_PORTFOLIO_TOTAL_END_COLUMN = 12


# 4. Securities Transactions
SECURITIES_TRANSACTIONS_START_STR = '4. Движение Ценных бумаг'
SECURITIES_TRANSACTIONS_STOP_STR = '(1*) - оплата проводится со своего (клиента) счета'


class Money:
    def __init__(self, sheet):
        self.sheet = sheet


class Deals:
    def __init__(self, sheet):
        self.sheet = sheet


class Securities:
    def __init__(self, sheet):
        self.class_name = self.__class__.__name__
        self.sheet = sheet
        self.start_column = EXCEL_START_COLUMN
        self.securities: list[dict] | None = None
        self.start_row: int | None = None
        self.stop_row: int | None = None
        self.portfolio_total_row: int | None = None
        self.portfolio_total_begin_column: int = SECURITIES_PORTFOLIO_TOTAL_BEGIN_COLUMN
        self.portfolio_total_end_column: int = SECURITIES_PORTFOLIO_TOTAL_END_COLUMN
        self.portfolio_total_value_begin_rub: float = 0
        self.portfolio_total_value_end_rub: float = 0

        self._find_boundaries()
        self._find_portfolio_total_row()
        self._extract_total_portfolio_values_rub()
        self.load()

    def _find_boundaries(self):
        for i in range(1, MAX_EXCEL_ROWS_NUM):
            cell = self.sheet.cell(row=i, column=self.start_column)
            if cell.value == SECURITIES_START_STR:
                self.start_row = cell.row
                break

        if self.start_row:
            for i in range(self.start_row, MAX_EXCEL_ROWS_NUM):
                cell = self.sheet.cell(row=i, column=self.start_column)
                if cell.value == SECURITIES_STOP_STR:
                    self.stop_row = cell.row - 1
                    break
        else:
            logger.error(f"{self.class_name}._find_boundaries(): "
                         f"Could not find a Securities start row index")

            raise Exception('Could not find a Securities start row index')

        if not self.stop_row:
            logger.error(f"{self.class_name}._find_boundaries(): "
                         f"Could not find a Securities stop row index")

            raise Exception('Could not find a Securities stop row index')

        logger.info(f"{self.class_name}._find_boundaries(): "
                    f'Securities boundaries found: {self.start_row}, {self.stop_row}')

    def _find_portfolio_total_row(self):
        if self.start_row and self.stop_row:
            for i in range(self.start_row, self.stop_row):
                cell = self.sheet.cell(row=i, column=self.start_column)
                if cell.value == SECURITIES_PORTFOLIO_TOTAL_STR:
                    self.portfolio_total_row = cell.row

        else:
            logger.error(f"{self.class_name}._find_portfolio_total_row(): "
                         f"No securities start or/and stop row(s) defined")

            raise Exception('No securities start or/and stop row(s) defined')

        logger.info(f"{self.class_name}._find_portfolio_total_row(): "
                    f'Securities portfolio total row found: {self.portfolio_total_row}')

    def _extract_total_portfolio_values_rub(self):
        if self.portfolio_total_row and self.portfolio_total_begin_column:
            cell_begin = self.sheet.cell(row=self.portfolio_total_row,
                                         column=self.portfolio_total_begin_column)
            self.portfolio_total_value_begin_rub = float(cell_begin.value)
            logger.info(f"{self.class_name}._extract_total_portfolio_values_rub(): "
                        f"portfolio_total_value_begin_rub = {self.portfolio_total_value_begin_rub}")

        if self.portfolio_total_row and self.portfolio_total_end_column:
            cell_end = self.sheet.cell(row=self.portfolio_total_row,
                                       column=self.portfolio_total_end_column)
            self.portfolio_total_value_end_rub = float(cell_end.value)
            logger.info(f"{self.class_name}._extract_total_portfolio_values_rub(): "
                        f"portfolio_total_value_end_rub = {self.portfolio_total_value_end_rub}")



    # def find_securities_entry(self) -> int | None:
    #     for i in range(1, MAX_EXCEL_ROWS_NUM):
    #         cell = self.sheet.cell(row=i, column=self.start_column)
    #         cell_next = self.sheet.cell(row=i+1, column=self.start_column)
    #         if cell.value == SECURITIES_ENTRY_STR and cell_next.value == 'Вид актива':
    #             entry_index = i + 2
    #             cell = self.sheet.cell(row=entry_index, column=self.start_column)
    #             print(entry_index)
    #             print(cell.row, cell.column)
    #             print(cell.coordinate)
    #             return cell.row
    #
    #     return None

    def load(self):
        pass


class SecuritiesTransactions:
    def __init__(self, sheet):
        self.sheet = sheet


class BrokerageMonthly:
    # 1. Money
    money: Money
    # 2. Deals
    deals: Deals
    # 3. Securities
    securities: Securities
    # 4. Securities Transactions
    securities_transactions: SecuritiesTransactions

    def __init__(self, report_path):
        self.stop_report_str = STOP_REPORT_STR
        self.class_name = self.__class__.__name__
        self.report_path: Path = report_path
        self.workbook = load_workbook(self.report_path)
        self.sheet = self.workbook[self.workbook.sheetnames[0]]

        self.money = Money(self.sheet)
        self.deals = Deals(self.sheet)
        self.securities = Securities(self.sheet)
        self.securities_transactions = SecuritiesTransactions(self.sheet)

    def feature_method(self):
        pass


