from openpyxl import load_workbook
from pathlib import Path

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
        self.sheet = sheet
        self.start_column = EXCEL_START_COLUMN
        self.securities: list[dict] | None = None
        self.start_pos: int | None = None
        self.stop_pos: int | None = None
        self.portfolio_total_pos: int | None = None
        self.find_boundaries()
        self.find_portfolio_total_pos()
        self.load()

    def find_boundaries(self):
        for i in range(1, MAX_EXCEL_ROWS_NUM):
            cell = self.sheet.cell(row=i, column=self.start_column)
            if cell.value == SECURITIES_START_STR:
                self.start_pos = cell.row
                break

        if self.start_pos:
            for i in range(self.start_pos, MAX_EXCEL_ROWS_NUM):
                cell = self.sheet.cell(row=i, column=self.start_column)
                if cell.value == SECURITIES_STOP_STR:
                    self.stop_pos = cell.row
                    break
        else:
            raise Exception('Could not find a Securities entry row index')

        if not self.stop_pos:
            raise Exception('Could not find a Securities end row index')

    def find_portfolio_total_pos(self):
        if self.start_pos and self.stop_pos:
            for i in range(self.start_pos, self.stop_pos):
                cell = self.sheet.cell(row=i, column=self.start_column)
                if cell.value == SECURITIES_PORTFOLIO_TOTAL_STR:
                    self.portfolio_total_pos = cell.row

        else:
            raise Exception('No securities start or/and stop position(s) defined')


    def find_securities_entry(self) -> int | None:
        for i in range(1, MAX_EXCEL_ROWS_NUM):
            cell = self.sheet.cell(row=i, column=self.start_column)
            cell_next = self.sheet.cell(row=i+1, column=self.start_column)
            if cell.value == SECURITIES_ENTRY_STR and cell_next.value == 'Вид актива':
                entry_index = i + 2
                cell = self.sheet.cell(row=entry_index, column=self.start_column)
                print(entry_index)
                print(cell.row, cell.column)
                print(cell.coordinate)
                return cell.row

        return None

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


