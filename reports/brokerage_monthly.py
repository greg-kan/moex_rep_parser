from openpyxl import load_workbook
from pathlib import Path

MAX_EXCEL_ROWS_NUM = 20000
EXCEL_START_COLUMN = 2

END_OF_REPORT_STR = 'Владелец: ООО "Компания БКС"'
SECURITIES_ENTRY_STR = 'Портфель по ценным бумагам, денежным средствам и ДМ (Рубль)'


class Securities:
    def __init__(self, workbook):
        self.workbook = workbook
        self.start_column = EXCEL_START_COLUMN
        self.securities: list[dict] | None = None
        self.sheet = self.workbook[self.workbook.sheetnames[0]]
        self.securities_entry_pos: int | None = self.find_securities_entry()
        self.load()

    def load(self):
        pass

    def find_securities_entry(self) -> int | None:
        for i in range(1, MAX_EXCEL_ROWS_NUM):
            cell = self.sheet.cell(row=i, column=self.start_column)
            cell_next = self.sheet.cell(row=i + 1, column=self.start_column)
            if cell.value == SECURITIES_ENTRY_STR and cell_next.value == 'Вид актива':
                entry_index = i + 2
                cell = self.sheet.cell(row=entry_index, column=self.start_column)
                print(entry_index)
                print(cell.row, cell.column)
                print(cell.coordinate)
                return cell

        return None


class BrokerageMonthly:
    securities: Securities

    def __init__(self, report_path):
        self.end_of_report_str = END_OF_REPORT_STR
        self.data: list[dict] | None = None
        self.class_name = self.__class__.__name__
        self.report_path: Path = report_path

        self.workbook = load_workbook(report_path)

        self.securities = Securities(self.workbook)

    def feature_method(self):
        pass
