from openpyxl import load_workbook
from pathlib import Path
import os
from datetime import datetime

from reports import BrokerageMonthly

REPORTS_DIR = 'investment/bcs_reports/brokerage'
reports_dir = Path.home() / REPORTS_DIR
report_file = 'B_k-3192641_ALL_23-12.xlsx'
report_path = reports_dir / report_file

END_OF_REPORT_STR = 'Владелец: ООО "Компания БКС"'


def routine():
    brokerage_2023_12 = BrokerageMonthly(report_path)
    print('securities_entry_pos =', brokerage_2023_12.securities.securities_entry_pos)


if __name__ == "__main__":
    routine()
