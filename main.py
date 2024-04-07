from openpyxl import load_workbook
from pathlib import Path
import os
from datetime import datetime

from logger import Logger
import settings as st
from core import *
from reports import BrokerageMonthly
from db.db_helper import *

logger = Logger('main', st.APPLICATION_LOG, write_to_stdout=st.DEBUG_MODE).get()

REPORTS_DIR = 'investment/bcs_reports/brokerage'
reports_dir = Path.home() / REPORTS_DIR
report_file = 'B_k-3192641_ALL_23-12.xlsx'
# report_file = 'B_k-3192641_ALL_23-04.xlsx'
report_path = reports_dir / report_file

# STOP_REPORT_STR = 'Владелец: ООО "Компания БКС"'


def routine():
    logger.info('Routine started')

    # # Test on all files
    # lst_files = os.listdir(reports_dir)
    # for fl in sorted(lst_files):
    #     if get_file_extension(fl) == '.xlsx':
    #         rep_path = reports_dir / fl
    #         print(rep_path)
    #         print()
    #         brokerage = BrokerageMonthly(rep_path)
    #         print()
    #         print()

    brokerage = BrokerageMonthly(report_path)
    print('securities_start_row =', brokerage.securities.start_row)

    logger.info('Routine stopped')


if __name__ == "__main__":
    trying_to_connect()
    # routine()
