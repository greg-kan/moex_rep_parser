from pathlib import Path
from logger import Logger
import settings as st
from core import *
from reports import BrokerageMonthly
from db import Base
from db import engine
from sqlalchemy.orm import Session

logger = Logger('main', st.APPLICATION_LOG, write_to_stdout=st.DEBUG_MODE).get()

REPORTS_DIR = 'investment/bcs_reports/brokerage'
reports_dir = Path.home() / REPORTS_DIR
report_file = 'B_k-3192641_ALL_23-12.xlsx'
report_path = reports_dir / report_file


def routine():
    logger.info('Routine started')
    print()

    Base.metadata.create_all(engine)
    session = Session(bind=engine)

    # Just one report
    brokerage = BrokerageMonthly(report_path)
    print(brokerage.securities.report)
    print(brokerage.money.operations[0].money)
    session.add(brokerage)
    # End of just one report

    # # The test on all files
    # lst_files = os.listdir(reports_dir)
    # for fl in sorted(lst_files):
    #     if get_file_extension(fl) == '.xlsx':
    #         rep_path = reports_dir / fl
    #         print(rep_path)
    #         print()
    #
    #         brokerage = BrokerageMonthly(rep_path)
    #         print(brokerage.securities.report)
    #         print(len(brokerage.money.operations))
    #         if len(brokerage.money.operations) > 0:
    #             print(brokerage.money.operations[0].money)
    #         session.add(brokerage)
    #
    #         print()
    #         print()
    # # End of the test on all files

    session.commit()

    logger.info('Routine finished')


if __name__ == "__main__":
    routine()
