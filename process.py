from datetime import datetime
from calendar import monthrange
import sys

import xlrd
from pprint import pprint


class Read_WB:
    def __init__(self):
        self.wb = xlrd.open_workbook(
            "docs/Demo_Assessment_Model_08.18.20.xlsx")
        self.sheet = self.wb.sheet_by_name("KPI Dashboard")
        self.last_col = self.sheet.ncols - 1
        self.last_row = self. sheet.nrows - 1

    def get_date(self, d):
        try:
            self.excel_date = datetime(
                *xlrd.xldate_as_tuple(d, self.wb.datemode))
            self.converted_date = self.excel_date.strftime(
                "%Y-%m-%dT%H:%M:%S.000+00:00")
            datetime.strptime(
                self.converted_date, f"%Y-%m-{monthrange(self.excel_date.year, self.excel_date.month)[1]}T00:00:00.000+00:00")
            return self.converted_date
        except ValueError:
            return False
        except TypeError:
            return False

    def is_category(self, idx):
        return True if self.get_date(self.sheet.cell_value(idx, 3)) else False

    def tests(self):
        print(self.is_category(3))
        print(self.get_date(43010.0))

    def process(self):
        pass


if __name__ == "__main__":
    Read_WB().tests()
