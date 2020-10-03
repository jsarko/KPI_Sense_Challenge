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

    def get_date(self, d, is_first_of_month=False):
        try:
            self.excel_date = datetime(
                *xlrd.xldate_as_tuple(d, self.wb.datemode))
            if is_first_of_month:
                day = "01"
            else:
                day = monthrange(self.excel_date.year,
                                 self.excel_date.month)[1]
            self.converted_date = self.excel_date.strftime(
                f"%Y-%m-{day}T%H:%M:%S.000+00:00")
            datetime.strptime(
                self.converted_date, f"%Y-%m-{day}T00:00:00.000+00:00")
            return self.converted_date
        except ValueError as ve:
            return False
        except TypeError:
            return False

    def has_sub(self, idx):
        col_b_value = self.sheet.cell_value(idx, 1)
        if col_b_value != "":
            return col_b_value
        else:
            return False

    def get_fields(self, index):
        # iterate over fields until a field evaluates to category
        fields = []
        index = index + 1
        while(index != self.last_row):
            cell_value = self.sheet.cell_value(index, 2)
            if self.is_category(index) is True:
                break
            elif cell_value != "":
                fields.append(cell_value.strip())
            index = index + 1

        return fields

    def is_category(self, idx):
        return True if self.get_date(self.sheet.cell_value(idx, 3)) else False

    def tests(self):
        print(self.is_category(3))
        print(self.get_date(43010.0))

    def parse_data(self):
        category_schema = []
        for index, value in enumerate(self.sheet.col_values(2)):
            temp = next(
                (item for item in category_schema if item['name'] == value), None)
            if self.is_category(index):
                if temp is None:
                    temp = {
                        "name": value,
                        "fields": self.get_fields(index),
                        "subsets": [["All", index]],
                        "start_date": self.get_date(self.sheet.cell_value(index, 3), is_first_of_month=True),
                        "end_date": self.get_date(self.sheet.cell_value(index, self.last_col)),
                        "data": []
                    }
                    category_schema.append(temp)
                else:
                    col_b_value = self.has_sub(index)
                    if col_b_value:
                        temp["subsets"].append([col_b_value, index])
        return category_schema


if __name__ == "__main__":
    d = Read_WB().parse_data()
    print(d)
