from datetime import datetime
import sys

import xlrd
from pprint import pprint


wb = xlrd.open_workbook("docs/Demo_Assessment_Model_08.18.20.xlsx")
sheet = wb.sheet_by_name("KPI Dashboard")
last_col = sheet.ncols - 1
last_row = sheet.nrows - 1
print(last_row)


def get_date(d):
    try:
        date = datetime(*xlrd.xldate_as_tuple(d, wb.datemode))\
            .strftime("%Y-%m-%d %H:%M:%S")
        datetime.strptime(date, "%Y-%m-31 00:00:00")
        return date
    except ValueError:
        return False
    except TypeError:
        return False


def is_category(idx):
    return True if get_date(sheet.cell_value(idx, 3)) else False


def has_sub(cell):
    col_b_value = sheet.cell_value(index, 1)
    if col_b_value != "":
        return col_b_value
    else:
        return False


def get_fields(index):
    # iterate over fields until a field evaluates to category
    fields = []
    index = index + 1
    while(index != last_row):
        cell_value = sheet.cell_value(index, 2)
        print(cell_value)
        if is_category(index) is True:
            break
        elif cell_value != "":
            fields.append(cell_value.strip())
        index = index + 1

    return fields


category_schema = []
for index, value in enumerate(sheet.col_values(2)):
    temp = next(
        (item for item in category_schema if item['name'] == value), None)
    if is_category(index):
        if temp is None:
            temp = {
                "name": value,
                "fields": get_fields(index),
                "subsets": ["all"],
                "start_date": get_date(sheet.cell_value(index, 3)),
                "end_date": get_date(sheet.cell_value(index, last_col))
            }
        col_b_value = has_sub(value)
        if col_b_value:
            temp["subsets"].append(col_b_value)
        category_schema.append(temp)

pprint(category_schema)


# if __name__ == "__main__":
#     import random
#     choice = input(
#         "What would you like to do?\r\n1. Run Program\r\n2. Run Tests\r\n")
#     if choice == "1":
#         main()
#     elif choice == "2":
#         def get_random_numbers_and_dates():
#             return datetime.date(datetime(random.choice([2018, 2019, 2020]), random.choice(range(1, 13)), random.choice(range(1, 29))))
#         rand = [get_random_numbers_and_dates() for x in range(30)]
#         for d in rand:
#             assert main.get_date(d) is False, "Should be False."
