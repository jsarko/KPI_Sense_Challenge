from datetime import datetime
from calendar import monthrange
import sys

import xlrd
from pprint import pprint


wb = xlrd.open_workbook("docs/Demo_Assessment_Model_08.18.20.xlsx")
sheet = wb.sheet_by_name("KPI Dashboard")
last_col = sheet.ncols - 1
last_row = sheet.nrows - 1


def get_date(d):
    try:
        excel_date = datetime(*xlrd.xldate_as_tuple(d, wb.datemode))
        converted_date = excel_date.strftime("%Y-%m-%dT%H:%M:%S.000+00:00")
        datetime.strptime(
            converted_date, f"%Y-%m-{monthrange(excel_date.year, excel_date.month)[1]}T00:00:00.000+00:00")
        return converted_date
    except ValueError:
        return False
    except TypeError:
        return False


def is_category(idx):
    return True if get_date(sheet.cell_value(idx, 3)) else False


def has_sub(idx):
    col_b_value = sheet.cell_value(idx, 1)
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
                "subsets": [["all", index]],
                "start_date": get_date(sheet.cell_value(index, 3)),
                "end_date": get_date(sheet.cell_value(index, last_col)),
                "data": []
            }
        col_b_value = has_sub(index)
        if col_b_value:
            temp["subsets"].append([col_b_value, index])
        category_schema.append(temp)

for category in category_schema:
    for column in range(3, last_col):
        date_row = sheet.cell_value(category["subsets"][0][1], column)
        temp = {
            "date": get_date(date_row),
            "values": [],
        }
        for index, field in enumerate(category["fields"]):
            temp_values = []
            for subset in category['subsets']:
                temp_values.append({
                    "name": field,
                    "subset": subset[0],
                    "value": sheet.cell_value(subset[1] + 1, column)
                })
                temp["values"].append(temp_values)

        category["data"].append(temp)

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
