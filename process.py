from datetime import datetime
from calendar import monthrange
from pprint import pprint
import os
import json

import xlrd


class Parse_WB:
    def __init__(self):
        self.wb = xlrd.open_workbook(
            "docs/Demo_Assessment_Model_08.18.20.xlsx")
        self.sheet = self.wb.sheet_by_name("KPI Dashboard")
        self.last_col = self.sheet.ncols - 1
        self.last_row = self. sheet.nrows - 1

    def get_date(self, d, is_first_of_month=False):
        """This function checks that the serialized dates read from the excel file are
            an actual date and not a fake date parsed from raw values. It does this by verifying the
            format is YYYY-mm-ddTHH:MM:SS.s+z and that the day is either the first or last of the month,
            whichever the function specified.
        """
        try:
            excel_date = datetime(*xlrd.xldate_as_tuple(d, self.wb.datemode))
            converted_date = excel_date.strftime("%Y-%m-%dT%H:%M:%S.000+00:00")
            datetime.strptime(
                converted_date, f"%Y-%m-{monthrange(excel_date.year, excel_date.month)[1]}T00:00:00.000+00:00")
            if is_first_of_month:
                return excel_date.strftime("%Y-%m-01T%H:%M:%S.000+00:00")
            else:
                return converted_date
        except ValueError:
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

    def run_tests(self, test_to_run):
        def test_date_function():
            serialized_dates = [34787, 34762, 37801, 37632, 40772, 39242, 37992, 41863,
                                42793, 33109, 43438, 32231, 43986, 38556, 43562, 34214, 37494,
                                41911, 42291, 41927, 41317, 32297, 34495, 33205, 33317, 40457,
                                38007, 36729, 33153, 39041, 39022]
            for date in serialized_dates:
                assert self.get_date(
                    date, is_first_of_month=False) is False, f"{date} should have returned False."
                print(f"Serialized value {date} passed.")
        if test_to_run == "1":
            print("This is to test whether the function 'get_date()' correctly parses serialized"
                  "dates to a proper datetime AND that it matches the format of: YYYY-mm-ddTHH:MM:SS.s+z"
                  "and that the day is either the first or last of the month.\r\n")
            input("Press enter to continue.")
            test_date_function()

    def parse_data(self):
        def get_category_schema():
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

        def parse_category_values(category_schema):
            for category in category_schema:
                for column in range(3, self.last_col):
                    date_row = self.sheet.cell_value(
                        category["subsets"][0][1], column)
                    temp = {
                        "date": self.get_date(date_row),
                        "values": [],
                    }
                    for field in category["fields"]:
                        temp_values = []
                        for subset in category['subsets']:
                            temp_values.append({
                                "name": field,
                                "subset": subset[0],
                                "value": self.sheet.cell_value(subset[1] + 1, column)
                            })
                        temp["values"].append(temp_values)

                    category["data"].append(temp)
            for category in category_schema:
                category["subsets"] = [subset[0]
                                       for subset in category["subsets"]]
            return category_schema

        category_schema = get_category_schema()
        data = parse_category_values(category_schema)
        return data


if __name__ == "__main__":
    print("\r\nWelcome!\r\n")

    def run_program():
        def get_mode_choice():
            print("Please choose one of the following options:")
            choice = input(
                "1. Parse data from default document.\r\n2. Run tests.\r\n")
            if choice not in ["1", "2"]:
                print("\r\n***Whoops!***")
                print(f"\r\n{choice} is not a valid option, please try again.")
                return get_mode_choice()
            return choice

        def get_test_choice():
            print("Which test would you like to run?\r\n")
            choice = input("1. Test Date Parse Function\r\n")
            if choice not in ["1", "2"]:
                print("\r\n***Whoops!***")
                print(f"\r\n{choice} is not a valid option, please try again.")
                return get_test_choice()
            return choice

        def get_parse_choice():
            print("Would you like to:\r\n")
            choice = input(
                "1. Print data to console\r\n2. Output to JSON\r\n3. Do both\r\n")
            if choice not in ["1", "2", "3"]:
                print("Please pick a valid choice\r\n")
                return get_parse_choice()
            return choice

        def output_to_json(data):
            loc = "data.json"
            with open(loc, "w") as output_file:
                json.dump(data, output_file, indent=4)
                output_file.close()
            print(f"File saved to: {os.path.abspath(loc)}")

        mode_choice = get_mode_choice()
        if mode_choice == "1":
            parse_choice = get_parse_choice()
            data = Parse_WB().parse_data()
            if parse_choice == "1":
                pprint(data)
            elif parse_choice == "2":
                output_to_json(data)
            else:
                pprint(data)
                output_to_json(data)
        elif mode_choice == "2":
            test_choice = get_test_choice()
            Parse_WB().run_tests(test_choice)

        run_again = input(
            "Would you like to run the program again? Y or N? ").upper()
        if run_again in "YES":
            return run_program()

    run_program()
