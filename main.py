from datetime import datetime
import sys

import xlrd
from pprint import pprint

wb = xlrd.open_workbook("docs/Demo_Assessment_Model_08.18.20.xlsx")
sheet = wb.sheet_by_name("KPI Dashboard")
last_col = sheet.ncols - 1
