{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": 67,
   "metadata": {},
   "outputs": [],
   "source": [
    "import sys\n",
    "import xlrd\n",
    "from datetime import datetime"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 68,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "<xlrd.sheet.Sheet at 0x1810e7dc748>"
      ]
     },
     "execution_count": 68,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "wb = xlrd.open_workbook(\"docs/Demo_Assessment_Model_08.18.20.xlsx\")\n",
    "sheet = wb.sheet_by_name(\"KPI Dashboard\")\n",
    "sheet"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "#### To get the categories, search column D for a valid date and return the index. That returned index can then be used to get top level categories in col C."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 80,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "[4, 22, 30, 38, 57, 66, 74, 82, 102, 110, 118]"
      ]
     },
     "execution_count": 80,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "ref = {}\n",
    "for col in range(0, 3):\n",
    "    temp = {}\n",
    "    for index, val in enumerate(sheet.col_values(col + 1)):\n",
    "        if val != \"\":\n",
    "            temp[index] = val\n",
    "    ref[col] = temp\n",
    "    \n",
    "def isDate(d):\n",
    "    try:\n",
    "        date = datetime(*xlrd.xldate_as_tuple(d, wb.datemode))\\\n",
    "            .strftime(\"%Y-%m-%d %H:%M:%S\")\n",
    "        datetime.strptime(date, \"%Y-%m-31 00:00:00\")\n",
    "        return True\n",
    "    except ValueError as ve:\n",
    "        return False\n",
    "    except TypeError as te:\n",
    "        return False\n",
    "dates_rows = []\n",
    "for i, v in enumerate(sheet.col_values(3)):\n",
    "    if isDate(v):\n",
    "        dates_rows.append(i)\n",
    "dates_rows"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 81,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "['Summary Financial Metrics',\n",
       " 'Customer Metrics',\n",
       " 'New Bookings ',\n",
       " 'MRR/ARR by Month',\n",
       " 'Customer Unit Economics',\n",
       " 'Customer Metrics',\n",
       " 'New Bookings ',\n",
       " 'MRR/ARR by Month',\n",
       " 'Customer Metrics',\n",
       " 'New Bookings ',\n",
       " 'MRR/ARR by Month']"
      ]
     },
     "execution_count": 81,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "cat_values = sheet.col_values(2)\n",
    "categories = [cat_values[r] for r in dates_rows]\n",
    "categories"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "#### If there is a string in col B, it means that category is a subcategory"
   ]
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.6.4"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 4
}
