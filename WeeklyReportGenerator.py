from __future__ import division
import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
from openpyxl import load_workbook
import WeeklyMovementReport as ItemMove
import pyodbc
import sys


"""
This script pull store and department information pulled from ODBC database used by ECRS.
Then, it uses the main method from WeeklyMovementReport.py to generate reports by each store and department in Excel file. 


@param start - the starting date of data input for the report. For example, 20170115
@param end - the ending date of data input for the report. For example, 20170120

How to run - 

python WeeklyReportGenerator.py start end

For example:
python WeeklyReportGenerator.py 20170115 20170120

"""

def main(start, end):

    # Connecting to the database
    cn = pyodbc.connect(r'DSN=;UID=;PWD=')
    cursor = cn.cursor()

    # fetching the list of stores 
    storeQuery = 'select sto_number from Stores'
    storeData = ItemMove.trimAllColumns(pd.read_sql(storeQuery, cn))
    # convert it into dataframe
    store = pd.DataFrame(storeData).iloc[:,0].tolist()

    # fetching the list of department
    deptQuery = 'select dpt_name from Departments'
    deptData = ItemMove.trimAllColumns(pd.read_sql(deptQuery, cn))
    # convert it into dataframe
    dept = pd.DataFrame(deptData).iloc[:,0].tolist()

    # make sure to end the connection when finished querying
    cn.close()


    # Create a weekly report file for each store
    # Start with 1 since the first index is HQ. No sell information in HQ.
    for x in range(1, len(store)):
        ItemMove.main(start, end, store[x], dept)

    # nothing to return
    return 0


"""
arg1 startDate
arg2 endDate
"""
if __name__ == '__main__':
    arg1 = sys.argv[1]
    arg2 = sys.argv[2]
    main(arg1, arg2)
