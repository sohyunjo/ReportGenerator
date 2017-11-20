from __future__ import division
import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
from openpyxl import load_workbook
import pyodbc
import sys



"""
This script creates an excel file of items sold with its metadata. The data is pulled from ODBC database used by ECRS.

Each column is selected in a specific order to organize the data so it is user friendly for management in excel format. The margin column is calculated from the data and is populated in the script to display profit. The Excel file will include worksheets called "summary", "all" and by each department. The function of each worksheet is explained in the code.

Each worksheet except for "summary" will have the following columns
A: Department
B: Supplier
C: UPC
D: Brad
E: Receipt Alias
F: Size 
G: Quantity Sold
H: Sales (Total sales of the item)
I: Last Cost
J: Base Price
K: Margin

@param start - the starting date of data input for the report
@param end - the ending date of data input for the report
@param store - which store location the data is from.
@param dept - an array that contains department name in order.

"""
def main(start, end, store,dept):
    # global variables
    global STARTDATE
    global ENDDATE

    # Setting up variables
    STARTDATE = '\'' + start + '\''
    ENDDATE = '\'' + end + '\''

    # connecting to the ODBC
    cn = pyodbc.connect(r'DSN=;UID=;PWD=')
    cursor = cn.cursor()

    # consistent file name
    fileName = 'ItemMovement_'+store+'_' + \
        STARTDATE.replace('/', '').replace('\'','') + '_' + ENDDATE.replace('/', '').replace('\'','') + '.xlsx'


    # Setting up the writer for excel file
    writer = pd.ExcelWriter(fileName)


    
    """
    Preparing for the SUMMARY and ALL worksheet

    SUMMARY: Displays the sales of each department in dollars and the percentage column depicts the department sales compared to total sales. Note that the department name includes department number and not all department will be shown depending on sale between the timeframe.

    ALL: Displays the sales of every item sold with its respective metadata. 
    
    """

    # passing -1 for getQuery not to specify the department
    data = pd.read_sql(getQuery(-1, store), cn)
    df = appendMargin(data)

    # summing up the sales by department
    summary = df.groupby('DEPT')['Sales'].sum().reset_index()
    # Total Sales
    salesTotal = summary.sum(numeric_only=True)[0]
    # calculating the department sales out of total sales. Appending the column at the end.
    summary['%'] = summary.Sales / salesTotal 

    # writing the SUMMARY worksheet to the excel file
    summary.to_excel(writer, 'SUMMARY', index=False)
    # writing the ALL worksheet to the excel file
    df.to_excel(writer, 'ALL', index=False)
    # formatting columns 
    formatColumns(writer, writer.sheets['SUMMARY'])
    formatColumns(writer, writer.sheets['ALL'])


    """ 
    Worksheet By Department 

    """
    for x in range(0, len(dept)):

        data = pd.read_sql(getQuery(x, store), cn)
        if(not data.empty):
            df = appendMargin(data)
            if('/' in dept[x]):
                dept[x] = dept[x].replace('/','')
            df.to_excel(writer, dept[x], index=False)
            worksheet = writer.sheets[dept[x]]
            formatColumns(writer, worksheet)

    
    
    # save
    writer.save()

    # closing connection
    cn.close()

    return 0

# end of main


"""

@param data is the original data fetched from database.
@return the dataframe with margin column appended at the end.

Converting the raw data into pandas dataframe.
Calculate margin from the existing data such as base price and last cost.
Once we have the list of margin, append to the existing dataframe at the end and return.

"""
def appendMargin(data):
    # create dataframe with the fetched data from the database
    pdData = pd.DataFrame(data)
    # sorting the data with quantity sold so that the top selling products will be placed from the top
    pdData = pdData.sort_values(by=['QTYSOLD'], ascending=[False])
    
    # if the data is not empty, calculate margin with Base Price 
    if(not pdData.empty):
        # create an array for margin
        margins = []
        # for each row, calculate.
        for x in range(0, pdData.shape[0]):
            # Base Price: how much we sell at the store
            baseP = pdData.iloc[x]['BasePrice']
            # Last Cost: how much we get from the supplier
            lastC = pdData.iloc[x]['LastCost']

            # If there is Base Price, calculate margin and append to the list.
            # If not, just append 0 for margin.
            # 100% margin means there was no Last Cost input. 
            if(baseP > 0):
                margin = (baseP - lastC) / (baseP) 
                margins.append(margin)
            else:
                margins.append(0)
        # Convert the list to the series so that we can get 
        se = pd.Series(margins)
        # creating the margin column at the end of dataframe
        pdData['Margin'] = se.values
        margins = []

    # return the dataframe
    return pdData


"""
Formatting columns to be user-friendly for $ and % values.

@param writer is the excel writer
@param worksheet is the worksheet of excel file
@return modified worksheet with correctly formatted columns

"""
def formatColumns(writer, worksheet):

    workbook = writer.book
    # 2 digits after period shown for $ amount
    monFormat = workbook.add_format({'num_format': '$0.00'})
    # moving the decimal value to show percentage format
    pctFormat = workbook.add_format({'num_format':'0.0%'})

    if worksheet.get_name() is 'SUMMARY':
        worksheet.set_column('B:B', None, monFormat)
        worksheet.set_column('C:C', None, pctFormat)
    else:
        worksheet.set_column('H:H', None, monFormat)
        worksheet.set_column('I:I', None, monFormat)
        worksheet.set_column('J:J', None, monFormat)
        worksheet.set_column('K:K', None, pctFormat)

    return worksheet



"""
Creates a tailored SQL query based on the parameter passed in.

@param num is for department number. -1 means no department. Meaning, Summary and ALL worksheet.
@param store to specify store number.
@return allsql is sql query

"""
def getQuery(num,store):
    deptStr = ''

    # make sure it is not -1 (summary)
    if(num >= 0):
        # 18th department has item id of 20 in the system. Since we are adding 1 below, add only 1 for now.
        if(num == 18):
            num = num + 1
        #num is from array, however, the department number starts with 1, not 0. So add 1.
        deptStr = 'v.DPT_Number like ' + str(num + 1) + ' and '

    allsql = 'select v.DPT_Name DEPT, m.ven_companyname Supplier, v.Scancode UPC, m.brd_name Brand, v.ReceiptAlias ReceiptAlias,   m.inv_size ItemSize, sum(v.SIT_Quantity) QTYSOLD, sum(v.SIT_Amount) Sales, m.inv_lastcost LastCost , m.sib_baseprice BasePrice from v_SummaryItems v LEFT OUTER JOIN  v_InventoryMaster m ON m.inv_scancode = v.Scancode  and m.STO_Number =v.STO_Number where ' + deptStr + 'v.STO_Number like \''+store+'%\' and v.ISG_StartTime between ' + STARTDATE + ' and ' + ENDDATE + \
        ' GROUP BY v.Scancode, v.DPT_Name,  v.STO_Number , m.brd_name, v.ReceiptAlias ,m.inv_lastcost,m.ven_companyname,m.sib_baseprice,m.inv_size'

    return allsql


def trimAllColumns(df):
    """
    Trim whitespace from ends of each value across all series in dataframe
    """
    trimStrings = lambda x: x.strip(' ][\'') if type(x) is str else x
    return df.applymap(trimStrings)


if __name__ == '__main__':
    arg1 = sys.argv[1]
    arg2 = sys.argv[2]
    arg3 = sys.argv[3]
    arg4 = sys.argv[4]
    main(arg1, arg2, arg3, arg4)
