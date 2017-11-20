# ReportGenerator
Creates Sale Reports using pandas framework from OBDC database used by ECRS.

# WeeklyReportGenerator.py

This script pull store and department information pulled from ODBC database used by ECRS.
Then, it uses the main method from WeeklyMovementReport.py to generate reports by each store and department in Excel file. This script is used to create files for all stores.


# WeeklyMovementReport.py
This script creates an excel file of items sold with its metadata. The data is pulled from ODBC database used by ECRS.

Each column is selected in a specific order to organize the data so it is user friendly for management in excel format. The margin column is calculated from the data and is populated in the script to display profit. The Excel file will include worksheets called "summary", "all" and by each department. The function of each worksheet is explained in the code.
This script is used to create a file for one store.

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




# How to run - 

Download both files and install necessary frameworks.

To create files for all stores:

Run with:
python WeeklyReportGenerator.py startDate endDate

For example:
python WeeklyReportGenerator.py 20170115 20170120



To create a file for a store:

Run with:
python WeeklyMovementReport.py startDate endDate storeID deptArr

For example:
python WeeklyMovementReport.py 20170115 20170120 RS1 ["01 Grocery", "02 Frozen", ..., "18 BULK"]


