tbts_grading is the grading program for Mxx.csv files that tested by TBTS.
This program generate the following output result files:
  1. testnumber-(xx).html    The DCD chart in html format
  2. testnumber-(xx).xlsx    The grading result table and printable result page
This program also save the result table in the MySQL database and print out
the result page for factory operation.

The files contained in this program:

TBTS_Grading.py     Main program with grading algorithm
defines.py          defines of the battery model and the color fill for cells
result_to_chart.py  Generate and save the DCD chart html file
result_to_excel.py  Generate and save the result table and printable result
                    page in excel file 
result_to_sql.py    Save the result table to MySQL database
printout.py         Printout the result page to the active printer
README.md           This file

Versions:

V22.0213.01
Date: 2022/02/13
 1. Initial Release for Python that is modified from VBA TBGR program
