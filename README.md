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
result_to_pdf.py    Save the result to pdf
README.md           This file

Python modules to be installed:
pandas, openpyxl, plotly, mysql_connector, reportlab, pyinstaller

Command line:
TBTS_Grading.py [printer name]


Versions:

V22.0213.01
Date: 2022/02/13
 1. Initial Release for Python that is modified from VBA TBGR program

V22.0215.01
Date: 2022/02/15
 1. TBTS_Grading.py, print the time tag on each tasks
 2. result_to_sql.py, if the test number exist in the SQL database, rename the
    test number to be -A, -B..., and insert the new result

V22.0216.01
Date: 2022/02/16
 1. Fix the bug the 'F' grade is not generated

V22.0225.01
Date: 2022/02/25
 1. Git branch to Result-to-pdf
 2. Remove printout.py
 3. Add result_to_pdf.py
 4. Add colors at terminal outout
 5. Fix the bug the model is '34M' in table

V22.0301.01
Date: 2022/03/01
 1. Add print pdf file in Windows. The command line add one argument for
    setting the printer name to print. If the argument is not specified, 
    it will print to default printer.
 2. When getting the date from csv, need to set the format to yyyy-mm-dd
    by strptime(line2[1][1:], "%m-%d-%Y").strftime("%Y-%m-%d") for correctly
    insert to sql.
