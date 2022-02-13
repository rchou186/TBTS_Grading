###############################################################################
### Store the result to MySQL database                                      ###
### Module to install: mysql_connector                                      ###
###############################################################################

import mysql.connector

def result_to_sql(table):
    #mydb = mysql.connector.connect(
    #    host = '192.168.1.84',
    #    user = 'richard',
    #    password = 'richardtbts',
    #    database = ''
    #)
    #mycursor = mydb.cursor()

    fieldlist = table.columns
    datalist = table.loc[0].values

    sql = 'INSERT INTO Table ('
    for i in range(len(fieldlist)):
        sql = sql + fieldlist[i]+','
    sql=sql[0:-1]+') VALUES ('
    for i in range(len(datalist)):
        sql = sql + "'"+str(datalist[i])+"',"
    sql=sql[0:-1]+')'
    print(sql)

    pass