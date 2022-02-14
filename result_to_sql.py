###############################################################################
### Store the result to MySQL database                                      ###
### Module to install: mysql_connector                                      ###
###############################################################################

import mysql.connector

MYSQL_DATABASE = 'Richard'   # TBR_Battery_Test
MYSQL_TABLE = 'Total'

def result_to_sql(table):
    mydb = mysql.connector.connect(
        host = '192.168.1.84',
        user = 'richard',
        password = 'richardtbts',
        database = MYSQL_DATABASE
    )
    mycursor = mydb.cursor()

    fieldlist = table.columns
    for i in range(len(table)):
        datalist = table.loc[i].values

        sql = "INSERT INTO " + MYSQL_TABLE + "("
        for j in range(len(fieldlist)):
            sql = sql + fieldlist[j] + ", "
        sql = sql[0:-2] + ") VALUES ("
        for j in range(len(datalist)):
            if datalist[j] == None:     # if null
                sql = sql + "Null, "            
            else:
                sql = sql + "'" + str(datalist[j]) + "', "
        sql = sql[0:-2] + ")"
       
        mycursor.execute(sql)
    
    mydb.commit()