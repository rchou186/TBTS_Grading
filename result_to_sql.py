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

    column_list = table.columns

    # check test number that is already in SQL database
    battery = table.loc[0, 'Battery']           # get the test number from table row 0
    sql = f"SELECT Battery FROM {MYSQL_TABLE} WHERE Battery LIKE '%{battery}%'"
    mycursor.execute(sql)
    myresult = sorted(mycursor.fetchall())      # feach all liked battery test number and sorted
    
    # if previous test number exist
    if len(myresult) != 0:

        # myresult is a list of tuple, myresult[-1][-1][-1] to get the last character of the tuple of the list
        mycharacter = myresult[-1][-1][-1]

        # if the last character is alphabet, change test number with next alphabet
        if mycharacter.isalpha():
            mycharacter = chr(ord(mycharacter) + 1)
            next_battery = battery + '-' + mycharacter

        # if the last character is numirical, change test number start with '-A'
        else:
            next_battery = battery + '-A'
        sql = f"UPDATE {MYSQL_TABLE} SET Battery = '{next_battery}' WHERE Battery = '{battery}'"
        mycursor.execute(sql)
        mydb.commit()
    
    # insert table to SQL database
    for row in range(len(table)):
        
        # get the row data to datalist
        data_list = table.loc[row].values
        sql1 = ""
        sql2 = ""
        for col in range(len(column_list)):
            sql1 = sql1 + column_list[col] + ", "
            if data_list[col] == None:     # if null
                sql2 = sql2 + "Null, "            
            else:
                sql2 = sql2 + "'" + str(data_list[col]) + "', "

        # remove the last comma of ", " in sql1 and sql2
        sql1 = sql1[0:-2]
        sql2 = sql2[0:-2]
        sql = f"INSERT INTO {MYSQL_TABLE} ({sql1}) VALUES ({sql2})"
       
        mycursor.execute(sql)
    
    mydb.commit()