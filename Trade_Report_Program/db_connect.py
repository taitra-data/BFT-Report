# MS SQL 2019 document:
# https://docs.microsoft.com/en-us/sql/connect/python/python-driver-for-sql-server?view=sql-server-ver15

"""
The functions for connections to iTrade analytical Database (MS SQL Server).
"""

from sqlite3 import Error
from secrets import *
import pyodbc


def create_connection():
    """ create a database connection to the iTrade analytical database
    :return: Connection object or None
    """
    cnxn = None
    try:
        server = '172.20.70.119'
        database = 'itrade_original'
        # username = ***  (update your username in "secrets.py" file)
        # password = ***  (update your password in "secrets.py" file)
        cnxn = pyodbc.connect(
            'DRIVER={ODBC Driver 17 for SQL Server};'
            'SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
        return cnxn

    except Error as e:
        print(e)
        return cnxn

def get_itrade_db_data(cnxn, select_table_sql):
    """ extract data from iTrade analytical DB.
    :param cnxn: SQL Server Connection object
    :param select_table_sql: a SELECT TABLE statement
    :return: selected data
    """

    try:
        cursor = cnxn.cursor()
        cursor.execute(select_table_sql)

    except Error as e:
        print(e)

    lst_data = [list(row) for row in cursor]
    num = (len(lst_data))
    cursor.close()
    print('Successfully extract', num, 'rows of data from DB')
    return lst_data


if __name__ == '__main__':
    # Test whether the functions can work successfully.

    year = 2021
    select_table_sql = '''
          SELECT TOP (50) *
          FROM [itrade_original].[dbo].[MOF_DATA_TXT]
          where year > {} 
          ORDER BY HSCode;
    '''.format(str(year))
    cnxn = create_connection()
    result = get_itrade_db_data(cnxn, select_table_sql)

    for i in result:
        print(i)

    cnxn.close()
