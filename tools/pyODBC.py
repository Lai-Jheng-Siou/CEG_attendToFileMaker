import pyodbc

from custTool import getConfig

def connect_odbc():
    config = getConfig()
    server = config['odbc']['server']
    database = config['odbc']['database']
    uid = config['odbc']['uid']
    pasd = config['odbc']['pasd']

    # ODBC 連接字串
    connection_string = f"Driver={{FileMaker ODBC}}; Server={server}; Database={database}; UID={uid}; PWD={pasd}"

    conn = pyodbc.connect(connection_string)

    return conn