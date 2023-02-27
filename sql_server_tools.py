import pyodbc


def sql_server_trusted_conn_insert_stmt(server_name, database_name, comm_str):
    conn = pyodbc.connect(f'Driver={{SQL Server Native Client 11.0}}; Server={server_name}; Database={database_name}; Trusted_Connection=yes;', autocommit=True, timeout=0)
    cursor = conn.cursor()
    cursor.execute(comm_str)
    print('insert')



