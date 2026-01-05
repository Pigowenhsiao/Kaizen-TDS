import pyodbc

def connect_sql_server(driver, server, database, username, password):
    """
    建立與 SQL Server 的連接。
    """
    try:
        connection_string = (
            f"DRIVER={{{driver}}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"UID={username};"
            f"PWD={password}"
        )
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        print("連接到 SQL Server 成功！")
        return conn, cursor
    except pyodbc.Error as err:
        print("連接到 SQL Server 失敗。")
        print(f"錯誤訊息: {err}")
        return None, None

def list_tables(cursor, schema='prime'):
    """
    列出指定 schema 下的所有資料表。
    """
    try:
        query = """
            SELECT TABLE_SCHEMA, TABLE_NAME
            FROM INFORMATION_SCHEMA.TABLES
            WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_SCHEMA = ?
            ORDER BY TABLE_NAME
        """
        cursor.execute(query, schema)
        tables = cursor.fetchall()
        return tables
    except pyodbc.Error as err:
        print("查詢資料表時發生錯誤。")
        print(f"錯誤訊息: {err}")
        return []

def list_views(cursor, schema='prime'):
    """
    列出指定 schema 下的所有視圖。
    """
    try:
        query = """
            SELECT TABLE_SCHEMA, TABLE_NAME
            FROM INFORMATION_SCHEMA.VIEWS
            WHERE TABLE_SCHEMA = ?
            ORDER BY TABLE_NAME
        """
        cursor.execute(query, schema)
        views = cursor.fetchall()
        return views
    except pyodbc.Error as err:
        print("查詢視圖時發生錯誤。")
        print(f"錯誤訊息: {err}")
        return []

def main():
    # 連接參數
    driver = 'ODBC Driver 17 for SQL Server'  # 根據實際情況調整
    server = '192.168.117.140'
    database = 'PrimeProd'
    username = 'prime-mfg'
    password = 'manufacturing'
    schema = 'prime'  # 根據實際情況調整 schema 名稱

    # 建立連接
    conn, cursor = connect_sql_server(driver, server, database, username, password)
    if not conn or not cursor:
        return

    # 列出資料表
    tables = list_tables(cursor, schema)
    print(f"\n資料庫 '{database}' 中 schema '{schema}' 下的資料表列表:")
    if tables:
        for table in tables:
            print(f"- {table.TABLE_SCHEMA}.{table.TABLE_NAME}")
    else:
        print("未找到任何資料表。")

    # 列出視圖
    views = list_views(cursor, schema)
    print(f"\n資料庫 '{database}' 中 schema '{schema}' 下的視圖列表:")
    if views:
        for view in views:
            print(f"- {view.TABLE_SCHEMA}.{view.TABLE_NAME}")
    else:
        print("未找到任何視圖。")

    # 關閉連接
    cursor.close()
    conn.close()
    print("\n已關閉與 SQL Server 的連接。")

if __name__ == "__main__":
    main()
