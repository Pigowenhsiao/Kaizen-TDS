import sys
sys.path.append('../MyModule')
import SQL

def main():
    conn, cursor = SQL.connSQL()
    if conn is None:
        print("DB connection failed.")
        return

    try:
        cursor.execute("select distinct ProductFamilyName from prime.v_TransactionData where ProductName like 'HL13B5-BFL%';")
        rows = cursor.fetchall()
        print("Distinct ProductFamilyName from prime.v_TransactionData:")
        for row in rows:
            print(row[0])
    except Exception as e:
        print(f"Error: {e}")
    finally:
        SQL.disconnSQL(conn, cursor)

if __name__ == "__main__":
    main()
