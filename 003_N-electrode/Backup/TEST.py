import logging
import shutil
import glob
import xlrd
import openpyxl  # 用來處理 Excel 文件
import os
import sys
from datetime import datetime, date


########## 自作関数の定義 ##########
sys.path.append('../MyModule')
import SQL
import Log
import Convert_Date

########## ログの設定 ##########
Log_FolderName = str(date.today())

if not os.path.exists("../Log/" + Log_FolderName):
    os.makedirs("../Log/" + Log_FolderName)

Log_file = '../Log/' + Log_FolderName + '/003_N-electrode.log'
Log.Log_Info(Log_file, 'Program Start')

########## Primeに接続し、全てのserial_numberを取得 ##########
import sys

def fetch_all_serial_numbers():
    try:
        # 連接到資料庫
        conn, cursor = SQL.connSQL()

        # 檢查是否成功連接到資料庫
        if conn is None:
            Log.Log_Error(Log_file, 'Connection with Prime Failed\n')
            sys.exit()

        # 查詢所有的 serial_number
        query = "select top 50 * from prime.v_LotStatus;" 
        query = "select ProductName, ContainerName from prime.v_LotStatus where ContainerName Like '____" + 'D1445'+ "'"
        cursor.execute(query)

        # 獲取所有查詢結果
        serial_numbers = cursor.fetchall()

        # 打印出所有的 serial_number
        for serial_number in serial_numbers:
            print(serial_number[0])

    except Exception as e:
        # 捕捉錯誤並記錄到日誌中
        Log.Log_Error(Log_file, f"SQL Query Failed: {e}")
    finally:
        # 確保無論成功或失敗都關閉資料庫連接
        if conn:
            SQL.disconnSQL(conn, cursor)

# 調用函數來撈取所有的 serial_number
fetch_all_serial_numbers()
