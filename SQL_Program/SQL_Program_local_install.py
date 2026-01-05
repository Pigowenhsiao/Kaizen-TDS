import pandas as pd
import cx_Oracle
import pyodbc
import glob
import sys
import os

from datetime import date

# 自作関数の定義
sys.path.append('../001_QC/MyModule')
import Log


# ----- windows拡張言語対応[UTF8] -----
os.environ['NLS_LANG'] = 'JAPANESE_JAPAN.AL32UTF8'


# ----- SQLサーバへの接続 -----
cle_conf = {
    "ID" : "readonly",
    "PASS" : "Re0d0n1y",
    "HOST" : "tdsprd08.c0uzg0vwy8aj.ap-northeast-1.rds.amazonaws.com",
    "SID" : "TDSPRD08",
    "PORT" : 1525
}

# ----- SQLサーバとの接続 -----
try:
    tns = cx_Oracle.makedsn( cle_conf["HOST"], cle_conf["PORT"], cle_conf["SID"])
    conn  = cx_Oracle.connect( cle_conf["ID"], cle_conf["PASS"], tns)
    cur = conn.cursor()
    print("Connection:OK")
except Exception as ex:
    print("Connection:NG")
    print(ex)
    sys.exit()

# ----- SQLファイルがある場所をカレントディレクトリとするし、SQLファイルを取得する -----
os.chdir('../SQL/')
sql_file_list = glob.glob('*.sql')


# ----- Logの作成(完了したものを書き込んでいく方式) -----
Log_Folder_Name = "-".join(str(date.today()).split("-"))
if not os.path.exists("../Log/" + Log_Folder_Name):
    os.makedirs("../Log/" + Log_Folder_Name)

Log_File = '../Log/' + Log_Folder_Name + '/SQL_Program.log'
Log.Log_Info(Log_File, "Program Start")


# ----- 全SQLファイルの実行 -----
for sql_file in sql_file_list:

        # ----- SQLファイルの読み込み -----
        with open(sql_file, 'r') as f:
            sql_query = f.read()

        # ----- SQLを実行し、返ってきたテーブルを保存 -----
        df = pd.read_sql_query(sql_query, con=conn)
        print(df)

        # ----- Excel, csv変換 -----
        #df.to_excel('C:\Users\fus86274\Desktop\Kaizen-TDS_bat\test_data\2022-11-18_19' + sql_file.replace('.sql', '.xlsx'), sheet_name='Sheet1', index=False)
        #df.to_csv('C:\Users\fus86274\Desktop\Kaizen-TDS_bat\test_data\2022-11-18_19' + sql_file.replace('.sql', '.csv'), index=False, encoding='cp932')

         ----- Log書き込み -----
        Log.Log_Info(Log_File, sql_file + " : OK")

# ----- SQLサーバとの接続を切る -----
cur.close()
conn.close()
print("Disconnection:OK")

# ----- Log書き込み -----
Log.Log_Info(Log_File, "Program End")