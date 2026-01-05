import os
import sys
import xlrd
import glob
import pyodbc
import shutil
import logging
import numpy as np
import pandas as pd
import openpyxl as px

from time import strftime, localtime
from datetime import date, timedelta, datetime
from dateutil.relativedelta import relativedelta


########## 自作関数の定義 ##########
sys.path.append('../MyModule')
import Log
import SQL
import Check
import Convert_Date
import Row_Number_Func


########## 全体パラメータ定義 ##########
Site = '350'
ProductFamily= 'SAG FAB'
Operation = 'PIX_SiN_Thickness'
TestStation = 'PIX'


########## Logの設定 ##########

# ----- ログファイルの作成 -----
Log_Folder_Name = str(date.today())
if not os.path.exists("../Log/" + Log_Folder_Name):
    os.makedirs("../Log/" + Log_Folder_Name)

Log_File = '../Log/' + Log_Folder_Name + '/026_PIX.log'
Log.Log_Info(Log_File, 'Program Start')


########## シート名の定義 ##########
Data_Sheet_Name = '相模原(PD220NL)'


########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
# Output_filepath = '../XML/026_PIX/'


########## 取得するデータの列番号を定義 ##########
Col_Start_Date_Time = 0
Col_Operator = 1
Col_Serial_Number = 2
Col_SiN_Thickness = 3
Col_SiN_Refraction = 4


########## 取得した項目と型の対応表を定義 ##########
key_type = {
    'key_Start_Date_Time' : str,
    'key_Operator' : str,
    'key_Serial_Number' : str,
    'key_Part_Number' : str,
    'key_SiN_Thickness' : float,
    'key_SiN_Refraction' : float,
    'key_STARTTIME_SORTED' : float,
    'key_SORTNUMBER' : float,
    "key_LotNumber_9": str
}


def Main():


    ########## Excelファイルをローカルにコピー ##########

    # ----- 正規表現で取出し、直近で変更があったファイルを取得する -----
    Log.Log_Info(Log_File, 'Excel File Copy')

    FilePath = 'Z:/CVD/P-CVD/'
    FileName = '*P-CVD着工記録*.xls*'
    Excel_File_List = []
    for file in glob.glob(FilePath + FileName):
        if '$' not in file:
            dt = strftime("%Y-%m-%d %H:%M:%S", localtime(os.path.getmtime(file)))
            Excel_File_List.append([file, dt])

    # ----- dt(更新日時)の降順で並び替える -----
    Excel_File_List = sorted(Excel_File_List, key=lambda x: x[1], reverse=True)
    Excel_File = shutil.copy(Excel_File_List[0][0], '../DataFile/026_PIX/')


    ########## DaraFrameの作成 ##########

    # ----- 取得開始行の取り出し -----
    Log.Log_Info(Log_File, 'Get The Starting Row Count')
    Start_Number = Row_Number_Func.start_row_number("PIX_StartRow.txt") - 500

    # ----- ExcelデータをDataFrameとして取得 -----
    Log.Log_Info(Log_File, 'Read Excel')
    df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="B:AU", skiprows=Start_Number)

    # ----- 必要なデータのみ取得 -----
    df = df.iloc[:, [0,1,5,26,29]]

    # ----- 列番号の振り直し -----
    Log.Log_Info(Log_File, 'Setting Columns Number')
    df.columns = range(df.shape[1])

    # ----- 末尾から欠損のデータを落としていく -----
    Getting_Row = len(df) - 1
    while Getting_Row >= 0 and df.isnull().any(axis=1)[Getting_Row]:
        Getting_Row -= 1

    df = df[:Getting_Row + 1]

    # ----- 次の開始行数をメモ -----
    Next_Start_Row = Start_Number + df.shape[0] + 1

    # ----- 日付欄に文字列が入っていたらNoneに置き換える -----
    for i in range(df.shape[0]):
        if not isinstance(df.iloc[i, 0], (pd.Timestamp, datetime)):
            df.iloc[i, 0] = np.nan

    # ----- 今日から1か月前のデータまでを取得する -----
    df[0] = pd.to_datetime(df[0])
    one_month_ago = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) - relativedelta(months=1)
    df = df[(df[0] >= one_month_ago)]


    ########## ループ前の設定 ##########

    # ----- 最終行数の取得 -----
    Log.Log_Info(Log_File, 'Get DataFrame End Index Number\n')
    row_end = len(df)

    # ----- 現在処理を行っている行数の定義 -----
    Row_Number = 0

    # ----- dfのindexリスト -----
    df_idx = df.index.values


    ########## データの取得 ##########

    while Row_Number < row_end:


        ########## 空欄判定 ##########

        # ----- 空欄が1つでも存在すればTrueを返す -----
        Log.Log_Info(Log_File, "Blank Check")
        if df.isnull().any(axis=1)[df_idx[Row_Number]]:
            Log.Log_Error(Log_File, "Blank Error\n")
            Row_Number += 1
            continue


        ########## 現在処理を行っている行のデータの取得 ##########

        # ----- 取得したデータを格納するデータ構造(辞書)を作成 -----
        Log.Log_Info(Log_File, 'Data Acquisition')
        data_dict = dict()

        # ----- ロット番号を取得 -----
        Serial_Number = str(df.iloc[Row_Number, Col_Serial_Number])
        if Serial_Number == "nan":
            Log.Log_Error(Log_File, "Lot Error\n")
            Row_Number += 1
            continue

        # ----- Primeに接続し、ロット番号に対応する品名を取り出す -----
        conn, cursor = SQL.connSQL()
        if conn is None:
            Log.Log_Error(Log_File, Serial_Number + ' : ' + 'Connection with Prime Failed')
            break
        Part_Number, Nine_Serial_Number = SQL.selectSQL(cursor, Serial_Number)
        SQL.disconnSQL(conn, cursor)

        # ----- 品名が None であれば処理を行わない -----
        if Part_Number is None:
            Log.Log_Error(Log_File, Serial_Number + ' : ' + "PartNumber Error\n")
            Row_Number += 1
            continue

        # ----- 品名が LDアレイ_ であれば処理を行わない -----
        if Part_Number == 'LDアレイ_':
            Row_Number += 1
            continue

        # ----- データの取得 -----
        data_dict = {
            'key_Start_Date_Time': df.iloc[Row_Number, Col_Start_Date_Time],
            'key_Operator': df.iloc[Row_Number, Col_Operator],
            'key_Serial_Number': Serial_Number,
            'key_Part_Number': Part_Number,
            "key_LotNumber_9": Nine_Serial_Number,
            'key_SiN_Thickness': df.iloc[Row_Number, Col_SiN_Thickness],
            'key_SiN_Refraction': df.iloc[Row_Number, Col_SiN_Refraction]
        }


        ########## 日付フォーマットの変換 ##########

        # ----- 日付を指定されたフォーマットに変換する -----
        Log.Log_Info(Log_File, 'Date Format Conversion')
        data_dict["key_Start_Date_Time"] = Convert_Date.Edit_Date(data_dict["key_Start_Date_Time"])

        # ----- 指定したフォーマットに変換出来たか確認 -----
        if len(data_dict["key_Start_Date_Time"]) != 19:
            Log.Log_Error(Log_File, data_dict["key_Serial_Number"] + ' : ' + "Date Error\n")
            Row_Number += 1
            continue


        ########## STARTTIME_SORTEDの追加 ##########

        # ----- 日付をExcel時間に変換する -----
        date = datetime.strptime(str(data_dict["key_Start_Date_Time"]).replace('T', ' ').replace('.', ':'), "%Y-%m-%d %H:%M:%S")
        date_excel_number = int(str(date - datetime(1899, 12, 30)).split()[0])

        # 行数を取得し、[行数/10^6]を行う
        excel_row = Start_Number + df_idx[Row_Number] + 1
        excel_row_div = excel_row / 10 ** 6

        # unix_timeに上記の値を加算する
        date_excel_number += excel_row_div

        # data_dictに登録する
        data_dict["key_STARTTIME_SORTED"] = date_excel_number
        data_dict["key_SORTNUMBER"] = excel_row


        ########## データ型の確認 ##########

        # ----- 数値データに入る箇所に文字列が入っていないか確認する -----
        Log.Log_Info(Log_File, "Check Data Type")
        Result = Check.Data_Type(key_type, data_dict)
        if Result == False:
            Log.Log_Error(Log_File, data_dict["key_Serial_Number"] + ' : ' + "Data Error\n")
            Row_Number += 1
            continue


        ########## XMLファイルの作成 ##########

        # ----- 保存するファイル名を定義 -----
        XML_File_Name = 'Site=' + Site + ',ProductFamily=' + ProductFamily + ',Operation=' + Operation + \
                        ',Partnumber=' + data_dict["key_Part_Number"] + ',Serialnumber=' + data_dict["key_Serial_Number"] + \
                        ',Testdate=' + data_dict["key_Start_Date_Time"] + '.xml'

        # ----- XMLファイルの作成 -----
        Log.Log_Info(Log_File, 'Excel File To XML File Conversion')

        f = open(Output_filepath + XML_File_Name, 'w', encoding="utf-8")

        f.write('<?xml version="1.0" encoding="utf-8"?>' + '\n' +
                '<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' + '\n' +
                '       <Result startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Result="Passed">' + '\n' +
                '               <Header SerialNumber=' + '"' + data_dict["key_Serial_Number"] + '"' + ' PartNumber=' + '"' + data_dict["key_Part_Number"] + '"' + ' Operation=' + '"' + Operation + '"' + ' TestStation=' + '"' + TestStation + '"' + ' Operator=' + '"' + data_dict["key_Operator"] + '"' + ' StartTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Site=' + '"' + Site + '"' + ' LotNumber=' + '"' + data_dict["key_Serial_Number"] + '"/>' + '\n' +
                '\n'
                '               <TestStep Name="SiN" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Thickness" Units="nm" Value=' + '"' + str(data_dict["key_SiN_Thickness"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Refraction" Units="percent" Value=' + '"' + str(data_dict["key_SiN_Refraction"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '\n'
                '               <TestStep Name="SORTED_DATA" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value=' + '"' + str(data_dict["key_STARTTIME_SORTED"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value=' + '"' + str(data_dict["key_SORTNUMBER"]) + '"/>' + '\n' +
                '                   <Data DataType="String" Name="LotNumber_5" Value=' + '"' + str(data_dict["key_Serial_Number"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                '                   <Data DataType="String" Name="LotNumber_9" Value=' + '"' + str(data_dict["key_LotNumber_9"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '\n'
                '               <TestEquipment>' + '\n' +
                '                   <Item DeviceName="P-CVD" DeviceSerialNumber="#1"/>' + '\n' +
                '               </TestEquipment>' + '\n' +
                '\n'
                '               <ErrorData/>' + '\n' +
                '               <FailureData/>' + '\n' +
                '               <Configuration/>' + '\n' +
                '       </Result>' + '\n' +
                '</Results>'
                )
        f.close()


        ########## XML変換完了時の処理 ##########

        Log.Log_Info(Log_File, data_dict["key_Serial_Number"] + ' : ' + "OK\n")
        Row_Number += 1


    ########## 次の開始行数の書き込み ##########
    Log.Log_Info(Log_File, 'Write the next starting line number')
    Row_Number_Func.next_start_row_number("PIX_StartRow.txt", Next_Start_Row)

    # ----- 最終行を書き込んだファイルをGドライブに転送 -----
    shutil.copy("PIX_StartRow.txt", 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/026_PIX/13_ProgramUsedFile/')


if __name__ == '__main__':

    Main()

# ----- ログ書込：Main処理の終了 -----
Log.Log_Info(Log_File, 'Program End')