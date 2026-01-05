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
from math import isnan


########## 自作関数の定義 ##########
sys.path.append('../MyModule')
import Log
import SQL
import Check
import Convert_Date
import Row_Number_Func


########## 全体パラメータ定義 ##########
Site = '350'
ProductFamily = 'SAG FAB'
Operation = 'ISO-EML_ISO_Step'
TestStation = 'ISO-EML'


########## Logの設定 ##########

# ----- ログファイルの作成 -----
Log_Folder_Name = str(date.today())
if not os.path.exists("../Log/" + Log_Folder_Name):
    os.makedirs("../Log/" + Log_Folder_Name)

Log_File = '../Log/' + Log_Folder_Name + '/024_ISO-EML.log'
Log.Log_Info(Log_File, 'Program Start')


########## シート名の定義 ##########
Data_Sheet_Name = 'ﾃﾞｰﾀ'


########## XML出力先ファイルパス ##########
#Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
# Output_filepath = '../XML/024_ISO-EML/'
Output_filepath = 'C:/Users/hsi67063/Box/00-home-pigo.hsiao/TEMP/XML/'

########## 取得するデータの列番号を定義 ##########
Col_Start_Date_Time = 0
Col_Operator = 1
Col_Serial_Number = 2
Col_Resist_Resist = 3
Col_Resist_Resist_Ave = 4
Col_Step_Step = 5
Col_Y1_Y1 = 6
Col_Y1_Y1_Ave = 7
Col_Y2_Y2 = 8
Col_Y2_Y2_Ave = 9
Col_Y1_ETCH_AMOUNT =32 #2025/02/09 New add
Col_Y2_ETCH_AMOUNT =33  #2025/02/09 New add


########## 取得した項目と型の対応表を定義 ##########
key_type = {
    'key_Start_Date_Time': str,
    'key_Operator': str,
    'key_Part_Number': str,
    'key_Serial_Number': str,
    'key_Resist_Resist1': float,
    'key_Resist_Resist2': float,
    'key_Resist_Resist3': float,
    'key_Resist_Resist4': float,
    'key_Resist_Resist_Ave': float,
    'key_Step_Step': float,
    'key_Y1_Y1_1': float,
    'key_Y1_Y1_2': float,
    'key_Y1_Y1_3': float,
    'key_Y1_Y1_4': float,
    'key_Y1_Y1_Ave': float,
    'key_Y2_Y2_1': float,
    'key_Y2_Y2_2': float,
    'key_Y2_Y2_3': float,
    'key_Y2_Y2_4': float,
    'key_Y2_Y2_Ave': float,
    'key_STARTTIME_SORTED' : float,
    'key_SORTNUMBER' : float,
    "key_LotNumber_9": str,
    'key_ETCH_AMOUNT_Y1' : float,  #2025/02/09 New add
    'key_ETCH_AMOUNT_Y2' : float    #2025/02/09 New add
}


def Main(FilePath, FileName, TextFile):


    ########## Excelファイルをローカルにコピー ##########

    # ----- 正規表現で取出し、直近で変更があったファイルを取得する -----
    Log.Log_Info(Log_File, 'Excel File Copy')

    Excel_File_List = []
    for file in glob.glob(FilePath + FileName):
        if '$' not in file:
            dt = strftime("%Y-%m-%d %H:%M:%S", localtime(os.path.getmtime(file)))
            Excel_File_List.append([file, dt])

    # ----- dt(更新日時)の降順で並び替える -----
    Excel_File_List = sorted(Excel_File_List, key=lambda x: x[1], reverse=True)
    Excel_File = shutil.copy(Excel_File_List[0][0], '../DataFile/024_ISO-EML/')


    ########## DaraFrameの作成 ##########

    # ----- 取得開始行の取り出し -----
    Log.Log_Info(Log_File, 'Get The Starting Row Count')
    Start_Number = Row_Number_Func.start_row_number(TextFile) - 500

    # ----- ExcelデータをDataFrameとして取得 -----
    Log.Log_Info(Log_File, 'Read Excel')
    df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="B:AZ", skiprows=Start_Number)

    # ----- 使わない列を落とす -----
    df = df.drop([3,5,8,9,10,11,12,13,14,16,17,18,19,22,23,24,25], axis=1)


    # ----- 列番号の振り直し -----
    Log.Log_Info(Log_File, 'Setting Columns Number')
    df.columns = range(df.shape[1])

    # ----- 末尾から欠損のデータを落としていく -----
    Getting_Row = len(df) - 1
    while Getting_Row >= 0 and str(df.iloc[Getting_Row, Col_Serial_Number]) == "nan":
        Getting_Row -= 1
    df = df[:Getting_Row + 3]

    # ----- OK/NGは使用しないのでnanに置き換える -----
    df = df.replace(["OK", "NG"], np.nan)

    # ----- 0もnanに置き換える(関数エラー) -----
    df = df.replace(0, np.nan)


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

        # ----- 日付/作業者/ロット番号に空欄が1つでも存在すればTrueを返す -----
        Log.Log_Info(Log_File, "Blank Check")
        if df.iloc[Row_Number, Col_Start_Date_Time] is np.nan or df.iloc[Row_Number, Col_Operator] is np.nan or df.iloc[Row_Number, Col_Serial_Number] is np.nan:
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
            'key_Part_Number': Part_Number,
            'key_Serial_Number': Serial_Number,
            "key_LotNumber_9": Nine_Serial_Number,
            'key_Resist_Resist1': df.iloc[Row_Number-1, Col_Resist_Resist],
            'key_Resist_Resist2': df.iloc[Row_Number, Col_Resist_Resist],
            'key_Resist_Resist3': df.iloc[Row_Number+1, Col_Resist_Resist],
            'key_Resist_Resist4': df.iloc[Row_Number+2, Col_Resist_Resist],
            'key_Resist_Resist_Ave': df.iloc[Row_Number, Col_Resist_Resist_Ave],
            'key_Step_Step': df.iloc[Row_Number-1, Col_Step_Step],
            'key_Y1_Y1_1': df.iloc[Row_Number-1, Col_Y1_Y1],
            'key_Y1_Y1_2': df.iloc[Row_Number, Col_Y1_Y1],
            'key_Y1_Y1_3': df.iloc[Row_Number+1, Col_Y1_Y1],
            'key_Y1_Y1_4': df.iloc[Row_Number+2, Col_Y1_Y1],
            'key_Y1_Y1_Ave': df.iloc[Row_Number, Col_Y1_Y1_Ave],
            'key_Y2_Y2_1': df.iloc[Row_Number-1, Col_Y2_Y2],
            'key_Y2_Y2_2': df.iloc[Row_Number, Col_Y2_Y2],
            'key_Y2_Y2_3': df.iloc[Row_Number+1, Col_Y2_Y2],
            'key_Y2_Y2_4': df.iloc[Row_Number+2, Col_Y2_Y2],
            'key_Y2_Y2_Ave': df.iloc[Row_Number, Col_Y2_Y2_Ave],
            'key_ETCH_AMOUNT_Y1' : df.iloc[Row_Number, Col_Y1_ETCH_AMOUNT], #2025/02/09 New add
            'key_ETCH_AMOUNT_Y2' : df.iloc[Row_Number, Col_Y2_ETCH_AMOUNT]  #2025/02/09 New add 
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
                '               <TestStep Name="Resist" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Resist1" Units="nm" Value=' + '"' + str(data_dict["key_Resist_Resist1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Resist2" Units="nm" Value=' + '"' + str(data_dict["key_Resist_Resist2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Resist3" Units="nm" Value=' + '"' + str(data_dict["key_Resist_Resist3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Resist4" Units="nm" Value=' + '"' + str(data_dict["key_Resist_Resist4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Resist_Ave" Units="nm" Value=' + '"' + str(data_dict["key_Resist_Resist_Ave"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Step" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Step" Units="nm" Value=' + '"' + str(data_dict["key_Step_Step"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Y1" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Y1_1" Units="nm" Value=' + '"' + str(data_dict["key_Y1_Y1_1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y1_2" Units="nm" Value=' + '"' + str(data_dict["key_Y1_Y1_2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y1_3" Units="nm" Value=' + '"' + str(data_dict["key_Y1_Y1_3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y1_4" Units="nm" Value=' + '"' + str(data_dict["key_Y1_Y1_4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y1_Ave" Units="nm" Value=' + '"' + str(data_dict["key_Y1_Y1_Ave"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y1_ETCH_AMOUNT" Units="nm" Value=' + '"' + str(data_dict["key_ETCH_AMOUNT_Y1"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Y2" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Y2_1" Units="nm" Value=' + '"' + str(data_dict["key_Y2_Y2_1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y2_2" Units="nm" Value=' + '"' + str(data_dict["key_Y2_Y2_2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y2_3" Units="nm" Value=' + '"' + str(data_dict["key_Y2_Y2_3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y2_4" Units="nm" Value=' + '"' + str(data_dict["key_Y2_Y2_4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y2_Ave" Units="nm" Value=' + '"' + str(data_dict["key_Y2_Y2_Ave"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y2_ETCH_AMOUNT" Units="nm" Value=' + '"' + str(data_dict["key_ETCH_AMOUNT_Y2"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="SORTED_DATA" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value=' + '"' + str(data_dict["key_STARTTIME_SORTED"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value=' + '"' + str(data_dict["key_SORTNUMBER"]) + '"/>' + '\n' +
                '                   <Data DataType="String" Name="LotNumber_5" Value=' + '"' + str(data_dict["key_Serial_Number"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                '                   <Data DataType="String" Name="LotNumber_9" Value=' + '"' + str(data_dict["key_LotNumber_9"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '\n'
                '               <TestEquipment>' + '\n' +
                '                   <Item DeviceName="Dektak" DeviceSerialNumber="-"/>' + '\n' +
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
    Row_Number_Func.next_start_row_number(TextFile, Start_Number + row_end)

    # ----- 最終行を書き込んだファイルをGドライブに転送 -----
    shutil.copy(TextFile, 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/024_ISO-EML/13_ProgramUsedFile/')


if __name__ == '__main__':

    Main('Z:/ホト・エッチング/製品/分離溝ｴｯﾁ/', '*HTL13★_ 分離溝ｴｯﾁ ﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsx', 'HTL13_StartRow.txt')
    Main('Z:/ホト・エッチング/製品/分離溝ｴｯﾁ/', '*HL13B6・E1_分離溝ｴｯﾁ ﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsx', 'HL13B6_StartRow.txt') #update file name 2025/01/09

# ----- ログ書込：Main処理の終了 -----
Log.Log_Info(Log_File, 'Program End')
