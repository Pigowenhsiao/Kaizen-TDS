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
from datetime import date, timedelta, datetime, time
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
ProductFamily = 'SAG FAB'
Operation = 'BJ2_Crystal_Depth'
TestStation = 'BJ2'


########## Logの設定 ##########

# ----- ログファイルの作成 -----
Log_Folder_Name = str(date.today())
if not os.path.exists("../Log/" + Log_Folder_Name):
    os.makedirs("../Log/" + Log_Folder_Name)

Log_File = '../Log/' + Log_Folder_Name + '/007_BJ2.log'
Log.Log_Info(Log_File, 'Program Start')


########## シート名の定義 ##########
Data_Sheet_Name = 'データ'
XY_Sheet_Name = 'ウェハ座標'


########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
# Output_filepath = '../XML/007_BJ2/'
#Output_filepath = 'C:/Users/hsi67063/Documents/TEMP/'  #for test


########## 取得するデータの列番号を定義 ##########
Col_Start_Date_Time = 0
Col_Operator = 1
Col_Serial_Number = 3
Col_Depth1_First1 = 4
Col_Depth1_First2 = 5
Col_Depth1_First3 = 6
Col_Depth1_First4 = 7
Col_Depth1_First5 = 8
Col_Depth1_First_Ave = 9
Col_Thickness1_First1 = 10
Col_Thickness1_First2 = 11
Col_Thickness1_First3 = 12
Col_Thickness1_First4 = 13
Col_Thickness1_First5 = 14
Col_Thickness1_First_Ave = 15
Col_First_Depth = 16
Col_First_Rate = 17
Col_Depth2_Second1 = 20
Col_Depth2_Second2 = 21
Col_Depth2_Second3 = 22
Col_Depth2_Second4= 23
Col_Depth2_Second5 = 24
Col_Depth2_Second_Ave = 25
Col_Thickness2_Second1 = 26
Col_Thickness2_Second2 = 27
Col_Thickness2_Second3 = 28
Col_Thickness2_Second4 = 29
Col_Thickness2_Second5 = 30
Col_Thickness2_Second_Ave = 31
Col_Second_Depth = 32
Col_Second_Rate = 33
Col_Second_Time = 19
Col_Final_Depth = 36
Col_Final_Error = 37
Col_X = 1
Col_Y = 2


########## 取得した項目と型の対応表を定義 ##########
key_type = {
    "key_Start_Date_Time" : str,
    "key_Part_Number" : str,
    "key_Serial_Number" : str,
    "key_Operator" : str,
    "key_Depth1_First1" : float,
    "key_Depth1_First2" : float,
    "key_Depth1_First3" : float,
    "key_Depth1_First4" : float,
    "key_Depth1_First5" : float,
    "key_Depth1_First_Ave" : float,
    "key_Thickness1_First1" : float,
    "key_Thickness1_First2" : float,
    "key_Thickness1_First3" : float,
    "key_Thickness1_First4" : float,
    "key_Thickness1_First5" : float,
    "key_Thickness1_First_Ave" : float,
    "key_First_Depth" : float,
    "key_First_Rate" : float,
    "key_Depth2_Second1": float,
    "key_Depth2_Second2": float,
    "key_Depth2_Second3": float,
    "key_Depth2_Second4": float,
    "key_Depth2_Second5": float,
    "key_Depth2_Second_Ave": float,
    "key_Thickness2_Second1": float,
    "key_Thickness2_Second2": float,
    "key_Thickness2_Second3": float,
    "key_Thickness2_Second4": float,
    "key_Thickness2_Second5": float,
    "key_Thickness2_Second_Ave": float,
    "key_Second_Depth" : float,
    "key_Second_Rate" : float,
    "key_Second_Time" : float,
    "key_Final_Depth" : float,
    "key_Final_Error" : float,
    "key_X1" : float,
    "key_X2" : float,
    "key_X3" : float,
    "key_X4" : float,
    "key_X5" : float,
    "key_Y1" : float,
    "key_Y2" : float,
    "key_Y3" : float,
    "key_Y4" : float,
    "key_Y5" : float,
    "key_STARTTIME_SORTED": float,
    "key_SORTNUMBER" : float,
    "key_LotNumber_9": str
}


def Main(FilePath, FileName, TextFile, Equipment):


    ########## Excelファイルをローカルにコピー ##########

    # ----- 正規表現で取出し、直近で変更があったファイルを取得する -----
    Log.Log_Info(Log_File, 'Excel File Copy')

    Excel_file_list = []
    for file in glob.glob(FilePath + FileName):
        if '$' not in file:
            dt = strftime("%Y-%m-%d %H:%M:%S", localtime(os.path.getmtime(file)))
            Excel_file_list.append([file, dt])

    # ----- dt(更新日時)の降順で並び替える -----
    Excel_file_list = sorted(Excel_file_list, key=lambda x: x[1], reverse=True)
    Excel_File = shutil.copy(Excel_file_list[0][0], '../DataFile/007_BJ2/')


    ########## DaraFrameの作成 ##########

    # ----- 取得開始行の取り出し -----
    Log.Log_Info(Log_File, 'Get The Starting Row Count')
    Start_Number = max(Row_Number_Func.start_row_number(TextFile)-500, 5)
    
    # ----- ExcelデータをDataFrameとして取得 -----
    Log.Log_Info(Log_File, 'Read Excel')
    Log.Log_Info(Log_File, Excel_File)
    df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="D:AS", skiprows=Start_Number)
    df_xy = pd.read_excel(Excel_File, header=None, sheet_name=XY_Sheet_Name, usecols="A:C")
    
    # ----- 使わない列を落とす -----
    df = df.drop(range(39, 43), axis=1)

    # ----- 列番号の振り直し -----
    Log.Log_Info(Log_File, 'Setting Columns Number')
    df.columns = range(df.shape[1])
    df_xy.columns = range(df_xy.shape[1])

    # ----- 00:00:00のデータをnp.nanで置き換える -----
    #df = df.replace(time(), np.nan).infer_objects(copy=False)
    pd.set_option('future.no_silent_downcasting', True) # the old method will not support in new version, update it
    df = df.replace(time(), np.nan)

    # ----- 末尾から未入力・欠損値のデータを落としていく -----
    Getting_Row = len(df) - 1
    while Getting_Row >= 0 and np.isnan(df.iloc[Getting_Row, Col_Final_Depth]):
        Getting_Row -= 1

    df = df[:Getting_Row + 1]

    # ----- 次の開始行数をメモ -----
    Next_Start_Row = Start_Number + df.shape[0] + 1

    # ----- 日付欄に文字列が入っていたらNoneに置き換える -----
    for i in range(df.shape[0]):
        if type(df.iloc[i, 0]) is not datetime:
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
        # ----- Operator -----
        if str(df.iloc[Row_Number, Col_Operator]) == 'nan':
            Log.Log_Error(Log_File, "Blank Error\n")
            Row_Number += 1
            continue
        
        # ----- FinalDepth -----
        if str(df.iloc[Row_Number, Col_Final_Depth]) == 'nan':
            Log.Log_Error(Log_File, "Blank Error\n")
            Row_Number += 1
            continue


        ########## 現在処理を行っている行のデータの取得 ##########

        # ----- 取得したデータを格納するデータ構造(辞書)を作成 -----
        Log.Log_Info(Log_File, 'Data Acquisition')
        data_dict = dict()

        # ----- ロット番号を取得 -----
        print(df.iloc[Row_Number, Col_Serial_Number])
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
            "key_Start_Date_Time": df.iloc[Row_Number, Col_Start_Date_Time],
            "key_Part_Number": Part_Number,
            "key_Serial_Number": Serial_Number,
            "key_LotNumber_9": Nine_Serial_Number,
            "key_Operator": df.iloc[Row_Number, Col_Operator],
            "key_Depth1_First1": df.iloc[Row_Number, Col_Depth1_First1],
            "key_Depth1_First2": df.iloc[Row_Number, Col_Depth1_First2],
            "key_Depth1_First3": df.iloc[Row_Number, Col_Depth1_First3],
            "key_Depth1_First4": df.iloc[Row_Number, Col_Depth1_First4],
            "key_Depth1_First5": df.iloc[Row_Number, Col_Depth1_First5],
            "key_Depth1_First_Ave": df.iloc[Row_Number, Col_Depth1_First_Ave],
            "key_Thickness1_First1": df.iloc[Row_Number, Col_Thickness1_First1],
            "key_Thickness1_First2": df.iloc[Row_Number, Col_Thickness1_First2],
            "key_Thickness1_First3": df.iloc[Row_Number, Col_Thickness1_First3],
            "key_Thickness1_First4": df.iloc[Row_Number, Col_Thickness1_First4],
            "key_Thickness1_First5": df.iloc[Row_Number, Col_Thickness1_First5],
            "key_Thickness1_First_Ave": df.iloc[Row_Number, Col_Thickness1_First_Ave],
            "key_First_Depth": df.iloc[Row_Number, Col_First_Depth],
            "key_First_Rate": df.iloc[Row_Number, Col_First_Rate],
            "key_Depth2_Second1": df.iloc[Row_Number, Col_Depth2_Second1],
            "key_Depth2_Second2": df.iloc[Row_Number, Col_Depth2_Second2],
            "key_Depth2_Second3": df.iloc[Row_Number, Col_Depth2_Second3],
            "key_Depth2_Second4": df.iloc[Row_Number, Col_Depth2_Second4],
            "key_Depth2_Second5": df.iloc[Row_Number, Col_Depth2_Second5],
            "key_Depth2_Second_Ave": df.iloc[Row_Number, Col_Depth2_Second_Ave],
            "key_Thickness2_Second1": df.iloc[Row_Number, Col_Thickness2_Second1],
            "key_Thickness2_Second2": df.iloc[Row_Number, Col_Thickness2_Second2],
            "key_Thickness2_Second3": df.iloc[Row_Number, Col_Thickness2_Second3],
            "key_Thickness2_Second4": df.iloc[Row_Number, Col_Thickness2_Second4],
            "key_Thickness2_Second5": df.iloc[Row_Number, Col_Thickness2_Second5],
            "key_Thickness2_Second_Ave": df.iloc[Row_Number, Col_Thickness2_Second_Ave],
            "key_Second_Depth": df.iloc[Row_Number, Col_Second_Depth],
            "key_Second_Rate": df.iloc[Row_Number, Col_Second_Rate],
            "key_Second_Time": df.iloc[Row_Number, Col_Second_Time],
            "key_Final_Depth": df.iloc[Row_Number, Col_Final_Depth],
            "key_Final_Error": df.iloc[Row_Number, Col_Final_Error],
            "key_X1": df_xy.iloc[1, Col_X],
            "key_X2": df_xy.iloc[2, Col_X],
            "key_X3": df_xy.iloc[3, Col_X],
            "key_X4": df_xy.iloc[4, Col_X],
            "key_X5": df_xy.iloc[5, Col_X],
            "key_Y1": df_xy.iloc[1, Col_Y],
            "key_Y2": df_xy.iloc[2, Col_Y],
            "key_Y3": df_xy.iloc[3, Col_Y],
            "key_Y4": df_xy.iloc[4, Col_Y],
            "key_Y5": df_xy.iloc[5, Col_Y],
            "key_Equipment": Equipment
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


        ########## 欠損値の変換 ##########
        
        # ----- FainalDepthに数値が入っていれば他に欠損値が含まれていても処理する -----
        for key, key_to_type in key_type.items():
            if data_dict[key] is np.nan or str(data_dict[key]) == "nan":
                Log.Log_Info(Log_File, 'key：' + key + 'Convert Nan to Blank')
                data_dict[key] = ""
        
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
                '               <TestStep Name="Depth1" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="First1" Units="nm" Value=' + '"' + str(data_dict["key_Depth1_First1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First2" Units="nm" Value=' + '"' + str(data_dict["key_Depth1_First2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First3" Units="nm" Value=' + '"' + str(data_dict["key_Depth1_First3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First4" Units="nm" Value=' + '"' + str(data_dict["key_Depth1_First4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First5" Units="nm" Value=' + '"' + str(data_dict["key_Depth1_First5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First_Ave" Units="nm" Value=' + '"' + str(data_dict["key_Depth1_First_Ave"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Thickness1" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="First1" Units="nm" Value=' + '"' + str(data_dict["key_Thickness1_First1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First2" Units="nm" Value=' + '"' + str(data_dict["key_Thickness1_First2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First3" Units="nm" Value=' + '"' + str(data_dict["key_Thickness1_First3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First4" Units="nm" Value=' + '"' + str(data_dict["key_Thickness1_First4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First5" Units="nm" Value=' + '"' + str(data_dict["key_Thickness1_First5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First_Ave" Units="nm" Value=' + '"' + str(data_dict["key_Thickness1_First_Ave"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="First" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Depth" Units="nm" Value=' + '"' + str(data_dict["key_First_Depth"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Rate" Units="nm/min" Value=' + '"' + str(data_dict["key_First_Rate"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Depth2" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Second1" Units="nm" Value=' + '"' + str(data_dict["key_Depth2_Second1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Second2" Units="nm" Value=' + '"' + str(data_dict["key_Depth2_Second2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Second3" Units="nm" Value=' + '"' + str(data_dict["key_Depth2_Second3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Second4" Units="nm" Value=' + '"' + str(data_dict["key_Depth2_Second4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Second5" Units="nm" Value=' + '"' + str(data_dict["key_Depth2_Second5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Second_Ave" Units="nm" Value=' + '"' + str(data_dict["key_Depth2_Second_Ave"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Thickness2" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Second1" Units="nm" Value=' + '"' + str(data_dict["key_Thickness2_Second1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Second2" Units="nm" Value=' + '"' + str(data_dict["key_Thickness2_Second2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Second3" Units="nm" Value=' + '"' + str(data_dict["key_Thickness2_Second3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Second4" Units="nm" Value=' + '"' + str(data_dict["key_Thickness2_Second4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Second5" Units="nm" Value=' + '"' + str(data_dict["key_Thickness2_Second5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Second_Ave" Units="nm" Value=' + '"' + str(data_dict["key_Thickness2_Second_Ave"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Second" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Depth" Units="nm" Value=' + '"' + str(data_dict["key_Second_Depth"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Rate" Units="nm/min" Value=' + '"' + str(data_dict["key_Second_Rate"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Time" Units="min" Value=' + '"' + str(data_dict["key_Second_Time"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Final" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Depth" Units="nm" Value=' + '"' + str(data_dict["key_Final_Depth"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Error" Units="nm" Value=' + '"' + str(data_dict["key_Final_Error"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Coordinate" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="X1" Units="um" Value=' + '"' + str(data_dict["key_X1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="X2" Units="um" Value=' + '"' + str(data_dict["key_X2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="X3" Units="um" Value=' + '"' + str(data_dict["key_X3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="X4" Units="um" Value=' + '"' + str(data_dict["key_X4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="X5" Units="um" Value=' + '"' + str(data_dict["key_X5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y1" Units="um" Value=' + '"' + str(data_dict["key_Y1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y2" Units="um" Value=' + '"' + str(data_dict["key_Y2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y3" Units="um" Value=' + '"' + str(data_dict["key_Y3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y4" Units="um" Value=' + '"' + str(data_dict["key_Y4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y5" Units="um" Value=' + '"' + str(data_dict["key_Y5"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="SORTED_DATA" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value=' + '"' + str(data_dict["key_STARTTIME_SORTED"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value=' + '"' + str(data_dict["key_SORTNUMBER"]) + '"/>' + '\n' +
                '                   <Data DataType="String" Name="LotNumber_5" Value=' + '"' + str(data_dict["key_Serial_Number"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                '                   <Data DataType="String" Name="LotNumber_9" Value=' + '"' + str(data_dict["key_LotNumber_9"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestEquipment>' + '\n' +
                '                   <Item DeviceName="DryEtch" DeviceSerialNumber=' + '"' + str(data_dict["key_Equipment"]) + '"/>' + '\n' +
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
    Row_Number_Func.next_start_row_number(TextFile, Next_Start_Row)

    # ----- 最終行を書き込んだファイルをGドライブに転送 -----
    shutil.copy(TextFile, 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/007_BJ2/13_ProgramUsedFile/')


if __name__ == '__main__':
    
    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ2号機/3インチ/HTL13★＊＊/', '*ドライ2号機_3インチ_HTL13★＊＊_BJ2ドライエッチプログラムシート*.xlsm', 'BJ2_HL13B6_StartRow_Dry2.txt', '#2') ##2025/01/09 New Add

    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ1号機/3インチ/HTL13★＊＊/', '*3ｲﾝﾁ_HTL13★＊＊_*BJ2ﾄﾞﾗｲｴｯﾁﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsm', 'BJ2_HTL13_StartRow_Dry1.txt', '#1')
    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ1号機/3インチ/HL13B6/', '*3ｲﾝﾁ_HL13B6-BT＊＊_*BJ2ﾄﾞﾗｲｴｯﾁﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsm', 'BJ2_HL13B6_StartRow_Dry1.txt', '#1')
    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ1号機/3インチ/HL15B5/', '*3ｲﾝﾁ_HL15B5-BT＊＊_*BJ2ﾄﾞﾗｲｴｯﾁﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsm', 'BJ2_HL15B5_StartRow_Dry1.txt', '#1')
    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ5号機/3インチ/HTL13★＊＊/', '*ドライ5号機_HTL13★＊＊_*BJ2ﾄﾞﾗｲｴｯﾁﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsx', 'BJ2_HTL13_StartRow_Dry5.txt', '#5')
    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ5号機/3インチ/HL13B6/', '*ドライ5号機_HL13B6-BT＊＊_*BJ2ﾄﾞﾗｲｴｯﾁﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsx', 'BJ2_HL13B6_StartRow_Dry5.txt', '#5')
    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ5号機/3インチ/HL15B5/', '*ドライ5号機_HL15B5-BT＊＊_*BJ2ﾄﾞﾗｲｴｯﾁﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsx', 'BJ2_HL15B5_StartRow_Dry5.txt', '#5')

# ----- ログ書込：Main処理の終了 -----
Log.Log_Info(Log_File, 'Program End')