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
Operation = 'MESA_Crystal_Depth'
TestStation = 'MESA'

########## Logの設定 ##########
Log_Folder_Name = str(date.today())
if not os.path.exists("../Log/" + Log_Folder_Name):
    os.makedirs("../Log/" + Log_Folder_Name)
Log_File = '../Log/' + Log_Folder_Name + '/013_MESA_ICP.log'
Log.Log_Info(Log_File, 'Program Start')

########## シート名の定義 ##########
Data_Sheet_Name = 'ﾃﾞｰﾀ'
XY_Sheet_Name = 'ウェハ座標'

########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
Output_filepath = 'C:/Users/hsi67063/Box/00-home-pigo.hsiao/TEMP/XML/'

########## 取得するデータの列番号を定義 ##########
Col_Start_date = 0
Col_Operator = 1
Col_Order = 3
Col_Serial_Number = 4
Col_Time_Time = 5
Col_Depth_First1 = 6
Col_Depth_First2 = 7
Col_Depth_First3 = 8
Col_Depth_First4 = 9
Col_Depth_First5 = 10
Col_Depth_First_Ave = 11
Col_Thickness_First1 = 12
Col_Thickness_First2 = 13
Col_Thickness_First3 = 14
Col_Thickness_First4 = 15
Col_Thickness_First5 = 16
Col_Thickness_First_Ave = 17
Col_Etching_Etching1 = 18
Col_Etching_Etching2 = 19
Col_Etching_Etching3 = 20
Col_Etching_Etching4 = 21
Col_Etching_Etching5 = 22
Col_Etching_Etching_Ave = 23
Col_Etching_Etching_Max_Min = 26
Col_Etching_Etching_3sigma = 27
Col_Etching_Etching_Rate = 28
Col_Etching_Etching_Error = 29
Col_X = 1
Col_Y = 2

########## 取得した項目と型の対応表を定義 ##########
key_type = {
    'key_Start_Date_Time': str,
    'key_Operator': str,
    'key_Order': str,
    'key_Serial_Number': str,
    'key_Part_Number': str,
    'key_Time_Time': float,
    'key_Depth_First1': float,
    'key_Depth_First2': float,
    'key_Depth_First3': float,
    'key_Depth_First4': float,
    'key_Depth_First5': float,
    'key_Depth_First_Ave': float,
    'key_Thickness_First1': float,
    'key_Thickness_First2': float,
    'key_Thickness_First3': float,
    'key_Thickness_First4': float,
    'key_Thickness_First5': float,
    'key_Thickness_First_Ave': float,
    'key_Etching_Etching1': float,
    'key_Etching_Etching2': float,
    'key_Etching_Etching3': float,
    'key_Etching_Etching4': float,
    'key_Etching_Etching5': float,
    'key_Etching_Etching_Ave': float,
    'key_Etching_Etching_Max-Min': float,
    'key_Etching_Etching_3sigma': float,
    'key_Etching_Etching_Rate': float,
    'key_Etching_Etching_Error': float,
    'key_X1': float,
    'key_X2': float,
    'key_X3': float,
    'key_X4': float,
    'key_X5': float,
    'key_Y1': float,
    'key_Y2': float,
    'key_Y3': float,
    'key_Y4': float,
    'key_Y5': float,
    'key_STARTTIME_SORTED' : float,
    'key_SORTNUMBER' : float,
    "key_LotNumber_9": str
}


def process_excel_file(Excel_File, Log_File):
    Log.Log_Info(Log_File, f"Processing file: {Excel_File}")
    Start_Number = Row_Number_Func.start_row_number("MESA_StartRow_ICP.txt") - 500

    # 讀取 Excel 檔案
    Log.Log_Info(Log_File, 'Read Excel')
    df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="B:AE", skiprows=Start_Number)
    df = df[df.iloc[:, 7] != 0]
    df_xy = pd.read_excel(Excel_File, header=None, sheet_name=XY_Sheet_Name, usecols="A:C")
    df.columns = range(df.shape[1])
    df_xy.columns = range(df_xy.shape[1])

    # 處理日期欄位與資料修正
    for i in range(df.shape[0]):
        if not isinstance(df.iloc[i, 0], (pd.Timestamp, datetime)):
            df.iloc[i, 0] = np.nan
        if df.iloc[i, 1] == time():
            df.iloc[i, 1] = np.nan

    Getting_Row = len(df) - 1
    while Getting_Row >= 0 and df.isnull().any(axis=1)[Getting_Row]:
        Getting_Row -= 1
    df = df[:Getting_Row + 1]

    Next_Start_Row = Start_Number + df.shape[0] + 1

    df[0] = pd.to_datetime(df[0])
    one_month_ago = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) - relativedelta(months=1)
    df = df[(df[0] >= one_month_ago)]
    print(df)
    sys.exit()
    row_end = len(df)
    Row_Number = 0
    df_idx = df.index.values

    while Row_Number < row_end:
        Log.Log_Info(Log_File, "Blank Check")
        if df.isnull().any(axis=1)[df_idx[Row_Number]]:
            Log.Log_Error(Log_File, "Blank Error\n")
            Row_Number += 1
            continue

        Log.Log_Info(Log_File, 'Data Acquisition')
        data_dict = dict()
        Serial_Number = str(df.iloc[Row_Number, Col_Serial_Number])
        if Serial_Number == "nan":
            Log.Log_Error(Log_File, "Lot Error\n")
            Row_Number += 1
            continue

        conn, cursor = SQL.connSQL()
        if conn is None:
            Log.Log_Error(Log_File, Serial_Number + ' : ' + 'Connection with Prime Failed')
            break
        Part_Number, Nine_Serial_Number = SQL.selectSQL(cursor, Serial_Number)
        SQL.disconnSQL(conn, cursor)
        if Part_Number is None:
            Log.Log_Error(Log_File, Serial_Number + ' : ' + "PartNumber Error\n")
            Row_Number += 1
            continue
        if Part_Number == 'LDアレイ_':
            Row_Number += 1
            continue

        data_dict = {
            'key_Start_Date_Time': df.iloc[Row_Number, Col_Start_date],
            'key_Operator': df.iloc[Row_Number, Col_Operator],
            'key_Order': df.iloc[Row_Number, Col_Order],
            'key_Serial_Number': df.iloc[Row_Number, Col_Serial_Number],
            'key_Part_Number': Part_Number,
            "key_LotNumber_9": Nine_Serial_Number,
            'key_Time_Time': df.iloc[Row_Number, Col_Time_Time],
            'key_Depth_First1': df.iloc[Row_Number, Col_Depth_First1],
            'key_Depth_First2': df.iloc[Row_Number, Col_Depth_First2],
            'key_Depth_First3': df.iloc[Row_Number, Col_Depth_First3],
            'key_Depth_First4': df.iloc[Row_Number, Col_Depth_First4],
            'key_Depth_First5': df.iloc[Row_Number, Col_Depth_First5],
            'key_Depth_First_Ave': df.iloc[Row_Number, Col_Depth_First_Ave],
            'key_Thickness_First1': df.iloc[Row_Number, Col_Thickness_First1],
            'key_Thickness_First2': df.iloc[Row_Number, Col_Thickness_First2],
            'key_Thickness_First3': df.iloc[Row_Number, Col_Thickness_First3],
            'key_Thickness_First4': df.iloc[Row_Number, Col_Thickness_First4],
            'key_Thickness_First5': df.iloc[Row_Number, Col_Thickness_First5],
            'key_Thickness_First_Ave': df.iloc[Row_Number, Col_Thickness_First_Ave],
            'key_Etching_Etching1': df.iloc[Row_Number, Col_Etching_Etching1],
            'key_Etching_Etching2': df.iloc[Row_Number, Col_Etching_Etching2],
            'key_Etching_Etching3': df.iloc[Row_Number, Col_Etching_Etching3],
            'key_Etching_Etching4': df.iloc[Row_Number, Col_Etching_Etching4],
            'key_Etching_Etching5': df.iloc[Row_Number, Col_Etching_Etching5],
            'key_Etching_Etching_Ave': df.iloc[Row_Number, Col_Etching_Etching_Ave],
            'key_Etching_Etching_Max-Min': df.iloc[Row_Number, Col_Etching_Etching_Max_Min],
            'key_Etching_Etching_3sigma': df.iloc[Row_Number, Col_Etching_Etching_3sigma],
            'key_Etching_Etching_Rate': df.iloc[Row_Number, Col_Etching_Etching_Rate],
            'key_Etching_Etching_Error': df.iloc[Row_Number, Col_Etching_Etching_Error],
            'key_X1': df_xy.iloc[1, Col_X],
            'key_X2': df_xy.iloc[2, Col_X],
            'key_X3': df_xy.iloc[3, Col_X],
            'key_X4': df_xy.iloc[4, Col_X],
            'key_X5': df_xy.iloc[5, Col_X],
            'key_Y1': df_xy.iloc[1, Col_Y],
            'key_Y2': df_xy.iloc[2, Col_Y],
            'key_Y3': df_xy.iloc[3, Col_Y],
            'key_Y4': df_xy.iloc[4, Col_Y],
            'key_Y5': df_xy.iloc[5, Col_Y],
            'Tool_name': 'ICP'
        }
        if 'ICPドライ2号機' in Excel_File:
            data_dict['Tool_name'] = 'ICP2'
        print(Excel_File, "Check")

        Log.Log_Info(Log_File, 'Date Format Conversion')
        data_dict["key_Start_Date_Time"] = Convert_Date.Edit_Date(data_dict["key_Start_Date_Time"])
        if len(data_dict["key_Start_Date_Time"]) != 19:
            Log.Log_Error(Log_File, data_dict["key_Serial_Number"] + ' : ' + "Date Error\n")
            Row_Number += 1
            continue

        date_obj = datetime.strptime(str(data_dict["key_Start_Date_Time"]).replace('T', ' ').replace('.', ':'), "%Y-%m-%d %H:%M:%S")
        date_excel_number = int(str(date_obj - datetime(1899, 12, 30)).split()[0])
        excel_row = Start_Number + df_idx[Row_Number] + 1
        excel_row_div = excel_row / 10 ** 6
        date_excel_number += excel_row_div
        data_dict["key_STARTTIME_SORTED"] = date_excel_number
        data_dict["key_SORTNUMBER"] = excel_row

        Log.Log_Info(Log_File, "Check Data Type")
        Result = Check.Data_Type(key_type, data_dict)
        if Result == False:
            Log.Log_Error(Log_File, data_dict["key_Serial_Number"] + ' : ' + "Data Error\n")
            Row_Number += 1
            continue

        XML_File_Name = 'Site=' + Site + ',ProductFamily=' + ProductFamily + ',Operation=' + Operation + \
                        ',Partnumber=' + data_dict["key_Part_Number"] + ',Serialnumber=' + data_dict["key_Serial_Number"] + \
                        ',Testdate=' + data_dict["key_Start_Date_Time"] + '.xml'
        Log.Log_Info(Log_File, 'Excel File To XML File Conversion: ' + XML_File_Name)
        with open(Output_filepath + XML_File_Name, 'w', encoding="utf-8") as f:
            f.write('<?xml version="1.0" encoding="utf-8"?>' + '\n' +
                '<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' + '\n' +
                '       <Result startDateTime="' + data_dict["key_Start_Date_Time"].replace(".", ":") + '" Result="Passed">' + '\n' +
                '               <Header SerialNumber="' + data_dict["key_Serial_Number"] + '" PartNumber="' + data_dict["key_Part_Number"] + '" Operation="' + Operation + '" TestStation="' + TestStation + '" Operator="' + data_dict["key_Operator"] + '" StartTime="' + data_dict["key_Start_Date_Time"].replace(".", ":") + '" Site="' + Site + '" LotNumber="' + data_dict["key_Serial_Number"] + '"/>' + '\n' +
                '               <TestStep Name="Order" startDateTime="' + data_dict["key_Start_Date_Time"].replace(".", ":") + '" Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Order" Units="No" Value="' + str(data_dict["key_Order"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Time" startDateTime="' + data_dict["key_Start_Date_Time"].replace(".", ":") + '" Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Time" Units="sec" Value="' + str(data_dict["key_Time_Time"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Depth" startDateTime="' + data_dict["key_Start_Date_Time"].replace(".", ":") + '" Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="First1" Units="nm" Value="' + str(data_dict["key_Depth_First1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First2" Units="nm" Value="' + str(data_dict["key_Depth_First2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First3" Units="nm" Value="' + str(data_dict["key_Depth_First3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First4" Units="nm" Value="' + str(data_dict["key_Depth_First4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First5" Units="nm" Value="' + str(data_dict["key_Depth_First5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First_Ave" Units="nm" Value="' + str(data_dict["key_Depth_First_Ave"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Thickness" startDateTime="' + data_dict["key_Start_Date_Time"].replace(".", ":") + '" Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="First1" Units="nm" Value="' + str(data_dict["key_Thickness_First1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First2" Units="nm" Value="' + str(data_dict["key_Thickness_First2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First3" Units="nm" Value="' + str(data_dict["key_Thickness_First3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First4" Units="nm" Value="' + str(data_dict["key_Thickness_First4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First5" Units="nm" Value="' + str(data_dict["key_Thickness_First5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="First_Ave" Units="nm" Value="' + str(data_dict["key_Thickness_First_Ave"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Etching" startDateTime="' + data_dict["key_Start_Date_Time"].replace(".", ":") + '" Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Etching1" Units="nm" Value="' + str(data_dict["key_Etching_Etching1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Etching2" Units="nm" Value="' + str(data_dict["key_Etching_Etching2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Etching3" Units="nm" Value="' + str(data_dict["key_Etching_Etching3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Etching4" Units="nm" Value="' + str(data_dict["key_Etching_Etching4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Etching5" Units="nm" Value="' + str(data_dict["key_Etching_Etching5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Etching_Ave" Units="nm" Value="' + str(data_dict["key_Etching_Etching_Ave"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Etching_Max-Min" Units="nm" Value="' + str(data_dict["key_Etching_Etching_Max-Min"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Etching_3sigma" Units="nm" Value="' + str(data_dict["key_Etching_Etching_3sigma"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Etching_Rate" Units="nm/min" Value="' + str(data_dict["key_Etching_Etching_Rate"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Etching_Error" Units="nm" Value="' + str(data_dict["key_Etching_Etching_Error"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Coordinate" startDateTime="' + data_dict["key_Start_Date_Time"].replace(".", ":") + '" Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="X1" Units="um" Value="' + str(data_dict["key_X1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="X2" Units="um" Value="' + str(data_dict["key_X2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="X3" Units="um" Value="' + str(data_dict["key_X3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="X4" Units="um" Value="' + str(data_dict["key_X4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="X5" Units="um" Value="' + str(data_dict["key_X5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y1" Units="um" Value="' + str(data_dict["key_Y1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y2" Units="um" Value="' + str(data_dict["key_Y2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y3" Units="um" Value="' + str(data_dict["key_Y3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y4" Units="um" Value="' + str(data_dict["key_Y4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y5" Units="um" Value="' + str(data_dict["key_Y5"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="SORTED_DATA" startDateTime="' + data_dict["key_Start_Date_Time"].replace(".", ":") + '" Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value="' + str(data_dict["key_STARTTIME_SORTED"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value="' + str(data_dict["key_SORTNUMBER"]) + '"/>' + '\n' +
                '                   <Data DataType="String" Name="LotNumber_5" Value="' + str(data_dict["key_Serial_Number"]) + '" CompOperation="LOG"/>' + '\n' +
                '                   <Data DataType="String" Name="LotNumber_9" Value="' + str(data_dict["key_LotNumber_9"]) + '" CompOperation="LOG"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestEquipment>' + '\n' +
                '                   <Item DeviceName="DryEtch" DeviceSerialNumber="' + data_dict['Tool_name'] + '"/>' + '\n' +
                '               </TestEquipment>' + '\n' +
                '               <ErrorData/>' + '\n' +
                '               <FailureData/>' + '\n' +
                '               <Configuration/>' + '\n' +
                '       </Result>' + '\n' +
                '</Results>'
            )
        Log.Log_Info(Log_File, data_dict["key_Serial_Number"] + ' : ' + "OK\n")
        Row_Number += 1

    ########## 次の開始行数の書き込み ##########
    Log.Log_Info(Log_File, 'Write the next starting line number')
    Row_Number_Func.next_start_row_number("MESA_StartRow_ICP.txt", Next_Start_Row)

    # ----- 最終行を書き込んだファイルをGドライブに転送 -----
    #shutil.copy("MESA_StartRow_ICP.txt", 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/013_MESA/13_ProgramUsedFile/')

def Main():
    Log.Log_Info(Log_File, 'Excel File Copy')
    #FilePath1 = 'Z:/ホト・エッチング/製品/ICPドライ/'
    #FileName1 = '*ICPドライ_HL13★xx_*ﾒｻﾄﾞﾗｲｴｯﾁﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsm'
    FilePath2 = 'Z:/ホト・エッチング/製品/ICPドライ#2/3インチ/'
    FileName2 = 'ICPドライ2号機_3ｲﾝﾁ_HL13★xx_ﾒｻﾄﾞﾗｲｴｯﾁﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsm'

    Excel_File_List = []

    # 檔案來源1
    #for file in glob.glob(FilePath1 + FileName1):
    #    if '$' not in file:
    #        dt = strftime("%Y-%m-%d %H:%M:%S", localtime(os.path.getmtime(file)))
    #        Excel_File_List.append([file, dt])
    # 檔案來源2
    for file in glob.glob(FilePath2 + FileName2):
        if '$' not in file:
            dt = strftime("%Y-%m-%d %H:%M:%S", localtime(os.path.getmtime(file)))
            Excel_File_List.append([file, dt])

    if len(Excel_File_List) > 0:
        # 先複製所有檔案到本地資料夾
        for file_info in Excel_File_List:
            shutil.copy(file_info[0], '../DataFile/013_MESA/')
            print(f"Copied file: {file_info[0]} to ../DataFile/013_MESA/")
        # 逐一處理每個 Excel 檔案
        for file_info in Excel_File_List:
            process_excel_file(file_info[0], Log_File)
    else:
        Log.Log_Error(Log_File, 'No Excel files found matching the specified patterns.')
        sys.exit()

if __name__ == '__main__':
    Main()

Log.Log_Info(Log_File, 'Program End')
