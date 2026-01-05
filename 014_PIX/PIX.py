#!/usr/bin/env python3
"""
本プログラムは configparser モジュールを利用して PIX_Config.ini 設定ファイルを読み込み、
設定ファイルのパラメータに基づいて以下の基本的な処理を実行します：
    1. 指定されたパスから Excel ファイルを検索し、ローカルフォルダにコピーします。
    2. Excel の特定のシートとカラムからデータを読み込み、前処理とチェックを行います。
    3. 異なる PIX タイプ (PIX1、PIX2、PIX3) に基づいて、データを統合した同一のデータ構造にまとめます。
    4. PIX3 の処理時に、統合されたデータを XML 形式に変換し、指定されたパスに出力します。
    5. 次回の実行に必要な開始行番号のファイルを更新し、指定された共有フォルダにコピーします。

注意：本プログラムで使用される Log、SQL、Check、Convert_Date および Row_Number_Func モジュールは、
すべて ../MyModule パスに配置されている必要があり、モジュールの存在と機能が正しいことを確認してください。
"""

import os
import sys
import glob
import shutil
import logging
from time import strftime, localtime
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
import configparser

import xlrd
import numpy as np
import pandas as pd
import openpyxl as px

# 自作モジュールディレクトリを追加し、必要なモジュールをインポートする（既に存在するものと仮定）
sys.path.append('../MyModule')
import Log
import SQL
import Check
import Convert_Date
import Row_Number_Func

# ---------------------------
# 補助関数：カンマで区切られた文字列を整数リストに変換する関数
def parse_int_list(s):
    return [int(item.strip()) for item in s.split(",") if item.strip() != ""]

# ---------------------------
# INI 設定ファイルの読み込み
config = configparser.ConfigParser(interpolation=None)
with open('./PIX_Config.ini', 'r', encoding='utf-8') as config_file:  # UTF-8 エンコーディングを指定
    config.read_file(config_file)

# ---------------------------
# [General] セクションからグローバルパラメータを取得
SITE = config.get("General", "Site")
PRODUCT_FAMILY = config.get("General", "ProductFamily")
OPERATION = config.get("General", "Operation")
TEST_STATION = config.get("General", "TestStation")
DATA_SHEET_NAME = config.get("General", "DataSheetName")

# ---------------------------
# [Logging] セクションからログ関連の設定を取得
LOG_BASE_DIR = config.get("Logging", "LogBaseDir")
LOG_FILE_NAME = config.get("Logging", "LogFileName")

# 今日の日付を名前とするログフォルダを作成
log_folder_name = str(date.today())
log_folder_path = os.path.join(LOG_BASE_DIR, log_folder_name)
if not os.path.exists(log_folder_path):
    os.makedirs(log_folder_path)
LOG_FILE = os.path.join(log_folder_path, LOG_FILE_NAME)
Log.Log_Info(LOG_FILE, 'Program Start')

# ---------------------------
# [XML] セクションから XML 出力パスなどの設定を取得
OUTPUT_FILE_PATH = config.get("XML", "OutputFilePath")

# ---------------------------
# [DateFormat] セクションから日付フォーマット設定を取得
INPUT_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"  # プログラム内で直接日付フォーマットを定義
OUTPUT_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# ---------------------------
# [Columns] セクションから Excel のカラム設定を取得
Col_PIX_Start_Date_Time = config.getint("Columns", "PIX_Start_Date_Time")
Col_PIX_Operator = config.getint("Columns", "PIX_Operator")
Col_PIX_Equipment = config.getint("Columns", "PIX_Equipment")
Col_PIX_Serial_Number = config.getint("Columns", "PIX_Serial_Number")
Col_PIX1_Step = parse_int_list(config.get("Columns", "PIX1_Step"))
Col_PIX2_Step = parse_int_list(config.get("Columns", "PIX2_Step"))
Col_PIX3_Step = parse_int_list(config.get("Columns", "PIX3_Step"))

# ---------------------------
# グローバル変数：各 Excel ファイルから読み込んだデータと関連するインデックスを格納
PIX_Data_List = list()     # 各データ（PIX1、PIX2、PIX3 のデータを含む）の2次元リスト
List_Index_Lot = dict()    # シリアル番号と PIX_Data_List のインデックスをマッピングするための辞書
Next_StartNumber = [0] * 3 # PIX1、PIX2、PIX3 の次回読み込み開始行番号を格納するリスト

# ---------------------------
# データ型チェック用の辞書 (元のプログラムと同じ内容)
key_type = {
    'key_Part_Number' : str,
    'key_Serial_Number' : str,
    'key_PIX1_Start_Date_Time' : str,
    'key_PIX1_Operator' : str,
    'key_PIX1_Equipment' : str,
    'key_PIX1_Step1' : float,
    'key_PIX1_Step2' : float,
    'key_PIX1_Step3' : float,
    'key_PIX1_Step_Ave' : float,
    'key_PIX1_Step_3sigma' : float,
    'key_PIX2_Start_Date_Time' : str,
    'key_PIX2_Operator' : str,
    'key_PIX2_Step1' : float,
    'key_PIX2_Step2' : float,
    'key_PIX2_Step3' : float,
    'key_PIX2_Step_Ave' : float,
    'key_PIX2_Step_3sigma' : float,
    'key_PIX3_Start_Date_Time' : str,
    'key_PIX3_Operator' : str,
    'key_PIX3_Step1' : float,
    'key_PIX3_Step2' : float,
    'key_PIX3_Step3' : float,
    'key_PIX3_Step_Ave' : float,
    'key_PIX3_Step_3sigma' : float,
    "key_STARTTIME_SORTED_PIX1" : float,
    "key_SORTNUMBER_PIX1" : float,
    "key_LotNumber_9": str
}

# ---------------------------
# メイン処理関数：指定された PIX タイプ（PIX1、PIX2、PIX3）に応じてファイル読み込み、データ処理、XML 出力を実施
def Main(FilePath, FileNamePattern, TextFile, PIX):
    # 1. 最新の Excel ファイルをローカルフォルダにコピーする処理
    Log.Log_Info(LOG_FILE, 'Excel File Copy')
    Excel_file_list = []
    for file in glob.glob(os.path.join(FilePath, FileNamePattern)):
        if '$' not in file:
            dt = strftime("%Y-%m-%d %H:%M:%S", localtime(os.path.getmtime(file)))
            Excel_file_list.append([file, dt])
    Excel_file_list = sorted(Excel_file_list, key=lambda x: x[1], reverse=True)
    if not Excel_file_list:
        Log.Log_Error(LOG_FILE, "No Excel file found for " + PIX)
        return
    # INI の PIX セクションからローカルデータ保存パスを取得
    local_data_dir = config.get(PIX, "LocalDataFileDir")
    Excel_File = shutil.copy(Excel_file_list[0][0], local_data_dir)
    
    # 2. Excel ファイルを読み込み、DataFrame を作成する処理
    Log.Log_Info(LOG_FILE, 'Get The Starting Row Count')
    Start_Number = Row_Number_Func.start_row_number(TextFile) - 500
    Log.Log_Info(LOG_FILE, 'Read Excel')
    df = pd.read_excel(Excel_File, header=None, sheet_name=DATA_SHEET_NAME, usecols="C:X", skiprows=Start_Number)
    df = df.dropna(how='all')
    Log.Log_Info(LOG_FILE, 'Setting Columns Number')
    df.columns = range(df.shape[1])
    for i in range(df.shape[0]):
        if not isinstance(df.iloc[i, 0], (pd.Timestamp, datetime)):
            df.iloc[i, 0] = np.nan
    Getting_Row = len(df) - 1
    while Getting_Row >= 0 and df.isnull().any(axis=1)[Getting_Row]:
        Getting_Row -= 1
    df = df[:Getting_Row + 1]
    df[0] = pd.to_datetime(df[0])
    one_month_ago = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) - relativedelta(months=3)
    df = df[(df[0] >= one_month_ago)]
    # PIX3 の場合、DataFrame の内容を出力する
    if PIX == "PIX3":
        print("DataFrame Content for PIX3:")
        print(df)
    Log.Log_Info(LOG_FILE, 'Get DataFrame End Index Number\n')
    row_end = len(df)
    Row_Number = 0
    if len(df) == 0:
        add_num = 0
    else:
        add_num = df.index.values[-1]
    # 次回の開始行番号を記録
    if PIX == "PIX1":
        Next_StartNumber[0] = Start_Number + add_num + 1
    elif PIX == "PIX2":
        Next_StartNumber[1] = Start_Number + add_num + 1
    else:
        Next_StartNumber[2] = Start_Number + add_num + 1
    df_idx = df.index.values

    # PIX タイプに応じた Step カラムの設定を選択
    if PIX == "PIX1":
        Col_Step = Col_PIX1_Step
    elif PIX == "PIX2":
        Col_Step = Col_PIX2_Step
    elif PIX == "PIX3":
        Col_Step = Col_PIX3_Step
    else:
        Log.Log_Error(LOG_FILE, "Unknown PIX type: " + PIX)
        return

    # 3. DataFrame の各行を読み込み、データチェックと保存を行う処理
    while Row_Number < row_end:
        Log.Log_Info(LOG_FILE, "Blank Check")
        if pd.isnull(df.iloc[Row_Number, [0, 4]]).any():
            print(Row_Number)
            Log.Log_Error(LOG_FILE, " Blank Error in Columns 0 or 4\n")
            Row_Number += 1
            continue

        Log.Log_Info(LOG_FILE, 'Data Acquisition')
        data_dict = dict()
        Serial_Number = str(df.iloc[Row_Number, Col_PIX_Serial_Number])
        if Serial_Number == "nan":
            Log.Log_Error(LOG_FILE, "Lot Error\n")
            Row_Number += 1
            continue

        if PIX == "PIX1":
            conn, cursor = SQL.connSQL()
            if conn is None:
                Log.Log_Error(LOG_FILE, Serial_Number + ' : ' + 'Connection with Prime Failed')
                break
            Part_Number, Nine_Serial_Number = SQL.selectSQL(cursor, Serial_Number)
            SQL.disconnSQL(conn, cursor)
            #Part_Number="TEST-PARTNUBER:"+PIX 
            #Nine_Serial_Number ="TEST9NUNBERSBER:"+PIX
            if Part_Number is None:
                Log.Log_Error(LOG_FILE, Serial_Number + ' : ' + "PartNumber Error\n")
                Row_Number += 1
                continue
            if Part_Number == 'LDアレイ_':
                Row_Number += 1
                continue
            PIX_Data_List.append([0] * 30)
            List_Index_Lot[Serial_Number] = len(PIX_Data_List) - 1
            col_index = 0
            index = -1
        else:
            if Serial_Number not in List_Index_Lot.keys():
                Log.Log_Error(LOG_FILE, Serial_Number + ' : ' + "Not in Dictionary\n")
                Row_Number += 1
                continue
            index = List_Index_Lot[Serial_Number]
            if PIX == "PIX2":
                col_index = 11
            else:
                col_index = 19
        # INI で設定されたカラム順に従い、PIX_Data_List にデータを書き込む
        print("PIX", PIX + "\n")
        PIX_Data_List[index][col_index] = Convert_Date.Edit_Date(df.iloc[Row_Number, Col_PIX_Start_Date_Time])
        PIX_Data_List[index][col_index + 1] = df.iloc[Row_Number, Col_PIX_Operator]
        print("After assigning Operator:", PIX_Data_List[index],"\n")
        if PIX == "PIX1":
            PIX_Data_List[index][col_index + 2] = df.iloc[Row_Number, Col_PIX_Serial_Number]
            PIX_Data_List[index][col_index + 3] = Part_Number
            PIX_Data_List[index][col_index + 4] = Nine_Serial_Number
            col_index += 3
        PIX_Data_List[index][col_index + 2] = str(df.iloc[Row_Number, Col_PIX_Equipment])[-2:]
        PIX_Data_List[index][col_index + 3] = df.iloc[Row_Number, Col_Step[0]]
        PIX_Data_List[index][col_index + 4] = df.iloc[Row_Number, Col_Step[1]]
        PIX_Data_List[index][col_index + 5] = df.iloc[Row_Number, Col_Step[2]]
        PIX_Data_List[index][col_index + 6] = df.iloc[Row_Number, Col_Step[3]] if len(df.columns) > Col_Step[3] else None
        PIX_Data_List[index][col_index + 7] = df.iloc[Row_Number, Col_Step[4]] if len(df.columns) > Col_Step[4] else None
        print("After assigning Step5:", PIX_Data_List[index]+"\n")
        Log.Log_Info(LOG_FILE, Serial_Number + ' : ' + "Check OK\n")
        Row_Number += 1

        # 各 PIX データの処理行番号を記録
        if PIX == "PIX1":
            PIX_Data_List[index][-3] = Row_Number
        elif PIX == "PIX2":
            PIX_Data_List[index][-2] = Row_Number
        else:
            PIX_Data_List[index][-1] = Row_Number

    # 4. PIX3 の場合、XML ファイルへの変換出力処理を実施
    if PIX == "PIX3":
        Log.Log_Info(LOG_FILE, "Data Organization")
        Last_index = float('inf')
        for i in range(len(PIX_Data_List)):
            if 0 in PIX_Data_List[i]:
                Log.Log_Error(LOG_FILE, "Data Incompleteness\n")
                continue
            data_dict = {
                'key_Part_Number': PIX_Data_List[i][3],
                'key_Serial_Number': PIX_Data_List[i][2],
                'key_PIX1_Start_Date_Time': PIX_Data_List[i][0],
                'key_PIX1_Operator': PIX_Data_List[i][1],
                'key_LotNumber_9': PIX_Data_List[i][4],
                'key_PIX1_Equipment': PIX_Data_List[i][5],
                'key_PIX1_Step1': PIX_Data_List[i][6],
                'key_PIX1_Step2': PIX_Data_List[i][7],
                'key_PIX1_Step3': PIX_Data_List[i][8],
                'key_PIX1_Step_Ave': PIX_Data_List[i][9],
                'key_PIX1_Step_3sigma': PIX_Data_List[i][10],
                'key_PIX2_Start_Date_Time': PIX_Data_List[i][11],
                'key_PIX2_Operator': PIX_Data_List[i][12],
                'key_PIX2_Equipment': PIX_Data_List[i][13],
                'key_PIX2_Step1': PIX_Data_List[i][14],
                'key_PIX2_Step2': PIX_Data_List[i][15],
                'key_PIX2_Step3': PIX_Data_List[i][16],
                'key_PIX2_Step_Ave': PIX_Data_List[i][17],
                'key_PIX2_Step_3sigma': PIX_Data_List[i][18],
                'key_PIX3_Start_Date_Time': PIX_Data_List[i][19],
                'key_PIX3_Operator': PIX_Data_List[i][20],
                'key_PIX3_Equipment': PIX_Data_List[i][21],
                'key_PIX3_Step1': PIX_Data_List[i][22],
                'key_PIX3_Step2': PIX_Data_List[i][23],
                'key_PIX3_Step3': PIX_Data_List[i][24],
                'key_PIX3_Step_Ave': PIX_Data_List[i][25],
                'key_PIX3_Step_3sigma': PIX_Data_List[i][26]
            }
            # 日付フォーマットのチェック
            if len(data_dict["key_PIX1_Start_Date_Time"]) != 19 or len(data_dict["key_PIX2_Start_Date_Time"]) != 19 or len(data_dict["key_PIX3_Start_Date_Time"]) != 19:
                Log.Log_Error(LOG_FILE, data_dict["key_Serial_Number"] + ' : ' + "Date Error\n")
                Row_Number += 1
                continue
            date_obj = datetime.strptime(str(data_dict["key_PIX1_Start_Date_Time"]).replace('T', ' ').replace('.', ':'), "%Y-%m-%d %H:%M:%S")
            date_excel_number = int(str(date_obj - datetime(1899, 12, 30)).split()[0])
            excel_row = PIX_Data_List[i][27] + Start_Number
            excel_row_div = excel_row / 10 ** 6
            date_excel_number += excel_row_div
            data_dict["key_STARTTIME_SORTED_PIX1"] = date_excel_number
            data_dict["key_SORTNUMBER_PIX1"] = excel_row
            Log.Log_Info(LOG_FILE, "Check Data Type")
            Result = Check.Data_Type(key_type, data_dict)
            if Result == False:
                Log.Log_Error(LOG_FILE, data_dict["key_Serial_Number"] + ' : ' + "Data Error\n")
                Row_Number += 1
                continue
            XML_File_Name = 'Site=' + SITE + ',ProductFamily=' + PRODUCT_FAMILY + ',Operation=' + OPERATION + \
                            ',Partnumber=' + data_dict["key_Part_Number"] + ',Serialnumber=' + data_dict["key_Serial_Number"] + \
                            ',Testdate=' + data_dict["key_PIX1_Start_Date_Time"] + '.xml'
            Log.Log_Info(LOG_FILE, 'Excel File To XML File Conversion')
            with open(os.path.join(OUTPUT_FILE_PATH, XML_File_Name), 'w', encoding="utf-8") as f:
                f.write('<?xml version="1.0" encoding="utf-8"?>' + '\n' +
                        '<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' + '\n' +
                        '       <Result startDateTime="' + data_dict["key_PIX1_Start_Date_Time"].replace(".", ":") + '" Result="Passed">' + '\n' +
                        '               <Header SerialNumber="' + data_dict["key_Serial_Number"] + '" PartNumber="' + data_dict["key_Part_Number"] + '" Operation="' + OPERATION + '" TestStation="' + TEST_STATION + '" Operator="' + data_dict["key_PIX1_Operator"] + '" StartTime="' + data_dict["key_PIX1_Start_Date_Time"].replace(".", ":") + '" Site="' + SITE + '" LotNumber="' + data_dict["key_Serial_Number"] + '"/>' + '\n' +
                        '               <HeaderMisc>' + '\n' +
                        '                   <Item Description="PIX1_Operator">' + str(data_dict["key_PIX1_Operator"]) + '</Item>' + '\n' +
                        '                   <Item Description="PIX2_Operator">' + str(data_dict["key_PIX2_Operator"]) + '</Item>' + '\n' +
                        '                   <Item Description="PIX3_Operator">' + str(data_dict["key_PIX3_Operator"]) + '</Item>' + '\n' +
                        '               </HeaderMisc>' + '\n' +
                        '\n'
                        '               <TestStep Name="PIX1" startDateTime="' + data_dict["key_PIX1_Start_Date_Time"].replace(".", ":") + '" Status="Passed">' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step1" Units="nm" Value="' + str(data_dict["key_PIX1_Step1"]) + '"/>' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step2" Units="nm" Value="' + str(data_dict["key_PIX1_Step2"]) + '"/>' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step3" Units="nm" Value="' + str(data_dict["key_PIX1_Step3"]) + '"/>' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step_Ave" Units="nm" Value="' + str(data_dict["key_PIX1_Step_Ave"]) + '"/>' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step_3sigma" Units="nm" Value="' + str(data_dict["key_PIX1_Step_3sigma"]) + '"/>' + '\n' +
                        '               </TestStep>' + '\n' +
                        '\n'
                        '               <TestStep Name="PIX2" startDateTime="' + data_dict["key_PIX2_Start_Date_Time"].replace(".", ":") + '" Status="Passed">' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step1" Units="nm" Value="' + str(data_dict["key_PIX2_Step1"]) + '"/>' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step2" Units="nm" Value="' + str(data_dict["key_PIX2_Step2"]) + '"/>' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step3" Units="nm" Value="' + str(data_dict["key_PIX2_Step3"]) + '"/>' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step_Ave" Units="nm" Value="' + str(data_dict["key_PIX2_Step_Ave"]) + '"/>' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step_3sigma" Units="nm" Value="' + str(data_dict["key_PIX2_Step_3sigma"]) + '"/>' + '\n' +
                        '               </TestStep>' + '\n' +
                        '\n'
                        '               <TestStep Name="PIX3" startDateTime="' + data_dict["key_PIX3_Start_Date_Time"].replace(".", ":") + '" Status="Passed">' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step1" Units="nm" Value="' + str(data_dict["key_PIX3_Step1"]) + '"/>' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step2" Units="nm" Value="' + str(data_dict["key_PIX3_Step2"]) + '"/>' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step3" Units="nm" Value="' + str(data_dict["key_PIX3_Step3"]) + '"/>' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step_Ave" Units="nm" Value="' + str(data_dict["key_PIX3_Step_Ave"]) + '"/>' + '\n' +
                        '                   <Data DataType="Numeric" Name="Step_3sigma" Units="nm" Value="' + str(data_dict["key_PIX3_Step_3sigma"]) + '"/>' + '\n' +
                        '               </TestStep>' + '\n' +
                        '\n'
                        '               <TestStep Name="SORTED_DATA" startDateTime="' + data_dict["key_PIX1_Start_Date_Time"].replace(".", ":") + '" Status="Passed">' + '\n' +
                        '                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value="' + str(data_dict["key_STARTTIME_SORTED_PIX1"]) + '"/>' + '\n' +
                        '                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value="' + str(data_dict["key_SORTNUMBER_PIX1"]) + '"/>' + '\n' +
                        '                   <Data DataType="String" Name="LotNumber_5" Value="' + str(data_dict["key_Serial_Number"]) + '" CompOperation="LOG"/>' + '\n' +
                        '                   <Data DataType="String" Name="LotNumber_9" Value="' + str(data_dict["key_LotNumber_9"]) + '" CompOperation="LOG"/>' + '\n' +
                        '               </TestStep>' + '\n' +
                        '\n'
                        '               <TestEquipment>' + '\n' +
                        '                   <Item DeviceName="DryEtch" DeviceSerialNumber="' + str(data_dict["key_PIX1_Equipment"]) + '"/>' + '\n' +
                        '               </TestEquipment>' + '\n' +
                        '\n'
                        '               <ErrorData/>' + '\n' +
                        '               <FailureData/>' + '\n' +
                        '               <Configuration/>' + '\n' +
                        '       </Result>' + '\n' +
                        '</Results>'
                        )
            Log.Log_Info(LOG_FILE, data_dict["key_Serial_Number"] + ' : ' + "OK\n" + OUTPUT_FILE_PATH + XML_File_Name + '\n')
            Last_index = i

        Log.Log_Info(LOG_FILE, 'Write the next starting line number')
        if Last_index != float('inf'):
            Row_Number_Func.next_start_row_number("PIX1_StartRow.txt", Next_StartNumber[0])
            Row_Number_Func.next_start_row_number("PIX2_StartRow.txt", Next_StartNumber[1])
            Row_Number_Func.next_start_row_number("PIX3_StartRow.txt", Next_StartNumber[2])
            # [FileTransfer] セクションから開始行番号ファイルの転送先パスを取得
            start_row_file_base = config.get("FileTransfer", "StartRowFileBaseDir")
            shutil.copy("PIX1_StartRow.txt", start_row_file_base)
            shutil.copy("PIX2_StartRow.txt", start_row_file_base)
            shutil.copy("PIX3_StartRow.txt", start_row_file_base)

# ---------------------------
# エントリポイント：INI の [PIX1]、[PIX2]、[PIX3] セクションの設定に基づき処理を実施
if __name__ == '__main__':
    PIX1_SourceFilePath = config.get("PIX1", "SourceFilePath")
    PIX1_FileNamePattern = config.get("PIX1", "SourceFileNamePattern")
    PIX1_StartRowFile = config.get("PIX1", "StartRowFile")

    PIX2_SourceFilePath = config.get("PIX2", "SourceFilePath")
    PIX2_FileNamePattern = config.get("PIX2", "SourceFileNamePattern")
    PIX2_StartRowFile = config.get("PIX2", "StartRowFile")

    PIX3_SourceFilePath = config.get("PIX3", "SourceFilePath")
    PIX3_FileNamePattern = config.get("PIX3", "SourceFileNamePattern")
    PIX3_StartRowFile = config.get("PIX3", "StartRowFile")

    Main(PIX1_SourceFilePath, PIX1_FileNamePattern, PIX1_StartRowFile, "PIX1")
    Main(PIX2_SourceFilePath, PIX2_FileNamePattern, PIX2_StartRowFile, "PIX2")
    Main(PIX3_SourceFilePath, PIX3_FileNamePattern, PIX3_StartRowFile, "PIX3")

    Log.Log_Info(LOG_FILE, 'Program End')
