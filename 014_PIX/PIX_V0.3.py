#!/usr/bin/env python3
"""
本程式利用 configparser 模組讀取 PIX_Config.ini 設定檔，
並根據設定檔參數執行原始流程：
    1. 從指定路徑尋找 Excel 檔案，複製至本地資料夾。
    2. 讀取 Excel 中特定工作表與欄位的數據，進行前處理與檢查。
    3. 依據不同 PIX 類型 (PIX1、PIX2、PIX3) 將數據整合到同一資料結構中。
    4. 當處理到 PIX3 時，將整合後的數據轉換為 XML 格式，並輸出至指定路徑。
    5. 更新下次執行所需的起始行數檔案並複製至指定共用資料夾。

注意：本程式中引用的 Log、SQL、Check、Convert_Date 與 Row_Number_Func 模組，
均需放置於 ../MyModule 路徑下，請確保模組存在且功能正確。
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

# 加入自訂模組目錄，並引入相關模組（假設這些模組已經存在）
sys.path.append('../MyModule')
import Log
import SQL
import Check
import Convert_Date
import Row_Number_Func

# ---------------------------
# 輔助函數：將以逗號分隔的字串轉為整數清單
def parse_int_list(s):
    return [int(item.strip()) for item in s.split(",") if item.strip() != ""]

# ---------------------------
# 讀取 INI 設定檔
config = configparser.ConfigParser(interpolation=None)
with open('./PIX_Config.ini', 'r', encoding='utf-8') as config_file:  # 指定使用 utf-8 編碼
    config.read_file(config_file)

# ---------------------------
# 從 [General] 區段讀取全局參數
SITE = config.get("General", "Site")
PRODUCT_FAMILY = config.get("General", "ProductFamily")
OPERATION = config.get("General", "Operation")
TEST_STATION = config.get("General", "TestStation")
DATA_SHEET_NAME = config.get("General", "DataSheetName")

# ---------------------------
# 從 [Logging] 區段讀取日誌相關設定
LOG_BASE_DIR = config.get("Logging", "LogBaseDir")
LOG_FILE_NAME = config.get("Logging", "LogFileName")

# 建立以今日日期為名稱的日誌資料夾
log_folder_name = str(date.today())
log_folder_path = os.path.join(LOG_BASE_DIR, log_folder_name)
if not os.path.exists(log_folder_path):
    os.makedirs(log_folder_path)
LOG_FILE = os.path.join(log_folder_path, LOG_FILE_NAME)
Log.Log_Info(LOG_FILE, 'Program Start')

# ---------------------------
# 從 [XML] 區段讀取 XML 輸出路徑等設定
OUTPUT_FILE_PATH = config.get("XML", "OutputFilePath")

# ---------------------------
# 從 [DateFormat] 區段讀取日期格式設定
INPUT_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"  # 直接在程式中定義日期格式
OUTPUT_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# ---------------------------
# 從 [Columns] 區段讀取 Excel 欄位設定
Col_PIX_Start_Date_Time = config.getint("Columns", "PIX_Start_Date_Time")
Col_PIX_Operator = config.getint("Columns", "PIX_Operator")
Col_PIX_Equipment = config.getint("Columns", "PIX_Equipment")
Col_PIX_Serial_Number = config.getint("Columns", "PIX_Serial_Number")
Col_PIX1_Step = parse_int_list(config.get("Columns", "PIX1_Step"))
Col_PIX2_Step = parse_int_list(config.get("Columns", "PIX2_Step"))
Col_PIX3_Step = parse_int_list(config.get("Columns", "PIX3_Step"))

# ---------------------------
# 全局變數：存放從各 Excel 檔讀取後的數據與相關索引
PIX_Data_List = list()     # 存放每筆資料（包含 PIX1、PIX2、PIX3 數據）的二維列表
List_Index_Lot = dict()    # 用於映射序號 (Serial Number) 與 PIX_Data_List 的索引
Next_StartNumber = [0] * 3 # 儲存 PIX1、PIX2、PIX3 下一次讀取數據的起始行數

# ---------------------------
# 資料型態檢查用的字典 (保持與原始程式碼相同)
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
# 主流程函數：依據指定 PIX 類型 (PIX1、PIX2、PIX3) 執行檔案讀取、數據處理與 XML 輸出
def Main(FilePath, FileNamePattern, TextFile, PIX):
    # 1. 複製最新的 Excel 檔案至本地資料夾
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
    # 從 INI 的 PIX 區段讀取本地資料存放路徑
    local_data_dir = config.get(PIX, "LocalDataFileDir")
    Excel_File = shutil.copy(Excel_file_list[0][0], local_data_dir)
    
    # 2. 讀取 Excel 檔案並建立 DataFrame
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
    one_month_ago = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) - relativedelta(months=2)
    df = df[(df[0] >= one_month_ago)]
    # 如果是 PIX3，則印出 DataFrame 的內容
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
    # 記錄下一次起始行數
    if PIX == "PIX1":
        Next_StartNumber[0] = Start_Number + add_num + 1
    elif PIX == "PIX2":
        Next_StartNumber[1] = Start_Number + add_num + 1
    else:
        Next_StartNumber[2] = Start_Number + add_num + 1
    df_idx = df.index.values

    # 根據 PIX 類型選擇對應的 Step 欄位設定
    if PIX == "PIX1":
        Col_Step = Col_PIX1_Step
    elif PIX == "PIX2":
        Col_Step = Col_PIX2_Step
    elif PIX == "PIX3":
        Col_Step = Col_PIX3_Step
    else:
        Log.Log_Error(LOG_FILE, "Unknown PIX type: " + PIX)
        return

    # 3. 逐行讀取 DataFrame 資料並進行資料檢查與存放
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
        # 寫入數據到 PIX_Data_List (依照 INI 設定的欄位順序)
        print("PIX",PIX + "\n")
        PIX_Data_List[index][col_index] = Convert_Date.Edit_Date(df.iloc[Row_Number, Col_PIX_Start_Date_Time])
        PIX_Data_List[index][col_index + 1] = df.iloc[Row_Number, Col_PIX_Operator]
        print("After assigning Operator:", PIX_Data_List[index]+"\n")
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

        # 記錄每個 PIX 資料的處理行數
        if PIX == "PIX1":
            PIX_Data_List[index][-3] = Row_Number
        elif PIX == "PIX2":
            PIX_Data_List[index][-2] = Row_Number
        else:
            PIX_Data_List[index][-1] = Row_Number

    # 4. 當處理到 PIX3 時，進行 XML 檔案轉換輸出
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
            # 日期格式檢查
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
                        '                   <Item DeviceName="DryEtch" DeviceSerialNumber="' + str(data_dict['key_PIX1_Equipment']) + '"/>' + '\n' +
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
            # 從 [FileTransfer] 區段讀取起始行數檔案轉存路徑
            start_row_file_base = config.get("FileTransfer", "StartRowFileBaseDir")
            shutil.copy("PIX1_StartRow.txt", start_row_file_base)
            shutil.copy("PIX2_StartRow.txt", start_row_file_base)
            shutil.copy("PIX3_StartRow.txt", start_row_file_base)

# ---------------------------
# 主程式入口：根據 INI 中 [PIX1]、[PIX2]、[PIX3] 區段的設定進行處理
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
