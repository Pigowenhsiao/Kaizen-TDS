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
Operation = 'PIX-DML_SiN_Step'
TestStation = 'PIX-DML'


########## Logの設定 ##########

# ----- ログファイルの作成 -----
Log_Folder_Name = str(date.today())
if not os.path.exists("../Log/" + Log_Folder_Name):
    os.makedirs("../Log/" + Log_Folder_Name)

Log_File = '../Log/' + Log_Folder_Name + '/042_PIX-DML.log'
Log.Log_Info(Log_File, 'Program Start')


########## シート名の定義 ##########
Data_Sheet_Name = 'データ'


########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
# Output_filepath = '../XML/042_PIX-DML/'


########## 取得するデータの列番号を定義 ##########
Col_PIX_Start_Date_Time = 0
Col_PIX_Operator = 1
Col_PIX_Equipment = 2
Col_PIX_Serial_Number = 3
Col_Step = [4,5,6,7,8]


########## 取得した項目と型の対応表を定義 ##########
key_type = {
    'key_Part_Number': str,
    'key_Serial_Number': str,
    'key_PIX1_Start_Date_Time': str,
    'key_PIX1_Operator': str,
    'key_PIX1_Equipment': str,
    'key_PIX1_Step1': float,
    'key_PIX1_Step2': float,
    'key_PIX1_Step3': float,
    'key_PIX1_Step_Ave': float,
    'key_PIX1_Step_3sigma': float,
    'key_PIX2_Start_Date_Time': str,
    'key_PIX2_Operator': str,
    'key_PIX2_Equipment': str,
    'key_PIX2_Step1': float,
    'key_PIX2_Step2': float,
    'key_PIX2_Step3': float,
    'key_PIX2_Step_Ave': float,
    'key_PIX2_Step_3sigma': float,
    'key_PIX3_Start_Date_Time': str,
    'key_PIX3_Operator': str,
    'key_PIX3_Equipment': str,
    'key_PIX3_Step1': float,
    'key_PIX3_Step2': float,
    'key_PIX3_Step3': float,
    'key_PIX3_Step_Ave': float,
    'key_PIX3_Step_3sigma': float,
    'key_PIX_Pre_Start_Date_Time': str,
    'key_PIX_Pre_Operator': str,
    'key_PIX_Pre_Equipment': str,
    'key_PIX_Pre_Step1' : float,
    'key_PIX_Pre_Step2' : float,
    'key_PIX_Pre_Step3' : float,
    'key_PIX_Pre_Step_Ave' : float,
    'key_PIX_Pre_Step_3sigma' : float,
    "key_STARTTIME_SORTED_PIX1" : float,
    "key_SORTNUMBER_PIX1" : float
}


########## 各ファイルから取得したデータを格納する二次元リスト(P1=PIX1, P2=PIX2, P3=PIX3, Pr=Pre, の、各測定項目5箇所) ##########
PIX_Data_List = list()


########## ロット番号とPIX_Data_Listの添え字を対応させる辞書 ##########
List_Index_Lot = dict()

# ----- 次の開始行数をメモ ----- #
Next_StartNumber = [0] * 4

def Main(FilePath, FileName, TextFile, PIX):


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
    Excel_File = shutil.copy(Excel_file_list[0][0], '../DataFile/042_PIX-DML/')


    ########## DaraFrameの作成 ##########

    # ----- 取得開始行の取り出し -----
    Log.Log_Info(Log_File, 'Get The Starting Row Count')
    Start_Number = Row_Number_Func.start_row_number(TextFile) - 50

    # ----- ExcelデータをDataFrameとして取得 -----
    Log.Log_Info(Log_File, 'Read Excel')
    
    
 
    # ----- それぞれ使う列だけ残す -----
    if PIX == "PIX1": 
        df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="C:V", skiprows=range(Start_Number))
        df = df.iloc[:, [0,1,2,4,15,16,17,18,19]]
    elif PIX == "PIX2": 
        df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="C:U", skiprows=range(Start_Number))
        df = df.iloc[:, [0,1,2,4,14,15,16,17,18]]
    elif PIX == "PIX3": 
        df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="C:X", skiprows=range(Start_Number))
        df = df.iloc[:, [0,1,2,4,17,18,19,20,21]]
    else: 
        df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="C:L", skiprows=range(Start_Number))
        df = df.iloc[:, [0,1,2,4,5,6,7,8,9]]

    # ----- 列番号の振り直し -----
    Log.Log_Info(Log_File, 'Setting Columns Number')
    df.columns = range(df.shape[1])

    # ----- pd.TimeStamp型以外のデータをnp.nanで置き換える -----
    for i in range(df.shape[0]):
        if not isinstance(df.iloc[i, 0], (pd.Timestamp, datetime)):
            df.iloc[i, 0] = np.nan
        #if df.iloc[i, 1] == time():
        #    df.iloc[i, 1] = np.nan

    # ----- 00:00:00のデータをnp.nanで置き換える -----
    #df = df.replace(time(), np.nan)

    # ----- 末尾から欠損のデータを落としていく -----
    Getting_Row = len(df) - 1
    while Getting_Row >= 0 and df.isnull().any(axis=1)[Getting_Row]:
        Getting_Row -= 1

    df = df[:Getting_Row + 1]

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

    # ----- 次の開始行数を保持する変数の定義 -----
    Next_Row_Number = Start_Number

    # ----- dfのindexリスト -----
    df_idx = df.index.values

    # ----- 次の開始行数を保持する変数の定義 -----
    if len(df_idx) == 0:
        add_num = 0
    else:
        add_num = df_idx[-1]
    if PIX == "PIX1":
        Next_StartNumber[0] = Start_Number + add_num + 1
    elif PIX == "PIX2":
        Next_StartNumber[1] = Start_Number + add_num + 1
    elif PIX == "PIX3":
        Next_StartNumber[2] = Start_Number + add_num + 1
    else:
        Next_StartNumber[3] = Start_Number + add_num + 1


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
        Serial_Number = str(df.iloc[Row_Number, Col_PIX_Serial_Number])
        if Serial_Number == "nan":
            Log.Log_Error(Log_File, "Lot Error\n")
            Row_Number += 1
            continue

        # ----- Primeに接続し、ロット番号に対応する品名を取り出す -----
        if PIX == "PIX1":
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
            print('Part_Number:',Part_Number)
            # ----- 品名が LDアレイ_ であれば処理を行わない -----
            if Part_Number == 'LDアレイ_':
                Row_Number += 1
                continue

            # ----- 現在持っているロット番号を格納する1次元リストの作成 -----
            PIX_Data_List.append([0] * 39)

            # ----- 上記リストの添え字とロット番号の対応を辞書に格納 -----
            List_Index_Lot[Serial_Number] = len(PIX_Data_List) - 1

            # ----- データを書き込む位置を指定 -----
            col_index = 0
            index = -1

        else:

            # ----- PIX2, PIX3の場合は、Serial_Numberをキーとしてリストの添え字を抜き出す -----
            if Serial_Number not in List_Index_Lot.keys():
                Log.Log_Error(Log_File, Serial_Number + ' : ' + "Not in Dictionary\n")
                Row_Number += 1
                continue

            index = List_Index_Lot[Serial_Number]

            # ----- リストの書き換えが発生する範囲の左端を定義 -----
            if PIX == "PIX2":
                col_index = 11
            elif PIX == 'PIX3':
                col_index = 19
            else:
                col_index = 27


        # ----- データの書き込み -----
        PIX_Data_List[index][col_index] = Convert_Date.Edit_Date(df.iloc[Row_Number, Col_PIX_Start_Date_Time])
        PIX_Data_List[index][col_index + 1] = df.iloc[Row_Number, Col_PIX_Operator]
        # ----- ロット番号と品名はPIX1のときしか書き込まない -----
        if PIX == "PIX1":
            PIX_Data_List[index][col_index + 2] = df.iloc[Row_Number, Col_PIX_Serial_Number]
            PIX_Data_List[index][col_index + 3] = Part_Number
            PIX_Data_List[index][col_index + 4] = Nine_Serial_Number
            col_index += 3
        PIX_Data_List[index][col_index + 2] = df.iloc[Row_Number, Col_PIX_Equipment][-2:]
        PIX_Data_List[index][col_index + 3] = df.iloc[Row_Number, Col_Step[0]]
        PIX_Data_List[index][col_index + 4] = df.iloc[Row_Number, Col_Step[1]]
        PIX_Data_List[index][col_index + 5] = df.iloc[Row_Number, Col_Step[2]]
        PIX_Data_List[index][col_index + 6] = df.iloc[Row_Number, Col_Step[3]]
        PIX_Data_List[index][col_index + 7] = df.iloc[Row_Number, Col_Step[4]]

        Log.Log_Error(Log_File, Serial_Number + ' : ' + "Check OK\n")
        Row_Number+=1

        # ----- 対象ロットの行数を格納 -----
        if PIX == "PIX1":
            PIX_Data_List[index][-4] = Row_Number + 1
        elif PIX == "PIX2":
            PIX_Data_List[index][-3] = Row_Number + 1
        elif PIX == "PIX3":
            PIX_Data_List[index][-2] = Row_Number + 1
        else:
            PIX_Data_List[index][-1] = Row_Number + 1


    ########## XML変換処理前準備 ##########
    if PIX == "PIX_Pre":

        # ----- 全データの処理 -----
        Log.Log_Info(Log_File, "Data Organization")

        Last_index = float('inf')
        for i in range(len(PIX_Data_List)):
            # ----- 0が含まれていた場合、そのロットに対する情報が欠落している(PIX2, PIX3のデータが足りていないと見なす) -----
            if 0 in PIX_Data_List[i]:
                Log.Log_Error(Log_File, "Data Incompleteness\n")
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
                'key_PIX3_Step_3sigma': PIX_Data_List[i][26],
                'key_PIX_Pre_Start_Date_Time': PIX_Data_List[i][27],
                'key_PIX_Pre_Operator': PIX_Data_List[i][28],
                'key_PIX_Pre_Equipment': PIX_Data_List[i][29],
                'key_PIX_Pre_Step1': PIX_Data_List[i][30],
                'key_PIX_Pre_Step2': PIX_Data_List[i][31],
                'key_PIX_Pre_Step3': PIX_Data_List[i][32],
                'key_PIX_Pre_Step_Ave': PIX_Data_List[i][33],
                'key_PIX_Pre_Step_3sigma': PIX_Data_List[i][34]
            }


            ########## 日付フォーマットの変換 ##########

            # ----- 指定したフォーマットに変換出来たか確認 -----
            if len(data_dict["key_PIX1_Start_Date_Time"]) != 19 or len(data_dict["key_PIX2_Start_Date_Time"]) != 19 or len(data_dict["key_PIX3_Start_Date_Time"]) != 19:
                Log.Log_Error(Log_File, data_dict["key_Serial_Number"] + ' : ' + "Date Error\n")
                Row_Number += 1
                continue


            ########## STARTTIME_SORTEDの追加 ##########

            # ----- 日付をExcel時間に変換する -----
            date = datetime.strptime(str(data_dict["key_PIX1_Start_Date_Time"]).replace('T', ' ').replace('.', ':'), "%Y-%m-%d %H:%M:%S")
            date_excel_number = int(str(date - datetime(1899, 12, 30)).split()[0])

            # 行数を取得し、[行数/10^6]を行う
            excel_row = Start_Number + PIX_Data_List[i][34]
            excel_row_div = excel_row / 10 ** 6

            # unix_timeに上記の値を加算する
            date_excel_number += excel_row_div

            # data_dictに登録する
            data_dict["key_STARTTIME_SORTED_PIX1"] = date_excel_number
            data_dict["key_SORTNUMBER_PIX1"] = excel_row


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
                            ',Testdate=' + data_dict["key_PIX1_Start_Date_Time"] + '.xml'

            # ----- XMLファイルの作成 -----
            Log.Log_Info(Log_File, 'Excel File To XML File Conversion'+Output_filepath + XML_File_Name )
            f = open(Output_filepath + XML_File_Name, 'w', encoding="utf-8")

            f.write('<?xml version="1.0" encoding="utf-8"?>' + '\n' +
                    '<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' + '\n' +
                    '       <Result startDateTime=' + '"' + data_dict["key_PIX1_Start_Date_Time"].replace(".", ":") + '"' + ' Result="Passed">' + '\n' +
                    '               <Header SerialNumber=' + '"' + data_dict["key_Serial_Number"] + '"' + ' PartNumber=' + '"' + data_dict["key_Part_Number"] + '"' + ' Operation=' + '"' + Operation + '"' + ' TestStation=' + '"' + TestStation + '"' + ' Operator=' + '"' + data_dict["key_PIX1_Operator"] + '"' + ' StartTime=' + '"' + data_dict["key_PIX1_Start_Date_Time"].replace(".", ":") + '"' + ' Site=' + '"' + Site + '"' + ' LotNumber=' + '"' + data_dict["key_Serial_Number"] + '"/>' + '\n' +
                    '               <HeaderMisc>' + '\n' +
                    '                   <Item Description="PIX1_Operator">' + str(data_dict["key_PIX1_Operator"]) + '</Item>' + '\n'
                    '                   <Item Description="PIX2_Operator">' + str(data_dict["key_PIX2_Operator"]) + '</Item>' + '\n'
                    '                   <Item Description="PIX3_Operator">' + str(data_dict["key_PIX3_Operator"]) + '</Item>' + '\n'
                    '                   <Item Description="PIX_Pre_Operator">' + str(data_dict["key_PIX_Pre_Operator"]) + '</Item>' + '\n'
                    '               </HeaderMisc>' + '\n' +
                    '\n'
                    '               <TestStep Name="PIX1" startDateTime=' + '"' + data_dict["key_PIX1_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step1" Units="nm" Value=' + '"' + str(data_dict["key_PIX1_Step1"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step2" Units="nm" Value=' + '"' + str(data_dict["key_PIX1_Step2"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step3" Units="nm" Value=' + '"' + str(data_dict["key_PIX1_Step3"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step_Ave" Units="nm" Value=' + '"' + str(data_dict["key_PIX1_Step_Ave"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step_3sigma" Units="nm" Value=' + '"' + str(data_dict["key_PIX1_Step_3sigma"]) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '\n'
                    '               <TestStep Name="PIX2" startDateTime=' + '"' + data_dict["key_PIX2_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step1" Units="nm" Value=' + '"' + str(data_dict["key_PIX2_Step1"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step2" Units="nm" Value=' + '"' + str(data_dict["key_PIX2_Step2"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step3" Units="nm" Value=' + '"' + str(data_dict["key_PIX2_Step3"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step_Ave" Units="nm" Value=' + '"' + str(data_dict["key_PIX2_Step_Ave"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step_3sigma" Units="nm" Value=' + '"' + str(data_dict["key_PIX2_Step_3sigma"]) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '\n'
                    '               <TestStep Name="PIX3" startDateTime=' + '"' + data_dict["key_PIX3_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step1" Units="nm" Value=' + '"' + str(data_dict["key_PIX3_Step1"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step2" Units="nm" Value=' + '"' + str(data_dict["key_PIX3_Step2"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step3" Units="nm" Value=' + '"' + str(data_dict["key_PIX3_Step3"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step_Ave" Units="nm" Value=' + '"' + str(data_dict["key_PIX3_Step_Ave"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step_3sigma" Units="nm" Value=' + '"' + str(data_dict["key_PIX3_Step_3sigma"]) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '\n'
                    '               <TestStep Name="PIX_Pre" startDateTime=' + '"' + data_dict["key_PIX_Pre_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step1" Units="nm" Value=' + '"' + str(data_dict["key_PIX_Pre_Step1"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step2" Units="nm" Value=' + '"' + str(data_dict["key_PIX_Pre_Step2"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step3" Units="nm" Value=' + '"' + str(data_dict["key_PIX_Pre_Step3"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step_Ave" Units="nm" Value=' + '"' + str(data_dict["key_PIX_Pre_Step_Ave"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Step_3sigma" Units="nm" Value=' + '"' + str(data_dict["key_PIX_Pre_Step_3sigma"]) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '\n'
                    '               <TestStep Name="SORTED_DATA" startDateTime=' + '"' + data_dict["key_PIX1_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value=' + '"' + str(data_dict["key_STARTTIME_SORTED_PIX1"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value=' + '"' + str(data_dict["key_SORTNUMBER_PIX1"]) + '"/>' + '\n' +
                    '                   <Data DataType="String" Name="LotNumber_5" Value=' + '"' + str(data_dict["key_Serial_Number"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                    '                   <Data DataType="String" Name="LotNumber_9" Value=' + '"' + str(data_dict["key_LotNumber_9"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '\n'
                    '               <TestEquipment>' + '\n' +
                    '                   <Item DeviceName="DryEtch" DeviceSerialNumber=' + '"' + str(data_dict['key_PIX1_Equipment']) + '"/>' + '\n' +
                    '               </TestEquipment>' + '\n' +
                    '\n'
                    '               <ErrorData/>' + '\n' +
                    '               <FailureData/>' + '\n' +
                    '               <Configuration/>' + '\n' +
                    '       </Result>' + '\n' +
                    '</Results>'
                    )
            f.close()

            Log.Log_Info(Log_File, data_dict["key_Serial_Number"] + ' : ' + "OK\n")
            Last_index = i


        ########## 次の開始行数の書き込み ##########
        Log.Log_Info(Log_File, 'Write the next starting line number')
        if Last_index != float('inf'):
            Row_Number_Func.next_start_row_number("PIX1_StartRow.txt", Next_StartNumber[0])
            Row_Number_Func.next_start_row_number("PIX2_StartRow.txt", Next_StartNumber[1])
            Row_Number_Func.next_start_row_number("PIX3_StartRow.txt", Next_StartNumber[2])
            Row_Number_Func.next_start_row_number("PIXPre_StartRow.txt", Next_StartNumber[3])

            # ----- 最終行を書き込んだファイルをGドライブに転送 -----
            shutil.copy("PIX1_StartRow.txt", 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/042_PIX-DML/13_ProgramUsedFile/')
            shutil.copy("PIX2_StartRow.txt", 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/042_PIX-DML/13_ProgramUsedFile/')
            shutil.copy("PIX3_StartRow.txt", 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/042_PIX-DML/13_ProgramUsedFile/')
            shutil.copy("PIXPre_StartRow.txt", 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/042_PIX-DML/13_ProgramUsedFile/')


if __name__ == '__main__':

    Main("Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ4号機/DML/", "*★直変系_PIX1ｴｯﾁ_ﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsm", "PIX1_StartRow.txt", "PIX1")
    Main("Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ4号機/DML/", "*★直変系_PIX2ｴｯﾁ_ﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsm", "PIX2_StartRow.txt", "PIX2")
    Main("Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ4号機/DML/", "*★直変系_PIX3ｴｯﾁ_ﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsm", "PIX3_StartRow.txt", "PIX3")
    Main("Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ4号機/DML/", "*★直変系_P電極蒸着前PIX段差_ﾌﾟﾛｸﾞﾗﾑｼｰﾄ*.xlsx", "PIXPre_StartRow.txt", "PIX_Pre")


# ----- Main処理の終了 -----
Log.Log_Info(Log_File, 'Program End')