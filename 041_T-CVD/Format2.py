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
X = "999999"
Y = "999999"


########## Logの設定 ##########

# ----- ログファイルの作成 -----
Log_Folder_Name = str(date.today())
if not os.path.exists("../Log/" + Log_Folder_Name):
    os.makedirs("../Log/" + Log_Folder_Name)

Log_File = '../Log/' + Log_Folder_Name + '/041_T-CVD_Format2.log'
Log.Log_Info(Log_File, 'Program Start')


########## シート名の定義 ##########
Data_Sheet_Name = '着工記録'


########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
# Output_filepath = '../XML/041_T-CVD/Format2/'


########## 取得するデータの列番号を定義 ##########
Col_Start_Date_Time = 0
Col_CVDNo = 1
Col_Operator = 2
Col_TypeOfDeposition = 3
Col_Process = 4
Col_Recipe = 5
Col_Serial_Number = [6, 7, 8, 9, 10, 11, 12, 13]
Col_Thickness_Thickness = 14
Col_HEAD1_O2 = 15
Col_HEAD1_N2_O2 = 16
Col_HEAD1_N2_SiH4 = 17
Col_HEAD1_SiH4_4percent = 18
Col_HEAD1_PH3_1percent = 19
Col_HEAD1_SiH4_100percent = 20
Col_HEAD1_PH3_100percent = 21
Col_HEAD1_P_PSi = 22
Col_HEAD1_Ratio_O2_Gas = 23
Col_HEAD2_O2 = 24
Col_HEAD2_N2_SiH4 = 25
Col_HEAD2_N2_O2 = 26
Col_HEAD2_SiH4_4percent = 27
Col_HEAD2_PH3_1percent = 28
Col_HEAD2_SiH4_100percent = 29
Col_HEAD2_Ratio_O2 = 30
Col_Temp_Target = 31
Col_Temp_H1 = 32
Col_Temp_H2 = 33
Col_Temp_H3 = 34
Col_Temp_H4 = 35
Col_Temp_H5 = 36
Col_Speed_Speed = 37
Col_Position_Start = 38
Col_Position_Stop = 39
Col_Position_Delta = 40


########## 取得した項目と型の対応表を定義 ##########
key_type = {
    "key_Start_Date_Time": str,
    "key_CVDNo" : str,
    "key_Operator": str,
    "key_TypeOfDeposition" : str,
    "key_Process" : str,
    "key_Recipe": str,
    "key_Thickness_Thickness" : float,
    "key_HEAD1_O2" : float,
    "key_HEAD1_N2_O2" : float,
    "key_HEAD1_N2_SiH4" : float,
    "key_HEAD1_SiH4_4percent" : float,
    "key_HEAD1_PH3_1percent" : float,
    "key_HEAD1_SiH4_100percent" : float,
    "key_HEAD1_PH3_100percent" : float,
    "key_HEAD1_P_PSi" : float,
    "key_HEAD1_Ratio_O2Gas" : float,
    "key_HEAD2_O2" : float,
    "key_HEAD2_N2_SiH4" : float,
    "key_HEAD2_N2_O2" : float,
    "key_HEAD2_SiH4_4percent" : float,
    "key_HEAD2_PH3_1percent" : float,
    "key_HEAD2_SiH4_100percent" : float,
    "key_HEAD2_Ratio_O2" : float,
    "key_Temp_Target" : float,
    "key_Temp_H1" : float,
    "key_Temp_H2" : float,
    "key_Temp_H3" : float,
    "key_Temp_H4" : float,
    "key_Temp_H5" : float,
    "key_Speed_Speed" : float,
    "key_Position_Start" : float,
    "key_Position_Stop" : float,
    "key_Position_Delta" : float,
    'key_STARTTIME_SORTED' : float,
    "key_SORTNUMBER" : float
}


def Main():


    ########## Excelファイルをローカルにコピー ##########

    # ----- 正規表現で取出し、直近で変更があったファイルを取得する -----
    Log.Log_Info(Log_File, 'Excel File Copy')

    FilePath = 'Z:/CVD/T-CVD/'
    FileName = '*T-CVD#2_r1_着工来歴*.xlsx'
    Excel_File_List = []
    for file in glob.glob(FilePath + FileName):
        if '$' not in file:
            dt = strftime("%Y-%m-%d %H:%M:%S", localtime(os.path.getmtime(file)))
            Excel_File_List.append([file, dt])

    # ----- dt(更新日時)の降順で並び替える -----
    Excel_File_List = sorted(Excel_File_List, key=lambda x: x[1], reverse=True)
    Excel_File = shutil.copy(Excel_File_List[0][0], '../DataFile/041_T-CVD/')


    ########## DaraFrameの作成 ##########

    # ----- 取得開始行の取り出し -----
    Log.Log_Info(Log_File, 'Get The Starting Row Count')
    Start_Number = Row_Number_Func.start_row_number("Format2_StartRow.txt") - 500

    # ----- ExcelデータをDataFrameとして取得 -----
    Log.Log_Info(Log_File, 'Read Excel')
    df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="A:BP", skiprows=Start_Number)

    # ----- 不要な列を落とす -----
    df = df.drop([4,6,16,17,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41], axis=1)

    # ----- 列番号の振り直し -----
    Log.Log_Info(Log_File, 'Setting Columns Number')
    df.columns = range(df.shape[1])

    # ----- 末尾から欠損のデータを落としていく -----
    Getting_Row = len(df) - 1
    while Getting_Row >= 0 and df.iloc[Getting_Row, Col_Start_Date_Time] is np.nan:
        Getting_Row -= 1

    df = df[:Getting_Row -22]

    # ----- 次の開始行数をメモ -----
    Next_Start_Row = Start_Number + df.shape[0] + 1

    # ----- 日付欄に文字列が入っていたらNoneに置き換える -----
    for i in range(df.shape[0]):
        if type(df.iloc[i, 0]) is not pd.Timestamp:
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

        # ----- ヘッダー情報に空欄が1つでも存在すればTrueを返す -----
        Log.Log_Info(Log_File, "Blank Check")
        if str(df.iloc[Row_Number, Col_Start_Date_Time]) == "nan" or str(df.iloc[Row_Number, Col_Operator]) == "nan" or str(df.iloc[Row_Number, Col_Serial_Number]) == "nan" or str(df.iloc[Row_Number, Col_CVDNo]) == "nan":
            Log.Log_Error(Log_File, "Blank Error\n")
            Row_Number += 1
            continue


        ########## 現在処理を行っている行のデータの取得 ##########

        # ----- 取得したデータを格納するデータ構造(辞書)を作成 -----
        Log.Log_Info(Log_File, 'Data Acquisition')
        data_dict = dict()

        # ----- 対応したOperationを探すキーの取得 -----
        Process = str(df.iloc[Row_Number, Col_Process])

        # 対応したOperationの処理工程に飛ぶ
        if "BJ1" in Process or "ＢＪ１" in Process:
            Operation, TestStation = "BJ1_T-CVD-Thickness", "BJ1"
        elif "BJ2" in Process:
            Operation, TestStation = "BJ2_T-CVD-Thickness", "BJ2"
        elif "回折格子" in Process or "GRT" in Process:
            Operation, TestStation = "GRATING_T-CVD-Thickness", "GRATING"
        elif "メサ" in Process or "MESA" in Process:
            Operation, TestStation = "MESA_T-CVD-Thickness", "MESA"
        elif 'ﾊﾟｯｼﾍﾞｰｼｮﾝ' in Process or 'パッシベーション' in Process:
            Operation, TestStation = "PASSI_T-CVD-Thickness", "PASSI"
        elif 'PAD' in Process or 'pad' in Process:
            Operation, TestStation = "PAD_T-CVD-Thickness", "PAD"
        else:
            Log.Log_Error(Log_File, "Operation Error\n")
            Row_Number += 1
            continue


        ########## データ取得 #########

        data_dict = {
            "key_Start_Date_Time": df.iloc[Row_Number, Col_Start_Date_Time],
            "key_CVDNo": df.iloc[Row_Number, Col_CVDNo],
            "key_Operator": df.iloc[Row_Number, Col_Operator],
            "key_TypeOfDeposition": df.iloc[Row_Number, Col_TypeOfDeposition],
            "key_Process": df.iloc[Row_Number, Col_Process],
            "key_Recipe": df.iloc[Row_Number, Col_Recipe],
            "key_Thickness_Thickness": df.iloc[Row_Number, Col_Thickness_Thickness],
            "key_HEAD1_O2": df.iloc[Row_Number, Col_HEAD1_O2],
            "key_HEAD1_N2_O2": df.iloc[Row_Number, Col_HEAD1_N2_O2],
            "key_HEAD1_N2_SiH4": df.iloc[Row_Number, Col_HEAD1_N2_SiH4],
            "key_HEAD1_SiH4_4percent": df.iloc[Row_Number, Col_HEAD1_SiH4_4percent],
            "key_HEAD1_PH3_1percent": df.iloc[Row_Number, Col_HEAD1_PH3_1percent],
            "key_HEAD1_SiH4_100percent": df.iloc[Row_Number, Col_HEAD1_SiH4_100percent],
            "key_HEAD1_PH3_100percent": df.iloc[Row_Number, Col_HEAD1_PH3_100percent],
            "key_HEAD1_P_PSi": df.iloc[Row_Number, Col_HEAD1_P_PSi],
            "key_HEAD1_Ratio_O2Gas": df.iloc[Row_Number, Col_HEAD1_Ratio_O2_Gas],
            "key_HEAD2_O2": df.iloc[Row_Number, Col_HEAD2_O2],
            "key_HEAD2_N2_SiH4": df.iloc[Row_Number, Col_HEAD2_N2_SiH4],
            "key_HEAD2_N2_O2": df.iloc[Row_Number, Col_HEAD2_N2_O2],
            "key_HEAD2_SiH4_4percent": df.iloc[Row_Number, Col_HEAD2_SiH4_4percent],
            "key_HEAD2_PH3_1percent": df.iloc[Row_Number, Col_HEAD2_PH3_1percent],
            "key_HEAD2_SiH4_100percent": df.iloc[Row_Number, Col_HEAD2_SiH4_100percent],
            "key_HEAD2_Ratio_O2": df.iloc[Row_Number, Col_HEAD2_Ratio_O2],
            "key_Temp_Target": df.iloc[Row_Number, Col_Temp_Target],
            "key_Temp_H1": df.iloc[Row_Number, Col_Temp_H1],
            "key_Temp_H2": df.iloc[Row_Number, Col_Temp_H2],
            "key_Temp_H3": df.iloc[Row_Number, Col_Temp_H3],
            "key_Temp_H4": df.iloc[Row_Number, Col_Temp_H4],
            "key_Temp_H5": df.iloc[Row_Number, Col_Temp_H5],
            "key_Speed_Speed": df.iloc[Row_Number, Col_Speed_Speed],
            "key_Position_Start": df.iloc[Row_Number, Col_Position_Start],
            "key_Position_Stop": df.iloc[Row_Number, Col_Position_Stop],
            "key_Position_Delta": df.iloc[Row_Number, Col_Position_Delta]
        }

        # ----- 空欄はnanで取得されるため、空欄に置き換える -----
        for key in data_dict.keys():
            if str(data_dict[key]) == "nan":
                data_dict[key] = ""


        ########## 日付フォーマットの変換 ##########

        # ----- 日付を指定されたフォーマットに変換する -----
        Log.Log_Info(Log_File, 'Date Format Conversion')
        data_dict["key_Start_Date_Time"] = Convert_Date.Edit_Date(data_dict["key_Start_Date_Time"])

        # ----- 指定したフォーマットに変換出来たか確認 -----
        if len(data_dict["key_Start_Date_Time"]) != 19:
            Log.Log_Error(Log_File, "Date Error\n")
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
            Log.Log_Error(Log_File, "Data Error\n")
            Row_Number += 1
            continue


        ########### Serial_Number でループを回す ##########
        for col_serial_number in Col_Serial_Number:

            # ----- Serial_Number の取得 -----
            Serial_Number = str(df.iloc[Row_Number, col_serial_number]).replace("'", "")

            # ----- Primeに接続し、ロット番号に対応する品名を取り出す -----
            conn, cursor = SQL.connSQL()
            if conn is None:
                Log.Log_Error(Log_File, Serial_Number + ' : ' + 'Connection with Prime Failed')
                break
            Part_Number, Nine_Serial_Number = SQL.selectSQL(cursor, Serial_Number)
            SQL.disconnSQL(conn, cursor)

            # ----- 品名が None であれば処理を行わない ------
            if Part_Number is None:
                Log.Log_Error(Log_File, Serial_Number + ' : ' + "PartNumber Error\n")
                continue

            # ----- 品名が LDアレイ_ であれば処理を行わない -----
            if Part_Number == 'LDアレイ_':
                continue

            # ----- 辞書に追加 -----
            data_dict["key_Serial_Number"] = Serial_Number
            data_dict["key_Part_Number"] = Part_Number
            data_dict["key_LotNumber_9"] =  Nine_Serial_Number

            ########## XMLファイルの作成 ##########

            # ----- 保存するファイル名を定義 -----
            XML_File_Name = 'Site=' + Site + ',ProductFamily=' + ProductFamily + ',Operation=' + Operation + \
                            ',Partnumber=' + Part_Number + ',Serialnumber=' + Serial_Number + \
                            ',Testdate=' + data_dict["key_Start_Date_Time"] + '.xml'

            # ----- XMLファイルを作成 -----
            f = open(Output_filepath + XML_File_Name, 'w', encoding="utf-8")
                
            f.write('<?xml version="1.0" encoding="utf-8"?>' + '\n' +
                    '<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' + '\n' +
                    '       <Result startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Result="Passed">' + '\n' +
                    '               <Header SerialNumber=' + '"' + data_dict["key_Serial_Number"] + '"' + ' PartNumber=' + '"' + data_dict["key_Part_Number"] + '"' + ' Operation=' + '"' + Operation + '"' + ' TestStation=' + '"' + TestStation + '"' + ' Operator=' + '"' + data_dict["key_Operator"] + '"' + ' StartTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Site=' + '"' + Site + '"' + ' BatchNumber=' + '"' + data_dict["key_CVDNo"] + '"' + ' LotNumber=' + '"' +data_dict["key_Serial_Number"] + '"/>' + '\n' +
                    '               <HeaderMisc>' + '\n' +
                    '                   <Item Description=' + '"' + "Group" + '">' + str(data_dict["key_Process"]) + '</Item>' + '\n'
                    '                   <Item Description=' + '"' + "CVD-No" + '">' + str(data_dict["key_CVDNo"]) + '</Item>' + '\n'
                    '                   <Item Description=' + '"' + "TypeOfDeposition" + '">' + str(data_dict["key_TypeOfDeposition"]) + '</Item>' + '\n'
                    '                   <Item Description=' + '"' + "Recipe" + '">' + str(data_dict["key_Recipe"]) + '</Item>' + '\n'
                    '               </HeaderMisc>' + '\n' +
                    '\n'
                    '               <TestStep Name="Coordinate" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + X + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + Y + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="Thickness" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="Thickness" Units="nm" Value=' + '"' + str(data_dict['key_Thickness_Thickness']) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="HEAD1" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="O2" Units="sccm" Value=' + '"' + str(data_dict['key_HEAD1_O2']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="N2_O2" Units="slm" Value=' + '"' + str(data_dict['key_HEAD1_N2_O2']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="N2_SiH4" Units="slm" Value=' + '"' + str(data_dict['key_HEAD1_N2_SiH4']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="SiH4_4percent" Units="sccm" Value=' + '"' + str(data_dict['key_HEAD1_SiH4_4percent']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="PH3_1percent" Units="sccm" Value=' + '"' + str(data_dict['key_HEAD1_PH3_1percent']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="SiH4_100percent" Units="sccm" Value=' + '"' + str(data_dict['key_HEAD1_SiH4_100percent']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="PH3_100percent" Units="sccm" Value=' + '"' + str(data_dict['key_HEAD1_PH3_100percent']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="P_P+Si" Units="percent" Value=' + '"' + str(data_dict['key_HEAD1_P_PSi']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Ratio_O2/Gas" Units="percent" Value=' + '"' + str(data_dict['key_HEAD1_Ratio_O2Gas']) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="HEAD2" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="O2" Units="sccm" Value=' + '"' + str(data_dict['key_HEAD2_O2']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="N2_SiH4" Units="slm" Value=' + '"' + str(data_dict['key_HEAD2_N2_SiH4']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="N2_O2" Units="slm" Value=' + '"' + str(data_dict['key_HEAD2_N2_O2']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="SiH4_4percent" Units="sccm" Value=' + '"' + str(data_dict['key_HEAD2_SiH4_4percent']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="PH3_1percent" Units="sccm" Value=' + '"' + str(data_dict['key_HEAD2_PH3_1percent']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="SiH4_100percent" Units="sccm" Value=' + '"' + str(data_dict['key_HEAD2_SiH4_100percent']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Ratio_O2" Units="percent" Value=' + '"' + str(data_dict['key_HEAD2_Ratio_O2']) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="Temp" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="Target" Units="degree" Value=' + '"' + str(data_dict['key_Temp_Target']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="H1" Units="degree" Value=' + '"' + str(data_dict['key_Temp_H1']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="H2" Units="degree" Value=' + '"' + str(data_dict['key_Temp_H2']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="H3" Units="degree" Value=' + '"' + str(data_dict['key_Temp_H3']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="H4" Units="degree" Value=' + '"' + str(data_dict['key_Temp_H4']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="H5" Units="degree" Value=' + '"' + str(data_dict['key_Temp_H5']) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="Speed" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="Speed" Units="mm/min" Value=' + '"' + str(data_dict['key_Speed_Speed']) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="Position" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="Start" Units="mm" Value=' + '"' + str(data_dict['key_Position_Start']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Stop" Units="mm" Value=' + '"' + str(data_dict['key_Position_Stop']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Delta" Units="mm" Value=' + '"' + str(data_dict['key_Position_Delta']) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="SORTED_DATA" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value=' + '"' + str(data_dict["key_STARTTIME_SORTED"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value=' + '"' + str(data_dict["key_SORTNUMBER"]) + '"/>' + '\n' +
                    '                   <Data DataType="String" Name="LotNumber_5" Value=' + '"' + str(data_dict["key_Serial_Number"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                    '                   <Data DataType="String" Name="LotNumber_9" Value=' + '"' + str(data_dict["key_LotNumber_9"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '\n'
                    '               <TestEquipment>' + '\n' +
                    '                   <Item DeviceName="Nanospec" DeviceSerialNumber="' + '2' + '"/>' + '\n' +
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

            Log.Log_Info(Log_File, Serial_Number + ' : ' + "OK\n")


        ########## 行数更新 ##########
        Row_Number += 1


    ########## 次の開始行数の書き込み ##########

    Log.Log_Info(Log_File, 'Write the next starting line number')
    Row_Number_Func.next_start_row_number("Format2_StartRow.txt", Next_Start_Row)

    # ----- 最終行を書き込んだファイルをGドライブに転送 -----
    shutil.copy("Format2_StartRow.txt", 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/041_T-CVD/13_ProgramUsedFile/')


if __name__ == '__main__':

    Main()

# ----- ログ書込：Main処理の終了 -----
Log.Log_Info(Log_File, 'Program End')