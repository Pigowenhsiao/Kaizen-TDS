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
ProductFamily = 'SAG FAB'
Operation = 'GRATING_CVD_EtchingRate'
TestStation = 'GRATING'


########## Logの設定 ##########

# ----- ログファイルの作成 -----
Log_Folder_Name = str(date.today())
if not os.path.exists("../Log/" + Log_Folder_Name):
    os.makedirs("../Log/" + Log_Folder_Name)

Log_File = '../Log/' + Log_Folder_Name + '/010_GRATING.log'
Log.Log_Info(Log_File, 'Program Start')


########## シート名の定義 ##########
Data_Sheet_Name = 'データ'


########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
# Output_filepath = '../XML/010_GRATING/'


########## 取得するデータの列番号を定義 ##########
Col_Start_Date_Time = 0
Col_operator = 1
Col_NanoSpec = 3
Col_Serial_Number = 4
Col_initial1 = 5
Col_initial2 = 6
Col_initial3 = 7
Col_initial4 = 8
Col_initial5 = 9
Col_initial_Ave = 10
Col_final1 = 11
Col_final2 = 12
Col_final3 = 13
Col_final4 = 14
Col_final5 = 15
Col_final_Ave = 16
Col_rate1 = 17
Col_rate2 = 18
Col_rate3 = 19
Col_rate4 = 20
Col_rate5 = 21
Col_rate_Ave = 22
Col_rate_3sigma = 23
Col_time = 24
Col_x1 = 25
Col_x2 = 26
Col_x3 = 27
Col_x4 = 28
Col_x5 = 29
Col_y1 = 30
Col_y2 = 31
Col_y3 = 32
Col_y4 = 33
Col_y5 = 34


########## 取得した項目と型の対応表を定義 ##########
key_type = {
    'key_Start_Date_Time' : str,
    'key_Part_Number' : str,
    'key_Serial_Number' : str,
    'key_Operator' : str,
    "key_Initial1": float,
    "key_Initial2": float,
    "key_Initial3": float,
    "key_Initial4": float,
    "key_Initial5": float,
    "key_Initial_ave": float,
    "key_Final1": float,
    "key_Final2": float,
    "key_Final3": float,
    "key_Final4": float,
    "key_Final5": float,
    "key_Final_ave": float,
    "key_Rate1": float,
    "key_Rate2": float,
    "key_Rate3": float,
    "key_Rate4": float,
    "key_Rate5": float,
    "key_Rate_ave": float,
    'key_Rate_3sigma' : float,
    "key_Time": float,
    "key_X1": float,
    "key_X2": float,
    "key_X3": float,
    "key_X4": float,
    "key_X5": float,
    "key_Y1": float,
    "key_Y2": float,
    "key_Y3": float,
    "key_Y4": float,
    "key_Y5": float,
    "key_STARTTIME_SORTED": float,
    "key_SORTNUMBER" : float,
    "key_LotNumber_9": str
}


def Main(FilePath, FileName, TextFile, Equipment, HeaderMisc):


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
    Excel_File = shutil.copy(Excel_file_list[0][0], '../DataFile/010_GRATING/')


    ########## DaraFrameの作成 ##########

    # ----- 取得開始行の取り出し -----
    Log.Log_Info(Log_File, 'Get The Starting Row Count')
    Start_Number = max(Row_Number_Func.start_row_number(TextFile)-500, 5)

    # ----- ExcelデータをDataFrameとして取得 -----
    Log.Log_Info(Log_File, 'Read Excel')
    df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="B:AJ", skiprows=Start_Number)

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
            'key_Part_Number': Part_Number,
            'key_Serial_Number': Serial_Number,
            "key_LotNumber_9": Nine_Serial_Number,
            'key_Operator': df.iloc[Row_Number, Col_operator],
            "key_Initial1": df.iloc[Row_Number, Col_initial1],
            "key_Initial2": df.iloc[Row_Number, Col_initial2],
            "key_Initial3": df.iloc[Row_Number, Col_initial3],
            "key_Initial4": df.iloc[Row_Number, Col_initial4],
            "key_Initial5": df.iloc[Row_Number, Col_initial5],
            "key_Initial_ave": df.iloc[Row_Number, Col_initial_Ave],
            "key_Final1": df.iloc[Row_Number, Col_final1],
            "key_Final2": df.iloc[Row_Number, Col_final2],
            "key_Final3": df.iloc[Row_Number, Col_final3],
            "key_Final4": df.iloc[Row_Number, Col_final4],
            "key_Final5": df.iloc[Row_Number, Col_final5],
            "key_Final_ave": df.iloc[Row_Number, Col_final_Ave],
            "key_Rate1": df.iloc[Row_Number, Col_rate1],
            "key_Rate2": df.iloc[Row_Number, Col_rate2],
            "key_Rate3": df.iloc[Row_Number, Col_rate3],
            "key_Rate4": df.iloc[Row_Number, Col_rate4],
            "key_Rate5": df.iloc[Row_Number, Col_rate5],
            "key_Rate_ave": df.iloc[Row_Number, Col_rate_Ave],
            'key_Rate_3sigma': df.iloc[Row_Number, Col_rate_3sigma],
            "key_Time": df.iloc[Row_Number, Col_time],
            "key_X1": df.iloc[Row_Number, Col_x1],
            "key_X2": df.iloc[Row_Number, Col_x2],
            "key_X3": df.iloc[Row_Number, Col_x3],
            "key_X4": df.iloc[Row_Number, Col_x4],
            "key_X5": df.iloc[Row_Number, Col_x5],
            "key_Y1": df.iloc[Row_Number, Col_y1],
            "key_Y2": df.iloc[Row_Number, Col_y2],
            "key_Y3": df.iloc[Row_Number, Col_y3],
            "key_Y4": df.iloc[Row_Number, Col_y4],
            "key_Y5": df.iloc[Row_Number, Col_y5],
            'key_TestEquipment_Nano': df.iloc[Row_Number, Col_NanoSpec],
            'key_TestEquipment_Dry': Equipment
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
                '               <HeaderMisc>' + '\n' +
                '                   <Item Description="CVD_Thickness">' + HeaderMisc + '</Item>' + '\n'
                '               </HeaderMisc>' + '\n' +
                '\n'
                '               <TestStep Name="Thickness1" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Initial1" Units="nm" Value=' + '"' + str(data_dict["key_Initial1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Initial2" Units="nm" Value=' + '"' + str(data_dict["key_Initial2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Initial3" Units="nm" Value=' + '"' + str(data_dict["key_Initial3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Initial4" Units="nm" Value=' + '"' + str(data_dict["key_Initial4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Initial5" Units="nm" Value=' + '"' + str(data_dict["key_Initial5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Initial_Ave" Units="nm" Value=' + '"' + str(data_dict["key_Initial_ave"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Thickness2" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Final1" Units="nm" Value=' + '"' + str(data_dict["key_Final1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Final2" Units="nm" Value=' + '"' + str(data_dict["key_Final2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Final3" Units="nm" Value=' + '"' + str(data_dict["key_Final3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Final4" Units="nm" Value=' + '"' + str(data_dict["key_Final4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Final5" Units="nm" Value=' + '"' + str(data_dict["key_Final5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Final_Ave" Units="nm" Value=' + '"' + str(data_dict["key_Final_ave"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Rate" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Rate1" Units="nm/min" Value=' + '"' + str(data_dict["key_Rate1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Rate2" Units="nm/min" Value=' + '"' + str(data_dict["key_Rate2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Rate3" Units="nm/min" Value=' + '"' + str(data_dict["key_Rate3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Rate4" Units="nm/min" Value=' + '"' + str(data_dict["key_Rate4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Rate5" Units="nm/min" Value=' + '"' + str(data_dict["key_Rate5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Rate_Ave" Units="nm/min" Value=' + '"' + str(data_dict["key_Rate_ave"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Rate_3sigma" Units="nm/min" Value=' + '"' + str(data_dict["key_Rate_3sigma"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Time" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="Time" Units="min" Value=' + '"' + str(data_dict["key_Time"]) + '"/>' + '\n' +
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
                '\n'
                '               <TestEquipment>' + '\n' +
                '                   <Item DeviceName="Nanospec" DeviceSerialNumber=' + '"' + str(data_dict["key_TestEquipment_Nano"]) + '"/>' + '\n' +
                '                   <Item DeviceName="DryEtch" DeviceSerialNumber=' + '"' + str(data_dict["key_TestEquipment_Dry"]) + '"/>' + '\n' +
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
    shutil.copy(TextFile, 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/010_GRATING/13_ProgramUsedFile/')


if __name__ == '__main__':

    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ3号機/3ｲﾝﾁ/', '*3インチ_回折格子SiO2膜ドライエッチプログラムシート*.xlsm', 'GRATING_100nm_Dry3.txt', '#3', '100nm')
    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ3号機/3ｲﾝﾁ/', '*3インチ_*回折格子SiO2*45nm薄膜品ドライエッチプログラムシート*.xlsm', 'GRATING_45nm_Dry3.txt', '#3', '45nm')
    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ3号機/3ｲﾝﾁ/', '*3インチ_*回折格子SiO2*55nm薄膜品ドライエッチプログラムシート*.xlsm', 'GRATING_55nm_Dry3.txt', '#3', '55nm')
    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ6号機/3ｲﾝﾁ/', '*3インチ_回折格子SiO2膜ドライエッチ(ドライ6号機)プログラムシート*.xlsm', 'GRATING_100nm_Dry6.txt', '#6', '100nm')
    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ6号機/3ｲﾝﾁ/', '*3インチ_回折格子SiO2*45nm薄膜品ドライエッチ(ドライ6号機)プログラムシート*.xlsm', 'GRATING_45nm_Dry6.txt', '#6', '45nm')
    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ8号機/3ｲﾝﾁ/', '*3インチ_回折格子SiO2 45nm薄膜品ドライエッチ(ドライ8号機)プログラムシート*.xlsm', 'GRATING_45nm_Dry8.txt', '#8', '45nm')
    Main('Z:/ホト・エッチング/製品/ﾄﾞﾗｲｴｯﾁ8号機/3ｲﾝﾁ/', '*3インチ_回折格子SiO2 55nm薄膜品ドライエッチ(ドライ8号機)プログラムシート*.xlsm', 'GRATING_55nm_Dry8.txt', '#8', '45nm')
# ----- ログ書込：Main処理の終了 -----
Log.Log_Info(Log_File, 'Program End')