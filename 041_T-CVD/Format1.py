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

Log_File = '../Log/' + Log_Folder_Name + '/041_T-CVD_Format1.log'
Log.Log_Info(Log_File, 'Program Start')


########## シート名の定義 ##########
Data_Sheet_Name = '着工記録'


########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
# Output_filepath = '../XML/041_T-CVD/Format1/'


########## 取得するデータの列番号を定義 ##########
Col_Start_Date_Time = 0
Col_Operator = 1
Col_HeaderMisc = 2
Col_Serial_Number = 3
Col_DepotTime_PSG = 4
Col_DepotTime_SiO2 = 5
Col_Temperature_Ts = 6
Col_Temperature_T = 7
Col_Temperature_T_Ts = 8
Col_Temperature_Plus14_Ts_Plus14 = 9
Col_Temperature_Plus14_T_Plus14 = 10
Col_Temperature_Plus14_T_Ts_Plus14 = 11
Col_ReplacementTime_ReplacementTime = 12
Col_GasFlowRate_O2 = 13
Col_GasFlowRate_N2 = 14
Col_GasFlowRate_PH3 = 15
Col_GasFlowRate_SiH4 = 16
Col_GasFlowRate_N2_carrier = 17
Col_Thickness_Thickness = 18


########## 取得した項目と型の対応表を定義 ##########
key_type = {
    "key_Start_Date_Time": str,
    "key_Operator": str,
    "key_HeaderMisc" : str,
    "key_Thickness_Thickness" : float,
    "key_DepotTime_PSG" : float,
    "key_DepotTime_SiO2" : float,
    "key_Temperature_Ts" : float,
    "key_Temperature_T" : float,
    "key_Temperature_T_Ts" : float,
    "key_Temperature_Plus14_Ts_Plus14" : float,
    "key_Temperature_Plus14_T_Plus14" : float,
    "key_Temperature_Plus14_T_Ts_Plus14" : float,
    "key_ReplacementTime_ReplacementTime" : float,
    "key_GasFlowRate_O2" : float,
    "key_GasFlowRate_N2" : float,
    "key_GasFlowRate_PH3" : float,
    "key_GasFlowRate_SiH4" : float,
    "key_GasFlowRate_N2_carrier" : float,
    'key_STARTTIME_SORTED' : float,
    'key_SORTNUMBER' : float
}


########## MARK工程とMESA工程のGroup定義 ##########
MARK = set(("ﾏｰｸ保護ﾎﾄ前", "ﾏｰｸﾎﾄ前", "マーク保護ホト前", "マーク保護ホト前CVD", "保護ホト前", "ｱﾗｲﾒﾝﾄﾏｰｸﾎﾄ前"))
MESA = set(("MESA", "MESA前", "ﾒｻCVD追加", "ﾒｻEBﾎﾄ前", "ﾒｻEBﾎﾄ前 (再）", "ﾒｻEBﾎﾄ前(再)", "ﾒｻEBﾎﾄ前(再生）", "ﾒｻEBﾎﾄ前3", "ﾒｻﾎﾄ前"))


def Main():


    ########## Excelファイルをローカルにコピー ##########

    # ----- 正規表現で取出し、直近で変更があったファイルを取得する -----
    Log.Log_Info(Log_File, 'Excel File Copy')

    FilePath = 'Z:/CVD/T-CVD/'
    #FilePath = 'C:/Users/hsi67063/Downloads/Python test/'
    FileName = '*T-CVD着工記録_相模原移転後*.xls*'
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
    Start_Number = Row_Number_Func.start_row_number("Format1_StartRow.txt") - 500

    # ----- ExcelデータをDataFrameとして取得 -----
    Log.Log_Info(Log_File, 'Read Excel')
    df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="A:AJ", skiprows=Start_Number)

    # ----- 必要な列のみ取り出す -----
    df = df.iloc[:, [0,1,4,5,19,20,21,22,23,24,25,26,28,29,30,31,32,33,35]]

    # ----- 列番号の振り直し -----
    Log.Log_Info(Log_File, 'Setting Columns Number')
    df.columns = range(df.shape[1])

    # ----- 末尾から欠損のデータを落としていく -----
    Getting_Row = len(df) - 1
    while Getting_Row >= 0 and (str(df.iloc[Getting_Row, Col_Start_Date_Time]) == "NaT" or str(df.iloc[Getting_Row, Col_Start_Date_Time]) == "nan"):
        Getting_Row -= 1

    df = df[:Getting_Row + 1]

    # ----- 次の開始行数をメモ -----
    Next_Start_Row = Start_Number + df.shape[0] + 1
    print(Start_Number,df.shape[0],Next_Start_Row)


    # ----- 日付欄に文字列が入っていたらNoneに置き換える -----
    for i in range(df.shape[0]):
        if type(df.iloc[i, 0]) is not pd.Timestamp:
            df.iloc[i, 0] = np.nan

    # ----- 今日から1か月前のデータまでを取得する -----
    df[0] = pd.to_datetime(df[0])
    one_month_ago = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) - relativedelta(months=10)
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
        if str(df.iloc[Row_Number, Col_Start_Date_Time])=="nan" or str(df.iloc[Row_Number, Col_Operator])=="nan" or str(df.iloc[Row_Number, Col_Serial_Number])=="nan" or str(df.iloc[Row_Number, Col_HeaderMisc])=="nan":
            Log.Log_Error(Log_File, "Blank Error\n")
            Row_Number += 1
            continue
        print(str(df.iloc[Row_Number, Col_Start_Date_Time]))

        ########## 現在処理を行っている行のデータの取得 ##########

        # ----- 取得したデータを格納するデータ構造(辞書)を作成 -----
        Log.Log_Info(Log_File, 'Data Acquisition')
        data_dict = dict()

        # ----- ロット番号を取得 -----
        Serial_Number_List = list(str(df.iloc[Row_Number, Col_Serial_Number]).split('/'))
        
        # ----- 対応したOperationを探すキーの取得 -----
        Process = str(df.iloc[Row_Number, Col_HeaderMisc])


        ########## OperationとTestStationの定義 ##########

        if "BJ1" in Process or "ＢＪ１" in Process:
            Operation, TestStation = "BJ1_T-CVD-Thickness", "BJ1"            
        elif "BJ2" in Process:
            Operation, TestStation = "BJ2_T-CVD-Thickness", "BJ2"            
        elif "回折格子" in Process:
            Operation, TestStation = "GRATING_T-CVD-Thickness", "GRATING"            
        elif Process in MARK:
            Operation, TestStation = "MARK-EML_T-CVD-Thickness", "MARK-EML"            
        elif Process in MESA:
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
            "key_Operator": df.iloc[Row_Number, Col_Operator],
            "key_HeaderMisc": df.iloc[Row_Number, Col_HeaderMisc],
            "key_Thickness_Thickness": df.iloc[Row_Number, Col_Thickness_Thickness],
            "key_DepotTime_PSG": df.iloc[Row_Number, Col_DepotTime_PSG],
            "key_DepotTime_SiO2": df.iloc[Row_Number, Col_DepotTime_SiO2],
            "key_Temperature_Ts": df.iloc[Row_Number, Col_Temperature_Ts],
            "key_Temperature_T": df.iloc[Row_Number, Col_Temperature_T],
            "key_Temperature_T_Ts": df.iloc[Row_Number, Col_Temperature_T_Ts],
            "key_Temperature_Plus14_Ts_Plus14": df.iloc[Row_Number, Col_Temperature_Plus14_Ts_Plus14],
            "key_Temperature_Plus14_T_Plus14": df.iloc[Row_Number, Col_Temperature_Plus14_T_Plus14],
            "key_Temperature_Plus14_T_Ts_Plus14": df.iloc[Row_Number, Col_Temperature_Plus14_T_Ts_Plus14],
            "key_ReplacementTime_ReplacementTime": df.iloc[Row_Number, Col_ReplacementTime_ReplacementTime],
            "key_GasFlowRate_O2": df.iloc[Row_Number, Col_GasFlowRate_O2],
            "key_GasFlowRate_N2": df.iloc[Row_Number, Col_GasFlowRate_N2],
            "key_GasFlowRate_PH3": df.iloc[Row_Number, Col_GasFlowRate_PH3],
            "key_GasFlowRate_SiH4": df.iloc[Row_Number, Col_GasFlowRate_SiH4],
            "key_GasFlowRate_N2_carrier": df.iloc[Row_Number, Col_GasFlowRate_N2_carrier]
        }


        ########## 日付フォーマットの変換 ##########

        # ----- 日付を指定されたフォーマットに変換する -----
        print(data_dict["key_Start_Date_Time"])
        if "2025/2/13;E" == data_dict["key_Start_Date_Time"]:
            Row_Number+=1
            continue
        Log.Log_Info(Log_File, 'Date Format Conversion')
        data_dict["key_Start_Date_Time"] = Convert_Date.Edit_Date(data_dict["key_Start_Date_Time"])

        # ----- 指定したフォーマットに変換出来たか確認 -----
        if len(data_dict["key_Start_Date_Time"]) != 19:
            Log.Log_Error(Log_File, "Date Error\n")
            Row_Number += 1
            continue


        ########## データ置換 ##########

        # ----- "key_Temperature_Plus14_Ts_Plus14"と"key_Temperature_Plus14_T_Plus14"が空欄であれば、""に置き換える -----
        if isnan(data_dict["key_Temperature_Plus14_Ts_Plus14"]): data_dict["key_Temperature_Plus14_Ts_Plus14"]=""
        if isnan(data_dict["key_Temperature_Plus14_T_Plus14"]): data_dict["key_Temperature_Plus14_T_Plus14"]=""

        # ----- 両方とも空欄であれば、T-Tsも空欄とする -----
        if data_dict["key_Temperature_Plus14_Ts_Plus14"] == data_dict["key_Temperature_Plus14_T_Plus14"] == "":
            data_dict["key_Temperature_Plus14_T_Ts_Plus14"] = ""


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


        ########## 全ロット番号の処理 ##########

        for Serial_Number in Serial_Number_List:

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

            data_dict["key_LotNumber_9"] = Nine_Serial_Number
                

            ########## XMLファイルの作成 ##########

            # ----- 保存するファイル名を定義 -----
            XML_File_Name = 'Site=' + Site + ',ProductFamily=' + ProductFamily + ',Operation=' + Operation + \
                            ',Partnumber=' + Part_Number + ',Serialnumber=' + Serial_Number + \
                            ',Testdate=' + data_dict["key_Start_Date_Time"] + '.xml'

            # ----- XMLファイルの作成 -----
            Log.Log_Info(Log_File, 'Excel File To XML File Conversion')            
            f = open(Output_filepath + XML_File_Name, 'w', encoding="utf-8")

            f.write('<?xml version="1.0" encoding="utf-8"?>' + '\n' +
                    '<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' + '\n' +
                    '       <Result startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Result="Passed">' + '\n' +
                    '               <Header SerialNumber=' + '"' + Serial_Number + '"' + ' PartNumber=' + '"' + Part_Number + '"' + ' Operation=' + '"' + Operation + '"' + ' TestStation=' + '"' + TestStation + '"' + ' Operator=' + '"' + data_dict["key_Operator"] + '"' + ' StartTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Site=' + '"' + Site + '"' + ' LotNumber=' + '"' + Serial_Number + '"/>' + '\n' +
                    '               <HeaderMisc>' + '\n' +
                    '                   <Item Description=' + '"' + "Group" + '">' + data_dict["key_HeaderMisc"] + '</Item>' + '\n'
                    '               </HeaderMisc>' + '\n' +
                    '\n'
                    '               <TestStep Name="Coordinate" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + X + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + Y + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="Thickness" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="Thickness" Units="nm" Value=' + '"' + str(data_dict['key_Thickness_Thickness']) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="DepotTime" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="PSG" Units="sec" Value=' + '"' + str(data_dict['key_DepotTime_PSG']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="SiO2" Units="sec" Value=' + '"' + str(data_dict['key_DepotTime_SiO2']) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="Temperature" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="Ts" Units="degree" Value=' + '"' + str(data_dict['key_Temperature_Ts']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="T" Units="degree" Value=' + '"' + str(data_dict['key_Temperature_T']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="T-Ts" Units="degree" Value=' + '"' + str(data_dict['key_Temperature_T_Ts']) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="Temperature_Plus14" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="Ts_Plus14" Units="degree" Value=' + '"' + str(data_dict['key_Temperature_Plus14_Ts_Plus14']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="T_Plus14" Units="degree" Value=' + '"' + str(data_dict['key_Temperature_Plus14_T_Plus14']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="T-Ts_Plus14" Units="degree" Value=' + '"' + str(data_dict['key_Temperature_Plus14_T_Ts_Plus14']) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="ReplacementTime" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="ReplacementTime" Units="sec" Value=' + '"' + str(data_dict['key_ReplacementTime_ReplacementTime']) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="GasFlowRate" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="O2" Units="L/min" Value=' + '"' + str(data_dict['key_GasFlowRate_O2']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="N2" Units="L/min" Value=' + '"' + str(data_dict['key_GasFlowRate_N2']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="PH3" Units="L/min" Value=' + '"' + str(data_dict['key_GasFlowRate_PH3']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="SiH4" Units="L/min" Value=' + '"' + str(data_dict['key_GasFlowRate_SiH4']) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="N2-carrier" Units="L/min" Value=' + '"' + str(data_dict['key_GasFlowRate_N2_carrier']) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '               <TestStep Name="SORTED_DATA" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value=' + '"' + str(data_dict["key_STARTTIME_SORTED"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value=' + '"' + str(data_dict["key_SORTNUMBER"]) + '"/>' + '\n' +
                    '                   <Data DataType="String" Name="LotNumber_5" Value=' + '"' + Serial_Number + '"' + ' CompOperation="LOG"/>' + '\n' +
                    '                   <Data DataType="String" Name="LotNumber_9" Value=' + '"' + str(data_dict["key_LotNumber_9"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '\n'
                    '               <TestEquipment>' + '\n' +
                    '                   <Item DeviceName="Nanospec" DeviceSerialNumber="' + '1' + '"/>' + '\n' +
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
    Row_Number_Func.next_start_row_number("Format1_StartRow.txt", Next_Start_Row)

    # ----- 最終行を書き込んだファイルをGドライブに転送 -----
    #shutil.copy("Format1_StartRow.txt", 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/041_T-CVD/13_ProgramUsedFile/')


if __name__ == '__main__':

    Main()

# ----- ログ書込：Main処理の終了 -----
Log.Log_Info(Log_File, 'Program End')