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


########## 全体パラメータの定義 ##########
Site = '350'
ProductFamily = 'SAG FAB'
Operation = 'MESA_CVD_Width'
TestStation = 'MESA'


########## Logの設定 ##########

# ----- ログファイルの作成 -----
Log_Folder_Name = str(date.today())
if not os.path.exists("../Log/" + Log_Folder_Name):
    os.makedirs("../Log/" + Log_Folder_Name)

Log_File = '../Log/' + Log_Folder_Name + '/002_MESA.log'
Log.Log_Info(Log_File, 'Program Start')


########## シート名の定義 ##########
Data_Sheet_Name = 'データ'
XY_Sheet_Name = 'ウェハ座標'
EA_DML_Sheet_Name = '品名一覧'


########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
# Output_filepath = '../XML/002_MESA/'


########## 取得するデータのdfの列番号を定義 ##########
Col_Start_Date_Time = 0
Col_Operator = 1
Col_Sem_Number = 2
Col_Serial_Number = 4
#add ICP tool_id 2024-12-04 Pigo
Col_ICP_Number = 5
#From Here    shift 1 column to right     2024-11-29 masaaki yoshida
Col_Width1 = 12 #before=11 after=12    the same applies below
Col_Width2 = 13
Col_Width3 = 14
Col_Width4 = 15
Col_Width5 = 16
Col_Result = 19
#To Here

########## 取得した項目と型の対応表を定義 ##########
key_type = {
    "key_Start_Date_Time": str,
    "key_Part_Number": str,
    "key_Sem_Number": str,
    "key_Serial_Number": str,
    "key_ICP_Number": str,
    "key_Operator": str,
    "key_Width1": float,
    "key_Width2": float,
    "key_Width3": float,
    "key_Width4": float,
    "key_Width5": float,
    "key_Result": str,
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
    "key_SORTNUMBER": float,
    "key_LotNumber_9" : str
}


def Main():


    ########## Excelファイルをローカルにコピー ##########

    # ----- 正規表現で取出し、直近で変更があったファイルを取得 -----
    Log.Log_Info(Log_File, 'Excel File Copy')

    FilePath = 'Z:/■QCデータシート/ドライエッチ/'
    FileName = '*ドライ処理後SiO2寸法*.xlsx'
    Excel_File_List = []
    for file in glob.glob(FilePath + FileName):
        if '$' not in file:
            dt = strftime("%Y-%m-%d %H:%M:%S", localtime(os.path.getmtime(file)))
            Excel_File_List.append([file, dt])

    # ----- dt(更新日時)の降順で並び替える -----
    Excel_File_List = sorted(Excel_File_List, key=lambda x: x[1], reverse=True)
    Excel_File = shutil.copy(Excel_File_List[0][0], '../DataFile/002_MESA/')


    ########## DaraFrameの作成 ##########

    # ----- 取得開始行の取り出し -----
    Log.Log_Info(Log_File, 'Get The Starting Row Count')
    Start_Number = Row_Number_Func.start_row_number("MESA_StartRow.txt") - 500

    # ----- ExcelデータをDataFrameとして取得 -----
    Log.Log_Info(Log_File, 'Read Excel')
    df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="B:U", skiprows=Start_Number)#B:T -> B:U     2024-11-29 masaaki yoshida
    df_EA_DML = pd.read_excel(Excel_File, header=None, sheet_name=EA_DML_Sheet_Name, usecols="H:P", skiprows=5)
    df_xy = pd.read_excel(Excel_File, header=None, sheet_name=XY_Sheet_Name, usecols="A:K")

    # ----- 列番号の振り直し -----
    Log.Log_Info(Log_File, 'Setting Columns Number')
    df.columns = range(df.shape[1])
    df_xy.columns = range(df_xy.shape[1])
    df_EA_DML.columns = range(df_EA_DML.shape[1])

    # ----- 末尾から欠損のデータを落としていく -----
    Getting_Row = len(df)-1
    while Getting_Row >= 0 and df.isnull().any(axis=1)[Getting_Row]:
        Getting_Row-=1

    df = df[:Getting_Row+1]
    # ----- 次の開始行数をメモ -----
    Next_Start_Row = Start_Number + df.shape[0] + 1

    # ----- 日付欄に文字列が入っていたらNoneに置き換える -----
    for i in range(df.shape[0]):
        if type(df.iloc[i, 0]) is not pd.Timestamp:
            df.iloc[i, 0] = np.nan

    # ----- 今日から1か月前のデータまでを取得する -----
    df[0] = pd.to_datetime(df[0])
    one_month_ago = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) - relativedelta(months=6)
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
        PartNumber, Nine_Serial_Number = SQL.selectSQL(cursor, Serial_Number)
        SQL.disconnSQL(conn, cursor)

        # ----- 品名が None であれば処理を行わない -----
        if PartNumber is None:
            Log.Log_Error(Log_File, Serial_Number + ' : ' + "PartNumber Error\n")
            Row_Number += 1
            continue

        # ----- 品名が LDアレイ_ であれば処理を行わない -----
        if PartNumber == 'LDアレイ_':
            Row_Number += 1
            continue

        # ----- ロット番号の左2文字を抜き出す -----
        Serial_Number_Initial = Serial_Number[:2]
        Start_Row, Initial_Col = 0, 5
        Serial_Number_Class = ""
        Search_End_Row = df_EA_DML.shape[0]

        # ----- ロット番号に対応する大分類(EML/DML)をサーチ -----
        while Start_Row < Search_End_Row:
            if df_EA_DML.iloc[Start_Row, Initial_Col] == Serial_Number_Initial:
                Serial_Number_Class = df_EA_DML.iloc[Start_Row, Initial_Col+3]
                break
            Start_Row+=1
        
        if Serial_Number_Class == "":
            Log.Log_Error(Log_File, Serial_Number + ' : ' + 'Classification Error')
            Row_Number+=1
            continue

        # ----- 大分類に対応する座標を抜き出す -----
        XY_Row, Initial_Col = 1, 0
        Search_End_Row = df_xy.shape[0]
        Flag = False
        
        # ----- 大分類に対応するXY座標をサーチ -----
        while XY_Row < Search_End_Row:
            print(df_xy.iloc[XY_Row, Initial_Col])
            if df_xy.iloc[XY_Row, Initial_Col] == Serial_Number_Class:
                Flag = True
                break
            XY_Row+=1
        
        if Flag == False:
            Log.Log_Error(Log_File, Serial_Number + ' : ' + 'Coordinate Error')
            Row_Number+=1
            continue

        # ----- データの取得 -----
        data_dict = {
            "key_Start_Date_Time": df.iloc[Row_Number, Col_Start_Date_Time],
            "key_Part_Number": PartNumber,
            "key_Sem_Number": df.iloc[Row_Number, Col_Sem_Number][-1],
            "key_Serial_Number": Serial_Number,
            "key_ICP_Number": df.iloc[Row_Number, Col_ICP_Number][-1],
            "key_LotNumber_9": Nine_Serial_Number,
            "key_Operator": df.iloc[Row_Number, Col_Operator],
            "key_Width1": df.iloc[Row_Number, Col_Width1],
            "key_Width2": df.iloc[Row_Number, Col_Width2],
            "key_Width3": df.iloc[Row_Number, Col_Width3],
            "key_Width4": df.iloc[Row_Number, Col_Width4],
            "key_Width5": df.iloc[Row_Number, Col_Width5],
            "key_Result": df.iloc[Row_Number, Col_Result],
            "key_X1": df_xy.iloc[XY_Row, 1],
            "key_X2": df_xy.iloc[XY_Row, 2],
            "key_X3": df_xy.iloc[XY_Row, 3],
            "key_X4": df_xy.iloc[XY_Row, 4],
            "key_X5": df_xy.iloc[XY_Row, 5],
            "key_Y1": df_xy.iloc[XY_Row, 6],
            "key_Y2": df_xy.iloc[XY_Row, 7],
            "key_Y3": df_xy.iloc[XY_Row, 8],
            "key_Y4": df_xy.iloc[XY_Row, 9],
            "key_Y5": df_xy.iloc[XY_Row, 10]
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

        # ----- Resultを定義 -----
        if data_dict["key_Result"] == "OK":
            tmp_Result = "Passed"
        elif data_dict["key_Result"] == "NG":
            tmp_Result = "Failed"
        else:
            tmp_Result = "Done"

        # ----- XMLファイルの作成 -----
        Log.Log_Info(Log_File, 'Excel File To XML File Conversion')
        
        f = open(Output_filepath + XML_File_Name, 'w', encoding="utf-8")

        f.write('<?xml version="1.0" encoding="utf-8"?>' + '\n' +
                '<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' + '\n' +
                '       <Result startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Result=' + '"' + tmp_Result + '">' + '\n' +
                '               <Header SerialNumber=' + '"' + data_dict["key_Serial_Number"] + '"' + ' PartNumber=' + '"' + data_dict["key_Part_Number"] + '"' + ' Operation=' + '"' + Operation + '"' + ' TestStation=' + '"' + TestStation + '"' + ' Operator=' + '"' + data_dict["key_Operator"] + '"' + ' StartTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Site=' + '"' + Site + '"' + ' LotNumber=' + '"' + data_dict["key_Serial_Number"] +  '"' +' MISCINFO=' + '"ICP' + data_dict["key_ICP_Number"]+ '"/>'+ '"/>' + '\n' +
                '\n'
                '               <TestStep Name="Width1" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status=' + '"' + tmp_Result + '">' + '\n' +
                '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y1"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Width" Units="nm" Value=' + '"' + str(data_dict["key_Width1"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Width2" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status=' + '"' + tmp_Result + '">' + '\n' +
                '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y2"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Width" Units="nm" Value=' + '"' + str(data_dict["key_Width2"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Width3" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status=' + '"' + tmp_Result + '">' + '\n' +
                '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y3"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Width" Units="nm" Value=' + '"' + str(data_dict["key_Width3"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Width4" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status=' + '"' + tmp_Result + '">' + '\n' +
                '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y4"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Width" Units="nm" Value=' + '"' + str(data_dict["key_Width4"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="Width5" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status=' + '"' + tmp_Result + '">' + '\n' +
                '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y5"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="Width" Units="nm" Value=' + '"' + str(data_dict["key_Width5"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '               <TestStep Name="SORTED_DATA" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status=' + '"' + tmp_Result + '">' + '\n' +
                '                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value=' + '"' + str(data_dict["key_STARTTIME_SORTED"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value=' + '"' + str(data_dict["key_SORTNUMBER"]) + '"/>' + '\n' +
                '                   <Data DataType="String" Name="LotNumber_5" Value=' + '"' + str(data_dict["key_Serial_Number"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                '                   <Data DataType="String" Name="LotNumber_9" Value=' + '"' + str(data_dict["key_LotNumber_9"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '\n'
                '               <TestEquipment>' + '\n' +
                '                   <Item DeviceName="SEM" DeviceSerialNumber=' + '"' + str(data_dict["key_Sem_Number"]) + '"/>' + '\n' +
                '                   <Item DeviceName="ICP" DeviceSerialNumber="' + str(data_dict['key_ICP_Number']) + '"/>' + '\n' +  # add ICP tool_id 2024-12-04 Pigo                
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
    Row_Number_Func.next_start_row_number("MESA_StartRow.txt", Next_Start_Row)

    # ----- 最終行を書き込んだファイルをGドライブに転送 ------
    #shutil.copy("MESA_StartRow.txt", 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/002_MESA/13_ProgramUsedFile/')


if __name__ == '__main__':
    Main()

# ----- ログ書込：Main処理の終了 -----
Log.Log_Info(Log_File, 'Program End')