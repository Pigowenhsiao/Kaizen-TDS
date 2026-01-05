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
Operation = 'N-electrode_DML_Deposition_Thickness'
TestStation = 'N-electrode'


########## Logの設定 ##########

# ----- ログファイルの作成 -----
Log_Folder_Name = str(date.today())
if not os.path.exists("../Log/" + Log_Folder_Name):
    os.makedirs("../Log/" + Log_Folder_Name)

Log_File = '../Log/' + Log_Folder_Name + '/034_N-electrode.log'
Log.Log_Info(Log_File, 'Program Start')


########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
# Output_filepath = '../XML/034_N-electrode/'


########## 取得するデータの列番号を定義 ##########
Col_Serial_Number = 0
Col_Operator = 1
Col_Start_Date_Time = 2
Col_Program_No = 3
Col_Thickness = 4


########## 取得した項目と型の対応表を定義 ##########
key_type = {
    "key_Start_Date_Time": str,
    "key_Part_Number": str,
    "key_Serial_Number": str,
    "key_Operator": str,
    "key_Program_No": float,
    "key_Thickness": float,
    "key_STARTTIME_SORTED" : float,
    'key_SORTNUMBER' : float,
    "key_LotNumber_9": str
}


########## 品種からEML,10G,25Gかどうかを返す ##########
def Part_Class_Check(Part_Number):

    # ----- Prime接続 -----
    conn, cursor = SQL.connSQL()
    if conn is None:
        Log.Log_Error(Log_File, Part_Number + ' : ' + 'Connection with Prime Failed')
        sys.exit()

    # ----- EML/10G/25Gの判定を行う -----
    Part_Class = None
    cursor.execute("select distinct ProductFamilyName from prime.v_TransactionData where ProductName like '" + Part_Number + "';")
    row = cursor.fetchone()
    while row:
        Part_Class = row[0]
        row = cursor.fetchone()

    # ----- Prime切断 -----
    SQL.disconnSQL(conn, cursor)

    return Part_Class


def Main(FilePath, FileName, TextFile, Data_Sheet_Name, Equipment):


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
    Excel_File = shutil.copy(Excel_File_List[0][0], '../DataFile/034_N-electrode/')


    ########## DaraFrameの作成 ##########

    # ----- 取得開始行の取り出し -----
    Log.Log_Info(Log_File, 'Get The Starting Row Count')
    Start_Number = Row_Number_Func.start_row_number(TextFile) - 500

    # ----- ExcelデータをDataFrameとして取得 -----
    Log.Log_Info(Log_File, 'Read Excel')
    df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="B:AM", skiprows=Start_Number)


    # ----- 使う行だけ残す -----
    if Data_Sheet_Name == 'EB蒸着着工記録':
        df = df.iloc[:, [1, 2, 5, 10, 36]]
    if Data_Sheet_Name == 'EB蒸着#2着工記録':
        df = df.iloc[:, [1, 2, 4, 9, 35]]
    if Data_Sheet_Name == 'EB蒸着#3着工記録':
        df = df.iloc[:, [1, 2, 5, 10, 36]]    

 
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
        if i <5:
            print(df.iloc[i, 0],df.iloc[i, 2])
        if not isinstance(df.iloc[i, 2], (pd.Timestamp, datetime)):
            df.iloc[i, 2] = np.nan



    # ----- 今日から1か月前のデータまでを取得する -----
    df[2] = pd.to_datetime(df[2])
    one_month_ago = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) - relativedelta(months=1)
    df = df[(df[2] >= one_month_ago)]


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
        Serial_Number_str = str(df.iloc[Row_Number, Col_Serial_Number])
        Serial_Number_List = Serial_Number_str.split()

        # ----- ロット番号がなければ次の行へ -----
        if Serial_Number_List[0] == "nan" or len(Serial_Number_List) == 0:
            print(Serial_Number_str)
            Log.Log_Error(Log_File, "Not Lot\n")
            Row_Number += 1
            continue

        # ----- 全ロットの処理 -----
        for Serial_Number in Serial_Number_List:
    
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
                continue
    
            # ----- 品名が LDアレイ_ であれば処理を行わない -----
            if Part_Number == 'LDアレイ_':
                continue
    
            # ----- データの取得 -----
            data_dict = {
                "key_Start_Date_Time": df.iloc[Row_Number, Col_Start_Date_Time],
                "key_Part_Number": Part_Number,
                "key_Serial_Number": Serial_Number,
                "key_LotNumber_9": Nine_Serial_Number,
                "key_Operator": df.iloc[Row_Number, Col_Operator],
                "key_Program_No": str(df.iloc[Row_Number, Col_Program_No]),
                "key_Thickness": df.iloc[Row_Number, Col_Thickness],
                "key_Equipment": Equipment
            }
            #print('data:',data_dict)

            # ----- 取得した品種が"25g-EA"かつプログラムNo.が29.0のときに処理を継続 -----
            Part_Class = Part_Class_Check(Part_Number)
            print('Part_Class:',Part_Class)
        
            if (Part_Class != "10G-DFB-LD" or (data_dict["key_Program_No"] != "15.0" and data_dict["key_Program_No"] != "15")) and (Part_Class != "25G-DFB" or (data_dict["key_Program_No"] != "29.0" and data_dict["key_Program_No"] != "29")):
                Log.Log_Error(Log_File, Serial_Number + ' : ' + "Not Covered\n")
                continue

    
            ########## 日付フォーマットの変換 ##########
    
            # ----- 日付を指定されたフォーマットに変換する -----
            Log.Log_Info(Log_File, 'Date Format Conversion')
            data_dict["key_Start_Date_Time"] = Convert_Date.Edit_Date(data_dict["key_Start_Date_Time"])
    
            # ----- 指定したフォーマットに変換出来たか確認 -----
            if len(data_dict["key_Start_Date_Time"]) != 19:
                Log.Log_Error(Log_File, data_dict["key_Serial_Number"] + ' : ' + "Date Error\n")
                break


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
                break


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
                    '               <Header SerialNumber=' + '"' + data_dict["key_Serial_Number"] + '"' + ' PartNumber=' + '"' + str(data_dict["key_Part_Number"]) + '"' + ' Operation=' + '"' + Operation + '"' + ' TestStation=' + '"' + TestStation + '"' + ' Operator=' + '"' + str(data_dict["key_Operator"]) + '"' + ' StartTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Site=' + '"' + Site + '"' + ' LotNumber=' + '"' + data_dict["key_Serial_Number"] + '"/>' + '\n' +
                    '               <HeaderMisc>' + '\n' +
                    '                   <Item Description=' + '"LotNumber">' + Serial_Number_str.replace("\n", "") + '</Item>' + '\n'
                    '                   <Item Description=' + '"ProgramNo">' + str(data_dict["key_Program_No"]) + '</Item>' + '\n'
                    '               </HeaderMisc>' + '\n' +
                    '\n'
                    '               <TestStep Name="Thickness" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="Thickness" Units="A" Value=' + '"' + str(data_dict["key_Thickness"]) + '"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '\n'
                    '               <TestStep Name="SORTED_DATA" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                    '                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value=' + '"' + str(data_dict["key_STARTTIME_SORTED"]) + '"/>' + '\n' +
                    '                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value=' + '"' + str(data_dict["key_SORTNUMBER"]) + '"/>' + '\n' +
                    '                   <Data DataType="String" Name="LotNumber_5" Value=' + '"' + str(data_dict["key_Serial_Number"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                    '                   <Data DataType="String" Name="LotNumber_9" Value=' + '"' + str(data_dict["key_LotNumber_9"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                    '               </TestStep>' + '\n' +
                    '\n'
                    '               <TestEquipment>' + '\n' +
                    '                   <Item DeviceName="Dektak" DeviceSerialNumber=' + '"' + str(data_dict["key_Equipment"]) + '"/>' + '\n' +
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
            Log.Log_Info(Log_File, data_dict["key_Serial_Number"] + ' : ' + "OK\n" + Output_filepath + XML_File_Name)

        Row_Number+=1


    ########## 次の開始行数の書き込み ##########
    Log.Log_Info(Log_File, 'Write the next starting line number')
    Row_Number_Func.next_start_row_number(TextFile, Next_Start_Row)

    # ----- 最終行を書き込んだファイルをGドライブに転送 -----
    #shutil.copy(TextFile, 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/034_N-electrode/13_ProgramUsedFile/')


if __name__ == '__main__':

    Main("Z:/電極/着工記録/EB蒸着#1号機/", "*EB蒸着作業着工記録*.xlsx", "EB1_StartRow.txt", "EB蒸着着工記録", '#1')
    Main("Z:/電極/着工記録/EB蒸着#2号機/", "*EB蒸着#2号機作業着工記録*.xlsx", "EB2_StartRow.txt", "EB蒸着#2着工記録", '#2')
    Main("Z:/電極/着工記録/EB蒸着#3号機/", "*EB蒸着#3号機作業着工記録*.xlsx", "EB3_StartRow.txt", "EB蒸着#3着工記録", '#3')

# ----- ログ書込：Main処理の終了 -----
Log.Log_Info(Log_File, 'Program End')