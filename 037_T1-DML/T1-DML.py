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
Operation = 'T1-DML'
TestStation = 'T1-DML'


########## Logの設定 ##########

# ----- ログファイルの作成 -----
Log_Folder_Name = str(date.today())
if not os.path.exists("../Log/" + Log_Folder_Name):
    os.makedirs("../Log/" + Log_Folder_Name)

Log_File = '../Log/' + Log_Folder_Name + '/037_T1-DML.log'
Log.Log_Info(Log_File, 'Program Start')


########## シート名の定義 ##########
Data_Sheet_Name = 'Growth Report'


########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
# Output_filepath = '../XML/037_T1-DML/'


########## 取得するデータの列番号を定義 ##########
Col_Start_Date_Time = 0
Col_ID_GrowthID = 1
Col_ID_SubstrateID = 2
Col_Serial_Number = 3
Col_PL_Wavelength_Center = 4
Col_PL_Wavelength_Average = 5
Col_PL_Wavelength_Median = 6
Col_PL_Wavelength_Sigma = 7
Col_PL_Wavelength_Delta = 8
Col_PL_Intensity_Rate_HH_LH = 9
Col_PL_Wavelength_HH_LH = 10
Col_PL_Intensity = 11
Col_PL_FWHM = 12
Col_PL_Intensity_Rate = 13
Col_XRD_Strain = 14
Col_XRD_Thickness = 15
Col_PL_Intensity_Rate_Calk = 16
Col_PL_Equipment = 17
Col_XRD_Equipment = 18


########## 取得した項目と型の対応表を定義 ##########
key_type = {
    'key_Start_Date_Time': str,
    'key_Serial_Number': str,
    'key_Part_Number': str,
    'key_PL_Equipment': str,
    'key_XRD_Equipment': str,
    'key_ID_GrowthID': str,
    'key_ID_SubstrateID': str,
    'key_PL_Wavelength_Center': float,
    'key_PL_Wavelength_Average': float,
    'key_PL_Wavelength_Median': float,
    'key_PL_Wavelength_Sigma': float,
    'key_PL_Wavelength_Delta': float,
    'key_PL_Intensity_Rate_HH_LH': float,
    'key_PL_Wavelength_HH_LH': float,
    'key_PL_Intensity': float,
    'key_PL_FWHM': float,
    'key_PL_Intensity_Rate': float,
    'key_PL_Intensity_Rate_Calk': float,
    'key_XRD_Strain': float,
    'key_XRD_Thickness': float,
    'key_STARTTIME_SORTED' : float,
    'key_SORTNUMBER' : float,
    "key_LotNumber_9": str
}


def Main():


    ########## Excelファイルをローカルにコピー ##########

    # ----- 正規表現で取出し、直近で変更があったファイルを取得する -----
    Log.Log_Info(Log_File, 'Excel File Copy')

    FilePath = 'Z:/MOCVD/CAS-T1多層/'
    FileName = 'Growth*Report*.xls*'
    Excel_File_List = []
    for file in glob.glob(FilePath + FileName):
        if '$' not in file:
            dt = strftime("%Y-%m-%d %H:%M:%S", localtime(os.path.getmtime(file)))
            Excel_File_List.append([file, dt])

    # ----- dt(更新日時)の降順で並び替える -----
    Excel_File_List = sorted(Excel_File_List, key=lambda x: x[1], reverse=True)
    Excel_File = shutil.copy(Excel_File_List[0][0], '../DataFile/037_T1-DML/')


    ########## DaraFrameの作成 ##########

    # ----- 取得開始行の取り出し -----
    Log.Log_Info(Log_File, 'Get The Starting Row Count')
    Start_Number = Row_Number_Func.start_row_number("T1-DML_StartRow.txt") - 500

    # ----- ExcelデータをDataFrameとして取得 -----
    Log.Log_Info(Log_File, 'Read Excel')
    df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="D:AX", skiprows=7)

    # ----- 必要な列のみ取出し -----
    # ----- 必要な列のみ取出し -----
    df = df.iloc[:, [0, 2, 3, 20, 25, 26, 27, 28, 29, 33, 34, 38, 39, 40, 42, 43, 44, 45, 46]]

    # ----- 丟棄df.iloc[6] 是空值的欄 -----
    df = df.dropna(subset=[df.columns[4]])



    # ----- 列番号の振り直し -----
    Log.Log_Info(Log_File, 'Setting Columns Number')
    df.columns = range(df.shape[1])

    # ----- PL/XRD列の空欄は #1 に置き換える -----
    df[17] = df[17].ffill()
    df[18] = df[18].ffill()

    print(df.head(20))
    print(df.tail(20))  
    # ----- 末尾から欠損のデータを落としていく -----

    Missing_Row = 0
    Getting_Row = len(df) - 1

    null_series = df.isnull().any(axis=1)  # 先獲取 Series

    while Getting_Row >= 0 and null_series.iloc[Getting_Row]:
        Missing_Row += 1
        Getting_Row -= 1

    df = df[:Getting_Row + 1]

    # ----- 次の開始行数をメモ -----
    Next_Start_Row = Start_Number + df.shape[0] + 1
  
    # ----- 日付欄に文字列が入っていたらNoneに置き換える -----
    for i in range(df.shape[0]):

        try:
            df.iloc[i, 0] = pd.to_datetime(df.iloc[i, 0], errors='coerce')
        except Exception as e:
            df.iloc[i, 0] = np.nan
    

    # ----- 今日から1か月前のデータまでを取得する -----
    df[0] = pd.to_datetime(df[0])
    one_month_ago = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) - relativedelta(months=3)
    df = df[(df[0] >= one_month_ago)]
        

    # ----- 行番号の振り直し -----
    df = df.reset_index(drop=True)


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
            'key_Serial_Number': Serial_Number,
            'key_Part_Number': Part_Number,
            "key_LotNumber_9": Nine_Serial_Number,
            'key_PL_Equipment': df.iloc[Row_Number, Col_PL_Equipment],
            'key_XRD_Equipment': df.iloc[Row_Number, Col_XRD_Equipment],
            'key_ID_GrowthID': df.iloc[Row_Number, Col_ID_GrowthID],
            'key_ID_SubstrateID': df.iloc[Row_Number, Col_ID_SubstrateID],
            'key_PL_Wavelength_Center': df.iloc[Row_Number, Col_PL_Wavelength_Center],
            'key_PL_Wavelength_Average': df.iloc[Row_Number, Col_PL_Wavelength_Average],
            'key_PL_Wavelength_Median': df.iloc[Row_Number, Col_PL_Wavelength_Median],
            'key_PL_Wavelength_Sigma': df.iloc[Row_Number, Col_PL_Wavelength_Sigma],
            'key_PL_Wavelength_Delta': df.iloc[Row_Number, Col_PL_Wavelength_Delta],
            'key_PL_Intensity_Rate_HH_LH': df.iloc[Row_Number, Col_PL_Intensity_Rate_HH_LH],
            'key_PL_Wavelength_HH_LH': df.iloc[Row_Number, Col_PL_Wavelength_HH_LH],
            'key_PL_Intensity': df.iloc[Row_Number, Col_PL_Intensity],
            'key_PL_FWHM': df.iloc[Row_Number, Col_PL_FWHM],
            'key_PL_Intensity_Rate': df.iloc[Row_Number, Col_PL_Intensity_Rate],
            'key_PL_Intensity_Rate_Calk': df.iloc[Row_Number, Col_PL_Intensity_Rate_Calk],
            'key_XRD_Strain': df.iloc[Row_Number, Col_XRD_Strain],
            'key_XRD_Thickness': df.iloc[Row_Number, Col_XRD_Thickness],
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
                '               <Header SerialNumber=' + '"' + data_dict["key_Serial_Number"] + '"' + ' PartNumber=' + '"' + data_dict["key_Part_Number"] + '"' + ' Operation=' + '"' + Operation + '"' + ' TestStation=' + '"' + TestStation + '"' + ' Operator="-" StartTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Site=' + '"' + Site + '"' + ' LotNumber=' + '"' + data_dict["key_Serial_Number"] + '"/>' + '\n' +
                '\n'
                '               <TestStep Name="ID" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="String" Name="GrowthID" Units="ID" Value=' + '"' + str(data_dict["key_ID_GrowthID"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                '                   <Data DataType="String" Name="SubstrateID" Units="ID" Value=' + '"' + str(data_dict["key_ID_SubstrateID"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '\n'
                '               <TestStep Name="PL" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="PL-Wavelength-Center" Units="nm" Value=' + '"' + str(data_dict["key_PL_Wavelength_Center"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="PL-Wavelength-Average" Units="nm" Value=' + '"' + str(data_dict["key_PL_Wavelength_Average"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="PL-Wavelength-Median" Units="nm" Value=' + '"' + str(data_dict["key_PL_Wavelength_Median"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="PL-Wavelength-Sigma" Units="nm" Value=' + '"' + str(data_dict["key_PL_Wavelength_Sigma"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="PL-Wavelength-Delta" Units="nm" Value=' + '"' + str(data_dict["key_PL_Wavelength_Delta"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="PL-Intensity-Rate_HH-LH" Units="a.u." Value=' + '"' + str(data_dict["key_PL_Intensity_Rate_HH_LH"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="PL-Wavelength_HH-LH" Units="nm" Value=' + '"' + str(data_dict["key_PL_Wavelength_HH_LH"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="PL-Intensity" Units="count" Value=' + '"' + str(data_dict["key_PL_Intensity"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="PL-FWHM" Units="meV" Value=' + '"' + str(data_dict["key_PL_FWHM"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="PL-Intensity_Rate" Units="a.u." Value=' + '"' + str(data_dict["key_PL_Intensity_Rate"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="PL-Intensity_Rate_Calk" Units="a.u." Value=' + '"' + str(data_dict["key_PL_Intensity_Rate_Calk"]) + '"/>' + '\n' +
                '               </TestStep>' + '\n' +
                '\n'
                '               <TestStep Name="XRD" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
                '                   <Data DataType="Numeric" Name="XRD-Strain" Units="percent" Value=' + '"' + str(data_dict["key_XRD_Strain"]) + '"/>' + '\n' +
                '                   <Data DataType="Numeric" Name="XRD-Thickness" Units="nm" Value=' + '"' + str(data_dict["key_XRD_Thickness"]) + '"/>' + '\n' +
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
                '                   <Item DeviceName="PL" DeviceSerialNumber="' + str(data_dict['key_PL_Equipment']) + '"/>' + '\n' +
                '                   <Item DeviceName="XRD" DeviceSerialNumber="' + str(data_dict['key_XRD_Equipment']) + '"/>' + '\n' +
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
    Row_Number_Func.next_start_row_number("T1-DML_StartRow.txt", Next_Start_Row)

    # ----- 最終行を書き込んだファイルをGドライブに転送 -----
    shutil.copy("T1-DML_StartRow.txt", 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/037_T1-DML/13_ProgramUsedFile/')


if __name__ == '__main__':

    Main()


# ----- ログ書込：Main処理の終了 -----
Log.Log_Info(Log_File, 'Program End')