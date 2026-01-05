import logging
import xlrd
import glob
import os
import sys

from datetime import datetime, timedelta, date
from time import strftime


########## 自作関数の定義 ##########
sys.path.append('../MyModule')
import Log
import Convert_Date
import Check


########## 全体パラメータ定義 ##########
Site = '350'
ProductFamily = 'SAG FAB'
Operation = 'N-electrode_Polish_MirrorPolishedThickness'
TestStation = 'N-electrode'


########## ログファイルの定義 ##########
Log_FolderName = str(date.today())
Log_file = '../Log/' + Log_FolderName + '/003_N-electrode.log'


########## シート名の定義 ##########
Data_Sheet_Name = '3ｲﾝﾁ用'
XY_Sheet_Name = 'ウェハ座標'


########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
# Output_filepath = '../XML/003_N-electrode/'


########## 取得するデータの列番号を定義 ##########
Row_Serial_Number = 3
Row_Start_Date_Time = 36
Row_Operator = 37
Row_Polish = 46
col_x = 1
col_y = 2


########## 取得した項目と型の対応表を定義 ##########
key_type = {
    "key_Start_Date_Time": str,
    "key_Part_Number": str,
    "key_Serial_Number": str,
    "key_Operator": str,
    "key_Polish1": float,
    "key_Polish2": float,
    "key_Polish3": float,
    "key_Polish4": float,
    "key_Polish5": float,
    "key_Polish6": float,
    "key_Polish7": float,
    "key_Polish8": float,
    "key_Polish9": float,
    "key_Polish10": float,
    "key_Polish11": float,
    "key_Polish12": float,
    "key_Polish13": float,
    "key_X1": float,
    "key_X2": float,
    "key_X3": float,
    "key_X4": float,
    "key_X5": float,
    "key_X6": float,
    "key_X7": float,
    "key_X8": float,
    "key_X9": float,
    "key_X10": float,
    "key_X11": float,
    "key_X12": float,
    "key_X13": float,
    "key_Y1": float,
    "key_Y2": float,
    "key_Y3": float,
    "key_Y4": float,
    "key_Y5": float,
    "key_Y6": float,
    "key_Y7": float,
    "key_Y8": float,
    "key_Y9": float,
    "key_Y10": float,
    "key_Y11": float,
    "key_Y12": float,
    "key_Y13": float
}


########## データの格納 ##########
def Open_Data_Sheet(filepath, Part_Number, Nine_Serial_Number):

    # ----- ログ書込：データの取得 -----
    Log.Log_Info(Log_file, 'Get Data')

    wb = xlrd.open_workbook(filepath, on_demand=True)
    sheet = wb.sheet_by_name(Data_Sheet_Name)
    Start_Date_Time = sheet.cell(Row_Start_Date_Time, 2).value

    Serial_Number = sheet.cell(Row_Serial_Number, 2).value
    Operator = sheet.cell(Row_Operator, 2).value

    # ----- ケモクロス研磨 最終ウェハ厚の格納 -----
    Polish = []
    for i in range(13):
        Polish.append(sheet.cell(Row_Polish, 3 + i).value)

    # ----- X/Y座標の格納 -----
    sheet = wb.sheet_by_name(XY_Sheet_Name)
    x, y = [], []
    for i in range(1, 14):
        x.append(sheet.cell(i, col_x).value)
        y.append(sheet.cell(i, col_y).value)

    wb.release_resources()

    # ----- 辞書型に格納 -----
    data_dict = {
        "key_Start_Date_Time": Start_Date_Time,
        "key_Part_Number": Part_Number,
        "key_Serial_Number": Serial_Number,
        "key_LotNumber_9": Nine_Serial_Number,
        "key_Operator": Operator,
        "key_Polish1": Polish[0],
        "key_Polish2": Polish[1],
        "key_Polish3": Polish[2],
        "key_Polish4": Polish[3],
        "key_Polish5": Polish[4],
        "key_Polish6": Polish[5],
        "key_Polish7": Polish[6],
        "key_Polish8": Polish[7],
        "key_Polish9": Polish[8],
        "key_Polish10": Polish[9],
        "key_Polish11": Polish[10],
        "key_Polish12": Polish[11],
        "key_Polish13": Polish[12],
        "key_X1": x[0],
        "key_X2": x[1],
        "key_X3": x[2],
        "key_X4": x[3],
        "key_X5": x[4],
        "key_X6": x[5],
        "key_X7": x[6],
        "key_X8": x[7],
        "key_X9": x[8],
        "key_X10": x[9],
        "key_X11": x[10],
        "key_X12": x[11],
        "key_X13": x[12],
        "key_Y1": y[0],
        "key_Y2": y[1],
        "key_Y3": y[2],
        "key_Y4": y[3],
        "key_Y5": y[4],
        "key_Y6": y[5],
        "key_Y7": y[6],
        "key_Y8": y[7],
        "key_Y9": y[8],
        "key_Y10": y[9],
        "key_Y11": y[10],
        "key_Y12": y[11],
        "key_Y13": y[12],
    }

    # ----- 着工者が空欄であれば'-'を入れる -----
    if data_dict["key_Operator"] == "":
        data_dict["key_Operator"] = '-'

    return data_dict


########## XML変換 ##########
def Output_XML(XML_File_Name, data_dict):

    # ----- ログ書込：XML変換 -----
    Log.Log_Info(Log_file, 'Excel -> XML')
    
    f = open(Output_filepath + XML_File_Name, 'w', encoding="utf-8")
        
    f.write('<?xml version="1.0" encoding="utf-8"?>' + '\n' +
            '<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' + '\n' +
            '       <Result startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Result="Done">' + '\n' +
            '               <Header SerialNumber=' + '"' + data_dict["key_Serial_Number"] + '"' + ' PartNumber=' + '"' + data_dict["key_Part_Number"] + '"' + ' Operation=' + '"' + Operation + '"' + ' TestStation=' + '"' + TestStation + '"' + ' Operator=' + '"' + data_dict["key_Operator"] + '"' + ' StartTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Site=' + '"' + Site + '"' + ' LotNumber=' + '"' + data_dict["key_Serial_Number"] + '"/>' + '\n' +
            '\n'
            '               <TestStep Name="Thickness1" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Done">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X1"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y1"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness" Units="um" Value=' + '"' + str(data_dict["key_Polish1"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="Thickness2" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Done">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X2"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y2"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness" Units="um" Value=' + '"' + str(data_dict["key_Polish2"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="Thickness3" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Done">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X3"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y3"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness" Units="um" Value=' + '"' + str(data_dict["key_Polish3"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="Thickness4" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Done">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X4"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y4"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness" Units="um" Value=' + '"' + str(data_dict["key_Polish4"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="Thickness5" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Done">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X5"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y5"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness" Units="um" Value=' + '"' + str(data_dict["key_Polish5"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="Thickness6" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Done">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X6"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y6"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness" Units="um" Value=' + '"' + str(data_dict["key_Polish6"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="Thickness7" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Done">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X7"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y7"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness" Units="um" Value=' + '"' + str(data_dict["key_Polish7"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="Thickness8" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Done">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X8"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y8"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness" Units="um" Value=' + '"' + str(data_dict["key_Polish8"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="Thickness9" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Done">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X9"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y9"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness" Units="um" Value=' + '"' + str(data_dict["key_Polish9"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="Thickness10" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Done">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X10"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y10"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness" Units="um" Value=' + '"' + str(data_dict["key_Polish10"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="Thickness11" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Done">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X11"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y11"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness" Units="um" Value=' + '"' + str(data_dict["key_Polish11"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="Thickness12" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Done">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X12"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y12"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness" Units="um" Value=' + '"' + str(data_dict["key_Polish12"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="Thickness13" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Done">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_X13"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Y13"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness" Units="um" Value=' + '"' + str(data_dict["key_Polish13"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="SORTED_DATA" startDateTime=' + '"' + data_dict["key_Start_Date_Time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="String" Name="LotNumber_5" Value=' + '"' + str(data_dict["key_Serial_Number"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
            '                   <Data DataType="String" Name="LotNumber_9" Value=' + '"' + str(data_dict["key_LotNumber_9"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '\n'
            '               <TestEquipment>' + '\n' +
            '                   <Item DeviceName="Stepmeter" DeviceSerialNumber="1"/>' + '\n' +
            '               </TestEquipment>' + '\n' +
            '\n'
            '               <ErrorData/>' + '\n' +
            '               <FailureData/>' + '\n' +
            '               <Configuration/>' + '\n' +
            '       </Result>' + '\n' +
            '</Results>'
            )
    f.close()


########## main処理 ##########
def main(File_Path, Part_Number, Nine_Serial_Number):

    # ----- ログ書込：オペレーション名 -----
    Log.Log_Info(Log_file, Operation)


    ########## データ取得 ##########
    data_dict = Open_Data_Sheet(File_Path, Part_Number, Nine_Serial_Number)


    ########## 日付フォーマット変換 ##########
    data_dict["key_Start_Date_Time"] = Convert_Date.Edit_Date(data_dict["key_Start_Date_Time"])


    ########## 空欄チェック ##########
    for val in data_dict.values():
        if val == "":
            Log.Log_Error(Log_file, data_dict["key_Serial_Number"] + ' : ' + "Blank Error\n")
            return


    ########## データ型の確認 ##########
    result = Check.Data_Type(key_type, data_dict)
    if result == False:
        Log.Log_Error(Log_file, data_dict["key_Serial_Number"] + ' : ' + "Data Error\n")
        return


    ########## XMLファイルの作成 ##########
    xml_file = 'Site=' + Site + ',ProductFamily=' + ProductFamily + ',Operation=' + Operation + \
               ',Partnumber=' + data_dict["key_Part_Number"] + ',Serialnumber=' + data_dict[
                   "key_Serial_Number"] + \
               ',Testdate=' + data_dict["key_Start_Date_Time"] + '.xml'

    Output_XML(xml_file, data_dict)