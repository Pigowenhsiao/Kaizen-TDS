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


########## 全体パラメータの定義 ##########
Site = '350'
ProductFamily = 'SAG FAB'
Operation = 'SEM-DML_MESA_Height'
TestStation = 'SEM-DML'


########## Logの設定 ##########
Log_FolderName = str(date.today())
Log_file = '../Log/' + Log_FolderName + '/032_SEM-DML.log'


########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
# Output_filepath = '../XML/032_SEM-DML/'


########## 取得するデータの列番号を定義 ##########
Row_Operator = 10
Row_Height_1 = 19
Row_Height_2 = 20
Row_Height_3 = 21
Row_Height_Ave = 22


########## 取得した項目と型の対応表を定義 ##########
key_type = {
    "key_Start_date_time": str,
    "key_Part_number": str,
    "key_Serial_number": str,
    "key_Operator": str,
    "key_Height_1": float,
    "key_Height_2": float,
    "key_Height_3": float,
    "key_Height_Ave": float
}


########## データの格納 ##########
def Open_Data_Sheet(filepath, start_date_time, part_number, serial_number, nine_serial_number):

    # ----- ログ書込：データの取得 -----
    Log.Log_Info(Log_file, 'Data Acquisition')

    wb = xlrd.open_workbook(filepath, on_demand=True)
    sheet = wb.sheet_by_index(0)

    # ----- 辞書型に格納 -----
    data_dict = {
        "key_Start_date_time": start_date_time,
        "key_Part_number": part_number,
        "key_Serial_number": serial_number,
        "key_LotNumber_9": nine_serial_number,
        "key_Operator": sheet.cell(Row_Operator, 4).value,
        "key_Height_1": sheet.cell(Row_Height_1, 4).value,
        "key_Height_2": sheet.cell(Row_Height_2, 4).value,
        "key_Height_3": sheet.cell(Row_Height_3, 4).value,
        "key_Height_Ave": sheet.cell(Row_Height_Ave, 4).value
    }

    wb.release_resources()

    if data_dict["key_Operator"] == "":
        data_dict["key_Operator"] = '-'


    return data_dict


########## XMLファイルの作成 ##########
def Output_XML(xml_file, data_dict):

    # ----- ログ書込：XML変換 -----
    Log.Log_Info(Log_file, 'Excel File To XML File Conversion')

    f = open(Output_filepath + xml_file, 'w', encoding="utf-8")
        
    f.write('<?xml version="1.0" encoding="utf-8"?>' + '\n' +
            '<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' + '\n' +
            '       <Result startDateTime=' + '"' + data_dict["key_Start_date_time"].replace(".", ":") + '"' + ' Result="Passed">' + '\n' +
            '               <Header SerialNumber=' + '"' + data_dict["key_Serial_number"] + '"' + ' PartNumber=' + '"' + data_dict["key_Part_number"] + '"' + ' Operation=' + '"' + Operation + '"' + ' TestStation=' + '"' + TestStation + '"' + ' Operator=' + '"' + data_dict["key_Operator"] + '"' + ' StartTime=' + '"' + data_dict["key_Start_date_time"].replace(".", ":") + '"' + ' Site=' + '"' + Site + '"' + ' LotNumber=' + '"' + data_dict["key_Serial_number"] + '"/>' + '\n' +
            '\n'
            '               <TestStep Name="MESA" startDateTime=' + '"' + data_dict["key_Start_date_time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="Height_1" Units="um" Value=' + '"' + str(data_dict["key_Height_1"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Height_2" Units="um" Value=' + '"' + str(data_dict["key_Height_2"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Height_3" Units="um" Value=' + '"' + str(data_dict["key_Height_3"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Height_Ave" Units="um" Value=' + '"' + str(data_dict["key_Height_Ave"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name="SORTED_DATA" startDateTime=' + '"' + data_dict["key_Start_date_time"].replace(".", ":") + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="String" Name="LotNumber_5" Value=' + '"' + str(data_dict["key_Serial_number"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
            '                   <Data DataType="String" Name="LotNumber_9" Value=' + '"' + str(data_dict["key_LotNumber_9"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '\n'
            '               <TestEquipment>' + '\n' +
            '                   <Item DeviceName="SEM" DeviceSerialNumber="#1"/>' + '\n' +
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
def main(file_path, start_date, part_number, serial_number, nine_serial_number):

    # ----- ログ書込：オペレーション名 -----
    Log.Log_Info(Log_file, Operation)

    # ----- データ取得 -----
    data_dict = Open_Data_Sheet(file_path, start_date, part_number, serial_number, nine_serial_number)

    # ----- データ型の確認 -----
    result = Check.Data_Type(key_type, data_dict)
    if result == False:
        Log.Log_Error(Log_file, data_dict["key_Serial_number"] + ' : ' + "Data Error\n")
        return

    # ----- XMLファイルの作成 -----
    xml_file = 'Site=' + Site + ',ProductFamily=' + ProductFamily + ',Operation=' + Operation + \
               ',Partnumber=' + data_dict["key_Part_number"] + ',Serialnumber=' + data_dict["key_Serial_number"] + \
               ',Testdate=' + data_dict["key_Start_date_time"] + '.xml'

    Output_XML(xml_file, data_dict)