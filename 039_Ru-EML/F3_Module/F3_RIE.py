import openpyxl as px
import logging
import shutil
import pyodbc
import xlrd
import glob
import sys
import os

from datetime import datetime, timedelta, date
from time import strftime

########## 自作関数の定義 ##########
sys.path.append('../MyModule')
import SQL
import Log
import Convert_Date
import Row_Number_Func
import Check


########## 全体パラメータの定義 ##########
Site = '350'
ProductFamily = 'SAG FAB'
Operation = 'Ru-EML_F3RIE'
TestStation = 'Ru-EML'
X = '999999'
Y = '999999'


########## Logフォルダ名の定義 ##########
Log_FolderName = str(date.today())
Log_File = '../Log/' + Log_FolderName + '/039_Ru-EML_F3.log'


########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
# Output_filepath = '../Test/'


########## TestStepの定義 ##########
teststep_dict = {
    'TestStep1' : 'Coordinate',
    'TestStep2' : 'Particles',
    'TestStep3' : 'BallastN2',
    'TestStep4' : 'MO-Temperature',
    'TestStep5' : 'Dulation',
    'TestStep6' : 'MO7-TMI',
    'TestStep7' : 'MO6-Ru',
    'TestStep8' : 'PH3-A-50percent',
    'TestStep9' : 'CH3Cl-5percent',
    'TestStep10': 'Temperature',
    'TestStep11': 'TMInPiezoconConc',
    'TestStep12': 'TMInPiezoconConcOffset',
    'TestStep13': 'CH3Cl_Conc',
    'TestStep14': 'Thickness',
    'TestStep15': 'SORTED_DATA'
}


########## HeaderMiscの定義 ##########
HeaderMisc_dict = {
    'HeaderMisc1' : 'RecipeName-Macro',
    'HeaderMisc2' : 'RecipeName-Program',
    'HeaderMisc3' : 'RecipeName-Folder'
}


########## 取得した項目と型の対応表を定義 ##########
key_type = {
    "key_start_date_time" : str,
    "key_serial_number" : str,
    "key_part_number" : str,
    "key_operator" : str,
    "key_batch_number" : str,
    "key_HeaderMisc1" : str,
    "key_HeaderMisc2" : str,
    "key_HeaderMisc3" : str,
    "key_Particles_Particles_All_Address" : float,
    "key_Particles_Particles_Center" : float,
    "key_Particles_Particles_NG_Count" : float,
    "key_Particles_Particles_NG_Address" : str,
    "key_Particles_Particles" : float,
    "key_BallastN2_BallastN2" : float,
    "key_MO-Temperature_MO7-TMI" : float,
    "key_MO-Temperature_MO8-Fe" : float,
    "key_MO-Temperature_MO6-Ru" : float,
    "key_Dulation_Step9" : float,
    "key_Dulation_Step10" : float,
    "key_MO7-TMI_Step9" : float,
    "key_MO7-TMI_Step10" : float,
    "key_MO6-Ru_Step9" : float,
    "key_MO6-Ru_Step10" : float,
    "key_PH3-A-50percent_Step9" : float,
    "key_PH3-A-50percent_Step10" : float,
    "key_CH3Cl-5percent_Step9" : float,
    "key_CH3Cl-5percent_Step10" : float,
    "key_Temperature_Step9" : float,
    "key_Temperature_Step10" : float,
    "key_TMInPiezoconConc_TMInPiezoconConc_ZeroValue" : float,
    "key_TMInPiezoconConc_TMInPiezoconConc_Check" : float,
    "key_TMInPiezoconConc_TMInPiezoconConc_EPI" : float,
    "key_TMInPiezoconConc_TMInPiezoconConc_CheckFlow" : float,
    "key_TMInPiezoconConcOffset_TMInPiezoconConcOffset_Check" : float,
    "key_TMInPiezoconConcOffset_TMInPiezoconConcOffset_EPI" : float,
    "key_TMInPiezoconConcOffset_TMInPiezoconConcOffset_CheckFlow" : float,
    "key_CH3Cl_Conc_CH3Cl_ConcBonbe" : float,
    "key_Thickness_Thickness_Total" : float,
    "key_Thickness_Thickness_Ru-InP" : float,
    "key_Thickness_Thickness_Cap" : float,
    "key_STARTTIME_SORTED": float,
    "key_SORTNUMBER" : float,
    "key_LotNumber_9" : str
}


########## 空欄チェック ##########
def Get_Cells_Info(file_path, SheetName):

    # ----- ログ書込：空欄判定 -----
    Log.Log_Info(Log_File, "Blank Check")

    # ----- Excelファイルを開く -----
    wb = px.load_workbook(file_path, read_only=True, data_only=True)
    Sheet = wb[SheetName]

    # ----- False -> 空欄がない -----
    is_cells_empty = False

    # ----- 空欄はNone表示となる -----
    if Sheet['I8'].value is None or \
        Sheet['Q7'].value is None or \
        Sheet['AD44'].value is None:
        is_cells_empty = True

    return is_cells_empty


########## データ取得 ##########
def Open_Data_Sheet(file_path, SheetName, Multi_CH_List):

    # ----- ログ書込：データ取得 -----
    Log.Log_Info(Log_File, 'Data Acquisition')

    # ----- Excelファイルを開く -----
    wb = px.load_workbook(file_path, read_only=True, data_only=True)
    sheet = wb[SheetName]

    # ----- 辞書の作成 -----
    data_dict = dict()

    # ----- Serial_NumberをもとにPrimeから品名を引き出す -----
    serial_number = sheet["M8"].value
    conn, cursor = SQL.connSQL()

    # ----- Prime接続できなかったときはNoneが返ってくる -----
    if conn is None:
        Log.Log_Error(Log_File, serial_number + ' : ' + 'Connection with Prime Failed')
        sys.exit()
    part_number, Nine_Serial_Number = SQL.selectSQL(cursor, serial_number)
    SQL.disconnSQL(conn, cursor)

    # ----- 異物数判定 -----
    if sheet['M48'].value is None and sheet["AD52"] is None:
        Particles = ""
    elif sheet['M48'].value is not None:
        Particles = sheet['M48'].value
    else:
        Particles = sheet['AD52'].value

    # ----- データの取得 -----
    data_dict = {
        "key_start_date_time" : str(sheet['Q7'].value).replace(" ", "T"),
        "key_serial_number" : serial_number,
        "key_part_number" : part_number,
        "key_operator" : '-',
        "key_LotNumber_9": Nine_Serial_Number,
        "key_batch_number" : sheet['I8'].value,
        "key_HeaderMisc1" : sheet['U3'].value,
        "key_HeaderMisc2" : sheet['U4'].value,
        "key_HeaderMisc3" : sheet['U5'].value,
        "key_Particles_Particles_All_Address" : sheet['AD48'].value,
        "key_Particles_Particles_Center" : sheet['AD49'].value,
        "key_Particles_Particles_NG_Count" : sheet['AD50'].value,
        "key_Particles_Particles_NG_Address" : sheet['AD51'].value,
        "key_Particles_Particles" : Particles,
        "key_BallastN2_BallastN2" : sheet['AD44'].value,
        "key_MO-Temperature_MO7-TMI" : sheet["K42"].value,
        "key_MO-Temperature_MO8-Fe" : sheet["L42"].value,
        "key_MO-Temperature_MO6-Ru" : sheet["M42"].value,
        "key_Dulation_Step9" : sheet['I22'].value,
        "key_Dulation_Step10" : sheet['I23'].value,
        "key_MO7-TMI_Step9" : sheet['K22'].value,
        "key_MO7-TMI_Step10" : sheet['K23'].value,
        "key_MO6-Ru_Step9" : sheet['M22'].value,
        "key_MO6-Ru_Step10" : sheet['M23'].value,
        "key_PH3-A-50percent_Step9" : sheet['T22'].value,
        "key_PH3-A-50percent_Step10" : sheet['T23'].value,
        "key_CH3Cl-5percent_Step9" : sheet['W22'].value,
        "key_CH3Cl-5percent_Step10" : sheet['W23'].value,
        "key_Temperature_Step9" : sheet["Z22"].value,
        "key_Temperature_Step10" : sheet["Z23"].value,
        "key_TMInPiezoconConc_TMInPiezoconConc_ZeroValue" : sheet["AA31"].value,
        "key_TMInPiezoconConc_TMInPiezoconConc_Check" : sheet["AA32"].value,
        "key_TMInPiezoconConc_TMInPiezoconConc_EPI" : sheet["AA33"].value,
        "key_TMInPiezoconConc_TMInPiezoconConc_CheckFlow" : sheet["AA35"].value,
        "key_TMInPiezoconConcOffset_TMInPiezoconConcOffset_Check" : sheet["AD32"].value,
        "key_TMInPiezoconConcOffset_TMInPiezoconConcOffset_EPI" : sheet["AD33"].value,
        "key_TMInPiezoconConcOffset_TMInPiezoconConcOffset_CheckFlow" : sheet["AD35"].value,
        "key_CH3Cl_Conc_CH3Cl_ConcBonbe" : sheet['AD37'].value,
        "key_Thickness_Thickness_Total" : Multi_CH_List[0],
        "key_Thickness_Thickness_Ru-InP" : Multi_CH_List[1],
        "key_Thickness_Thickness_Cap" : Multi_CH_List[2],
        "key_Equipment_SEM" : Multi_CH_List[3],
        "key_Equipment_MOCVD" : "3"
    }

    # ----- 空欄箇所はNoneとして取得される。Noneは文字列に変換できないため、空欄("")に置き換える -----
    for keys in data_dict:
        if data_dict[keys] is None or data_dict[keys] == "None" or data_dict[keys] == "#VALUE!" or (data_dict[keys] == '-' and keys != "key_operator"):
            data_dict[keys] = ""

    return data_dict


########## XML変換 ##########
def Output_XML(xml_file, data_dict):

    # ----- ログ書込：XML変換 -----
    Log.Log_Info(Log_File, 'Excel File To XML File Conversion')
    
    f = open(Output_filepath + xml_file, 'w', encoding="utf-8")

    f.write('<?xml version="1.0" encoding="utf-8"?>' + '\n' +
            '<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' + '\n' +
            '       <Result startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Result="Passed">' + '\n' +
            '               <Header SerialNumber=' + '"' + data_dict["key_serial_number"] + '"' + ' PartNumber=' + '"' + data_dict["key_part_number"] + '"' + ' Operation=' + '"' + Operation + '"' + ' TestStation=' + '"' + TestStation + '"' + ' Operator=' + '"' + data_dict["key_operator"] + '"' + ' StartTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Site=' + '"' + Site + '"' + ' BatchNumber=' + '"' + data_dict["key_batch_number"] + '"' + ' LotNumber=' + '"' + data_dict["key_serial_number"] + '"/>' + '\n' +
            '               <HeaderMisc>' + '\n' +
            '                   <Item Description=' + '"' + HeaderMisc_dict["HeaderMisc1"] + '">' + data_dict["key_HeaderMisc1"] + '</Item>' + '\n'
            '                   <Item Description=' + '"' + HeaderMisc_dict["HeaderMisc2"] + '">' + data_dict["key_HeaderMisc2"] + '</Item>' + '\n'
            '                   <Item Description=' + '"' + HeaderMisc_dict["HeaderMisc3"] + '">' + data_dict["key_HeaderMisc3"] + '</Item>' + '\n'
            '               </HeaderMisc>' + '\n' +
            '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep1"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + X + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + Y + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep2"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="Particles_All_Address" Units="count" Value=' + '"' + str(data_dict["key_Particles_Particles_All_Address"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Particles_Center" Units="count" Value=' + '"' + str(data_dict["key_Particles_Particles_Center"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Particles_NG_Count" Units="count" Value=' + '"' + str(data_dict["key_Particles_Particles_NG_Count"]) + '"/>' + '\n' +
            '                   <Data DataType="String" Name="Particles_NG_Address" Value=' + '"' + str(data_dict["key_Particles_Particles_NG_Address"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Particles" Units="count" Value=' + '"' + str(data_dict["key_Particles_Particles"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep3"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="BallastN2" Units="slm" Value=' + '"' + str(data_dict["key_BallastN2_BallastN2"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep4"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="MO7-TMI" Units="degree" Value=' + '"' + str(data_dict["key_MO-Temperature_MO7-TMI"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="MO8-Fe" Units="degree" Value=' + '"' + str(data_dict["key_MO-Temperature_MO8-Fe"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="MO6-Ru" Units="degree" Value=' + '"' + str(data_dict["key_MO-Temperature_MO6-Ru"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep5"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="Step9" Units="min" Value=' + '"' + str(data_dict["key_Dulation_Step9"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Step10" Units="min" Value=' + '"' + str(data_dict["key_Dulation_Step10"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep6"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="Step9" Units="sccm" Value=' + '"' + str(data_dict["key_MO7-TMI_Step9"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Step10" Units="sccm" Value=' + '"' + str(data_dict["key_MO7-TMI_Step10"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep7"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="Step9" Units="sccm" Value=' + '"' + str(data_dict["key_MO6-Ru_Step9"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Step10" Units="sccm" Value=' + '"' + str(data_dict["key_MO6-Ru_Step10"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep8"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="Step9" Units="sccm" Value=' + '"' + str(data_dict["key_PH3-A-50percent_Step9"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Step10" Units="sccm" Value=' + '"' + str(data_dict["key_PH3-A-50percent_Step10"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep9"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="Step9" Units="sccm" Value=' + '"' + str(data_dict["key_CH3Cl-5percent_Step9"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Step10" Units="sccm" Value=' + '"' + str(data_dict["key_CH3Cl-5percent_Step10"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep10"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="Step9" Units="degree" Value=' + '"' + str(data_dict["key_Temperature_Step9"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Step10" Units="degree" Value=' + '"' + str(data_dict["key_Temperature_Step10"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep11"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="TMInPiezoconConc_ZeroValue" Units="percent" Value=' + '"' + str(data_dict["key_TMInPiezoconConc_TMInPiezoconConc_ZeroValue"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="TMInPiezoconConc_Check" Units="percent" Value=' + '"' + str(data_dict["key_TMInPiezoconConc_TMInPiezoconConc_Check"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="TMInPiezoconConc_EPI" Units="percent" Value=' + '"' + str(data_dict["key_TMInPiezoconConc_TMInPiezoconConc_EPI"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="TMInPiezoconConc_CheckFlow" Units="sccm" Value=' + '"' + str(data_dict["key_TMInPiezoconConc_TMInPiezoconConc_CheckFlow"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep12"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="TMInPiezoconConcOffset_Check" Units="percent" Value=' + '"' + str(data_dict["key_TMInPiezoconConcOffset_TMInPiezoconConcOffset_Check"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="TMInPiezoconConcOffset_EPI" Units="percent" Value=' + '"' + str(data_dict["key_TMInPiezoconConcOffset_TMInPiezoconConcOffset_EPI"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="TMInPiezoconConcOffset_CheckFlow" Units="sccm" Value=' + '"' + str(data_dict["key_TMInPiezoconConcOffset_TMInPiezoconConcOffset_CheckFlow"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep13"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="CH3Cl_ConcBonbe" Units="percent" Value=' + '"' + str(data_dict["key_CH3Cl_Conc_CH3Cl_ConcBonbe"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep14"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness_Total" Units="nm" Value=' + '"' + str(data_dict["key_Thickness_Thickness_Total"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness_Ru-InP" Units="nm" Value=' + '"' + str(data_dict["key_Thickness_Thickness_Ru-InP"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="Thickness_Cap" Units="nm" Value=' + '"' + str(data_dict["key_Thickness_Thickness_Cap"]) + '"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '               <TestStep Name=' + '"' + teststep_dict["TestStep15"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' +
            '                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value=' + '"' + str(data_dict["key_STARTTIME_SORTED"]) + '"/>' + '\n' +
            '                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value=' + '"' + str(data_dict["key_SORTNUMBER"]) + '"/>' + '\n' +
            '                   <Data DataType="String" Name="LotNumber_5" Value=' + '"' + str(data_dict["key_serial_number"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
            '                   <Data DataType="String" Name="LotNumber_9" Value=' + '"' + str(data_dict["key_LotNumber_9"]) + '"' + ' CompOperation="LOG"/>' + '\n' +
            '               </TestStep>' + '\n' +
            '\n'
            '               <TestEquipment>' + '\n' +
            '                   <Item DeviceName="SEM" DeviceSerialNumber="' + data_dict["key_Equipment_SEM"] + '"/>' + '\n' +
            '                   <Item DeviceName="MOCVD" DeviceSerialNumber="' + data_dict["key_Equipment_MOCVD"] + '"/>' + '\n' +
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
def main(file_path, SheetName, Multi_CH_List):

    # ----- 空欄チェック -----
    if Get_Cells_Info(file_path, SheetName):
        Log.Log_Error(Log_File, "Blank Error\n")
        return
    else:
        Log.Log_Info(Log_File, "OK")

    # ----- データ取得 -----
    data_dict = Open_Data_Sheet(file_path, SheetName, Multi_CH_List)

    # ----- 品名が正しいかチェック -----
    if len(data_dict["key_part_number"]) == 0:
        Log.Log_Error(Log_File, data_dict["key_serial_number"] + ' : ' + "Part Number Error\n")
        return -1

    # ----- 日付フォーマット変換 -----
    if len(data_dict['key_start_date_time']) != 19 or '年' in data_dict['key_start_date_time']:
        Log.Log_Error(Log_File, data_dict["key_serial_number"] + ' : ' + "Date Error\n")
        return -1

    # ----- STARTTIME_SORTEDの追加 -----

    # 日付をExcel時間に変換する
    date = datetime.strptime(str(data_dict["key_start_date_time"]).replace('T', ' ').replace('.', ':'), "%Y-%m-%d %H:%M:%S")
    date_excel_number = int(str(date - datetime(1899, 12, 30)).split()[0])

    # エピ番号の数値部だけを取得する
    Epi_Number = 0
    for i in data_dict["key_batch_number"]:
        try:
            if 0<=int(i)<=9:
                Epi_Number = Epi_Number*10+int(i)
        except:
            pass

    # エピ番号を10^6で割って、excel時間に加算する
    date_excel_number += Epi_Number/10**6

    # data_dictに登録する
    data_dict["key_STARTTIME_SORTED"] = date_excel_number
    data_dict["key_SORTNUMBER"] = Epi_Number

    # ----- データ型の確認 -----
    result = Check.Data_Type(key_type, data_dict)
    if result == False:
        Log.Log_Error(Log_File, data_dict["key_serial_number"] + ' : ' + "Data Error\n")
        return

    # ----- XMLファイルの作成 -----
    xml_file = 'Site=' + Site + ',ProductFamily=' + ProductFamily + ',Operation=' + Operation + \
               ',Partnumber=' + data_dict["key_part_number"] + ',Serialnumber=' + data_dict["key_serial_number"] + \
               ',Testdate=' + data_dict["key_start_date_time"].replace(':', '.') + '.xml'

    Output_XML(xml_file, data_dict)
    Log.Log_Info(Log_File, data_dict["key_serial_number"] + ' : ' + "OK\n")