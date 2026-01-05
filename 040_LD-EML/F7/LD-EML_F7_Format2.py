import openpyxl as px
import pandas as pd
import logging
import shutil
import pyodbc
import xlrd
import glob
import sys
import os

from datetime import datetime, timedelta, date
from time import strftime, localtime


########## 自作関数の定義 ##########
sys.path.append('../../MyModule')
import SQL
import Log
import ExpandExp
import Convert_Date
import Row_Number_Func
import MOCVD_OldFileSearch
import Check


########## 全体パラメータの定義 ##########
Site = '350'
ProductFamily = 'SAG FAB'
Operation = 'LD-EML_F7_Format2'
TestStation = 'LD-EML'
X = '999999'
Y = '999999'


########## Logフォルダ名の定義 ##########
Log_FolderName = "-".join(str(date.today()).split("-"))

# ----- 格納するLogフォルダがなければ作成する -----
if not os.path.exists("../../Log/" + Log_FolderName):
    os.makedirs("../../Log/" + Log_FolderName)

# ----- ログ書き込み先ファイルパス -----
Log_file = '../../Log/' + Log_FolderName + '/040_LD-EML_F7_Format2.log'

# ----- ログ書込：プログラムの開始 -----
Log.Log_Info(Log_file, 'Program Start')


########## XML出力先ファイルパス ##########
Output_filepath = '//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/'
#Output_filepath = '../../XML/'
# Output_filepath = '../../XML/040_LD-EML/F7_Format2/'


########## シート名定義 ##########
Data_sheet_name = "まとめ"


########## TestStepの定義 ##########
teststep_dict = {
    'TestStep1': 'Coordinate',
    'TestStep2': 'Mapper_3sigma',
    'TestStep3': 'Mapper_Average',
    'TestStep4': 'Mapper_Adjust',
    'TestStep5': 'Center',
    'TestStep6': 'A',
    'TestStep7': 'B',
    'TestStep8': 'C',
    'TestStep9': 'D',
    'TestStep10': 'AB',
    'TestStep11': 'AC',
    'TestStep12': 'BC',
    'TestStep13': 'CD',
    'TestStep14': 'AA',
    'TestStep15': 'BB',
    'TestStep16': 'CC',
    'TestStep17': 'DD',
    'TestStep18': 'Measurements',
    'TestStep19': 'SORTED_DATA'
}


########## 取得した項目と型の対応表を定義 ##########
key_type = {
    "key_start_date_time": str,
    "key_part_number": str,
    "key_serial_number": str,
    "key_operator": str,
    "key_batch_number": str,
    "key_Mapper_3sigma_Mapper_PL_Lambda": float,
    "key_Mapper_3sigma_Mapper_PL_Intensity": float,
    "key_Mapper_3sigma_Mapper_PL_FWHM": float,
    "key_Mapper_Average_Mapper_PL_Lambda": float,
    "key_Mapper_Average_Mapper_PL_Intensity": float,
    "key_Mapper_Average_Mapper_PL_FWHM": float,
    "key_Mapper_Average_Mapper_PL_Tails": float, #Add New Item 2025/1/8 M.Yoshida
    "key_Mapper_Average_Lambda_diff": float,  # Add New Item 2025/2/7 M.Yoshida
    "key_Mapper_Average_Mapper_PL_StrengthRatio": float,
    "key_Mapper_Adjust_Mapper_TargetWavelength": float,
    "key_Mapper_Adjust_Mapper_Wavelength(1.3um)": float,
    "key_Mapper_Adjust_Mapper_CheckingAdjustValue": float,
    "key_Center_X": float,
    "key_Center_Y": float,
    "key_Center_Lambda": float,
    "key_Center_Intensity": float,
    "key_Center_FWHM": float,
    "key_Center_Tails": float,
    "key_Center_DeltaLambda": float,
    "key_A_X": float,
    "key_A_Y": float,
    "key_A_Lambda": float,
    "key_A_Intensity": float,
    "key_A_FWHM": float,
    "key_A_Tails": float,
    "key_A_DeltaLambda": float,
    "key_B_X": float,
    "key_B_Y": float,
    "key_B_Lambda": float,
    "key_B_Intensity": float,
    "key_B_FWHM": float,
    "key_B_Tails": float,
    "key_B_DeltaLambda": float,
    "key_C_X": float,
    "key_C_Y": float,
    "key_C_Lambda": float,
    "key_C_Intensity": float,
    "key_C_FWHM": float,
    "key_C_Tails": float,
    "key_C_DeltaLambda": float,
    "key_D_X": float,
    "key_D_Y": float,
    "key_D_Lambda": float,
    "key_D_Intensity": float,
    "key_D_FWHM": float,
    "key_D_Tails": float,
    "key_D_DeltaLambda": float,
    "key_AB_X": float,
    "key_AB_Y": float,
    "key_AB_Lambda": float,
    "key_AB_Intensity": float,
    "key_AB_FWHM": float,
    "key_AB_Tails": float,
    "key_AB_DeltaLambda": float,
    "key_AC_X": float,
    "key_AC_Y": float,
    "key_AC_Lambda": float,
    "key_AC_Intensity": float,
    "key_AC_FWHM": float,
    "key_AC_Tails": float,
    "key_AC_DeltaLambda": float,
    "key_BC_X": float,
    "key_BC_Y": float,
    "key_BC_Lambda": float,
    "key_BC_Intensity": float,
    "key_BC_FWHM": float,
    "key_BC_Tails": float,
    "key_BC_DeltaLambda": float,
    "key_CD_X": float,
    "key_CD_Y": float,
    "key_CD_Lambda": float,
    "key_CD_Intensity": float,
    "key_CD_FWHM": float,
    "key_CD_Tails": float,
    "key_CD_DeltaLambda": float,
    "key_AA_X": float,
    "key_AA_Y": float,
    "key_AA_Lambda": float,
    "key_AA_Intensity": float,
    "key_AA_FWHM": float,
    "key_AA_Tails": float,
    "key_AA_DeltaLambda": float,
    "key_BB_X": float,
    "key_BB_Y": float,
    "key_BB_Lambda": float,
    "key_BB_Intensity": float,
    "key_BB_FWHM": float,
    "key_BB_Tails": float,
    "key_BB_DeltaLambda": float,
    "key_CC_X": float,
    "key_CC_Y": float,
    "key_CC_Lambda": float,
    "key_CC_Intensity": float,
    "key_CC_FWHM": float,
    "key_CC_Tails": float,
    "key_CC_DeltaLambda": float,
    "key_DD_X": float,
    "key_DD_Y": float,
    "key_DD_Lambda": float,
    "key_DD_Intensity": float,
    "key_DD_FWHM": float,
    "key_DD_Tails": float,
    "key_DD_DeltaLambda": float,
    "key_Measurements_Addr32GravityLambda": float,
    "key_Measurements_AreaStandards_1": float,
    "key_Measurements_AreaStandards_2": float,
    'key_Cmb_Date_Time': 'datetime',
    'key_Cmb_Spectral_Range': str,
    'key_Cmb_Peak_Position': float,
    'key_Cmb_Peak_Intensity': float,
    'key_Cmb_FWHM(nm)': float,
    'key_Cmb_FWHM(meV)': float,
    'key_Cmb_SideBandHight': float,
    'key_Cmb_X': float,
    'key_Cmb_Y': float,
    "key_STARTTIME_SORTED": float,
    "key_SORTNUMBER" : float,
    "key_LotNumber_9": str
}


########## 対象フォルダの全ファイルを格納 ##########
ALL_FILES = []


########## 取得しないフォルダ名 ##########
NotUsedDir = {'F1_ref', 'old', 'TAK-Ru trial', 'TEG', 'trial2', 'trial3', '多層check'}


########## 対象ロット番号のイニシャルを書込したファイルを取得する ##########
Log.Log_Info(Log_file, 'Get SerialNumber Initial List ')
try:
#    with open('T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/004_T2-EML/SerialNumber_Initial.txt', 'r', encoding='utf-8') as textfile:
    with open('../../SerialNumber_Initial.txt', 'r', encoding='utf-8') as textfile:
        SerialNumber_list = {s.strip() for s in textfile.readlines()}
except:
    with open('C:/Users/hsi67063/Downloads/SerialNumber_Initial.txt', 'r', encoding='utf-8') as textfile:    
        SerialNumber_list = {s.strip() for s in textfile.readlines()}


########## 空欄チェック ##########
def Get_Cells_Info(file_path):

    # ----- ログ書込：空欄判定 -----
    Log.Log_Info(Log_file, "Blank Check")

    # ----- ファイルとシートを指定し開く -----
    wb = px.load_workbook(file_path, read_only=True, data_only=True)
    sheet = wb[Data_sheet_name]

    # ----- False -> 空欄がない -----
    is_cells_empty = False

    # ----- 空欄はNone表示となる -----
    if sheet['H3'].value is None or sheet['H4'].value is None:
        is_cells_empty = True

    wb.close()

    return is_cells_empty


########## データを取得 ##########
def Open_Data_Sheet(file_path):

    # ----- ログ書込：データ取得 -----
    Log.Log_Info(Log_file, 'Data Acquisition')

    # ----- 辞書の初期化 -----
    data_dict = dict()

    # ----- Excelファイルとシートの定義 -----
    wb = px.load_workbook(file_path, read_only=True, data_only=True)
    sheet = wb[Data_sheet_name]
    date_sheet = wb['データファイル']

    # ----- ロット番号を取得し、Primeにロット番号に対応する品種があるか調べる -----
    serial_number = sheet['H4'].value

    # ----- ロット番号が数値の時があるので、数値の時はエラーを出す -----
    if type(serial_number) is not str:
        return None, None

    # ----- ログ書込：処理対象ロット番号 -----
    Log.Log_Info(Log_file, serial_number)

    # ----- Prime接続 -----
    conn, cursor = SQL.connSQL()
    if conn is None:
        Log.Log_Error(Log_file, serial_number + ' : ' + 'Connection with Prime Failed')
        sys.exit()

    # ----- 品名を取得 -----
    part_number, Nine_Serial_Number = SQL.selectSQL(cursor, serial_number)
    SQL.disconnSQL(conn, cursor)

    # ----- [データファイル]シートのA5から日付を抜き出す -----
    start_date = str(date_sheet['A5'].value).replace(' ', 'T')

    # ----- PLmapper装置Noの判定 -----
    PLmapper = '1'
    if "#2" in str(sheet['J4'].value):
        PLmapper = '2'
    if "#3" in str(sheet['J4'].value):
        PLmapper = '3'  
    if "#4" in str(sheet['J4'].value):
        PLmapper = '4'    

    # ----- データの取得 -----
    data_dict = {
        "key_start_date_time": start_date,
        "key_part_number": part_number,
        "key_serial_number": serial_number,
        "key_operator": sheet['F4'].value,
        "key_batch_number": sheet["H3"].value,
        "key_LotNumber_9": Nine_Serial_Number,
        "key_Mapper_3sigma_Mapper_PL_Lambda": sheet['O3'].value,
        "key_Mapper_3sigma_Mapper_PL_Intensity": sheet['O4'].value,
        "key_Mapper_3sigma_Mapper_PL_FWHM": sheet['O5'].value,
        "key_Mapper_Average_Mapper_PL_Lambda": sheet['N3'].value,
        "key_Mapper_Average_Mapper_PL_Intensity": sheet['N4'].value,
        "key_Mapper_Average_Mapper_PL_FWHM": sheet['N5'].value,
        "key_Mapper_Average_Mapper_PL_Tails": sheet['M27'].value, #Add New Item 2025/1/8 M.Yoshida
        "key_Mapper_Average_Lambda_diff": sheet['N3'].value-sheet['E27'].value,  # Add New Item 2025/2/7 M.Yoshida
#        "key_Mapper_Average_Mapper_PL_StrengthRatio": sheet['N7'].value,
        "key_Mapper_Average_Mapper_PL_StrengthRatio": "",
        "key_Mapper_Adjust_Mapper_TargetWavelength": sheet['E27'].value,
        "key_Mapper_Adjust_Mapper_Wavelength(1.3um)": sheet['E28'].value,
        "key_Mapper_Adjust_Mapper_CheckingAdjustValue": sheet['E29'].value,
        "key_Center_X": sheet['D9'].value,
        "key_Center_Y": sheet['E9'].value,
        "key_Center_Lambda": sheet['F9'].value,
        "key_Center_Intensity": sheet['G9'].value,
        "key_Center_FWHM": sheet['H9'].value,
        "key_Center_Tails": sheet['I9'].value,
        "key_Center_DeltaLambda": sheet['J9'].value,
        "key_A_X": sheet['D10'].value,
        "key_A_Y": sheet['E10'].value,
        "key_A_Lambda": sheet['F10'].value,
        "key_A_Intensity": sheet['G10'].value,
        "key_A_FWHM": sheet['H10'].value,
        "key_A_Tails": sheet['I10'].value,
        "key_A_DeltaLambda": sheet['J10'].value,
        "key_B_X": sheet['D11'].value,
        "key_B_Y": sheet['E11'].value,
        "key_B_Lambda": sheet['F11'].value,
        "key_B_Intensity": sheet['G11'].value,
        "key_B_FWHM": sheet['H11'].value,
        "key_B_Tails": sheet['I11'].value,
        "key_B_DeltaLambda": sheet['J11'].value,
        "key_C_X": sheet['D12'].value,
        "key_C_Y": sheet['E12'].value,
        "key_C_Lambda": sheet['F12'].value,
        "key_C_Intensity": sheet['G12'].value,
        "key_C_FWHM": sheet['H12'].value,
        "key_C_Tails": sheet['I12'].value,
        "key_C_DeltaLambda": sheet['J12'].value,
        "key_D_X": sheet['D13'].value,
        "key_D_Y": sheet['E13'].value,
        "key_D_Lambda": sheet['F13'].value,
        "key_D_Intensity": sheet['G13'].value,
        "key_D_FWHM": sheet['H13'].value,
        "key_D_Tails": sheet['I13'].value,
        "key_D_DeltaLambda": sheet['J13'].value,
        "key_AB_X": sheet['D14'].value,
        "key_AB_Y": sheet['E14'].value,
        "key_AB_Lambda": sheet['F14'].value,
        "key_AB_Intensity": sheet['G14'].value,
        "key_AB_FWHM": sheet['H14'].value,
        "key_AB_Tails": sheet['I14'].value,
        "key_AB_DeltaLambda": sheet['J14'].value,
        "key_AC_X": sheet['D15'].value,
        "key_AC_Y": sheet['E15'].value,
        "key_AC_Lambda": sheet['F15'].value,
        "key_AC_Intensity": sheet['G15'].value,
        "key_AC_FWHM": sheet['H15'].value,
        "key_AC_Tails": sheet['I15'].value,
        "key_AC_DeltaLambda": sheet['J15'].value,
        "key_BC_X": sheet['D16'].value,
        "key_BC_Y": sheet['E16'].value,
        "key_BC_Lambda": sheet['F16'].value,
        "key_BC_Intensity": sheet['G16'].value,
        "key_BC_FWHM": sheet['H16'].value,
        "key_BC_Tails": sheet['I16'].value,
        "key_BC_DeltaLambda": sheet['J16'].value,
        "key_CD_X": sheet['D17'].value,
        "key_CD_Y": sheet['E17'].value,
        "key_CD_Lambda": sheet['F17'].value,
        "key_CD_Intensity": sheet['G17'].value,
        "key_CD_FWHM": sheet['H17'].value,
        "key_CD_Tails": sheet['I17'].value,
        "key_CD_DeltaLambda": sheet['J17'].value,
        "key_AA_X": sheet['D18'].value,
        "key_AA_Y": sheet['E18'].value,
        "key_AA_Lambda": sheet['F18'].value,
        "key_AA_Intensity": sheet['G18'].value,
        "key_AA_FWHM": sheet['H18'].value,
        "key_AA_Tails": sheet['I18'].value,
        "key_AA_DeltaLambda": sheet['J18'].value,
        "key_BB_X": sheet['D19'].value,
        "key_BB_Y": sheet['E19'].value,
        "key_BB_Lambda": sheet['F19'].value,
        "key_BB_Intensity": sheet['G19'].value,
        "key_BB_FWHM": sheet['H19'].value,
        "key_BB_Tails": sheet['I19'].value,
        "key_BB_DeltaLambda": sheet['J19'].value,
        "key_CC_X": sheet['D20'].value,
        "key_CC_Y": sheet['E20'].value,
        "key_CC_Lambda": sheet['F20'].value,
        "key_CC_Intensity": sheet['G20'].value,
        "key_CC_FWHM": sheet['H20'].value,
        "key_CC_Tails": sheet['I20'].value,
        "key_CC_DeltaLambda": sheet['J20'].value,
        "key_DD_X": sheet['D21'].value,
        "key_DD_Y": sheet['E21'].value,
        "key_DD_Lambda": sheet['F21'].value,
        "key_DD_Intensity": sheet['G21'].value,
        "key_DD_FWHM": sheet['H21'].value,
        "key_DD_Tails": sheet['I21'].value,
        "key_DD_DeltaLambda": sheet['J21'].value,
        "key_Measurements_Addr32GravityLambda": sheet['L19'].value,
        "key_Measurements_AreaStandards_1": sheet['L20'].value,
        "key_Measurements_AreaStandards_2": sheet['L21'].value,
        "key_Equipment_PLmapper": PLmapper
    }
    #Temporary variable (ex:"F9" means cell F9 value)    2025-03-11 yoshida
    F9=sheet['F9'].value
    F14=sheet['F14'].value
    F15=sheet['F15'].value
    F16=sheet['F16'].value
    F17=sheet['F17'].value

    wb.close()

    # ----- 空欄箇所はNoneとして取得される。Noneは文字列に変換できないため、空欄("")に置き換える -----
    for keys in data_dict:
        if data_dict[keys] is None or data_dict[keys] == "#DIV/0!" or data_dict[keys] == '-':
            data_dict[keys] = ""
        # ----- 指数表記を展開する -----
        if type(data_dict[keys]) is float and 'e' in str(data_dict[keys]) and keys != "key_start_date_time":
            data_dict[keys] = ExpandExp.Expand(data_dict[keys])

    # ----- データファイルシートからデータを取得する -----
    Data_file_sheet_array = []
    df = pd.read_excel(file_path, header=None, sheet_name='データファイル', usecols="A:I", skiprows=1)
    row, col = df.shape
    for i in range(row):
        tmp = []
        for j in range(col):
            tmp.append(str(df.iloc[i, j]))
        Data_file_sheet_array.append(tmp)

    # ----- 1行目のデータだけ、データ型確認用として辞書に追加する -----
    data_dict['key_Cmb_Date_Time'] = Data_file_sheet_array[0][0]
    data_dict['key_Cmb_Spectral_Range'] = Data_file_sheet_array[0][1]
    data_dict['key_Cmb_Peak_Position'] = Data_file_sheet_array[0][2]
    data_dict['key_Cmb_Peak_Intensity'] = Data_file_sheet_array[0][3]
    data_dict['key_Cmb_FWHM(nm)'] = Data_file_sheet_array[0][4]
    data_dict['key_Cmb_FWHM(meV)'] = Data_file_sheet_array[0][5]
    data_dict['key_Cmb_SideBandHight'] = Data_file_sheet_array[0][6]
    data_dict['key_Cmb_X'] = Data_file_sheet_array[0][7]
    data_dict['key_Cmb_Y'] = Data_file_sheet_array[0][8]
    # Remove microseconds from 'key_Cmb_Date_Time' if present 2025/02/19 add this function to match format requirement
    if '.' in data_dict['key_Cmb_Date_Time']:
        data_dict['key_Cmb_Date_Time'] = data_dict['key_Cmb_Date_Time'].split('.')[0]

    data_dict['key_Lambda_Diff']= data_dict["key_Mapper_Average_Mapper_PL_Lambda"]-data_dict["key_Mapper_Adjust_Mapper_TargetWavelength"]
    data_dict['key_Lambda2_Diff'] = max(abs(F14-F9),abs(F15-F9),abs(F16-F9),abs(F17-F9)) #2025-03-11 yoshida
    
    return data_dict, Data_file_sheet_array


########## XML変換 ##########
def Output_XML(xml_file, data_dict, Data_file_sheet_array):

    # ----- ログ書込：XML変換 -----
    Log.Log_Info(Log_file, 'Excel File To XML File Conversion')
    XML = '<?xml version="1.0" encoding="utf-8"?>' + '\n' + \
        '<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' + '\n' + \
        '       <Result startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Result="Passed">' + '\n' + \
        '               <Header SerialNumber=' + '"' + data_dict["key_serial_number"] + '"' + ' PartNumber=' + '"' + data_dict["key_part_number"] + '"' + ' Operation=' + '"' + Operation + '"' + ' TestStation=' + '"' + TestStation + '"' + ' Operator=' + '"' + data_dict["key_operator"] + '"' + ' StartTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Site=' + '"' + Site + '"' + ' BatchNumber=' + '"' + data_dict["key_batch_number"] + '"' + ' LotNumber=' + '"' + data_dict["key_serial_number"] + '"/>' + '\n' + \
        '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep1"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + X + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + Y + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep2"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="Mapper_PL_Lambda" Units="nm" Value=' + '"' + str(data_dict["key_Mapper_3sigma_Mapper_PL_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Mapper_PL_Intensity" Units="mV" Value=' + '"' + str(data_dict["key_Mapper_3sigma_Mapper_PL_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Mapper_PL_FWHM" Units="meV" Value=' + '"' + str(data_dict["key_Mapper_3sigma_Mapper_PL_FWHM"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep3"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="Mapper_PL_Lambda" Units="nm" Value=' + '"' + str(data_dict["key_Mapper_Average_Mapper_PL_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Mapper_PL_Intensity" Units="mV" Value=' + '"' + str(data_dict["key_Mapper_Average_Mapper_PL_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Mapper_PL_FWHM" Units="meV" Value=' + '"' + str(data_dict["key_Mapper_Average_Mapper_PL_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Mapper_PL_Tails" Units="AU" Value=' + '"' + str(data_dict["key_Mapper_Average_Mapper_PL_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda_diff" Units="AU" Value=' + '"' + str(data_dict["key_Mapper_Average_Lambda_diff"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Mapper_PL_StrengthRatio" Units="degree" Value=' + '"' + str(data_dict["key_Mapper_Average_Mapper_PL_StrengthRatio"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep4"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="Mapper_TargetWavelength" Units="nm" Value=' + '"' + str(data_dict["key_Mapper_Adjust_Mapper_TargetWavelength"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Mapper_Wavelength(1.3um)" Units="nm" Value=' + '"' + str(data_dict["key_Mapper_Adjust_Mapper_Wavelength(1.3um)"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Mapper_CheckingAdjustValue" Units="nm" Value=' + '"' + str(data_dict["key_Mapper_Adjust_Mapper_CheckingAdjustValue"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep5"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_Center_X"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_Center_Y"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda" Units="nm" Value=' + '"' + str(data_dict["key_Center_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda_Diff" Units="nm" Value=' + '"' + str(data_dict["key_Lambda_Diff"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda2_Diff" Units="nm" Value=' + '"' + str(data_dict["key_Lambda2_Diff"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Intensity" Units="counts" Value=' + '"' + str(data_dict["key_Center_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="FWHM" Units="meV" Value=' + '"' + str(data_dict["key_Center_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Tails" Units="AU" Value=' + '"' + str(data_dict["key_Center_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="DeltaLambda" Units="nm" Value=' + '"' + str(data_dict["key_Center_DeltaLambda"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep6"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_A_X"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_A_Y"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda" Units="nm" Value=' + '"' + str(data_dict["key_A_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Intensity" Units="counts" Value=' + '"' + str(data_dict["key_A_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="FWHM" Units="meV" Value=' + '"' + str(data_dict["key_A_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Tails" Units="AU" Value=' + '"' + str(data_dict["key_A_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="DeltaLambda" Units="nm" Value=' + '"' + str(data_dict["key_A_DeltaLambda"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep7"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_B_X"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_B_Y"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda" Units="nm" Value=' + '"' + str(data_dict["key_B_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Intensity" Units="counts" Value=' + '"' + str(data_dict["key_B_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="FWHM" Units="meV" Value=' + '"' + str(data_dict["key_B_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Tails" Units="AU" Value=' + '"' + str(data_dict["key_B_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="DeltaLambda" Units="nm" Value=' + '"' + str(data_dict["key_B_DeltaLambda"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep8"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_C_X"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_C_Y"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda" Units="nm" Value=' + '"' + str(data_dict["key_C_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Intensity" Units="counts" Value=' + '"' + str(data_dict["key_C_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="FWHM" Units="meV" Value=' + '"' + str(data_dict["key_C_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Tails" Units="AU" Value=' + '"' + str(data_dict["key_C_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="DeltaLambda" Units="nm" Value=' + '"' + str(data_dict["key_C_DeltaLambda"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep9"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_D_X"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_D_Y"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda" Units="nm" Value=' + '"' + str(data_dict["key_D_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Intensity" Units="counts" Value=' + '"' + str(data_dict["key_D_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="FWHM" Units="meV" Value=' + '"' + str(data_dict["key_D_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Tails" Units="AU" Value=' + '"' + str(data_dict["key_D_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="DeltaLambda" Units="nm" Value=' + '"' + str(data_dict["key_D_DeltaLambda"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep10"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_AB_X"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_AB_Y"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda" Units="nm" Value=' + '"' + str(data_dict["key_AB_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Intensity" Units="counts" Value=' + '"' + str(data_dict["key_AB_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="FWHM" Units="meV" Value=' + '"' + str(data_dict["key_AB_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Tails" Units="AU" Value=' + '"' + str(data_dict["key_AB_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="DeltaLambda" Units="nm" Value=' + '"' + str(data_dict["key_AB_DeltaLambda"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep11"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_AC_X"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_AC_Y"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda" Units="nm" Value=' + '"' + str(data_dict["key_AC_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Intensity" Units="counts" Value=' + '"' + str(data_dict["key_AC_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="FWHM" Units="meV" Value=' + '"' + str(data_dict["key_AC_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Tails" Units="AU" Value=' + '"' + str(data_dict["key_AC_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="DeltaLambda" Units="nm" Value=' + '"' + str(data_dict["key_AC_DeltaLambda"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep12"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_BC_X"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_BC_Y"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda" Units="nm" Value=' + '"' + str(data_dict["key_BC_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Intensity" Units="counts" Value=' + '"' + str(data_dict["key_BC_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="FWHM" Units="meV" Value=' + '"' + str(data_dict["key_BC_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Tails" Units="AU" Value=' + '"' + str(data_dict["key_BC_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="DeltaLambda" Units="nm" Value=' + '"' + str(data_dict["key_BC_DeltaLambda"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep13"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_CD_X"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_CD_Y"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda" Units="nm" Value=' + '"' + str(data_dict["key_CD_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Intensity" Units="counts" Value=' + '"' + str(data_dict["key_CD_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="FWHM" Units="meV" Value=' + '"' + str(data_dict["key_CD_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Tails" Units="AU" Value=' + '"' + str(data_dict["key_CD_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="DeltaLambda" Units="nm" Value=' + '"' + str(data_dict["key_CD_DeltaLambda"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep14"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_AA_X"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_AA_Y"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda" Units="nm" Value=' + '"' + str(data_dict["key_AA_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Intensity" Units="counts" Value=' + '"' + str(data_dict["key_AA_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="FWHM" Units="meV" Value=' + '"' + str(data_dict["key_AA_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Tails" Units="AU" Value=' + '"' + str(data_dict["key_AA_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="DeltaLambda" Units="nm" Value=' + '"' + str(data_dict["key_AA_DeltaLambda"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep15"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_BB_X"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_BB_Y"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda" Units="nm" Value=' + '"' + str(data_dict["key_BB_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Intensity" Units="counts" Value=' + '"' + str(data_dict["key_BB_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="FWHM" Units="meV" Value=' + '"' + str(data_dict["key_BB_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Tails" Units="AU" Value=' + '"' + str(data_dict["key_BB_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="DeltaLambda" Units="nm" Value=' + '"' + str(data_dict["key_BB_DeltaLambda"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep16"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_CC_X"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_CC_Y"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda" Units="nm" Value=' + '"' + str(data_dict["key_CC_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Intensity" Units="counts" Value=' + '"' + str(data_dict["key_CC_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="FWHM" Units="meV" Value=' + '"' + str(data_dict["key_CC_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Tails" Units="AU" Value=' + '"' + str(data_dict["key_CC_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="DeltaLambda" Units="nm" Value=' + '"' + str(data_dict["key_CC_DeltaLambda"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep17"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="X" Units="um" Value=' + '"' + str(data_dict["key_DD_X"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Y" Units="um" Value=' + '"' + str(data_dict["key_DD_Y"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Lambda" Units="nm" Value=' + '"' + str(data_dict["key_DD_Lambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Intensity" Units="counts" Value=' + '"' + str(data_dict["key_DD_Intensity"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="FWHM" Units="meV" Value=' + '"' + str(data_dict["key_DD_FWHM"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="Tails" Units="AU" Value=' + '"' + str(data_dict["key_DD_Tails"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="DeltaLambda" Units="nm" Value=' + '"' + str(data_dict["key_DD_DeltaLambda"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep18"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="Addr32GravityLambda" Units="um" Value=' + '"' + str(data_dict["key_Measurements_Addr32GravityLambda"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="AreaStandards_1" Units="percent" Value=' + '"' + str(data_dict["key_Measurements_AreaStandards_1"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="AreaStandards_2" Units="percent" Value=' + '"' + str(data_dict["key_Measurements_AreaStandards_2"]) + '"/>' + '\n' + \
        '               </TestStep>' + '\n' + \
        '               <TestStep Name=' + '"' + teststep_dict["TestStep19"] + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
        '                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value=' + '"' + str(data_dict["key_STARTTIME_SORTED"]) + '"/>' + '\n' + \
        '                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value=' + '"' + str(data_dict["key_SORTNUMBER"]) + '"/>' + '\n' + \
        '                   <Data DataType="String" Name="BATCHNUMBER_SORTED" Value=' + '"' + str(data_dict["key_batch_number"]) + '"' + ' CompOperation="LOG"/>' + '\n' + \
        '                   <Data DataType="String" Name="LotNumber_5" Value=' + '"' + str(data_dict["key_serial_number"]) + '"' + ' CompOperation="LOG"/>' + '\n' + \
        '                   <Data DataType="String" Name="LotNumber_9" Value=' + '"' + str(data_dict["key_LotNumber_9"]) + '"' + ' CompOperation="LOG"/>' + '\n' + \
        '               </TestStep>' + '\n'

    # Cmbの付与
    for i in range(len(Data_file_sheet_array)):
        TestStep = 'Cmb' + str(i + 1)
        XML += '               <TestStep Name=' + '"' + TestStep + '"' + ' startDateTime=' + '"' + data_dict["key_start_date_time"] + '"' + ' Status="Passed">' + '\n' + \
               '                   <Data DataType="String" Name="Cmb_Date_Time" Value=' + '"' + str(Data_file_sheet_array[i][0]) + '"' + ' CompOperation="LOG"/>' + '\n' + \
               '                   <Data DataType="String" Name="Cmb_Spectral_Range" Value=' + '"' + str(Data_file_sheet_array[i][1]) + '"' + ' CompOperation="LOG"/>' + '\n' + \
               '                   <Data DataType="Numeric" Name="Cmb_Peak_Position" Units="nm" Value=' + '"' + str(Data_file_sheet_array[i][2]) + '"/>' + '\n' + \
               '                   <Data DataType="Numeric" Name="Cmb_Peak_Intensity" Units="counts" Value=' + '"' + str(Data_file_sheet_array[i][3]) + '"/>' + '\n' + \
               '                   <Data DataType="Numeric" Name="Cmb_FWHM(nm)" Units="nm" Value=' + '"' + str(Data_file_sheet_array[i][4]) + '"/>' + '\n' + \
               '                   <Data DataType="Numeric" Name="Cmb_FWHM(meV)" Units="meV" Value=' + '"' + str(Data_file_sheet_array[i][5]) + '"/>' + '\n' + \
               '                   <Data DataType="Numeric" Name="Cmb_SideBandHight" Units="counts" Value=' + '"' + str(Data_file_sheet_array[i][6]) + '"/>' + '\n' + \
               '                   <Data DataType="Numeric" Name="Cmb_X" Units="um" Value=' + '"' + str(Data_file_sheet_array[i][7]) + '"/>' + '\n' + \
               '                   <Data DataType="Numeric" Name="Cmb_Y" Units="um" Value=' + '"' + str(Data_file_sheet_array[i][8]) + '"/>' + '\n' + \
               '               </TestStep>' + '\n'
    
    XML +=  '\n' \
            '               <TestEquipment>' + '\n' + \
            '                   <Item DeviceName="PLmapper" DeviceSerialNumber="' + str(data_dict["key_Equipment_PLmapper"]) + '"/>' + '\n' + \
            '                   <Item DeviceName="MOCVD" DeviceSerialNumber="' + "7" + '"/>' + '\n' + \
            '               </TestEquipment>' + '\n' + \
            '\n' \
            '               <ErrorData/>' + '\n' + \
            '               <FailureData/>' + '\n' + \
            '               <Configuration/>' + '\n' + \
            '       </Result>' + '\n' + \
            '</Results>'

    f = open(Output_filepath + xml_file, 'w', encoding="utf-8")
            
    f.write(XML)

    f.close()


########## 指定パス内のすべてのファイルを取得 ##########
def ALL_FILE_FETCH(Path):
    global ALL_FILES
    for path, dir, file in os.walk(Path):
        dir[:] = [d for d in dir if d not in NotUsedDir]
        for f in file:
            if f.endswith('.xlsx') or f.endswith('.xlsm') or f.endswith('.xls'):
                ALL_FILES.append(os.path.join(path, f))

########## main処理 ##########
if __name__ == '__main__':

    # ----- 対象フォルダ内の拡張子が[.xl*]のものをすべて取り出す -----
    Path = [
        "//sagqnap01.li.lumentuminc.net/K11Data/PL-MAP#4/_PLマッパー判定/HL13B5/F7LD/",
        "//sagqnap01.li.lumentuminc.net/K11Data/PL-MAP/★PLﾏｯﾊﾟｰ判定/3ｲﾝﾁHTL13B4/F7 LD/",
        "//sagqnap01.li.lumentuminc.net/K11Data/PL-MAP/★PLﾏｯﾊﾟｰ判定/3ｲﾝﾁHTL13B5/F7 LD/",
        "//sagqnap01.li.lumentuminc.net/K11Data/PL-MAP#2/★PLﾏｯﾊﾟｰ判定/HL13B4/F7LD/",
        "//sagqnap01.li.lumentuminc.net/K11Data/PL-MAP#2/★PLﾏｯﾊﾟｰ判定/HL13B5/F7LD/",
        "//sagqnap01.li.lumentuminc.net/K11Data/PL-MAP#3/★PLﾏｯﾊﾟｰ判定/HL13B5/F7LD/ﾛｯﾄ採番済",
        "//sagqnap01.li.lumentuminc.net/K11Data/PL-MAP#3/★PLﾏｯﾊﾟｰ判定/HL13B5/F7LD/LD合格品",
        "//sagqnap01.li.lumentuminc.net/K11Data/PL-MAP#3/★PLﾏｯﾊﾟｰ判定/4インチ_HL13B5/F7LD/LD合格品",
    ]

    for p in Path:
        ALL_FILE_FETCH(p)
  

    # ----- 処理を行ったファイル名をテキストファイルから呼び出す -----
    with open('EndsFile_F7_Format2.txt', 'r', encoding='utf-8') as textfile:
        EndFiles = set(s.strip() for s in textfile.readlines())

    # ----- 取得した全ファイルの処理 -----
    for FilePath in ALL_FILES:
        
        # ----- ファイルパスからファイル名を取り出し定義 -----
        file_mod_time = datetime.fromtimestamp(os.path.getmtime(FilePath))
        print(FilePath,file_mod_time,datetime.now() - file_mod_time )
        if (datetime.now() - file_mod_time).days > 30:
            Log.Log_Info(Log_file, f"File is older than 30 days, skipping.")
            continue
        File = os.path.basename(FilePath)
        print(File, "read ready")
        # ----- 処理を行ったファイルであれば、処理を行わず次のファイルへ -----
        if File in EndFiles or '~$' in File:
            print(EndFiles,"temp file or uploaded file skip")
            continue

        Log.Log_Info(Log_file, FilePath)

        # ---- 空欄チェック -----
        if Get_Cells_Info(FilePath):
            Log.Log_Error(Log_file, "Blank Error\n")
            continue

        # ----- ここを通ったら処理を行ったとみなす -----
        EndFiles.add(File)

        # ----- データの取得 -----
        data_dict, Data_file_sheet_array = Open_Data_Sheet(FilePath)

        # ----- ロット番号エラー -----
        if data_dict is None:
            Log.Log_Error(Log_file, "Lot No Error\n")
            continue

        # ----- Primeにロット番号に対応する品名が入ってなければエラー処理を行う -----
        if data_dict['key_part_number'] == "":
            Log.Log_Error(Log_file, data_dict["key_serial_number"] + ' : ' + "Part Number Error\n")
            continue

        # ----- ログ書込：日付フォーマットの変換 -----
        Log.Log_Info(Log_file, 'Date Format Conversion')

        # ----- 日付フォーマットの変換を行い、辞書型に上書きする -----
        if len(data_dict['key_start_date_time']) != 19 or data_dict['key_start_date_time'][10] != 'T' or data_dict['key_start_date_time'][4] != '-':
            data_dict['key_start_date_time'] = Convert_Date.Edit_Date(data_dict['key_start_date_time'])
            #change the date/Time format for . -> : to mapping IEEE XML format requirement 2025/02/19 New add pigo
            data_dict['key_start_date_time'] = data_dict['key_start_date_time'].replace('.', ':')

        # ----- 日付変換が失敗した(空欄で返ったきた)とき、エラー処理を行う -----
        if data_dict['key_start_date_time'] == "":
            Log.Log_Error(Log_file, data_dict["key_serial_number"] + ' : ' + "Date Error\n")
            continue

        # ----- 作業者が空だった場合、'-'とする -----
        if data_dict['key_operator'] == "":
            data_dict["key_operator"] = "-"

        # ----- STARTTIME_SORTEDの追加 -----

        # 日付をExcel時間に変換する
        date = datetime.strptime(str(data_dict["key_start_date_time"]).replace('T', ' ').replace('.', ':'), "%Y-%m-%d %H:%M:%S")
        date_excel_number = int(str(date - datetime(1899, 12, 30)).split()[0])
        #print("date:",date)
        # エピ番号の数値部だけを取得する
        Epi_Number = 0
        for i in data_dict["key_batch_number"]:
            try:
                if 0 <= int(i) <= 9:
                    Epi_Number = Epi_Number * 10 + int(i)
            except:
                pass

        # エピ番号を10^6で割って、excel時間に加算する
        date_excel_number += Epi_Number/10**6

        # data_dictに登録する
        data_dict["key_STARTTIME_SORTED"] = date_excel_number
        data_dict["key_SORTNUMBER"] = Epi_Number

        # ---- データ型の確認 -----
        result = Check.Data_Type(key_type, data_dict)
        if result == False:
            Log.Log_Error(Log_file, data_dict["key_serial_number"] + ' : ' + "Data Error\n")
            continue

        # ----- XMLファイルの作成 -----
        xml_file = 'Site=' + Site + ',ProductFamily=' + ProductFamily + ',Operation=' + Operation + \
                   ',Partnumber=' + data_dict["key_part_number"] + ',Serialnumber=' + data_dict["key_serial_number"] + \
                   ',Testdate=' + data_dict["key_start_date_time"].replace(':', '.') + '.xml'
        
        try:
            Output_XML(xml_file, data_dict, Data_file_sheet_array)
            print(xml_file, "is created", data_dict["key_serial_number"])
        except Exception as e:
            Log.Log_Error(Log_file, f"Failed to create XML for {data_dict['key_serial_number']}: {e}")
            continue


    # ----- 処理が完了したファイルをテキストファイルに書き込む -----
    EndFiles_list = sorted(list(EndFiles))
    EndFiles_str = "\n".join(EndFiles_list)
    with open('EndsFile_F7_Format2.txt', 'w', encoding='utf-8') as textfile:
        textfile.write(EndFiles_str)

    # ----- 最終行を書き込んだファイルをGドライブに転送 -----
    #shutil.copy("EndsFile_F7_Format2.txt", 'T:/04_プロセス関係/10_共通/91_KAIZEN-TDS/01_開発/040_LD-EML/F7/13_ProgramUsedFile/')


########## ログ書込：プログラムの終了 ##########
Log.Log_Info(Log_file, 'Program End')