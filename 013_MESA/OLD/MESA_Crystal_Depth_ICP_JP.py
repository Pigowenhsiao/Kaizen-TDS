import os
import sys
import glob
import shutil
import logging
import numpy as np
import pandas as pd
from datetime import datetime, date, timedelta
from configparser import ConfigParser
from pathlib import Path
import traceback

# --- バージョン検証 ---
print("--- Running v6 (with date filtering feature) ---")
# -----------------

# --- カスタムモジュールのインポート ---
sys.path.append('../MyModule')
import Log
import SQL
import Convert_Date
import Row_Number_Func

# --- INIファイルの設定を格納するためのクラス ---
class IniSettings:
    def __init__(self):
        self.site = ""
        self.product_family = ""
        self.operation = ""
        self.test_station = ""
        self.file_name_patterns = []
        self.input_paths = []
        self.output_path = ""
        self.intermediate_data_path = ""
        self.log_path = ""
        self.running_rec = ""
        self.sheet_name = ""
        self.data_columns = ""
        self.xy_sheet_name = ""
        self.xy_columns = ""
        self.skip_rows = 500
        self.key_col_idx = 7
        self.serial_col_idx = 4
        self.xy_x_idx = 1
        self.xy_y_idx = 2
        self.xy_points = 5
        self.tool_name_map = {}
        self.field_map = {}
        # 新機能：データの保持期間（日数）
        self.retention_date = 0


# --- ロギングを設定する関数 ---
def setup_logging(log_dir):
    log_folder = os.path.join(log_dir, str(date.today()))
    os.makedirs(log_folder, exist_ok=True)
    log_file = os.path.join(log_folder, '013_MESA_ICP.log')
    logging.basicConfig(filename=log_file, level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')
    return log_file

# --- INI設定ファイルを読み込んで解析する ---
def _read_and_parse_ini_config(config_file_path):
    config = ConfigParser(interpolation=None)
    config.read(config_file_path, encoding='utf-8')
    return config

# --- データフィールドのマッピング設定を解析する ---
def _parse_fields_map_from_lines(fields_lines):
    fields = {}
    for line in fields_lines:
        if ':' in line and not line.strip().startswith('#'):
            try:
                key, col_str, dtype_str = map(str.strip, line.split(':', 2))
                
                if dtype_str == 'float':
                    dtype = float
                elif dtype_str == 'str':
                    dtype = str
                else: 
                    dtype = dtype_str
                
                fields[key] = {'col': col_str, 'type': dtype}
            except ValueError:
                continue
    return fields

# --- 解析された設定オブジェクトから各設定値を抽出し、IniSettingsインスタンスに格納する ---
def _extract_settings_from_config(config):
    s = IniSettings()
    s.site = config.get('Basic_info', 'Site')
    # 新機能：retention_date設定の読み込み
    s.retention_date = config.getint('Basic_info', 'retantion_date')
    s.product_family = config.get('Basic_info', 'ProductFamily')
    s.operation = config.get('Basic_info', 'Operation')
    s.test_station = config.get('Basic_info', 'TestStation')
    s.file_name_patterns = [x.strip() for x in config.get('Basic_info', 'file_name_patterns').split(',')]
    s.input_paths = [x.strip() for x in config.get('Paths', 'input_paths').split(',')]
    s.output_path = config.get('Paths', 'output_path')
    s.intermediate_data_path = config.get('Paths', 'intermediate_data_path')
    s.log_path = config.get('Paths', 'log_path')
    s.running_rec = config.get('Paths', 'running_rec')
    s.sheet_name = config.get('Excel', 'sheet_name')
    s.data_columns = config.get('Excel', 'data_columns')
    s.xy_sheet_name = config.get('Excel', 'xy_sheet_name')
    s.xy_columns = config.get('Excel', 'xy_columns')
    s.skip_rows = config.getint('Excel', 'main_skip_rows')
    s.key_col_idx = config.getint('Excel', 'main_dropna_key_col_idx')
    s.serial_col_idx = config.getint('Excel', 'serial_number_source_column_idx')
    s.xy_x_idx = config.getint('Excel', 'xy_coord_x_col_idx')
    s.xy_y_idx = config.getint('Excel', 'xy_coord_y_col_idx')
    s.xy_points = config.getint('Excel', 'xy_num_points')
    s.tool_name_map = dict(config.items('ToolNameMapping'))
    fields_lines = config.get('DataFields', 'fields').splitlines()
    s.field_map = _parse_fields_map_from_lines(fields_lines)
    return s

# --- ファイル名に含まれるキーワードからツール名を検出する ---
def detect_tool_name(filename, tool_map):
    filename_str = str(filename)
    for keyword, tool in tool_map.items():
        if keyword != 'default' and keyword in filename_str:
            return tool
    return tool_map.get('default', 'UNKNOWN')

# --- 辞書に格納されたデータの型が期待通りか検証する ---
def Data_Type(key_to_type, data_dict):
    for key, expected_type in key_to_type.items():
        value = data_dict.get(key)
        if isinstance(expected_type, str):
            expected_type_name = expected_type
        else:
            expected_type_name = expected_type.__name__
        if value is None or pd.isna(value):
            print(f"DATA_TYPE_ERROR: Value for key '{key}' is empty.")
            return False
        try:
            if expected_type is float:
                converted_value = float(value)
                if 'e' in str(converted_value):
                    data_dict[key] = int(converted_value)
                else:
                    data_dict[key] = converted_value
            elif expected_type == 'datetime':
                clean_value = str(value).replace('T', ' ').replace('.', ':')
                datetime.strptime(clean_value, '%Y-%m-%d %H:%M:%S')
            elif expected_type is str:
                pass
            elif not isinstance(value, expected_type):
                print(f"DATA_TYPE_ERROR: Incorrect type for key '{key}', expected {expected_type_name}, but got {type(value).__name__}")
                return False
        except (ValueError, TypeError) as e:
            print(f"DATA_TYPE_ERROR: Value '{value}' for key '{key}' cannot be processed as {expected_type_name}. Error message: {e}")
            return False
    return True

# --- 処理されたデータからXMLファイルを生成する ---
def generate_xml(output_path, site, product_family, operation, test_station, data_dict, log_file):
    start_time_str = str(data_dict.get('key_Start_Date_Time', ''))
    safe_test_date = start_time_str.replace(':', '.').replace(' ', 'T')
    filename = f"Site={site},ProductFamily={product_family},Operation={operation},Partnumber={data_dict.get('key_Part_Number', 'NA')},Serialnumber={data_dict.get('key_Serial_Number', 'NA')},Testdate={safe_test_date}.xml"
    filepath = Path(output_path) / filename

    try:
        with open(filepath, 'w', encoding="utf-8") as f:
            f.write('<?xml version="1.0" encoding="utf-8"?>\n')
            f.write('<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">\n')
            f.write(f'  <Result startDateTime="{start_time_str}" Result="Passed">\n')
            f.write(f'    <Header SerialNumber="{data_dict.get("key_Serial_Number", "")}" PartNumber="{data_dict.get("key_Part_Number", "")}" Operation="{operation}" TestStation="{test_station}" Operator="{data_dict.get("key_Operator", "NA")}" StartTime="{start_time_str}" Site="{site}" LotNumber="{data_dict.get("key_Serial_Number", "")}"/>\n')
            
            f.write(f'    <TestStep Name="Order" startDateTime="{start_time_str}" Status="Passed">\n')
            f.write(f'      <Data DataType="String" Name="Order" Units="No" Value="{data_dict.get("key_Order", "")}"/>\n')
            f.write(f'    </TestStep>\n')
            f.write(f'    <TestStep Name="Time" startDateTime="{start_time_str}" Status="Passed">\n')
            f.write(f'      <Data DataType="Numeric" Name="Time" Units="sec" Value="{data_dict.get("key_Time_Time", "")}"/>\n')
            f.write(f'    </TestStep>\n')
            f.write(f'    <TestStep Name="Depth" startDateTime="{start_time_str}" Status="Passed">\n')
            f.write(f'      <Data DataType="Numeric" Name="First1" Units="nm" Value="{data_dict.get("key_Depth_First1", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="First2" Units="nm" Value="{data_dict.get("key_Depth_First2", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="First3" Units="nm" Value="{data_dict.get("key_Depth_First3", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="First4" Units="nm" Value="{data_dict.get("key_Depth_First4", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="First5" Units="nm" Value="{data_dict.get("key_Depth_First5", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="First_Ave" Units="nm" Value="{data_dict.get("key_Depth_First_Ave", "")}"/>\n')
            f.write(f'    </TestStep>\n')
            f.write(f'    <TestStep Name="Thickness" startDateTime="{start_time_str}" Status="Passed">\n')
            f.write(f'      <Data DataType="Numeric" Name="First1" Units="nm" Value="{data_dict.get("key_Thickness_First1", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="First2" Units="nm" Value="{data_dict.get("key_Thickness_First2", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="First3" Units="nm" Value="{data_dict.get("key_Thickness_First3", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="First4" Units="nm" Value="{data_dict.get("key_Thickness_First4", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="First5" Units="nm" Value="{data_dict.get("key_Thickness_First5", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="First_Ave" Units="nm" Value="{data_dict.get("key_Thickness_First_Ave", "")}"/>\n')
            f.write(f'    </TestStep>\n')
            f.write(f'    <TestStep Name="Etching" startDateTime="{start_time_str}" Status="Passed">\n')
            f.write(f'      <Data DataType="Numeric" Name="Etching1" Units="nm" Value="{data_dict.get("key_Etching_Etching1", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Etching2" Units="nm" Value="{data_dict.get("key_Etching_Etching2", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Etching3" Units="nm" Value="{data_dict.get("key_Etching_Etching3", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Etching4" Units="nm" Value="{data_dict.get("key_Etching_Etching4", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Etching5" Units="nm" Value="{data_dict.get("key_Etching_Etching5", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Etching_Ave" Units="nm" Value="{data_dict.get("key_Etching_Etching_Ave", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Etching_Max-Min" Units="nm" Value="{data_dict.get("key_Etching_Etching_Max-Min", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Etching_3sigma" Units="nm" Value="{data_dict.get("key_Etching_Etching_3sigma", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Etching_Rate" Units="nm/min" Value="{data_dict.get("key_Etching_Etching_Rate", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Etching_Error" Units="nm" Value="{data_dict.get("key_Etching_Etching_Error", "")}"/>\n')
            f.write(f'    </TestStep>\n')
            f.write(f'    <TestStep Name="Coordinate" startDateTime="{start_time_str}" Status="Passed">\n')
            f.write(f'      <Data DataType="Numeric" Name="X1" Units="um" Value="{data_dict.get("key_X1", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="X2" Units="um" Value="{data_dict.get("key_X2", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="X3" Units="um" Value="{data_dict.get("key_X3", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="X4" Units="um" Value="{data_dict.get("key_X4", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="X5" Units="um" Value="{data_dict.get("key_X5", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Y1" Units="um" Value="{data_dict.get("key_Y1", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Y2" Units="um" Value="{data_dict.get("key_Y2", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Y3" Units="um" Value="{data_dict.get("key_Y3", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Y4" Units="um" Value="{data_dict.get("key_Y4", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Y5" Units="um" Value="{data_dict.get("key_Y5", "")}"/>\n')
            f.write(f'    </TestStep>\n')
            f.write(f'    <TestStep Name="SORTED_DATA" startDateTime="{start_time_str}" Status="Passed">\n')
            f.write(f'      <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value="{data_dict.get("key_STARTTIME_SORTED", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value="{data_dict.get("key_SORTNUMBER", "")}"/>\n')
            f.write(f'      <Data DataType="String" Name="LotNumber_5" Value="{data_dict.get("key_Serial_Number", "")}" CompOperation="LOG"/>\n')
            f.write(f'      <Data DataType="String" Name="LotNumber_9" Value="{data_dict.get("key_LotNumber_9", "")}" CompOperation="LOG"/>\n')
            f.write('    </TestStep>\n')
            f.write(f'    <TestEquipment>\n')
            f.write(f'      <Item DeviceName="DryEtch" DeviceSerialNumber="{data_dict.get("Tool_name", "")}"/>\n')
            f.write(f'    </TestEquipment>\n')
            f.write('    <ErrorData/>\n')
            f.write('    <FailureData/>\n')
            f.write('    <Configuration/>\n')
            f.write('  </Result>\n')
            f.write('</Results>')
        Log.Log_Info(log_file, f"XML file written: {filepath}")
        return True
    except Exception as e:
        Log.Log_Error(log_file, f"Error writing XML: {e}")
        return False

# --- 単一のExcelファイルを処理するメインロジック ---
def process_single_excel_file(filepath, settings, log_file):
    Log.Log_Info(log_file, f"Processing file: {filepath}")
    start_row = Row_Number_Func.start_row_number(settings.running_rec) - settings.skip_rows
    df = pd.read_excel(filepath, header=None, sheet_name=settings.sheet_name, usecols=settings.data_columns, skiprows=start_row)

    # --- 新機能：日付によるデータフィルタリング ---
    try:
        # INI設定から日付列のインデックスを取得
        date_column_index = int(settings.field_map['key_Start_Date_Time']['col'])
        
        # 日付列をdatetime型に変換（エラーはNaTとして処理）
        date_series = pd.to_datetime(df.iloc[:, date_column_index], errors='coerce')

        # カットオフ日を計算
        today = datetime.now()
        cutoff_date = today - timedelta(days=settings.retention_date)
        
        # カットオフ日より新しい、有効な日付のデータのみを保持
        df = df[date_series.notna() & (date_series >= cutoff_date)]
        Log.Log_Info(log_file, f"After date filtering, file {filepath.name} has {len(df)} valid records left.")

    except (KeyError, ValueError, IndexError) as e:
        Log.Log_Error(log_file, f"Date filtering failed. Please check the 'key_Start_Date_Time' setting in the INI file. Error: {e}")
        # 日付フィルタリングに失敗した場合、このファイルをスキップする
        return

    # キーとなる列に値が存在する行のみを保持
    df = df[df.iloc[:, settings.key_col_idx].notna()]
    
    # XY座標データを別のシートから読み込む
    df_xy = pd.read_excel(filepath, header=None, sheet_name=settings.xy_sheet_name, usecols=settings.xy_columns)
    df.columns = range(df.shape[1])
    
    if df_xy.shape[0] < settings.xy_points + 1:
        Log.Log_Error(log_file, f"Insufficient XY coordinate data, less than {settings.xy_points} points. Actual rows: {df_xy.shape[0]}. Skipping: {filepath}")
        return

    row_count = 0
    # データフレームの各行をループ処理
    for idx in df.index:
        data_dict = {}
        
        # 設定されたマッピングに基づき、各データを抽出
        for key, mapping in settings.field_map.items():
            col_str = mapping['col']
            if 'xy_' in col_str:
                try:
                    parts = col_str.split('_')
                    row_index = int(parts[1]) - 1 
                    col_index = int(parts[2]) - 1
                    data_dict[key] = df_xy.iloc[row_index, col_index]
                except (IndexError, ValueError) as e:
                    data_dict[key] = None
                    Log.Log_Error(log_file, f"Failed to read XY coordinate: Key '{key}' (setting: {col_str}). Error: {e}")
                continue
            if col_str == '-1':
                continue
            try:
                col_index = int(col_str)
                data_dict[key] = df.loc[idx].iloc[col_index]
            except (ValueError, IndexError):
                data_dict[key] = None
                Log.Log_Error(log_file, f"Failed to read column: Key '{key}' (column index: {col_str})")

        serial = data_dict.get('key_Serial_Number')
        if not serial or pd.isna(serial): continue
        serial = str(serial)

        # シリアル番号を使用してSQLデータベースに問い合わせる
        conn, cursor = SQL.connSQL()
        if conn is None: continue
        part, lot9 = SQL.selectSQL(cursor, serial)
        SQL.disconnSQL(conn, cursor)
        if part is None or part == 'LDアレイ_': continue
        
        data_dict['key_Part_Number'] = part
        data_dict['key_LotNumber_9'] = lot9
        data_dict['Tool_name'] = detect_tool_name(filepath, settings.tool_name_map)
        
        raw_date = data_dict.get('key_Start_Date_Time')
        if raw_date and not isinstance(raw_date, str):
             data_dict['key_Start_Date_Time'] = pd.to_datetime(raw_date).strftime('%Y-%m-%d %H:%M:%S')
        
        # データの型を検証
        key_to_type = {k: v['type'] for k, v in settings.field_map.items() if v['col'] != '-1'}
        if not Data_Type(key_to_type, data_dict):
            Log.Log_Error(log_file, serial + ' : ' + 'Data Error')
            continue

        try:
            # ソート用のExcel日付シリアル値を計算する
            clean_datetime_str = str(data_dict.get('key_Start_Date_Time', '')).replace('T', ' ').replace('.', ':')
            date_obj = datetime.strptime(clean_datetime_str, "%Y-%m-%d %H:%M:%S")
            date_excel_number = int(str(date_obj - datetime(1899, 12, 30)).split()[0])
            excel_row = idx 
            date_excel_number += (excel_row + 1) / 10**6 
            data_dict['key_STARTTIME_SORTED'] = date_excel_number
            data_dict['key_SORTNUMBER'] = excel_row + 1 
        except Exception as e:
            Log.Log_Error(log_file, f"Date conversion failed for {serial}: {e}")
            continue

        # XML生成関数を呼び出す
        if generate_xml(settings.output_path, settings.site, settings.product_family, settings.operation, settings.test_station, data_dict, log_file):
            Log.Log_Info(log_file, f"{serial} : OK")
            row_count += 1
        else:
            Log.Log_Error(log_file, f"{serial} : XML Generation Failed")
            continue

    # 次に処理を開始する行番号を更新する
    next_start_row = start_row + df.shape[0] + 1
    Row_Number_Func.next_start_row_number(settings.running_rec, next_start_row)

# --- スクリプトのメインエントリーポイント ---
def main():
    # 作業ディレクトリをスクリプトのある場所に変更
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # カレントディレクトリにある全ての.iniファイルを検索
    ini_files = [f for f in os.listdir('.') if f.endswith('.ini')]
    if not ini_files:
        print("No .ini files found in current directory.")
        return

    # 各iniファイルに対して処理ループを実行
    for ini_path in ini_files:
        try:
            config = _read_and_parse_ini_config(ini_path)
            settings = _extract_settings_from_config(config)
            log_file = setup_logging(settings.log_path)
            Log.Log_Info(log_file, f"Start processing INI config file: {ini_path}")
            input_paths = [Path(p) for p in settings.input_paths]
            intermediate_path = Path(settings.intermediate_data_path)
            intermediate_path.mkdir(parents=True, exist_ok=True)
            for input_p in input_paths:
                for pattern in settings.file_name_patterns:
                    matched_files = [p for p in input_p.glob(pattern) if not p.name.startswith('~$')]
                    for src_path in matched_files:
                        dst_path = intermediate_path / src_path.name
                        try:
                            # ソースファイルを中間ディレクトリにコピー
                            shutil.copy(src_path, dst_path)
                            # ファイル処理関数を呼び出す
                            process_single_excel_file(dst_path, settings, log_file)
                        # 例外処理
                        except FileNotFoundError as fnf:
                            error_message = f"File not found during copy: {src_path} -> {dst_path} | {fnf}"
                            Log.Log_Error(log_file, error_message)
                        except Exception as e:
                            tb_str = traceback.format_exc()
                            print("\n--- DETAILED TRACEBACK ---")
                            print(tb_str)
                            print("--------------------------\n")
                            error_message = f"Unexpected error for file {src_path}: [{type(e).__name__}] {e}"
                            Log.Log_Error(log_file, error_message)
                            Log.Log_Error(log_file, tb_str)
            
            Log.Log_Info(log_file, f"Finished processing INI file: {ini_path}")

        except Exception as e:
            tb_str = traceback.format_exc()
            print("\n--- FATAL TRACEBACK ---")
            print(tb_str)
            print("-----------------------\n")
            print(f"FATAL Error processing INI file {ini_path}: [{type(e).__name__}] {e}")

    print("✅ All .ini settings have been processed.")

# --- スクリプトが直接実行された場合にmain関数を呼び出す ---
if __name__ == '__main__':
    main()