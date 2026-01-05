
import os
import sys
import shutil
import logging
import numpy as np
import pandas as pd
from datetime import datetime, date, timedelta
from configparser import ConfigParser
from pathlib import Path
import traceback

# 從 MyModule 匯入必要的模組
sys.path.append('../MyModule')
import Log
import SQL
import Convert_Date
import Row_Number_Func

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
        self.key_col_idx = 11
        self.serial_col_idx = 4
        self.xy_x_idx = 1
        self.xy_y_idx = 2
        self.xy_points = 5
        self.tool_name_map = {}
        self.field_map = {}
        self.retention_date = 0

def setup_logging(log_dir):
    """設定日誌記錄器"""
    log_folder = os.path.join(log_dir, str(date.today()))
    os.makedirs(log_folder, exist_ok=True)
    # 將日誌檔名客製化為 Dry
    log_file = os.path.join(log_folder, '013_MESA_Dry.log')
    logging.basicConfig(filename=log_file, level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')
    return log_file

def _read_and_parse_ini_config(config_file_path):
    """讀取並解析 INI 設定檔"""
    config = ConfigParser(interpolation=None)
    config.read(config_file_path, encoding='utf-8')
    return config

def _parse_fields_map_from_lines(fields_lines):
    """從設定檔行中解析欄位對應"""
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

def _extract_settings_from_config(config):
    """從 config 物件中提取設定到 IniSettings 類別"""
    s = IniSettings()
    s.site = config.get('Basic_info', 'Site')
    s.retention_date = config.getint('Basic_info', 'retention_date')
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

def detect_tool_name(filename, tool_map):
    """根據檔名偵測工具名稱"""
    filename_str = str(filename)
    for keyword, tool in tool_map.items():
        if keyword != 'default' and keyword in filename_str:
            return tool
    return tool_map.get('default', 'UNKNOWN')

def Data_Type(key_to_type, data_dict):
    """驗證資料字典中的每個值的型態"""
    for key, expected_type in key_to_type.items():
        value = data_dict.get(key)
        if isinstance(expected_type, str):
            expected_type_name = expected_type
        else:
            expected_type_name = expected_type.__name__
        if value is None or pd.isna(value):
            print(f"資料型態錯誤: Key '{key}' 的值為空。")
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
                print(f"資料型態錯誤: Key '{key}' 的型態錯誤，期望 {expected_type_name}，實際為 {type(value).__name__}")
                return False
        except (ValueError, TypeError) as e:
            print(f"資料型態錯誤: Key '{key}' 的值 '{value}' 無法被處理為 {expected_type_name}。錯誤訊息: {e}")
            return False
    return True

def generate_xml(output_path, site, product_family, operation, test_station, data_dict, log_file):
    """根據資料字典生成 XML 檔案"""
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
            
            # 根據 field_map 動態生成 TestStep
            test_steps = {}
            for key, value in data_dict.items():
                if key.startswith('key_'):
                    parts = key.split('_')
                    if len(parts) > 2:
                        step_name = parts[1]
                        data_name = "_".join(parts[2:])
                        if step_name not in test_steps:
                            test_steps[step_name] = []
                        
                        # 決定 DataType 和 Units
                        data_type = "String" if isinstance(value, str) else "Numeric"
                        units = "No"
                        if "Time" in data_name: units = "sec"
                        if "Depth" in step_name or "Thickness" in step_name or "Etching" in step_name: units = "nm"
                        if "Rate" in data_name: units = "nm/min"
                        if key.startswith("key_X") or key.startswith("key_Y"): units = "um"
                        if "SORTED" in step_name: units = ""

                        test_steps[step_name].append(f'      <Data DataType="{data_type}" Name="{data_name}" Units="{units}" Value="{value}"/>\n')

            for step_name, data_lines in test_steps.items():
                f.write(f'    <TestStep Name="{step_name}" startDateTime="{start_time_str}" Status="Passed">\n')
                f.writelines(data_lines)
                f.write(f'    </TestStep>\n')

            f.write(f'    <TestEquipment>\n')
            # 將設備名稱客製化
            f.write(f'      <Item DeviceName="DryProcess" DeviceSerialNumber="{data_dict.get("Tool_name", "")}"/>\n')
            f.write(f'    </TestEquipment>\n')
            f.write('    <ErrorData/>\n')
            f.write('    <FailureData/>\n')
            f.write('    <Configuration/>\n')
            f.write('  </Result>\n')
            f.write('</Results>')
        Log.Log_Info(log_file, f"XML 檔案已生成: {filepath}")
        return True
    except Exception as e:
        Log.Log_Error(log_file, f"生成 XML 時發生錯誤: {e}")
        return False

def process_single_excel_file(filepath, settings, log_file):
    """處理單一 Excel 檔案的核心邏輯"""
    Log.Log_Info(log_file, f"正在處理檔案: {filepath}")

    # --- 偵錯 ---
    # 檢查 skiprows 的計算邏輯。這段邏輯可能是想從上次中斷的地方繼續，
    # 但計算方式可能導致跳過所有資料，進而引發錯誤。
    start_row_from_rec = Row_Number_Func.start_row_number(settings.running_rec)
    start_row = start_row_from_rec - settings.skip_rows
    print(f"\n--- 偵錯資訊 for {filepath.name} ---")
    print(f"從 {settings.running_rec} 讀取到的起始行號: {start_row_from_rec}")
    print(f"減去 INI 中的 main_skip_rows ({settings.skip_rows})")
    print(f"最終計算給 pandas 的 skiprows 參數值為: {start_row}")

    df = pd.read_excel(filepath, header=None, sheet_name=settings.sheet_name, usecols=settings.data_columns, skiprows=start_row)
    print(f"讀取 Excel 後，DataFrame 的維度 (shape): {df.shape}")
    if df.empty:
        print("⚠️  警告: 讀取到的 DataFrame 是空的。這很可能是錯誤的根源。請檢查 Excel 檔案內容或 skiprows 設定。")
    print("------------------------------------")

    # 日期篩選邏輯
    try:
        date_column_index = int(settings.field_map['key_Start_Date_Time']['col'])
        date_series = pd.to_datetime(df.iloc[:, date_column_index], errors='coerce')
        cutoff_date = datetime.now() - timedelta(days=settings.retention_date)
        df = df[date_series.notna() & (date_series >= cutoff_date)]
        Log.Log_Info(log_file, f"經過日期篩選後，檔案 {filepath.name} 剩餘 {len(df)} 筆有效資料。")
    except (KeyError, ValueError, IndexError) as e:
        Log.Log_Error(log_file, f"日期篩選失敗，請檢查 INI 中 key_Start_Date_Time 的設定。錯誤: {e}")
        return

    df = df[df.iloc[:, settings.key_col_idx].notna()]
    df_xy = pd.read_excel(filepath, header=None, sheet_name=settings.xy_sheet_name, usecols=settings.xy_columns)
    df.columns = range(df.shape[1])
    
    if df_xy.shape[0] < settings.xy_points + 1:
        Log.Log_Error(log_file, f"XY 座標資料不足 {settings.xy_points} 個點，實際僅 {df_xy.shape[0]} 行，跳過：{filepath}")
        return

    row_count = 0
    for idx in df.index:
        data_dict = {}
        
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
                    Log.Log_Error(log_file, f"讀取 XY 座標失敗: Key '{key}' (設定: {col_str}). 錯誤: {e}")
                continue
            if col_str == '-1':
                continue
            try:
                col_index = int(col_str)
                data_dict[key] = df.loc[idx].iloc[col_index]
            except (ValueError, IndexError):
                data_dict[key] = None
                Log.Log_Error(log_file, f"讀取欄位失敗: Key '{key}' (欄位索引: {col_str})")

        serial = data_dict.get('key_Serial_Number')
        if not serial or pd.isna(serial): continue
        serial = str(serial)

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
        
        key_to_type = {k: v['type'] for k, v in settings.field_map.items() if v['col'] != '-1'}
        if not Data_Type(key_to_type, data_dict):
            Log.Log_Error(log_file, serial + ' : ' + '資料錯誤')
            continue

        try:
            clean_datetime_str = str(data_dict.get('key_Start_Date_Time', '')).replace('T', ' ').replace('.', ':')
            date_obj = datetime.strptime(clean_datetime_str, "%Y-%m-%d %H:%M:%S")
            date_excel_number = int(str(date_obj - datetime(1899, 12, 30)).split()[0])
            excel_row = idx 
            date_excel_number += (excel_row + 1) / 10**6 
            data_dict['key_STARTTIME_SORTED'] = date_excel_number
            data_dict['key_SORTNUMBER'] = excel_row + 1 
        except Exception as e:
            Log.Log_Error(log_file, f"日期轉換失敗 {serial}: {e}")
            continue

        if generate_xml(settings.output_path, settings.site, settings.product_family, settings.operation, settings.test_station, data_dict, log_file):
            Log.Log_Info(log_file, f"{serial} : OK")
            row_count += 1
        else:
            Log.Log_Error(log_file, f"{serial} : XML 生成失敗")
            continue

    next_start_row = start_row + df.shape[0] + 1
    Row_Number_Func.next_start_row_number(settings.running_rec, next_start_row)

def main():
    """主執行函式"""
    # 將工作目錄切換到此腳本所在的目錄
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    # 指定要處理的設定檔名稱
    ini_path = "MESA_Crystal_Depth_Dry.txt"

    if not os.path.exists(ini_path):
        print(f"指定的設定檔 '{ini_path}' 不存在。")
        return

    try:
        config = _read_and_parse_ini_config(ini_path)
        settings = _extract_settings_from_config(config)
        log_file = setup_logging(settings.log_path)
        Log.Log_Info(log_file, f"開始處理設定檔: {ini_path}")
        
        input_paths = [Path(p) for p in settings.input_paths]
        intermediate_path = Path(settings.intermediate_data_path)
        intermediate_path.mkdir(parents=True, exist_ok=True)
        
        for input_p in input_paths:
            for pattern in settings.file_name_patterns:
                # --- 偵錯資訊 ---
                print(f"\nℹ️  正在路徑 '{input_p}' 中搜尋檔案...")
                print(f"ℹ️  使用模式: '{pattern}'")
                if not input_p.exists():
                    print(f"⚠️  警告: 路徑不存在或無法存取。請檢查設定檔中的 'input_paths'。")
                    continue
                
                matched_files = [p for p in input_p.glob(pattern) if not p.name.startswith('~$')]
                print(f"✅ 找到 {len(matched_files)} 個相符的檔案。")
                if not matched_files:
                    print(f"⚠️  警告: 在指定路徑下找不到任何與模式相符的檔案。請檢查 'file_name_patterns' 是否正確。")

                for src_path in matched_files:
                    dst_path = intermediate_path / src_path.name
                    try:
                        shutil.copy(src_path, dst_path)
                        process_single_excel_file(dst_path, settings, log_file)
                    except FileNotFoundError as fnf:
                        error_message = f"複製檔案時找不到檔案: {src_path} -> {dst_path} | {fnf}"
                        Log.Log_Error(log_file, error_message)
                    except Exception as e:
                        tb_str = traceback.format_exc()
                        print("\n--- 詳細錯誤追蹤 ---")
                        print(tb_str)
                        print("----------------------\n")
                        error_message = f"處理檔案 {src_path} 時發生未預期錯誤: [{type(e).__name__}] {e}"
                        Log.Log_Error(log_file, error_message)
                        Log.Log_Error(log_file, tb_str)
        
        Log.Log_Info(log_file, f"完成處理設定檔: {ini_path}")

    except Exception as e:
        tb_str = traceback.format_exc()
        print("\n--- 嚴重錯誤追蹤 ---")
        print(tb_str)
        print("-----------------------\n")
        print(f"處理設定檔 {ini_path} 時發生嚴重錯誤: [{type(e).__name__}] {e}")

    print(f"✅ 設定檔 '{ini_path}' 處理完畢。")

if __name__ == '__main__':
    main()
