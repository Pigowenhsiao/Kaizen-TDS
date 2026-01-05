import os
import sys
import glob
import shutil
import logging
import csv
import numpy as np
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
from configparser import ConfigParser
from pathlib import Path
import traceback

# 確保 MyModule 模組在 Python 的搜尋路徑中
sys.path.append('../MyModule')
import Log
import SQL
import Convert_Date
import Row_Number_Func

class IniSettings:
    """用來存放從 INI 檔案讀取的所有設定"""
    def __init__(self):
        self.site = ""
        self.product_family = ""
        self.operation = ""
        self.test_station = ""
        self.tool_name = ""
        self.retention_date = 30
        self.file_name_patterns = []
        self.input_paths = []
        self.output_path = ""
        self.csv_path = "" 
        self.intermediate_data_path = ""
        self.log_path = ""
        self.running_rec = ""
        self.backup_running_rec_path = ""
        self.sheet_name = ""
        self.data_columns = ""
        self.skip_rows = 500
        self.field_map = {}

def setup_logging(log_dir):
    """設定日誌記錄功能"""
    log_folder = os.path.join(log_dir, str(date.today()))
    os.makedirs(log_folder, exist_ok=True)
    log_file = os.path.join(log_folder, '012_MESA_CVD.log')
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    logging.basicConfig(filename=log_file, level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')
    return log_file

def _read_and_parse_ini_config(config_file_path):
    """讀取並解析INI設定檔"""
    config = ConfigParser()
    config.read(config_file_path, encoding='utf-8')
    return config

def _parse_fields_map_from_lines(fields_lines):
    """從[DataFields]區塊解析欄位對應"""
    fields = {}
    for line in fields_lines:
        if ':' in line and not line.strip().startswith('#'):
            try:
                key, col_str, dtype_str = map(str.strip, line.split(':', 2))
                fields[key] = {'col': col_str}
            except ValueError:
                continue
    return fields

def _extract_settings_from_config(config):
    """從解析後的config物件中提取所有設定"""
    s = IniSettings()
    s.site = config.get('Basic_info', 'Site')
    s.product_family = config.get('Basic_info', 'ProductFamily')
    s.operation = config.get('Basic_info', 'Operation')
    s.test_station = config.get('Basic_info', 'TestStation')
    s.tool_name = config.get('Basic_info', 'Tool_Name')
    s.retention_date = config.getint('Basic_info', 'retention_date')
    s.file_name_patterns = [x.strip() for x in config.get('Basic_info', 'file_name_patterns').split(',')]
    
    s.input_paths = [x.strip() for x in config.get('Paths', 'input_paths').split(',')]
    s.output_path = config.get('Paths', 'output_path', fallback=None)
    s.csv_path = config.get('Paths', 'CSV_path', fallback=None)
    s.intermediate_data_path = config.get('Paths', 'intermediate_data_path')
    s.log_path = config.get('Paths', 'log_path')
    s.running_rec = config.get('Paths', 'running_rec')
    s.backup_running_rec_path = config.get('Paths', 'backup_running_rec_path', fallback=None)

    s.sheet_name = config.get('Excel', 'sheet_name')
    s.data_columns = config.get('Excel', 'data_columns')
    s.skip_rows = config.getint('Excel', 'main_skip_rows')
    
    fields_lines = config.get('DataFields', 'fields').splitlines()
    s.field_map = _parse_fields_map_from_lines(fields_lines)
    return s

def write_to_csv(csv_filepath, dataframe, log_file):
    """將 DataFrame 附加到指定的 CSV 檔案中"""
    try:
        file_exists = os.path.isfile(csv_filepath)
        dataframe.to_csv(csv_filepath, mode='a', header=not file_exists, index=False, encoding='utf-8-sig')
        return True
    except Exception as e:
        Log.Log_Error(log_file, f"Error writing DataFrame to CSV {csv_filepath}: {e}")
        return False

def generate_pointer_xml(output_path, csv_path, settings, log_file):
    """產生指向CSV檔案的指標XML"""
    try:
        os.makedirs(output_path, exist_ok=True)
        now_iso = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        serial_no = Path(csv_path).stem
        
        xml_file_name = (
            f"Site={settings.site},"
            f"ProductFamily={settings.product_family},"
            f"Operation={settings.operation},"
            f"Partnumber=UNKNOWPN,"
            f"Serialnumber={serial_no},"
            f"Testdate={now_iso}.xml"
        ).replace(":", ".")
        
        xml_file_path = os.path.join(output_path, xml_file_name)

        results = ET.Element("Results", {"xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance", "xmlns:xsd": "http://www.w3.org/2001/XMLSchema"})
        result = ET.SubElement(results, "Result", startDateTime=now_iso, endDateTime=now_iso, Result="Passed")
        ET.SubElement(result, "Header",
            SerialNumber=serial_no, PartNumber="UNKNOWPN",
            Operation=settings.operation, TestStation=settings.test_station,
            Operator="NA", StartTime=now_iso, Site=settings.site, LotNumber=""
        )
        test_step = ET.SubElement(result, "TestStep", Name=settings.operation, startDateTime=now_iso, endDateTime=now_iso, Status="Passed")
        ET.SubElement(test_step, "Data", DataType="Table", Name=f"tbl_{settings.operation.upper()}", Value=str(csv_path), CompOperation="LOG")
        
        xml_str = minidom.parseString(ET.tostring(results)).toprettyxml(indent="  ", encoding="utf-8")
        with open(xml_file_path, "wb") as f:
            f.write(xml_str)

        Log.Log_Info(log_file, f"Pointer XML generated successfully at: {xml_file_path}")
    except Exception as e:
        Log.Log_Error(log_file, f"Failed to generate pointer XML. Error: {e}")

def process_excel_file(filepath_str, settings, log_file, csv_filepath):
    """以批次化、向量化的方式處理單一Excel檔案"""
    filepath = Path(filepath_str)
    Log.Log_Info(log_file, f"Processing file: {filepath.name}")
    start_row = max(Row_Number_Func.start_row_number(settings.running_rec) - settings.skip_rows, 4)
    
    try:
        df = pd.read_excel(filepath, header=None, sheet_name=settings.sheet_name, usecols=settings.data_columns, skiprows=start_row)
        
        ini_keys_by_col_index = {int(mapping['col']): key for key, mapping in settings.field_map.items()}
        df.columns = [ini_keys_by_col_index.get(i, f'unused_{i}') for i in range(df.shape[1])]
        
        if 'key_Start_Date_Time' not in df.columns:
            raise KeyError("INI mapping failed to create 'key_Start_Date_Time'. Check Excel columns and INI.")

        date_series = pd.to_datetime(df['key_Start_Date_Time'], errors='coerce')
        df = df[date_series.notna()]
        df = df[date_series >= (datetime.now() - relativedelta(days=settings.retention_date))]
        df.dropna(subset=['key_Serial_Number'], inplace=True)
    except Exception as e:
        Log.Log_Error(log_file, f"Failed to read or filter Excel file {filepath.name}. Error: {e}")
        return

    if df.empty:
        Log.Log_Info(log_file, f"No data left after initial filtering for {filepath.name}.")
        return

    conn, cursor = None, None
    try:
        conn, cursor = SQL.connSQL()
        if conn is None:
            Log.Log_Error(log_file, "Database connection failed.")
            return

        def get_db_info(serial):
            return pd.Series(SQL.selectSQL(cursor, str(serial)))

        df[['key_Part_Number', 'key_LotNumber_9']] = df['key_Serial_Number'].apply(get_db_info)
        df.dropna(subset=['key_Part_Number'], inplace=True)
        df = df[df['key_Part_Number'] != 'LDアレイ_']
    finally:
        if conn: SQL.disconnSQL(conn, cursor)

    if df.empty:
        Log.Log_Info(log_file, f"No data left after database lookup for {filepath.name}.")
        return

    def clean_date(raw_date):
        try:
            return pd.to_datetime(Convert_Date.Edit_Date(raw_date).replace('T', ' ').replace('.', ':'))
        except (ValueError, TypeError): return pd.NaT
            
    df['datetime_obj'] = df['key_Start_Date_Time'].apply(clean_date)
    df.dropna(subset=['datetime_obj'], inplace=True)
    
    base_date = datetime(1899, 12, 30)
    df['date_excel_number'] = (df['datetime_obj'] - base_date).dt.days
    df['excel_row'] = start_row + df.index + 1
    df['key_STARTTIME_SORTED'] = df['date_excel_number'] + (df['excel_row'] / 10**6)
    df['key_SORTNUMBER'] = df['excel_row']
    
    df['Operation'] = settings.operation
    df['TestStation'] = settings.test_station
    df['Site'] = settings.site
    df['key_TestEquipment_Dry'] = settings.tool_name
    df['key_Start_Date_Time'] = df['datetime_obj'].dt.strftime('%Y-%m-%d %H:%M:%S')

    # --- 1. 動態生成欄位對應字典 (Rename Map) ---
    rename_map = {}
    special_renames = {
        'key_Serial_Number': 'Serial_Number', 'key_Part_Number': 'Part_Number',
        'key_Start_Date_Time': 'Start_Date_Time', 'key_TestEquipment_Nano': 'Nanospec_DeviceSerialNumber',
        'key_TestEquipment_Dry': 'DryEtch_DeviceSerialNumber'
    }
    for key in settings.field_map.keys():
        if key in special_renames:
            rename_map[key] = special_renames[key]
        else:
            rename_map[key] = key.replace('key_', '', 1)
    rename_map.update({'Operation': 'Operation', 'TestStation': 'TestStation', 'Site': 'Site'})
    
    # --- 2. 動態建立欄位順序列表 ---
    # 核心欄位，順序固定
    dynamic_column_order = [
        'Serial_Number', 'Part_Number', 'Start_Date_Time', 'Operation', 'TestStation', 'Site'
    ]
    # 依據 INI [DataFields] 的順序，添加剩餘欄位
    for key in settings.field_map.keys():
        final_header = rename_map.get(key)
        if final_header and final_header not in dynamic_column_order:
            dynamic_column_order.append(final_header)
    # 將計算欄位添加到末尾
    dynamic_column_order.extend(['STARTTIME_SORTED', 'SORTNUMBER'])
    
    # --- 3. 整理 DataFrame 並寫入 CSV ---
    df_renamed = df.rename(columns=rename_map)
    final_columns = [col for col in dynamic_column_order if col in df_renamed.columns]
    df_to_csv = df_renamed[final_columns]

    if csv_filepath:
        if write_to_csv(csv_filepath, df_to_csv, log_file):
            Log.Log_Info(log_file, f"Successfully wrote {len(df_to_csv)} rows to CSV for file {filepath.name}")
        else:
            Log.Log_Error(log_file, f"Failed to write DataFrame to CSV for file {filepath.name}")

    original_row_count = pd.read_excel(filepath_str, header=None, sheet_name=settings.sheet_name).shape[0]
    next_start_row = start_row + original_row_count + 1
    Row_Number_Func.next_start_row_number(settings.running_rec, next_start_row)
    if settings.backup_running_rec_path:
        try:
            shutil.copy(settings.running_rec, settings.backup_running_rec_path)
        except Exception as e:
            Log.Log_Error(log_file, f"Failed to backup running_rec file. Error: {e}")

def main():
    """主函式，尋找並處理所有INI設定檔"""
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    ini_files = [f for f in os.listdir('.') if f.endswith('.ini')]
    if not ini_files:
        print("No .ini files found in the current directory.")
        return

    csv_filepath = None
    log_file = None
    first_settings = None

    for ini_path in ini_files:
        try:
            print(f"--- Processing config: {ini_path} ---")
            config = _read_and_parse_ini_config(ini_path)
            settings = _extract_settings_from_config(config)
            
            if first_settings is None: first_settings = settings
            if log_file is None: log_file = setup_logging(settings.log_path)
            
            Log.Log_Info(log_file, f"Start processing INI config file: {ini_path}")
            
            if csv_filepath is None and settings.csv_path:
                Path(settings.csv_path).mkdir(parents=True, exist_ok=True)
                timestamp = datetime.now().strftime('%Y_%m_%dT%H.%M.%S')
                filename = f"{settings.operation}_{timestamp}.csv"
                csv_filepath = Path(settings.csv_path) / filename
                Log.Log_Info(log_file, f"CSV output for this run will be: {csv_filepath}")

            intermediate_path = Path(settings.intermediate_data_path)
            intermediate_path.mkdir(parents=True, exist_ok=True)
            source_files_found = False
            for input_p_str in settings.input_paths:
                input_p = Path(input_p_str)
                for pattern in settings.file_name_patterns:
                    files = [p for p in input_p.glob(pattern) if not p.name.startswith('~$')]
                    if not files: continue
                    source_files_found = True
                    latest_file = max(files, key=os.path.getmtime)
                    try:
                        dst_path = shutil.copy(latest_file, intermediate_path)
                        process_excel_file(dst_path, settings, log_file, csv_filepath)
                    except Exception:
                        Log.Log_Error(log_file, f"Error processing file {latest_file.name}: {traceback.format_exc()}")

            if not source_files_found:
                Log.Log_Info(log_file, "No matching source files found for this configuration.")

        except Exception:
            error_message = f"FATAL Error with INI {ini_path}: {traceback.format_exc()}"
            print(error_message)
            if log_file: Log.Log_Error(log_file, error_message)

    if csv_filepath and os.path.exists(csv_filepath) and first_settings and first_settings.output_path:
        Log.Log_Info(log_file, f"Generating final pointer XML for {csv_filepath}...")
        generate_pointer_xml(
            output_path=first_settings.output_path,
            csv_path=csv_filepath,
            settings=first_settings,
            log_file=log_file
        )
    
    print("✅ All .ini configurations have been processed.")

if __name__ == '__main__':
    main()