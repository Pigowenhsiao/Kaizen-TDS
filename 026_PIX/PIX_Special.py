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
import re

# Ensure MyModule is in the Python search path
sys.path.append('../MyModule')
import Log
import SQL
import Convert_Date
import Row_Number_Func
import Check

class IniSettings:
    """Class to hold all settings read from the INI file (Universal Version)"""
    def __init__(self):
        # Common settings
        self.site = ""
        self.product_family = ""
        self.operation = ""
        self.test_station = ""
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
        # CVD-specific
        self.tool_name = ""
        # ICP/Dry-specific
        self.xy_sheet_name = ""
        self.xy_columns = ""
        self.tool_name_map = {}

def setup_logging(log_dir, operation_name):
    """Sets up the logging feature."""
    log_folder = os.path.join(log_dir, str(date.today()))
    os.makedirs(log_folder, exist_ok=True)
    log_file = os.path.join(log_folder, f'{operation_name}.log')
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    logging.basicConfig(filename=log_file, level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')
    return log_file

def _read_and_parse_ini_config(config_file_path):
    """Reads and parses the INI configuration file."""
    config = ConfigParser()
    config.read(config_file_path, encoding='utf-8')
    return config

def _parse_fields_map_from_lines(fields_lines):
    """Parses the field mapping from the [DataFields] section."""
    fields = {}
    for line in fields_lines:
        if ':' in line and not line.strip().startswith('#'):
            try:
                key, col_str, dtype_str = map(str.strip, line.split(':', 2))
                fields[key] = {'col': col_str, 'dtype': dtype_str}
            except ValueError:
                continue
    return fields

def _extract_settings_from_config(config):
    """Extracts all settings from the parsed config object."""
    s = IniSettings()
    # Basic Info
    s.site = config.get('Basic_info', 'Site')
    s.product_family = config.get('Basic_info', 'ProductFamily')
    s.operation = config.get('Basic_info', 'Operation')
    s.test_station = config.get('Basic_info', 'TestStation')
    s.retention_date = config.getint('Basic_info', 'retention_date', fallback=30)
    s.file_name_patterns = [x.strip() for x in config.get('Basic_info', 'file_name_patterns').split(',')]
    s.tool_name = config.get('Basic_info', 'Tool_Name', fallback=None) # CVD

    # Paths
    s.input_paths = [x.strip() for x in config.get('Paths', 'input_paths').split(',')]
    s.output_path = config.get('Paths', 'output_path', fallback=None)
    s.csv_path = config.get('Paths', 'CSV_path', fallback=None)
    s.intermediate_data_path = config.get('Paths', 'intermediate_data_path')
    s.log_path = config.get('Paths', 'log_path')
    s.running_rec = config.get('Paths', 'running_rec')
    s.backup_running_rec_path = config.get('Paths', 'backup_running_rec_path', fallback=None)

    # Excel
    s.sheet_name = config.get('Excel', 'sheet_name')
    s.data_columns = config.get('Excel', 'data_columns')
    s.skip_rows = config.getint('Excel', 'main_skip_rows')
    s.xy_sheet_name = config.get('Excel', 'xy_sheet_name', fallback=None) # ICP/Dry
    s.xy_columns = config.get('Excel', 'xy_columns', fallback=None) # ICP/Dry

    # DataFields and ToolNameMapping
    fields_lines = config.get('DataFields', 'fields').splitlines()
    s.field_map = _parse_fields_map_from_lines(fields_lines)
    if config.has_section('ToolNameMapping'): # ICP/Dry
        s.tool_name_map = dict(config.items('ToolNameMapping'))
    return s

def write_to_csv(csv_filepath, dataframe, log_file):
    """Appends a DataFrame to the specified CSV file."""
    Log.Log_Info(log_file, "Executing function write_to_csv...")
    try:
        file_exists = os.path.isfile(csv_filepath)
        dataframe.to_csv(csv_filepath, mode='a', header=not file_exists, index=False, encoding='utf-8-sig')
        Log.Log_Info(log_file, "Function write_to_csv executed successfully.")
        return True
    except Exception as e:
        Log.Log_Error(log_file, f"Function write_to_csv failed: {e}")
        return False

def generate_pointer_xml(output_path, csv_path, settings, log_file, df):
    """Generates the pointer XML file that points to the CSV."""
    Log.Log_Info(log_file, "Executing function generate_pointer_xml...")
    try:
        os.makedirs(output_path, exist_ok=True)
        # 這裡我們使用 DataFrame 的第一筆資料來生成 XML 檔名和部分內容，這是一個代表性的批次指標
        if not df.empty:
            serial_no = df['Serial_Number'].iloc[0]
            part_no = df['Part_Number'].iloc[0]
            start_time = df['Start_Date_Time'].iloc[0]
            operator = df['Operator'].iloc[0]
            start_time_iso = pd.to_datetime(start_time).strftime("%Y-%m-%dT%H:%M:%S")
        else:
            # 如果 df 是空的，則使用預設值
            serial_no = "UNKNOWN"
            part_no = "UNKNOWN"
            start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            operator = "NA"
            start_time_iso = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        
        xml_file_name = (
            f"Site={settings.site},"
            f"ProductFamily={settings.product_family},"
            f"Operation={settings.operation},"
            f"Partnumber={part_no},"
            f"Serialnumber={serial_no},"
            f"Testdate={start_time_iso.replace(':', '.')}.xml"
        )
        
        xml_file_path = os.path.join(output_path, xml_file_name)

        results = ET.Element("Results", {"xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance", "xmlns:xsd": "http://www.w3.org/2001/XMLSchema"})
        result = ET.SubElement(results, "Result", startDateTime=start_time_iso, endDateTime=start_time_iso, Result="Passed")
        
        # Header 內容符合你的範例格式，並使用第一筆資料的資訊
        ET.SubElement(result, "Header",
            SerialNumber=serial_no, PartNumber=part_no,
            Operation=settings.operation, TestStation=settings.test_station,
            Operator=operator, StartTime=start_time_iso, Site=settings.site, LotNumber=serial_no
        )
        
        # TestStep 欄位指向 CSV 檔案
        test_step = ET.SubElement(result, "TestStep", Name=settings.operation, startDateTime=start_time_iso, endDateTime=start_time_iso, Status="Passed")
        ET.SubElement(test_step, "Data", DataType="Table", Name=f"tbl_{settings.operation.upper()}", Value=str(csv_path).replace("\\", "/"), CompOperation="LOG")
        
        # TestEquipment 欄位
        test_equipment_step = ET.SubElement(result, "TestEquipment")
        ET.SubElement(test_equipment_step, "Item", DeviceName="P-CVD", DeviceSerialNumber="#1")

        xml_str = minidom.parseString(ET.tostring(results)).toprettyxml(indent="  ", encoding="utf-8")
        with open(xml_file_path, "wb") as f:
            f.write(xml_str)

        Log.Log_Info(log_file, f"Pointer XML generated successfully at: {xml_file_path}")
    except Exception as e:
        Log.Log_Error(log_file, f"Function generate_pointer_xml failed: {e}")

def _process_serial_numbers(df, log_file, conn, cursor):
    """
    Splits serial numbers, filters for valid ones, and performs database lookups.
    """
    # Step 1: Split serial numbers by multiple delimiters and expand the DataFrame
    Log.Log_Info(log_file, "Step 3a: Splitting multi-value serial numbers...")
    df['key_Serial_Number'] = df['key_Serial_Number'].astype(str).str.strip()
    df['key_Serial_Number'] = df['key_Serial_Number'].str.replace(' ', ';')
    df['key_Serial_Number'] = df['key_Serial_Number'].str.replace(',', ';')
    df = df.assign(key_Serial_Number=df['key_Serial_Number'].str.split(';')).explode('key_Serial_Number')
    df['key_Serial_Number'] = df['key_Serial_Number'].astype(str).str.strip()
    
    # Step 2: Filter out non-alphabetic starting serial numbers
    Log.Log_Info(log_file, "Step 3b: Filtering out non-alphabetic serial numbers...")
    df = df[df['key_Serial_Number'].str[0].str.isalpha()]
    
    # Step 3: Perform vectorized database lookup
    Log.Log_Info(log_file, "Step 3c: Performing vectorized database query...")
    serial_list = df['key_Serial_Number'].astype(str).tolist()
    part_lot_map = {serial: SQL.selectSQL(cursor, serial) for serial in serial_list}
    df[['key_Part_Number', 'key_LotNumber_9']] = df['key_Serial_Number'].astype(str).apply(lambda s: pd.Series(part_lot_map.get(s, (None, None))))
    
    df.dropna(subset=['key_Part_Number'], inplace=True)
    df = df[df['key_Part_Number'] != 'LDアレイ_']
    
    Log.Log_Info(log_file, f"Serial number processing complete. {df.shape[0]} valid rows remaining.")
    return df

def process_excel_file(filepath_str, settings, log_file, csv_filepath):
    """Processes a single Excel file in a batched, vectorized manner (Universal Version)."""
    filepath = Path(filepath_str)
    Log.Log_Info(log_file, f"--- Start processing file: {filepath.name} ---")
    
    start_row = Row_Number_Func.start_row_number(settings.running_rec) - settings.skip_rows
    if start_row < 0: start_row = 0
    
    try:
        # Step 1: Read the main Excel worksheet
        df = pd.read_excel(filepath, header=None, sheet_name=settings.sheet_name, usecols=settings.data_columns, skiprows=start_row)
        Log.Log_Info(log_file, f"Step 1: Successfully read main sheet '{settings.sheet_name}', {df.shape[0]} rows loaded.")
        
        df.dropna(how='all', inplace=True)
        if df.empty:
            Log.Log_Info(log_file, "No data left after initial filtering. Ending process for this file.")
            return

        excel_col_to_key = {int(v['col']): k for k, v in settings.field_map.items() if v['col'] != '-1'}
        column_indices = sorted(excel_col_to_key.keys())
        df = df[df.columns[column_indices]]
        df.columns = [excel_col_to_key[i] for i in column_indices]

        # Step 2: Initial filtering
        date_series = pd.to_datetime(df['key_Start_Date_Time'], errors='coerce')
        df = df[date_series.notna() & (date_series >= (datetime.now() - relativedelta(days=settings.retention_date)))]
        
        Log.Log_Info(log_file, f"Step 2: Initial filtering (date) complete. {df.shape[0]} rows remaining.")
        
    except Exception as e:
        Log.Log_Error(log_file, f"Step 1/2 failed: Error during Excel read or filter. Error: {e}")
        return

    if df.empty:
        Log.Log_Info(log_file, "No data left after initial filtering. Ending process for this file.")
        return

    # Step 3: Database query and data type check
    conn, cursor = None, None
    try:
        Log.Log_Info(log_file, "Step 3: Starting database connection...")
        conn, cursor = SQL.connSQL()
        if conn is None: 
            Log.Log_Error(log_file, "Database connection failed.")
            return
        
        # Process serial numbers and perform database lookup
        df = _process_serial_numbers(df, log_file, conn, cursor)

        # 這裡就是修正後的 Step 4 邏輯
        Log.Log_Info(log_file, "Step 4: Checking data types and replacing invalid values...")
        try:
            # 使用 pd.to_numeric() 將欄位轉換為數字，並將無法轉換的設定為 NaN
            df['key_SiN_Thickness'] = pd.to_numeric(df['key_SiN_Thickness'], errors='coerce').fillna(0)
            df['key_SiN_Refraction'] = pd.to_numeric(df['key_SiN_Refraction'], errors='coerce').fillna(0)
        except KeyError:
            Log.Log_Info(log_file, "Thickness or Refraction columns not found, skipping fillna.")
            
        # 取得 INI 中所有浮點數型態的欄位，並過濾出那些在 DataFrame 中**確實存在**的欄位
        float_fields_from_ini = [k for k, v in settings.field_map.items() if v['dtype'] == 'float']
        existing_float_fields = [k for k in float_fields_from_ini if k in df.columns]
        
        # 僅對那些存在的浮點數欄位進行空值刪除
        df.dropna(subset=existing_float_fields, inplace=True)
        
        Log.Log_Info(log_file, f"Data type check and fillna complete. {df.shape[0]} rows remaining.")

    finally:
        if conn:
            SQL.disconnSQL(conn, cursor)
            Log.Log_Info(log_file, "Database connection closed.")
    
    if df.empty:
        Log.Log_Info(log_file, "No data left after database lookup and type check. Ending process for this file.")
        return

    # Step 5: Data transformation and calculation
    Log.Log_Info(log_file, "Step 5: Starting data transformation and calculation...")
    df['datetime_obj'] = df['key_Start_Date_Time'].apply(Convert_Date.Edit_Date).apply(lambda x: pd.to_datetime(x.replace('T', ' ').replace('.', ':'), errors='coerce'))
    df.dropna(subset=['datetime_obj'], inplace=True)
    
    base_date = datetime(1899, 12, 30)
    df['date_excel_number'] = (df['datetime_obj'] - base_date).dt.days
    df['excel_row'] = start_row + df.index + 1
    df['key_STARTTIME_SORTED'] = df['date_excel_number'] + (df['excel_row'] / 10**6)
    df['key_SORTNUMBER'] = df['excel_row']
    Log.Log_Info(log_file, "Date and SORTED field calculations complete.")
    
    # Step 6: Append additional info
    Log.Log_Info(log_file, "Step 6: Appending additional info (Operation, ToolName, etc.) complete.")
    df['Operation'] = settings.operation
    df['TestStation'] = settings.test_station
    df['Site'] = settings.site
    df['key_Start_Date_Time'] = df['datetime_obj'].dt.strftime('%Y-%m-%d %H:%M:%S')
    df['key_Tool_name'] = "P-CVD"
    
    # Step 7: Dynamically generate columns and write to CSV and XML
    Log.Log_Info(log_file, f"Step 7: Preparing to write {len(df)} rows to CSV...")
    
    rename_map = {
        'key_Start_Date_Time': 'Start_Date_Time',
        'key_Operator': 'Operator',
        'key_Serial_Number': 'Serial_Number',
        'key_Part_Number': 'Part_Number',
        'key_SiN_Thickness': 'SiN_Thickness',
        'key_SiN_Refraction': 'SiN_Refraction',
        'key_LotNumber_9': 'LotNumber_9',
        'key_Tool_name': 'Tool_name',
        'key_STARTTIME_SORTED': 'STARTTIME_SORTED',
        'key_SORTNUMBER': 'SORTNUMBER'
    }
    df_renamed = df.rename(columns=rename_map)
    final_columns = [rename_map.get(key, key) for key in settings.field_map.keys() if rename_map.get(key, key) in df_renamed.columns]
    final_columns.extend(['Operation', 'TestStation', 'Site'])
    df_to_csv = df_renamed[final_columns]
    
    # Check if a CSV file path was provided
    if csv_filepath:
        write_to_csv(csv_filepath, df_to_csv, log_file)
        Log.Log_Info(log_file, f"CSV created at: {csv_filepath}")
        
        # 這裡只呼叫 generate_pointer_xml 來生成單一 XML 指標檔
        if settings.output_path:
            generate_pointer_xml(settings.output_path, csv_filepath, settings, log_file, df_to_csv)


    # Step 8: Update the starting row record
    original_row_count = pd.read_excel(filepath_str, header=None, sheet_name=settings.sheet_name).shape[0]
    next_start_row = start_row + original_row_count + 1
    Row_Number_Func.next_start_row_number(settings.running_rec, next_start_row)
    Log.Log_Info(log_file, f"Step 8: Updating next start row to {next_start_row}")
    if settings.backup_running_rec_path:
        try: shutil.copy(settings.running_rec, settings.backup_running_rec_path)
        except Exception as e: Log.Log_Error(log_file, f"Failed to backup running_rec file: {e}")
    
    Log.Log_Info(log_file, f"--- Function process_excel_file executed successfully ---")

def main():
    """Main function to find and process all INI files."""
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    log_file = setup_logging('../Log/', 'UniversalScript_Init')
    Log.Log_Info(log_file, "===== Universal Script Start =====")

    ini_files = [f for f in os.listdir('.') if f.endswith('.ini')]
    if not ini_files:
        Log.Log_Info(log_file, "No .ini or .txt config files found in the current directory. Exiting.")
        print("No config files (.ini or .txt) found in the current directory.")
        return
    Log.Log_Info(log_file, f"Found {len(ini_files)} config file(s): {', '.join(ini_files)}")

    for ini_path in ini_files:
        try:
            print(f"--- Processing config: {ini_path} ---")
            config = _read_and_parse_ini_config(ini_path)
            settings = _extract_settings_from_config(config)
            
            # Set up a specific log file for this operation
            log_file = setup_logging(settings.log_path, settings.operation)
            Log.Log_Info(log_file, f"--- Start processing config file: {ini_path} ---")

            # 創建一個獨特的 CSV 檔案路徑
            csv_filepath_for_this_ini = None
            if settings.csv_path:
                Path(settings.csv_path).mkdir(parents=True, exist_ok=True)
                timestamp = datetime.now().strftime('%Y_%m_%dT%H.%M.%S')
                filename = f"{settings.operation}_{timestamp}.csv"
                csv_filepath_for_this_ini = Path(settings.csv_path) / filename
                Log.Log_Info(log_file, f"CSV output for this config will be: {csv_filepath_for_this_ini}")

            intermediate_path = Path(settings.intermediate_data_path)
            intermediate_path.mkdir(parents=True, exist_ok=True)
            source_files_found = False
            for input_p_str in settings.input_paths:
                input_p = Path(input_p_str)
                for pattern in settings.file_name_patterns:
                    Log.Log_Info(log_file, f"Searching in path '{input_p}' with pattern '{pattern}'")
                    files = [p for p in input_p.glob(pattern) if not p.name.startswith('~$')]
                    if not files: continue
                    source_files_found = True
                    latest_file = max(files, key=os.path.getmtime)
                    Log.Log_Info(log_file, f"Found latest source file: {latest_file.name}")
                    try:
                        dst_path = shutil.copy(latest_file, intermediate_path)
                        Log.Log_Info(log_file, f"File copied successfully -> {dst_path}")
                        # 修正: 將 CSV 路徑傳入 process_excel_file
                        process_excel_file(dst_path, settings, log_file, csv_filepath_for_this_ini)
                    except Exception:
                        Log.Log_Error(log_file, f"Error processing file {latest_file.name}: {traceback.format_exc()}")

            if not source_files_found:
                Log.Log_Info(log_file, "No matching source files found for this configuration.")

            Log.Log_Info(log_file, f"--- Finished processing config file: {ini_path} ---")

        except Exception:
            error_message = f"FATAL Error with INI {ini_path}: {traceback.format_exc()}"
            print(error_message)
            if log_file: Log.Log_Error(log_file, error_message)

    Log.Log_Info(log_file, "===== Universal Script End =====")
    print("✅ All .ini configurations have been processed.")

if __name__ == '__main__':
    main()