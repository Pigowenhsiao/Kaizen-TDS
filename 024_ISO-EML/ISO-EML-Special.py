# -*- coding: utf-8 -*-
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

# Ensure MyModule is in the Python search path
sys.path.append('../MyModule')
import Log
import SQL
import Convert_Date
import Row_Number_Func

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
                fields[key] = {'col': col_str}
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

def detect_tool_name(filename, tool_map):
    """Detects tool name based on filename (for ICP/Dry)."""
    filename_str = str(filename)
    for keyword, tool in tool_map.items():
        if keyword != 'default' and keyword in filename_str:
            return tool
    return tool_map.get('default', 'UNKNOWN')

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

def generate_pointer_xml(output_path, csv_path, settings, log_file):
    """Generates the pointer XML file that points to the CSV."""
    Log.Log_Info(log_file, "Executing function generate_pointer_xml...")
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
        Log.Log_Error(log_file, f"Function generate_pointer_xml failed: {e}")

# ===================== DF 展開成 sample.csv 欄位 =====================
def df_expand_to_sample(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    規則（經與使用者確認）：
    1) INI 使用到的欄位若有空白，先用 ffill 補值（合併儲存格）
    2) 以第 6 欄(索引5) 的測定位置 A/B/C/D 篩出有效列並展開
    3) Resist: 值=第 7 欄(索引6)，AVG=第 8 欄(索引7)
       Y1    : 值=第21 欄(索引20)，AVG=第22 欄(索引21)
       Y2    : 值=第48 欄(索引47)，AVG=第49 欄(索引48)
       Step  : 第16 欄(索引15)
       Start_Date_Time: 第2欄(索引1)
       Operator       : 第3欄(索引2)
       Part_Number    : 第4欄(索引3)
       Serial_Number  : 第5欄(索引4)
    4) 缺值用對應 AVG 補齊（不重算 AVG）
    5) 以 (Serial_Number, Start_Date_Time) 去重（保留第一列）
    6) 欄位順序對齊 sample.csv
    """

    POS_COL = 5
    RES_VAL, RES_AVG = 6, 7
    Y1_VAL,  Y1_AVG  = 20, 21
    Y2_VAL,  Y2_AVG  = 47, 48
    STEP_COL = 15
    SDT_COL, OPR_COL, PART_COL, SNR_COL = 1, 2, 3, 4

    # 先針對可能用到的欄位做 ffill（合併儲存格補值）
    for col in [SDT_COL, OPR_COL, PART_COL, SNR_COL, STEP_COL, RES_AVG, Y1_AVG, Y2_AVG]:
        if col in df_raw.columns:
            df_raw[col] = df_raw[col].ffill()

    # 只保留 A/B/C/D
    pos_rows = df_raw[df_raw[POS_COL].isin(list('ABCD'))].copy()
    if pos_rows.empty:
        return pd.DataFrame(columns=[
            'Serial_Number','Part_Number','Start_Date_Time','Operator',
            'Operation','TestStation',
            'Resist1','Resist2','Resist3','Resist4','Resist_AVG',
            'Step',
            'Y1_1','Y1_2','Y1_3','Y1_4','Y1_AVG',
            'Y2_1','Y2_2','Y2_3','Y2_4','Y2_AVG',
            'STARTTIME_SORTED','SORTNUMBER','LotNumber_5','LotNumber_9'
        ])

    # 位置映射
    pos_idx = {'A': 1, 'B': 2, 'C': 3, 'D': 4}
    pos_rows['pos_i'] = pos_rows[POS_COL].map(pos_idx)

    def assemble(group: pd.DataFrame) -> pd.Series:
        """
        群組內 pivot 取值，避免逐列迭代造成跨群組沾黏。
        """
        out = {}
        g = group.copy()

        # Resist
        res_vals = g.set_index('pos_i')[RES_VAL] if RES_VAL in g.columns else pd.Series(dtype=float)
        res_avg  = g[RES_AVG].dropna().iloc[0] if (RES_AVG in g.columns and g[RES_AVG].notna().any()) else np.nan
        for i in (1, 2, 3, 4):
            out[f'Resist{i}'] = res_vals.get(i, np.nan)
            if pd.isna(out[f'Resist{i}']) and pd.notna(res_avg):
                out[f'Resist{i}'] = res_avg
        out['Resist_AVG'] = res_avg

        # Y1
        y1_vals = g.set_index('pos_i')[Y1_VAL] if Y1_VAL in g.columns else pd.Series(dtype=float)
        y1_avg  = g[Y1_AVG].dropna().iloc[0] if (Y1_AVG in g.columns and g[Y1_AVG].notna().any()) else np.nan
        for i in (1, 2, 3, 4):
            out[f'Y1_{i}'] = y1_vals.get(i, np.nan)
            if pd.isna(out[f'Y1_{i}']) and pd.notna(y1_avg):
                out[f'Y1_{i}'] = y1_avg
        out['Y1_AVG'] = y1_avg

        # Y2
        y2_vals = g.set_index('pos_i')[Y2_VAL] if Y2_VAL in g.columns else pd.Series(dtype=float)
        y2_avg  = g[Y2_AVG].dropna().iloc[0] if (Y2_AVG in g.columns and g[Y2_AVG].notna().any()) else np.nan
        for i in (1, 2, 3, 4):
            out[f'Y2_{i}'] = y2_vals.get(i, np.nan)
            if pd.isna(out[f'Y2_{i}']) and pd.notna(y2_avg):
                out[f'Y2_{i}'] = y2_avg
        out['Y2_AVG'] = y2_avg

        # 基本欄位：取第一個非空即可
        out['Serial_Number'] = g[SNR_COL].dropna().iloc[0] if SNR_COL in g.columns else None
        out['Part_Number']   = g[PART_COL].dropna().iloc[0] if PART_COL in g.columns else None
        out['Operator']      = g[OPR_COL].dropna().iloc[0] if OPR_COL in g.columns else None

        sdt_raw = g[SDT_COL].dropna().iloc[0] if SDT_COL in g.columns else None
        sdt = pd.to_datetime(sdt_raw, errors='coerce')
        if pd.notna(sdt):
            out['Start_Date_Time'] = sdt.strftime('%Y/%-m/%-d %H:%M') if sys.platform != 'win32' else sdt.strftime('%Y/%#m/%#d %H:%M')
        else:
            out['Start_Date_Time'] = str(sdt_raw) if sdt_raw is not None else ''

        out['Step'] = g[STEP_COL].dropna().iloc[0] if (STEP_COL in g.columns and g[STEP_COL].notna().any()) else None

        # 保留欄位位置（後段流程再補）
        out['Operation']        = None
        out['TestStation']      = None
        out['STARTTIME_SORTED'] = None
        out['SORTNUMBER']       = None
        out['LotNumber_5']      = None
        out['LotNumber_9']      = None
        return pd.Series(out)

    # 以 (Serial_Number, Start_Date_Time-原始欄位) 群組
    grouped_records = []
    for (snr, sdt), g in pos_rows.groupby([SNR_COL, SDT_COL], dropna=False):
        grouped_records.append(assemble(g))
    wide = pd.DataFrame(grouped_records)

    # 依 (Serial_Number, Start_Date_Time) 去重（保留第一列）
    wide = wide.drop_duplicates(subset=['Serial_Number', 'Start_Date_Time'], keep='first')

    # 對齊 sample.csv 欄位順序
    final_cols = [
        'Serial_Number','Part_Number','Start_Date_Time','Operator',
        'Operation','TestStation',
        'Resist1','Resist2','Resist3','Resist4','Resist_AVG',
        'Step',
        'Y1_1','Y1_2','Y1_3','Y1_4','Y1_AVG',
        'Y2_1','Y2_2','Y2_3','Y2_4','Y2_AVG',
        'STARTTIME_SORTED','SORTNUMBER','LotNumber_5','LotNumber_9'
    ]
    return wide.reindex(columns=final_cols)
# ====================================================================

def process_excel_file(filepath_str, settings, log_file, csv_filepath):
    """Processes a single Excel file (DF-preview mode: stop after DF processing)."""
    filepath = Path(filepath_str)
    Log.Log_Info(log_file, f"--- Start processing file: {filepath.name} ---")
    start_row = max(Row_Number_Func.start_row_number(settings.running_rec) - settings.skip_rows, 4)
    start_row = int(20)  # 保持與原程式相同的行為（暫不調整）

    try:
        # Step 1: Read the main Excel worksheet（保持原參數）
        df = pd.read_excel(filepath, header=None, sheet_name=settings.sheet_name,
                           usecols=settings.data_columns, skiprows=start_row)
        Log.Log_Info(log_file, f"Step 1: Successfully read main sheet '{settings.sheet_name}', {df.shape[0]} rows loaded.")

        # DF-only 展開：ffill + 展開 + 去重 + 對齊 sample.csv
        df_sample_like = df_expand_to_sample(df)

        # 完整印出供人工確認
        with pd.option_context('display.max_rows', None, 'display.max_columns', None, 'display.width', 200):
            print("\n===== DF after processing (aligned to sample.csv columns) =====")
            print(df_sample_like)

        # 記錄到 log
        Log.Log_Info(log_file, f"DF-only processing completed. Rows: {len(df_sample_like)}. Printing to console for verification.")

        # 暫停程式，等待人工確認
        input("\n(Program paused) Press <Enter> after you finish checking the DF... ")

        # 依你的要求：到此暫停並返回，不往下做 DB/CSV/XML
        Log.Log_Info(log_file, "Per request, stopping after DF processing (no DB/CSV/XML in this run).")
        return

        # ===== 以下保留原邏輯，但目前不會執行（因為上面 return）=====
        # ini_keys_by_col_index = {int(v['col']): k for k, v in settings.field_map.items() if not v['col'].startswith('xy_')}
        # df.columns = [ini_keys_by_col_index.get(i, f'unused_{i}') for i in range(df.shape[1])]
        # ...（初篩、DB查詢、重命名、寫CSV、產XML等—全部維持原狀）
        # ============================================================

    except Exception as e:
        Log.Log_Error(log_file, f"Step 1/DF failed: Error during Excel read or DF processing. Error: {e}")
        print(f"[ERROR] DF processing failed for {filepath.name}: {e}")
        return

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
            
            # Create a unique CSV file path（此輪不輸出 CSV，但保留展示）
            csv_filepath_for_this_ini = None
            if settings.csv_path:
                Path(settings.csv_path).mkdir(parents=True, exist_ok=True)
                timestamp = datetime.now().strftime('%Y_%m_%dT%H.%M.%S')
                filename = f"{settings.operation}_{timestamp}.csv"
                csv_filepath_for_this_ini = Path(settings.csv_path) / filename
                Log.Log_Info(log_file, f"(Preview mode) CSV output (unused in this run) would be: {csv_filepath_for_this_ini}")

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
                        process_excel_file(dst_path, settings, log_file, csv_filepath_for_this_ini)
                    except Exception:
                        Log.Log_Error(log_file, f"Error processing file {latest_file.name}: {traceback.format_exc()}")

            if not source_files_found:
                Log.Log_Info(log_file, "No matching source files found for this configuration.")

            # Pointer XML 保留但此輪不執行（process_excel_file 已 return）
            if False and csv_filepath_for_this_ini and os.path.exists(csv_filepath_for_this_ini) and settings.output_path:
                Log.Log_Info(log_file, f"--- Generating pointer XML for {ini_path} ---")
                generate_pointer_xml(
                    output_path=settings.output_path,
                    csv_path=csv_filepath_for_this_ini,
                    settings=settings,
                    log_file=log_file
                )
            
            Log.Log_Info(log_file, f"--- Finished processing config file: {ini_path} ---")

        except Exception:
            error_message = f"FATAL Error with INI {ini_path}: {traceback.format_exc()}"
            print(error_message)
            if log_file: Log.Log_Error(log_file, error_message)

    Log.Log_Info(log_file, "===== Universal Script End =====")
    print("✅ All .ini configurations have been processed (DF-preview mode).")

if __name__ == '__main__':
    main()
