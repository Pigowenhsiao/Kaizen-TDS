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
from configparser import ConfigParser, MissingSectionHeaderError, NoSectionError
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
        self.retention_date = 90
        self.file_name_patterns = []
        self.input_paths = []
        self.output_path = ""
        self.csv_path = "" 
        self.intermediate_data_path = ""
        self.log_path = ""
        self.running_rec = ""
        self.backup_running_rec_path = ""
        self.sheet_name = ""
        self.header_rows = 5
        self.column_map = {}
        self.fill_cols_indices = []

def setup_logging(log_dir, operation_name):
    """設定日誌記錄功能"""
    log_folder = os.path.join(log_dir, str(date.today()))
    os.makedirs(log_folder, exist_ok=True)
    log_file = os.path.join(log_folder, f'{operation_name}_Special.log')
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

def _extract_settings_from_config(config):
    """從解析後的config物件中提取所有設定"""
    s = IniSettings()
    s.site = config.get('Basic_info', 'Site')
    s.product_family = config.get('Basic_info', 'ProductFamily')
    s.operation = config.get('Basic_info', 'Operation')
    s.test_station = config.get('Basic_info', 'TestStation')
    s.retention_date = config.getint('Basic_info', 'retention_date', fallback=90)
    s.file_name_patterns = [x.strip() for x in config.get('Basic_info', 'file_name_patterns').split(',')]
    
    s.input_paths = [x.strip() for x in config.get('Paths', 'input_paths').split(',')]
    s.output_path = config.get('Paths', 'output_path', fallback=None)
    s.csv_path = config.get('Paths', 'CSV_path', fallback=None)
    s.intermediate_data_path = config.get('Paths', 'intermediate_data_path')
    s.log_path = config.get('Paths', 'log_path')
    s.running_rec = config.get('Paths', 'running_rec')
    s.backup_running_rec_path = config.get('Paths', 'backup_running_rec_path', fallback=None)

    s.sheet_name = config.get('Excel', 'sheet_name')
    s.header_rows = config.getint('Excel', 'header_rows', fallback=5)
    
    temp_column_map = {}
    temp_fill_cols_str = ""
    for key, value in config.items('SpecialFormat'):
        if key == 'forward_fill_cols':
            temp_fill_cols_str = value
        else:
            temp_column_map[key] = ord(value.upper()) - 65
    s.column_map = temp_column_map
    
    if temp_fill_cols_str:
        s.fill_cols_indices = [ord(c.strip().upper()) - 65 for c in temp_fill_cols_str.split(',')]
    
    return s

def write_to_csv(csv_filepath, dataframe, log_file):
    """將 DataFrame 附加到指定的 CSV 檔案中"""
    try:
        file_exists = os.path.isfile(csv_filepath)
        dataframe.to_csv(csv_filepath, mode='a', header=not file_exists, index=False, encoding='utf-8-sig')
        return True
    except Exception as e:
        Log.Log_Error(log_file, f"函式 write_to_csv 執行失敗: {e}")
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

        Log.Log_Info(log_file, f"指標 XML 已生成於: {xml_file_path}")
    except Exception as e:
        Log.Log_Error(log_file, f"函式 generate_pointer_xml 執行失敗: {e}")

def process_special_format_excel(filepath_str, settings, log_file, csv_filepath):
    """專門處理特殊一對多格式 Excel 的核心函式"""
    filepath = Path(filepath_str)
    Log.Log_Info(log_file, f"--- 開始使用特別版邏輯處理檔案: {filepath.name} ---")
    
    start_row = max(Row_Number_Func.start_row_number(settings.running_rec), settings.header_rows)

    try:
        df_full = pd.read_excel(filepath, sheet_name=settings.sheet_name, header=None, skiprows=start_row)
        Log.Log_Info(log_file, f"讀取 Excel 成功，共讀取 {df_full.shape[0]} 行 (從第 {start_row + 1} 行開始)")
        
        if df_full.empty:
            Log.Log_Info(log_file, "檔案沒有新資料")
            return
            
        df_full.replace(r'^\s*$', np.nan, regex=True, inplace=True)
        
        date_col_idx = settings.column_map['date_col']
        df_full[date_col_idx] = pd.to_datetime(df_full[date_col_idx], errors='coerce')
        
        df_full[settings.fill_cols_indices] = df_full[settings.fill_cols_indices].ffill()
        Log.Log_Info(log_file, f"已對指定欄位執行前向填充")

        serial_no_col_idx = settings.column_map['serial_no_col']
        df_processed = df_full.dropna(subset=[serial_no_col_idx]).copy()
        Log.Log_Info(log_file, f"以 E 欄為基準，篩選出 {len(df_processed)} 筆有效資料列")
        
        if df_processed.empty:
            Log.Log_Info(log_file, "篩選後無有效資料列")
            return

        final_data = pd.DataFrame()
        final_data['Start_Date_Time'] = df_processed[date_col_idx]
        final_data['Part_Number'] = df_processed[settings.column_map['part_no_col']]
        final_data['Serial_Number'] = df_processed[serial_no_col_idx]
        final_data['Operator'] = df_processed[settings.column_map['operator_col']]
        final_data['Tool_Name'] = 'EVP' + df_processed[settings.column_map['tool_name_col']].astype(str)
        final_data['Thickness'] = df_processed[settings.column_map['thickness_col']]
        final_data['Stress'] = df_processed[settings.column_map['stress_col']]
        final_data['Result'] = df_processed[settings.column_map['result_col']]
        Log.Log_Info(log_file, "欄位賦值完成")

        # --- 執行順序調整：先進行日期篩選 ---
        final_data['datetime_obj'] = pd.to_datetime(final_data['Start_Date_Time'])
        final_data = final_data[final_data['datetime_obj'] >= (datetime.now() - relativedelta(days=settings.retention_date))]
        if final_data.empty:
            Log.Log_Info(log_file, "所有紀錄都因超過 retention_date 而被過濾")
            return
        Log.Log_Info(log_file, f"日期篩選完成，剩餘 {len(final_data)} 筆紀錄")

        # --- 在日期篩選後，才進行資料庫查詢 ---
        conn, cursor = SQL.connSQL()
        if conn:
            def get_lot9(serial):
                _, lot9 = SQL.selectSQL(cursor, str(serial))
                return lot9
            final_data['LotNumber_9'] = final_data['Serial_Number'].apply(get_lot9)
            SQL.disconnSQL(conn, cursor)
            Log.Log_Info(log_file, "資料庫查詢 LotNumber_9 完成")
        else:
            final_data['LotNumber_9'] = ''
            Log.Log_Error(log_file, "資料庫連線失敗，LotNumber_9 欄位為空")

        # 計算 SORTED 欄位
        base_date = datetime(1899, 12, 30)
        final_data['date_excel_number'] = (final_data['datetime_obj'] - base_date).dt.days
        # df_processed.index 繼承自 df_full，是相對於 start_row 的相對索引
        final_data['excel_row'] = final_data.index + start_row + 1 
        final_data['STARTTIME_SORTED'] = final_data['date_excel_number'] + (final_data['excel_row'] / 10**6)
        final_data['SORTNUMBER'] = final_data['excel_row']
        Log.Log_Info(log_file, "SORTED 欄位計算完成")

        # 最終過濾規則
        final_data['Thickness'] = pd.to_numeric(final_data['Thickness'], errors='coerce')
        final_data['Stress'] = pd.to_numeric(final_data['Stress'], errors='coerce')
        final_data.dropna(subset=['Thickness', 'Stress'], inplace=True)
        final_data = final_data[~((final_data['Thickness'] == 0) & (final_data['Stress'] == 0))]
        Log.Log_Info(log_file, f"過濾無效數據 (0 或錯誤值) 後，剩餘 {len(final_data)} 筆紀錄")

        if final_data.empty: return
        
        # 附加額外資訊
        final_data['Operation'] = settings.operation
        final_data['TestStation'] = settings.test_station
        final_data['Site'] = settings.site
        final_data['Start_Date_Time'] = final_data['datetime_obj'].dt.strftime('%Y-%m-%d %H:%M:%S')
        Log.Log_Info(log_file, "附加額外資訊完成")
        
        if csv_filepath:
            final_columns = [
                'Start_Date_Time', 'Part_Number', 'Serial_Number', 'Operator', 'Tool_Name',
                'Thickness', 'Stress', 'Result', 'LotNumber_9',
                'STARTTIME_SORTED', 'SORTNUMBER',
                'Operation', 'TestStation', 'Site'
            ]
            final_columns_exist = [col for col in final_columns if col in final_data.columns]
            write_to_csv(csv_filepath, final_data[final_columns_exist], log_file)
        
        next_start_row = start_row + df_full.shape[0]
        Row_Number_Func.next_start_row_number(settings.running_rec, next_start_row)
        Log.Log_Info(log_file, f"更新下次起始行號為 {next_start_row}")

    except Exception as e:
        Log.Log_Error(log_file, f"處理特殊格式檔案時發生錯誤: {traceback.format_exc()}")

def main():
    """主函式"""
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    ini_files = [f for f in os.listdir('.') if f.endswith('.ini')]
    
    log_file = setup_logging('../Log/', 'SpecialFormat_Init')
    Log.Log_Info(log_file, f"===== 特別版腳本開始執行 =====")

    if not ini_files:
        Log.Log_Info(log_file, "找不到任何 .ini 設定檔，程式結束")
        print("Error: No .ini config files found.")
        return
    Log.Log_Info(log_file, f"找到 {len(ini_files)} 個 .ini 設定檔: {', '.join(ini_files)}")

    for ini_path in ini_files:
        try:
            print(f"--- Processing config: {ini_path} ---")
            config = _read_and_parse_ini_config(ini_path)
            settings = _extract_settings_from_config(config)
            
            log_file = setup_logging(settings.log_path, settings.operation)
            Log.Log_Info(log_file, f"成功讀取設定檔: {ini_path}")
            
            csv_filepath = None
            if settings.csv_path:
                Path(settings.csv_path).mkdir(parents=True, exist_ok=True)
                timestamp = datetime.now().strftime('%Y_%m_%dT%H.%M.%S')
                filename = f"{settings.operation}_{timestamp}.csv"
                csv_filepath = Path(settings.csv_path) / filename
                Log.Log_Info(log_file, f"CSV 輸出檔案為: {csv_filepath}")

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
                    Log.Log_Info(log_file, f"找到來源檔案: {latest_file.name}")
                    try:
                        dst_path = shutil.copy(latest_file, intermediate_path)
                        process_special_format_excel(dst_path, settings, log_file, csv_filepath)
                    except Exception:
                        Log.Log_Error(log_file, f"處理檔案時發生錯誤: {traceback.format_exc()}")

            if not source_files_found:
                Log.Log_Info(log_file, "找不到任何相符的來源檔案")

            if csv_filepath and os.path.exists(csv_filepath) and settings.output_path:
                Log.Log_Info(log_file, "開始生成指標 XML...")
                generate_pointer_xml(
                    output_path=settings.output_path,
                    csv_path=csv_filepath,
                    settings=settings,
                    log_file=log_file
                )
        
        except Exception:
            error_message = f"處理過程中發生嚴重錯誤: {traceback.format_exc()}"
            print(error_message)
            if log_file: Log.Log_Error(log_file, error_message)

    Log.Log_Info(log_file, f"===== 特別版腳本執行完畢 =====")
    print("✅ Special format processing complete.")

if __name__ == '__main__':
    main()