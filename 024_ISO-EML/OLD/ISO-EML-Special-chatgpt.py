import os
import sys
import glob
import shutil
import logging
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
from configparser import ConfigParser
from pathlib import Path
import traceback
from math import isnan
import numpy as np

# Set pandas option to avoid FutureWarning
pd.set_option('future.no_silent_downcasting', True)

# Ensure MyModule is in the Python search path
sys.path.append('../MyModule')
import Log
import SQL
import Convert_Date
import Row_Number_Func
import Check

class IniSettings:
    """Class to hold all settings read from the INI file"""
    def __init__(self):
        self.site = ""
        self.product_family = ""
        self.operation = ""
        self.test_station = ""
        self.retention_date = 7
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
        self.header_row = 0
        self.serial_number_source_column_idx = 0
        self.field_map = {}
        self.rename_map = {}

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

def _extract_settings_from_config(config):
    """Extracts all settings from the parsed config object."""
    s = IniSettings()
    s.site = config.get('Basic_info', 'Site')
    s.product_family = config.get('Basic_info', 'ProductFamily')
    s.operation = config.get('Basic_info', 'Operation')
    s.test_station = config.get('Basic_info', 'TestStation')
    s.retention_date = config.getint('Basic_info', 'retention_date', fallback=7)
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
    s.serial_number_source_column_idx = config.getint('Excel', 'serial_number_source_column_idx')

    fields_lines = config.get('DataFields', 'fields').splitlines()
    for line in fields_lines:
        if ':' in line and not line.strip().startswith('#'):
            key, col_str, dtype_str = map(str.strip, line.split(':', 2))
            s.field_map[key] = {'col': int(col_str), 'dtype': dtype_str}

    rename_map_items = dict(config.items('ColumnMapping'))['rename_map'].split(',')
    s.rename_map = {item.split(':')[0].strip(): item.split(':')[1].strip() for item in rename_map_items}

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

def _get_part_and_lots(serial_number, log_file):
    """從 SQL 取得 Part Number 與 9 碼批號，並推導 LotNumber_5。抓不到則回傳 (None, None, None)。"""
    try:
        conn, cursor = SQL.connSQL()
        if conn is None:
            Log.Log_Error(log_file, f"Connection with Prime Failed for {serial_number}")
            return None, None, None
        part_number, nine_serial_number = SQL.selectSQL(cursor, str(serial_number))
        SQL.disconnSQL(conn, cursor)
        lot5 = (nine_serial_number[:5] if isinstance(nine_serial_number, str) and len(nine_serial_number) >= 5 else None)
        return part_number, lot5, nine_serial_number
    except Exception as e:
        Log.Log_Error(log_file, f"SQL lookup failed for {serial_number}: {e}")
        return None, None, None

def process_iso_eml_file(filepath_str, settings, log_file, csv_filepath, config=None):
    """
    新版處理順序（依你的需求）：
      1) 讀取 Excel
      2) 只保留 INI 指定的欄位（key_*）
      3) 先展開為「每個 Serial_Number 單行」：Resist1~4 / Y1_1~4 / Y2_1~4 與 AVG
      4) 再開始篩選（必要欄位、日期、型別）與計算排序欄位
    其他流程（找檔、log、running_rec、XML）維持不變。
    """
    filepath = Path(filepath_str)
    Log.Log_Info(log_file, f"--- Start processing ISO-EML file: {filepath.name} ---")

    start_number = max(Row_Number_Func.start_row_number(settings.running_rec) - settings.skip_rows, 4)

    try:
        # 1) 讀取 Excel（依 INI）
        df_raw = pd.read_excel(
            filepath,
            header=None,
            sheet_name=settings.sheet_name,
            usecols=settings.data_columns,
            skiprows=start_number
        )
        Log.Log_Info(log_file, f"Step 1: read sheet='{settings.sheet_name}', shape={df_raw.shape}")

        # 依 INI DataFields 動態命名欄位
        col_indices_map = {v['col'] - 1: k for k, v in settings.field_map.items() if v['col'] != -1}
        col_names = [col_indices_map.get(i, f'col_{i}') for i in range(df_raw.shape[1])]
        df_raw.columns = col_names

        # 2) 僅保留 INI 指定欄位
        ini_cols = [k for k, v in settings.field_map.items() if v['col'] != -1]
        keep_cols = [c for c in ini_cols if c in df_raw.columns]
        df_keep = df_raw[keep_cols].copy()
        Log.Log_Info(log_file, f"Step 2: keep ini columns -> {len(keep_cols)} cols")

        # 針對合併儲存格的關鍵欄位先做 ffill/bfill（避免空白被濾掉）
        ffill_candidates = [
            'key_Start_Date_Time', 'key_Operator', 'key_Serial_Number',
            'key_Part_Number', 'key_Lot_Sheet', 'key_LotNumber_9'
        ]
        ffill_cols = [c for c in ffill_candidates if c in df_keep.columns]
        if ffill_cols:
            df_keep[ffill_cols] = df_keep[ffill_cols].ffill().bfill()
        Log.Log_Info(log_file, f"Step 2.5: ffill/bfill on {ffill_cols}")

        # 3) 先展開成「每個 Serial 單行」
        def _safe_get(r, col):
            try:
                return df_keep.iloc[r, df_keep.columns.get_loc(col)]
            except Exception:
                return np.nan

        records = []
        idx_min, idx_max = (df_keep.index.min(), df_keep.index.max()) if len(df_keep) else (0, -1)
        for i in range(idx_min, idx_max + 1):
            serial = _safe_get(i, 'key_Serial_Number')
            # 先建立一列，不馬上過濾，讓你能在展開後再做統一篩選
            resist_vals = [ _safe_get(i-1,'key_Resist_Resist'), _safe_get(i,'key_Resist_Resist'), _safe_get(i+1,'key_Resist_Resist'), _safe_get(i+2,'key_Resist_Resist') ]
            y1_vals     = [ _safe_get(i-1,'key_Y1_Y1'),       _safe_get(i,'key_Y1_Y1'),       _safe_get(i+1,'key_Y1_Y1'),       _safe_get(i+2,'key_Y1_Y1') ]
            y2_vals     = [ _safe_get(i-1,'key_Y2_Y2'),       _safe_get(i,'key_Y2_Y2'),       _safe_get(i+1,'key_Y2_Y2'),       _safe_get(i+2,'key_Y2_Y2') ]

            rec = {
                'Serial_Number': serial,
                'Start_Date_Time': _safe_get(i, 'key_Start_Date_Time'),
                'Operator': _safe_get(i, 'key_Operator'),
                'Step': _safe_get(i, 'key_Step_Step'),
                'Resist1': resist_vals[0], 'Resist2': resist_vals[1], 'Resist3': resist_vals[2], 'Resist4': resist_vals[3],
                'Resist_AVG': np.nanmean(resist_vals),
                'Y1_1': y1_vals[0], 'Y1_2': y1_vals[1], 'Y1_3': y1_vals[2], 'Y1_4': y1_vals[3], 'Y1_AVG': np.nanmean(y1_vals),
                'Y2_1': y2_vals[0], 'Y2_2': y2_vals[1], 'Y2_3': y2_vals[2], 'Y2_4': y2_vals[3], 'Y2_AVG': np.nanmean(y2_vals),
                # 暫存 Excel 列號，用於後續排序欄位
                '_EXCEL_ROW': start_number + i + 1
            }
            records.append(rec)

        df_expanded = pd.DataFrame(records)
        Log.Log_Info(log_file, f"Step 3: expanded to per-serial rows, shape={df_expanded.shape}")

        # 3.5) （選擇性）清理：將 'OK'/'NG'/0 → NaN，以免干擾數值
        df_expanded = df_expanded.replace(["OK", "NG"], np.nan).replace(0, np.nan)

        # 4) 開始篩選邏輯（在『展開』之後）
        # 4-1) 必要欄位齊全
        req_cols = ['Serial_Number', 'Start_Date_Time', 'Operator']
        mask_required = pd.Series(True, index=df_expanded.index)
        for c in req_cols:
            mask_required &= df_expanded[c].notna()
        df_filt = df_expanded[mask_required].copy()
        Log.Log_Info(log_file, f"Step 4.1: required fields filter -> {df_filt.shape}")

        # 4-2) 日期處理：先走 Convert_Date，失敗再容錯；並計算排序欄位
        def _fmt19(dt):
            return dt.strftime('%Y-%m-%d %H:%M:%S')

        start_sorted = []
        sort_number = []
        start_dt_out = []
        keep_idx = []
        for idx, row in df_filt.iterrows():
            raw_dt = row['Start_Date_Time']
            edited = None
            try:
                tmp = Convert_Date.Edit_Date(raw_dt)
                if isinstance(tmp, str) and len(tmp) == 19:
                    edited = tmp
            except Exception:
                pass
            if not edited:
                # Excel 序號
                try:
                    val = float(raw_dt)
                    base = datetime(1899, 12, 30)
                    edited = _fmt19(base + timedelta(days=val))
                except Exception:
                    pass
            if not edited:
                try:
                    raw_str = str(raw_dt).replace('T',' ').replace('.',':').strip()
                    dt_obj = pd.to_datetime(raw_str, errors='coerce')
                    if pd.notna(dt_obj):
                        edited = _fmt19(dt_obj.to_pydatetime())
                except Exception:
                    pass
            if not edited:
                for fmt in ("%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M"):
                    try:
                        edited = _fmt19(datetime.strptime(str(raw_dt).strip(), fmt))
                        break
                    except Exception:
                        continue

            if not edited or len(edited) != 19:
                Log.Log_Error(log_file, f"{row['Serial_Number']} : Date Error -> raw={repr(raw_dt)}, edited={repr(edited)}")
                continue  # 丟掉此列

            start_dt_out.append(edited)
            d = datetime.strptime(edited, "%Y-%m-%d %H:%M:%S")
            excel_num = int((d - datetime(1899,12,30)).days)
            sr = row['_EXCEL_ROW'] if not pd.isna(row['_EXCEL_ROW']) else 0
            start_sorted.append(excel_num + (sr / 10**6))
            sort_number.append(sr)
            keep_idx.append(idx)

        df_kept = df_filt.loc[keep_idx].copy()
        df_kept['Start_Date_Time'] = start_dt_out
        df_kept['STARTTIME_SORTED'] = start_sorted
        df_kept['SORTNUMBER'] = sort_number
        Log.Log_Info(log_file, f"Step 4.2: date & sorting computed -> {df_kept.shape}")

        # 4-3) 先完成所有篩選後，再進行「單次連線、批次查詢」SQL 以取得 PN/Lot
        #     規則：
        #       - 整段期間只建立一次 SQL 連線；失敗則記錄並結束此檔處理（避免寫出不含 Part_Number 的資料）。
        #       - 僅保留 "查得到 Part_Number" 的列；查無者剔除，不輸出到 CSV。
        #
        #     實作：
        #       1) 準備唯一 Serial 清單
        #       2) 建立單一連線，逐一 selectSQL(cursor, serial)
        #       3) 將查回結果對應回 df_kept，並剔除 Part_Number 為空的列
        serial_series = df_kept['Serial_Number'].astype(str)
        unique_serials = serial_series.dropna().unique().tolist()

        part_map = {}   # serial -> part_number
        lot9_map = {}   # serial -> lot_number_9

        conn = None
        cursor = None
        try:
            conn, cursor = SQL.connSQL()
            if conn is None:
                Log.Log_Error(log_file, "[SQL] Connection failed. Abort processing this file.")
                return

            for sn in unique_serials:
                try:
                    pn, lot9 = SQL.selectSQL(cursor, str(sn))
                    part_map[str(sn)] = pn
                    lot9_map[str(sn)] = lot9
                except Exception as e:
                    Log.Log_Error(log_file, f"[SQL] select failed for SN={sn}: {e}")
                    part_map[str(sn)] = None
                    lot9_map[str(sn)] = None
        except Exception as e:
            Log.Log_Error(log_file, f"[SQL] connection/select error: {e}")
            return
        finally:
            try:
                if conn is not None and cursor is not None:
                    SQL.disconnSQL(conn, cursor)
            except Exception:
                pass

        # 對應回 df_kept
        df_kept['Part_Number'] = serial_series.map(lambda s: part_map.get(str(s)))
        df_kept['LotNumber_9'] = serial_series.map(lambda s: lot9_map.get(str(s)))
        df_kept['LotNumber_5'] = df_kept['LotNumber_9'].apply(lambda x: x[:5] if isinstance(x, str) and len(x) >= 5 else None)

        # 只保留查得到 Part_Number 的列
        before_rows = len(df_kept)
        df_kept = df_kept[df_kept['Part_Number'].notna() & (df_kept['Part_Number'] != '')].copy()
        Log.Log_Info(log_file, f"Step 4.3: SQL mapped & filtered by Part_Number -> {df_kept.shape} (removed {before_rows - len(df_kept)})")

        # 4-4) 最終輸出欄位
        expected_cols = [
            'Serial_Number','Part_Number','Start_Date_Time','Operator','Operation','TestStation',
            'Resist1','Resist2','Resist3','Resist4','Resist_AVG',
            'Step',
            'Y1_1','Y1_2','Y1_3','Y1_4','Y1_AVG',
            'Y2_1','Y2_2','Y2_3','Y2_4','Y2_AVG',
            'STARTTIME_SORTED','SORTNUMBER','LotNumber_5','LotNumber_9'
        ]
        df_kept['Operation'] = settings.operation
        df_kept['TestStation'] = settings.test_station

        df_to_csv = df_kept.reindex(columns=expected_cols)
        Log.Log_Info(log_file, f"Step 4.4: final df_to_csv shape={df_to_csv.shape}")
        write_to_csv(csv_filepath, df_to_csv, log_file)

    except Exception as e:
        Log.Log_Error(log_file, f"Processing failed: {e}")
        return

    # 維持原有進度與備份邏輯
    try:
        total_rows = pd.read_excel(filepath_str, header=None, sheet_name=settings.sheet_name).shape[0]
        next_start_row = start_number + total_rows + 1
        Row_Number_Func.next_start_row_number(settings.running_rec, next_start_row)
        Log.Log_Info(log_file, f"Step 5: Updating next start row to {next_start_row}")
        if settings.backup_running_rec_path:
            shutil.copy(settings.running_rec, settings.backup_running_rec_path)
    except Exception as e:
        Log.Log_Error(log_file, f"Failed to update or backup running_rec file: {e}")

    Log.Log_Info(log_file, f"--- Function process_iso_eml_file executed successfully ---")


def main():

    """Main function to find and process all ISO-EML-related INI files."""
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    log_file = setup_logging('../Log/', 'ISO-EML-Special_Init')
    Log.Log_Info(log_file, "===== ISO-EML-Special Script Start =====")

    ini_files = [f for f in os.listdir('.') if f.endswith('.ini')]
    if not ini_files:
        Log.Log_Info(log_file, "No .ini config files found in the current directory. Exiting.")
        print("No config files (.ini) found in the current directory.")
        return
    Log.Log_Info(log_file, f"Found {len(ini_files)} config file(s): {', '.join(ini_files)}")

    for ini_path in ini_files:
        try:
            print(f"--- Processing config: {ini_path} ---")
            config = _read_and_parse_ini_config(ini_path)
            settings = _extract_settings_from_config(config)
            
            log_file = setup_logging(settings.log_path, settings.operation)
            Log.Log_Info(log_file, f"--- Start processing config file: {ini_path} ---")
            
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
                    if not files:
                        continue
                    source_files_found = True
                    latest_file = max(files, key=os.path.getmtime)
                    Log.Log_Info(log_file, f"Found latest source file: {latest_file.name}")
                    try:
                        dst_path = shutil.copy(latest_file, intermediate_path)
                        Log.Log_Info(log_file, f"File copied successfully -> {dst_path}")
                        process_iso_eml_file(dst_path, settings, log_file, csv_filepath_for_this_ini)
                    except Exception:
                        Log.Log_Error(log_file, f"Error processing file {latest_file.name}: {traceback.format_exc()}")

            if not source_files_found:
                Log.Log_Info(log_file, "No matching source files found for this configuration.")

            if csv_filepath_for_this_ini and os.path.exists(csv_filepath_for_this_ini) and settings.output_path:
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

    Log.Log_Info(log_file, "===== ISO-EML-Special Script End =====")
    print("✅ All ISO-EML configurations have been processed.")
    
if __name__ == '__main__':
    main()
