# -*- coding: utf-8 -*-
import os, sys, shutil, logging, traceback
from pathlib import Path
from datetime import datetime, date
from configparser import ConfigParser
from dateutil.relativedelta import relativedelta
import numpy as np
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom

# keep original module layout
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
        self.tool_name = ""
        self.xy_sheet_name = ""
        self.xy_columns = ""
        self.tool_name_map = {}

def setup_logging(log_dir, operation_name):
    folder = os.path.join(log_dir, str(date.today()))
    os.makedirs(folder, exist_ok=True)
    log_file = os.path.join(folder, f'{operation_name}.log')
    for h in logging.root.handlers[:]:
        logging.root.removeHandler(h)
    logging.basicConfig(filename=log_file, level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')
    return log_file

def _read_and_parse_ini_config(path):
    cfg = ConfigParser()
    cfg.read(path, encoding='utf-8')
    return cfg

def _parse_fields_map_from_lines(lines):
    m = {}
    for line in lines:
        if ':' in line and not line.strip().startswith('#'):
            try:
                k, col, dt = map(str.strip, line.split(':', 2))
                m[k] = {'col': col, 'dtype': dt}
            except ValueError:
                pass
    return m

def _extract_settings_from_config(cfg):
    s = IniSettings()
    s.site = cfg.get('Basic_info', 'Site')
    s.product_family = cfg.get('Basic_info', 'ProductFamily')
    s.operation = cfg.get('Basic_info', 'Operation')
    s.test_station = cfg.get('Basic_info', 'TestStation')
    s.retention_date = cfg.getint('Basic_info', 'retention_date', fallback=30)
    s.file_name_patterns = [x.strip() for x in cfg.get('Basic_info', 'file_name_patterns').split(',')]
    s.tool_name = cfg.get('Basic_info', 'Tool_Name', fallback=None)

    s.input_paths = [x.strip() for x in cfg.get('Paths', 'input_paths').split(',')]
    s.output_path = cfg.get('Paths', 'output_path', fallback=None)
    s.csv_path = cfg.get('Paths', 'CSV_path', fallback=None)
    s.intermediate_data_path = cfg.get('Paths', 'intermediate_data_path')
    s.log_path = cfg.get('Paths', 'log_path')
    s.running_rec = cfg.get('Paths', 'running_rec')
    s.backup_running_rec_path = cfg.get('Paths', 'backup_running_rec_path', fallback=None)

    s.sheet_name = cfg.get('Excel', 'sheet_name')
    s.data_columns = cfg.get('Excel', 'data_columns')
    s.skip_rows = cfg.getint('Excel', 'main_skip_rows')
    s.xy_sheet_name = cfg.get('Excel', 'xy_sheet_name', fallback=None)
    s.xy_columns = cfg.get('Excel', 'xy_columns', fallback=None)

    fields_lines = cfg.get('DataFields', 'fields').splitlines()
    s.field_map = _parse_fields_map_from_lines(fields_lines)
    if cfg.has_section('ToolNameMapping'):
        s.tool_name_map = dict(cfg.items('ToolNameMapping'))
    return s

def write_to_csv(csv_filepath, df, log_file):
    Log.Log_Info(log_file, "Writing CSV...")
    try:
        df.to_csv(csv_filepath, index=False, encoding='utf-8-sig')
        Log.Log_Info(log_file, f"CSV written: {csv_filepath}")
        return True
    except Exception as e:
        Log.Log_Error(log_file, f"CSV write failed: {e}")
        return False

def generate_pointer_xml(output_path, csv_path, settings, log_file):
    Log.Log_Info(log_file, "Generating pointer XML...")
    try:
        os.makedirs(output_path, exist_ok=True)
        now_iso = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        serial_no = Path(csv_path).stem
        xml_name = (
            f"Site={settings.site},ProductFamily={settings.product_family},"
            f"Operation={settings.operation},Partnumber=UNKNOWPN,"
            f"Serialnumber={serial_no},Testdate={now_iso}.xml"
        ).replace(":", ".")
        xml_path = os.path.join(output_path, xml_name)

        root = ET.Element("Results", {"xmlns:xsi":"http://www.w3.org/2001/XMLSchema-instance","xmlns:xsd":"http://www.w3.org/2001/XMLSchema"})
        result = ET.SubElement(root, "Result", startDateTime=now_iso, endDateTime=now_iso, Result="Passed")
        ET.SubElement(result, "Header",
                      SerialNumber=serial_no, PartNumber="UNKNOWPN",
                      Operation=settings.operation, TestStation=settings.test_station,
                      Operator="NA", StartTime=now_iso, Site=settings.site, LotNumber="")
        ts = ET.SubElement(result, "TestStep", Name=settings.operation, startDateTime=now_iso, endDateTime=now_iso, Status="Passed")
        ET.SubElement(ts, "Data", DataType="Table", Name=f"tbl_{settings.operation.upper()}", Value=str(csv_path), CompOperation="LOG")
        xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="  ", encoding="utf-8")
        with open(xml_path, "wb") as f:
            f.write(xml_str)
        Log.Log_Info(log_file, f"Pointer XML written: {xml_path}")
    except Exception as e:
        Log.Log_Error(log_file, f"Pointer XML failed: {e}")

# ---------- helpers for DataFields ----------
def _normalize_indexing(indices, ncols):
    idx = [int(i) for i in indices if str(i).strip() not in ("", "-1")]
    if not idx: return 0, {}
    one_ok = all(1 <= i <= ncols for i in idx)
    zero_ok = all(0 <= i < ncols for i in idx)
    if one_ok:  return -1, {i: i-1 for i in idx}
    if zero_ok: return 0,  {i: i    for i in idx}
    fixed = {}; used = -1
    for i in idx:
        j = i - 1
        if 0 <= j < ncols: fixed[i] = j
    if not fixed:
        used = 0
        for i in idx:
            if 0 <= i < ncols: fixed[i] = i
    return used, fixed

def _col_from_map(df, settings, ini_key, pos_map, warns):
    ent = settings.field_map.get(ini_key)
    if not ent:
        warns.append(f"[WARN] Missing DataFields key: {ini_key}")
        return None
    col_str = ent['col'].strip()
    if col_str == "-1": return None
    try:
        ini_idx = int(col_str)
    except ValueError:
        warns.append(f"[WARN] Non-integer col for {ini_key}: {col_str}")
        return None
    if ini_idx not in pos_map:
        warns.append(f"[WARN] Out-of-range index for {ini_key}: {ini_idx}")
        return None
    return pos_map[ini_idx]

def _detect_position_col(df):
    valid = set(list("ABCD"))
    for c in df.columns:
        s = df[c].dropna().astype(str).str.strip()
        if len(s)==0: continue
        u = set(s.unique())
        if u.issubset(valid) and len(u) >= 2:
            return c
    return None

# ---------- SQL enrichment ----------
def _excel_serial(dt_series):
    origin = pd.Timestamp('1899-12-30')  # Excel serial date origin
    diff = (pd.to_datetime(dt_series, errors='coerce') - origin)
    return diff.dt.total_seconds() / 86400.0

def enrich_via_sql(df: pd.DataFrame, settings, log_file):
    """
    SQL-first for: STARTTIME_SORTED, SORTNUMBER, LotNumber_5, LotNumber_9.
    Lot numbers: if SQL cannot return reliable result, keep BLANK (no local fallback).
    """
    # STARTTIME_SORTED
    try:
        if hasattr(SQL, 'get_starttime_sorted'):
            s = SQL.get_starttime_sorted(df)
            if s is not None:
                df['STARTTIME_SORTED'] = s
                Log.Log_Info(log_file, "Filled STARTTIME_SORTED via SQL.get_starttime_sorted")
    except Exception as e:
        Log.Log_Error(log_file, f"SQL.get_starttime_sorted failed: {e}")
    finally:
        if 'STARTTIME_SORTED' not in df or df['STARTTIME_SORTED'].isna().all():
            # local safe fallback still allowed for STARTTIME_SORTED
            df['STARTTIME_SORTED'] = _excel_serial(df['Start_Date_Time'])

    # SORTNUMBER
    try:
        if hasattr(SQL, 'get_sortnumber'):
            s = SQL.get_sortnumber(df, settings.site, settings.operation)
            if s is not None:
                df['SORTNUMBER'] = s
                Log.Log_Info(log_file, "Filled SORTNUMBER via SQL.get_sortnumber")
    except Exception as e:
        Log.Log_Error(log_file, f"SQL.get_sortnumber failed: {e}")
    finally:
        if 'SORTNUMBER' not in df or df['SORTNUMBER'].isna().all():
            # local safe fallback still allowed for SORTNUMBER
            dt = pd.to_datetime(df['Start_Date_Time'], errors='coerce')
            df['SORTNUMBER'] = dt.rank(method='first').astype('Int64')

    # Lot numbers (strict SQL-only, else blank)
    lot5 = None
    lot9 = None
    try:
        if hasattr(SQL, 'get_lotnumbers'):
            lots = SQL.get_lotnumbers(df)
            if isinstance(lots, pd.DataFrame):
                if 'LotNumber_5' in lots.columns:
                    lot5 = lots['LotNumber_5']
                if 'LotNumber_9' in lots.columns:
                    lot9 = lots['LotNumber_9']
                Log.Log_Info(log_file, "Lot numbers loaded via SQL.get_lotnumbers")
    except Exception as e:
        Log.Log_Error(log_file, f"SQL.get_lotnumbers failed: {e}")

    # Write lot numbers only if SQL returns; otherwise keep blank
    if lot5 is not None:
        df['LotNumber_5'] = lot5
    else:
        df['LotNumber_5'] = ""  # blank

    if lot9 is not None:
        df['LotNumber_9'] = lot9
    else:
        df['LotNumber_9'] = ""  # blank

    return df

# ---------- DF building ----------
def df_expand_to_sample(df_raw: pd.DataFrame, settings: IniSettings, log_file) -> pd.DataFrame:
    ncols = df_raw.shape[1]
    ini_indices = []
    for _, v in settings.field_map.items():
        cs = str(v.get('col','')).strip()
        if cs not in ("", "-1"):
            try: ini_indices.append(int(cs))
            except: pass
    used_shift, pos_map = _normalize_indexing(ini_indices, ncols)
    print("[INFO] INI indices detected as 1-based -> shifted to 0-based." if used_shift==-1 else "[INFO] INI indices treated as 0-based.")
    Log.Log_Info(log_file, "Field indices normalized.")

    warns = []
    c_Start  = _col_from_map(df_raw, settings, 'key_Start_Date_Time', pos_map, warns)
    c_Oper   = _col_from_map(df_raw, settings, 'key_Operator', pos_map, warns)
    c_SN     = _col_from_map(df_raw, settings, 'key_Serial_Number', pos_map, warns)
    c_Step   = _col_from_map(df_raw, settings, 'key_Step_Step', pos_map, warns)
    c_ResVal = _col_from_map(df_raw, settings, 'key_Resist_Resist', pos_map, warns)
    c_ResAvg = _col_from_map(df_raw, settings, 'key_Resist_Resist_Ave', pos_map, warns)
    c_Y1Val  = _col_from_map(df_raw, settings, 'key_Y1_Y1', pos_map, warns)
    c_Y1Avg  = _col_from_map(df_raw, settings, 'key_Y1_Y1_Ave', pos_map, warns)
    c_Y2Val  = _col_from_map(df_raw, settings, 'key_Y2_Y2', pos_map, warns)
    c_Y2Avg  = _col_from_map(df_raw, settings, 'key_Y2_Y2_Ave', pos_map, warns)
    c_Part   = _col_from_map(df_raw, settings, 'key_Part_Number', pos_map, warns)
    c_Pos    = _col_from_map(df_raw, settings, 'key_Position', pos_map, warns)

    if c_Pos is None:
        c_Pos = _detect_position_col(df_raw)
        if c_Pos is None:
            raise RuntimeError("Cannot locate position column (A/B/C/D). Please add key_Position in [DataFields].")
        else:
            print(f"[INFO] Auto-detected position column at index {c_Pos} (A/B/C/D)")
            Log.Log_Info(log_file, f"Auto-detected position column={c_Pos}")

    # forward-fill identifiers and *_AVG from merged cells
    for c in [c_Start, c_Oper, c_SN, c_Part, c_Step, c_ResAvg, c_Y1Avg, c_Y2Avg]:
        if c is not None and c in df_raw.columns:
            df_raw[c] = df_raw[c].ffill()

    # keep only rows with A/B/C/D in position column
    pos_mask = df_raw[c_Pos].astype(str).str.strip().isin(list('ABCD'))
    pos_rows = df_raw[pos_mask].copy()
    if pos_rows.empty:
        print("[WARN] No rows with A/B/C/D found; returning empty frame.")
        return pd.DataFrame(columns=[
            'Serial_Number','Part_Number','Start_Date_Time','Operator',
            'Operation','TestStation',
            'Resist1','Resist2','Resist3','Resist4','Resist_AVG',
            'Step',
            'Y1_1','Y1_2','Y1_3','Y1_4','Y1_AVG',
            'Y2_1','Y2_2','Y2_3','Y2_4','Y2_AVG',
            'STARTTIME_SORTED','SORTNUMBER','LotNumber_5','LotNumber_9'
        ])

    # segmentation by Start_Date_Time row anchor: boundary at row ABOVE each Start row
    if c_Start is not None and c_Start in df_raw.columns:
        start_mark = df_raw[c_Start].notna()
        boundary = start_mark.shift(1, fill_value=True)  # first row treated as boundary
        record_id_full = boundary.cumsum()
        pos_rows['record_id'] = record_id_full.loc[pos_rows.index].values
    else:
        pos_rows['pos_chr'] = pos_rows[c_Pos].astype(str).str.strip()
        pos_rows['record_id'] = (pos_rows['pos_chr'] == 'A').cumsum()

    pos_idx = {'A':1,'B':2,'C':3,'D':4}
    pos_rows['pos_i'] = pos_rows[c_Pos].astype(str).str.strip().map(pos_idx).astype('Int64')

    def first_valid_or(series, default=np.nan):
        if series is None: return default
        s = series.dropna()
        return s.iloc[0] if len(s) else default

    def series_by_pos(group, value_col):
        if value_col is None or value_col not in group.columns:
            return pd.Series(dtype=float)
        g = group[['pos_i', value_col]].dropna(subset=['pos_i']).copy()
        if g.empty: return pd.Series(dtype=float)
        def first_non_null(s):
            s2 = s.dropna()
            return s2.iloc[0] if len(s2) else np.nan
        s = g.groupby('pos_i', as_index=True)[value_col].apply(first_non_null)
        try: s.index = s.index.astype(int)
        except: pass
        return s

    def assemble(group):
        out = {}
        # Resist
        res_vals = series_by_pos(group, c_ResVal)
        res_avg  = first_valid_or(group[c_ResAvg] if c_ResAvg in group.columns else None, np.nan)
        for i in (1,2,3,4):
            v = res_vals.get(i, np.nan)
            out[f'Resist{i}'] = res_avg if (pd.isna(v) and pd.notna(res_avg)) else v
        out['Resist_AVG'] = res_avg
        # Y1
        y1_vals = series_by_pos(group, c_Y1Val)
        y1_avg  = first_valid_or(group[c_Y1Avg] if c_Y1Avg in group.columns else None, np.nan)
        for i in (1,2,3,4):
            v = y1_vals.get(i, np.nan)
            out[f'Y1_{i}'] = y1_avg if (pd.isna(v) and pd.notna(y1_avg)) else v
        out['Y1_AVG'] = y1_avg
        # Y2
        y2_vals = series_by_pos(group, c_Y2Val)
        y2_avg  = first_valid_or(group[c_Y2Avg] if c_Y2Avg in group.columns else None, np.nan)
        for i in (1,2,3,4):
            v = y2_vals.get(i, np.nan)
            out[f'Y2_{i}'] = y2_avg if (pd.isna(v) and pd.notna(y2_avg)) else v
        out['Y2_AVG'] = y2_avg
        # Others
        out['Serial_Number'] = first_valid_or(group[c_SN]   if c_SN   in group.columns else None, None)
        out['Part_Number']   = first_valid_or(group[c_Part] if c_Part in group.columns else None, None)
        out['Operator']      = first_valid_or(group[c_Oper] if c_Oper in group.columns else None, None)
        sdt_raw = first_valid_or(group[c_Start] if c_Start in group.columns else None, None)
        sdt = pd.to_datetime(sdt_raw, errors='coerce')
        if pd.notna(sdt):
            out['Start_Date_Time'] = sdt.strftime('%Y/%-m/%-d %H:%M') if sys.platform!='win32' else sdt.strftime('%Y/%#m/%#d %H:%M')
        else:
            out['Start_Date_Time'] = '' if (sdt_raw is None or (isinstance(sdt_raw,float) and np.isnan(sdt_raw))) else str(sdt_raw)
        out['Step'] = first_valid_or(group[c_Step] if (c_Step is not None and c_Step in group.columns) else None, None)
        out['Operation']   = settings.operation
        out['TestStation'] = settings.test_station
        return pd.Series(out)

    recs = [assemble(g) for _, g in pos_rows.groupby('record_id', dropna=False)]
    wide = pd.DataFrame(recs).drop_duplicates(subset=['Serial_Number','Start_Date_Time'], keep='first')

    final_cols = [
        'Serial_Number','Part_Number','Start_Date_Time','Operator',
        'Operation','TestStation',
        'Resist1','Resist2','Resist3','Resist4','Resist_AVG',
        'Step',
        'Y1_1','Y1_2','Y1_3','Y1_4','Y1_AVG',
        'Y2_1','Y2_2','Y2_3','Y2_4','Y2_AVG',
        'STARTTIME_SORTED','SORTNUMBER','LotNumber_5','LotNumber_9'
    ]
    result = wide.reindex(columns=final_cols)

    # SQL-first enrichment (Lot numbers are SQL-only; blank if not found)
    result = enrich_via_sql(result, settings, log_file)

    if warns:
        print("\n=== DataFields vs Excel mismatch warnings ===")
        for m in warns: print(m)
        print("============================================\n")
        Log.Log_Info(log_file, " | ".join(warns))
    return result

# ---------- main pipeline ----------
def process_excel_file(filepath_str, settings, log_file, csv_filepath):
    filepath = Path(filepath_str)
    Log.Log_Info(log_file, f"Start file: {filepath.name}")
    start_row = max(Row_Number_Func.start_row_number(settings.running_rec) - settings.skip_rows, 4)
    start_row = int(20)  # keep original behavior

    try:
        df = pd.read_excel(filepath, header=None, sheet_name=settings.sheet_name,
                           usecols=settings.data_columns, skiprows=start_row)
        Log.Log_Info(log_file, f"Main sheet read: '{settings.sheet_name}', rows={df.shape[0]}.")

        df_out = df_expand_to_sample(df, settings, log_file)

        # retention filter
        cut_dt = datetime.now() - relativedelta(days=settings.retention_date)
        ts = pd.to_datetime(df_out['Start_Date_Time'], errors='coerce', infer_datetime_format=True)
        before = len(df_out)
        df_out = df_out[ts >= cut_dt].reset_index(drop=True)
        after = len(df_out)
        print(f"[INFO] retention_date={settings.retention_date} days; kept >= {cut_dt:%Y-%m-%d %H:%M:%S}. rows {before} -> {after}")
        Log.Log_Info(log_file, f"Retention filter kept >= {cut_dt:%Y-%m-%d %H:%M:%S}. rows {before}->{after}")

        # CSV + Pointer XML
        if csv_filepath:
            ok = write_to_csv(str(csv_filepath), df_out, log_file)
            if ok and settings.output_path:
                generate_pointer_xml(settings.output_path, str(csv_filepath), settings, log_file)
        else:
            print("[WARN] CSV_path not configured; skip CSV/XML.")
            Log.Log_Info(log_file, "CSV_path not configured; skipped CSV/XML.")

    except Exception as e:
        Log.Log_Error(log_file, f"DF/IO failed: {e}")
        print(f"[ERROR] DF/IO failed for {filepath.name}: {e}")

def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    log_file = setup_logging('../Log/', 'UniversalScript_Init')
    Log.Log_Info(log_file, "===== Universal Script Start =====")

    ini_files = [f for f in os.listdir('.') if f.endswith('.ini')]
    if not ini_files:
        Log.Log_Info(log_file, "No INI found.")
        print("No INI found in current directory.")
        return
    Log.Log_Info(log_file, f"Found {len(ini_files)} INI: {', '.join(ini_files)}")

    for ini_path in ini_files:
        try:
            print(f"--- Processing config: {ini_path} ---")
            cfg = _read_and_parse_ini_config(ini_path)
            settings = _extract_settings_from_config(cfg)

            log_file = setup_logging(settings.log_path, settings.operation)
            Log.Log_Info(log_file, f"Begin INI: {ini_path}")

            csv_target = None
            if settings.csv_path:
                Path(settings.csv_path).mkdir(parents=True, exist_ok=True)
                ts = datetime.now().strftime('%Y_%m_%dT%H.%M.%S')
                csv_target = Path(settings.csv_path) / f"{settings.operation}_{ts}.csv"
                Log.Log_Info(log_file, f"CSV target: {csv_target}")

            interm = Path(settings.intermediate_data_path); interm.mkdir(parents=True, exist_ok=True)
            found = False
            for p_str in settings.input_paths:
                p = Path(p_str)
                for pat in settings.file_name_patterns:
                    Log.Log_Info(log_file, f"Search path='{p}' pattern='{pat}'")
                    files = [x for x in p.glob(pat) if not x.name.startswith('~$')]
                    if not files: continue
                    found = True
                    latest = max(files, key=os.path.getmtime)
                    Log.Log_Info(log_file, f"Latest source: {latest.name}")
                    try:
                        dst = shutil.copy(latest, interm)
                        Log.Log_Info(log_file, f"Copied to: {dst}")
                        process_excel_file(dst, settings, log_file, csv_target)
                    except Exception:
                        Log.Log_Error(log_file, f"Process error {latest.name}: {traceback.format_exc()}")

            if not found:
                Log.Log_Info(log_file, "No matching source file for this INI.")
            Log.Log_Info(log_file, f"End INI: {ini_path}")
        except Exception:
            err = f"FATAL INI {ini_path}: {traceback.format_exc()}"
            print(err)
            if log_file: Log.Log_Error(log_file, err)

    Log.Log_Info(log_file, "===== Universal Script End =====")
    print("✅ All INIs processed (CSV & XML enabled).")

if __name__ == '__main__':
    main()
