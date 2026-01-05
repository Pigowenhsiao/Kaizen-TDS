import os
import sys
import glob
import shutil
import logging
import numpy as np
import pandas as pd
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
from configparser import ConfigParser
from pathlib import Path
import traceback

# Import necessary modules from MyModule
# Ensure MyModule is in the python path
sys.path.append('../MyModule')
import Log
import SQL
import Convert_Date
import Row_Number_Func

class IniSettings:
    """Class to hold all settings read from the INI file."""
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
        self.intermediate_data_path = ""
        self.log_path = ""
        self.running_rec = ""
        self.backup_running_rec_path = ""
        self.sheet_name = ""
        self.data_columns = ""
        self.skip_rows = 500
        self.key_col_idx = 4
        self.serial_col_idx = 4
        self.field_map = {}

def setup_logging(log_dir):
    """Sets up logging for the script execution."""
    log_folder = os.path.join(log_dir, str(date.today()))
    os.makedirs(log_folder, exist_ok=True)
    log_file = os.path.join(log_folder, '012_MESA_CVD.log')
    # Clear previous handlers
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    logging.basicConfig(filename=log_file, level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')
    return log_file

def _read_and_parse_ini_config(config_file_path):
    """Reads and parses the INI configuration file."""
    config = ConfigParser(interpolation=None)
    config.read(config_file_path, encoding='utf-8')
    return config

def _parse_fields_map_from_lines(fields_lines):
    """Parses the field mapping from the [DataFields] section."""
    fields = {}
    for line in fields_lines:
        if ':' in line and not line.strip().startswith('#'):
            try:
                key, col_str, dtype_str = map(str.strip, line.split(':', 2))
                
                if dtype_str == 'float':
                    dtype = float
                elif dtype_str == 'str':
                    dtype = str
                elif dtype_str == 'datetime':
                    dtype = 'datetime' # Use string marker for datetime
                else: 
                    dtype = dtype_str
                
                fields[key] = {'col': col_str, 'type': dtype}
            except ValueError:
                continue
    return fields

def _extract_settings_from_config(config):
    """Extracts all settings from the parsed config object into an IniSettings instance."""
    s = IniSettings()
    # Basic Info
    s.site = config.get('Basic_info', 'Site')
    s.product_family = config.get('Basic_info', 'ProductFamily')
    s.operation = config.get('Basic_info', 'Operation')
    s.test_station = config.get('Basic_info', 'TestStation')
    s.tool_name = config.get('Basic_info', 'Tool_Name')
    s.retention_date = config.getint('Basic_info', 'retention_date')
    s.file_name_patterns = [x.strip() for x in config.get('Basic_info', 'file_name_patterns').split(',')]
    
    # Paths
    s.input_paths = [x.strip() for x in config.get('Paths', 'input_paths').split(',')]
    s.output_path = config.get('Paths', 'output_path')
    s.intermediate_data_path = config.get('Paths', 'intermediate_data_path')
    s.log_path = config.get('Paths', 'log_path')
    s.running_rec = config.get('Paths', 'running_rec')
    s.backup_running_rec_path = config.get('Paths', 'backup_running_rec_path', fallback=None)

    # Excel
    s.sheet_name = config.get('Excel', 'sheet_name')
    s.data_columns = config.get('Excel', 'data_columns')
    s.skip_rows = config.getint('Excel', 'main_skip_rows')
    s.key_col_idx = config.getint('Excel', 'main_dropna_key_col_idx')
    s.serial_col_idx = config.getint('Excel', 'serial_number_source_column_idx')
    
    # DataFields
    fields_lines = config.get('DataFields', 'fields').splitlines()
    s.field_map = _parse_fields_map_from_lines(fields_lines)
    return s

def data_type_check(key_to_type, data_dict, log_file, serial):
    """Performs data type validation for the extracted data."""
    for key, expected_type in key_to_type.items():
        if key not in data_dict:
            continue
        value = data_dict.get(key)
        if pd.isna(value) or value is None:
            Log.Log_Error(log_file, f"{serial} : DATA_TYPE_ERROR: Value for key '{key}' is empty.")
            return False
        
        try:
            if expected_type is float:
                float(value)
            elif expected_type == 'datetime':
                # The format is already validated during conversion
                pass
            elif expected_type is str:
                str(value)
        except (ValueError, TypeError) as e:
            Log.Log_Error(log_file, f"{serial} : DATA_TYPE_ERROR: Value '{value}' for key '{key}' cannot be processed as {expected_type}. Error: {e}")
            return False
    return True

def generate_xml(output_path, site, product_family, operation, test_station, data_dict, log_file):
    """Generates the XML file from the data dictionary."""
    # For XML content, use the standard "YYYY-MM-DD HH:MM:SS" format
    start_time_str = data_dict.get('key_Start_Date_Time', '')
    
    # For the filename, use the original non-standard format to match the old script
    filename_date_str = data_dict.get('key_Start_Date_Time_For_Filename', '')

    serial_number = data_dict.get('key_Serial_Number', 'NA')
    part_number = data_dict.get('key_Part_Number', 'NA')
    
    # Use filename_date_str for the filename, as per the original script's behavior
    filename = f"Site={site},ProductFamily={product_family},Operation={operation},Partnumber={part_number},Serialnumber={serial_number},Testdate={filename_date_str}.xml"
    filepath = Path(output_path) / filename
    
    try:
        with open(filepath, 'w', encoding="utf-8") as f:
            # The XML content uses the clean, standard 'start_time_str'
            f.write('<?xml version="1.0" encoding="utf-8"?>\n')
            f.write('<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">\n')
            f.write(f'  <Result startDateTime="{start_time_str}" Result="Passed">\n')
            f.write(f'    <Header SerialNumber="{serial_number}" PartNumber="{part_number}" Operation="{operation}" TestStation="{test_station}" Operator="{data_dict.get("key_Operator", "")}" StartTime="{start_time_str}" Site="{site}" LotNumber="{serial_number}"/>\n')
            f.write(f'    <TestStep Name="Thickness1" startDateTime="{start_time_str}" Status="Passed">\n')
            f.write(f'      <Data DataType="Numeric" Name="Initial1" Units="nm" Value="{data_dict.get("key_Initial1", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Initial2" Units="nm" Value="{data_dict.get("key_Initial2", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Initial3" Units="nm" Value="{data_dict.get("key_Initial3", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Initial4" Units="nm" Value="{data_dict.get("key_Initial4", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Initial5" Units="nm" Value="{data_dict.get("key_Initial5", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Initial_Ave" Units="nm" Value="{data_dict.get("key_Initial_ave", "")}"/>\n')
            f.write(f'    </TestStep>\n')
            f.write(f'    <TestStep Name="Thickness2" startDateTime="{start_time_str}" Status="Passed">\n')
            f.write(f'      <Data DataType="Numeric" Name="Final1" Units="nm" Value="{data_dict.get("key_Final1", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Final2" Units="nm" Value="{data_dict.get("key_Final2", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Final3" Units="nm" Value="{data_dict.get("key_Final3", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Final4" Units="nm" Value="{data_dict.get("key_Final4", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Final5" Units="nm" Value="{data_dict.get("key_Final5", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Final_Ave" Units="nm" Value="{data_dict.get("key_Final_ave", "")}"/>\n')
            f.write(f'    </TestStep>\n')
            f.write(f'    <TestStep Name="Rate" startDateTime="{start_time_str}" Status="Passed">\n')
            f.write(f'      <Data DataType="Numeric" Name="Rate1" Units="nm/min" Value="{data_dict.get("key_Rate1", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Rate2" Units="nm/min" Value="{data_dict.get("key_Rate2", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Rate3" Units="nm/min" Value="{data_dict.get("key_Rate3", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Rate4" Units="nm/min" Value="{data_dict.get("key_Rate4", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Rate5" Units="nm/min" Value="{data_dict.get("key_Rate5", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Rate_Ave" Units="nm/min" Value="{data_dict.get("key_Rate_ave", "")}"/>\n')
            f.write(f'      <Data DataType="Numeric" Name="Rate_3sigma" Units="nm/min" Value="{data_dict.get("key_Rate_3sigma", "")}"/>\n')
            f.write(f'    </TestStep>\n')
            f.write(f'    <TestStep Name="Time" startDateTime="{start_time_str}" Status="Passed">\n')
            f.write(f'      <Data DataType="Numeric" Name="Time" Units="min" Value="{data_dict.get("key_Time", "")}"/>\n')
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
            f.write(f'      <Data DataType="String" Name="LotNumber_5" Value="{serial_number}" CompOperation="LOG"/>\n')
            f.write(f'      <Data DataType="String" Name="LotNumber_9" Value="{data_dict.get("key_LotNumber_9", "")}" CompOperation="LOG"/>\n')
            f.write(f'    </TestStep>\n')
            f.write(f'    <TestEquipment>\n')
            f.write(f'      <Item DeviceName="Nanospec" DeviceSerialNumber="{data_dict.get("key_TestEquipment_Nano", "")}"/>\n')
            f.write(f'      <Item DeviceName="DryEtch" DeviceSerialNumber="{data_dict.get("key_TestEquipment_Dry", "")}"/>\n')
            f.write(f'    </TestEquipment>\n')
            f.write('    <ErrorData/>\n')
            f.write('    <FailureData/>\n')
            f.write('    <Configuration/>\n')
            f.write('  </Result>\n')
            f.write('</Results>')
        Log.Log_Info(log_file, f"XML file written: {filepath}")
        return True
    except Exception as e:
        Log.Log_Error(log_file, f"Error writing XML for {serial_number}: {e}")
        return False

def process_excel_file(filepath, settings, log_file):
    """Main processing logic for a single Excel file."""
    Log.Log_Info(log_file, f"Processing file: {filepath.name}")
    start_row = max(Row_Number_Func.start_row_number(settings.running_rec) - settings.skip_rows, 4)

    try:
        df = pd.read_excel(filepath, header=None, sheet_name=settings.sheet_name, usecols=settings.data_columns, skiprows=start_row)
        df.columns = range(df.shape[1])
    except Exception as e:
        Log.Log_Error(log_file, f"Failed to read Excel file {filepath.name}. Error: {e}")
        return

    # Date filtering
    try:
        date_col_idx = int(settings.field_map['key_Start_Date_Time']['col'])
        # Convert date column, coercing errors to NaT (Not a Time)
        date_series = pd.to_datetime(df.iloc[:, date_col_idx], errors='coerce')
        
        # Keep only rows where date is not NaT
        df = df[date_series.notna()]
        date_series = date_series[date_series.notna()] # filter NaTs from the series too

        # Calculate cutoff date
        cutoff_date = datetime.now() - relativedelta(days=settings.retention_date)
        
        # Keep rows newer than or equal to cutoff
        df = df[date_series >= cutoff_date]
        Log.Log_Info(log_file, f"After date filtering, {len(df)} records remain in {filepath.name}.")
    except Exception as e:
        Log.Log_Error(log_file, f"Date filtering failed for {filepath.name}. Check 'key_Start_Date_Time' settings. Error: {e}")
        return

    # Drop rows where the key column is NaN
    df.dropna(subset=[settings.key_col_idx], inplace=True)
    
    # Process each valid row
    for idx in df.index:
        data_dict = {}
        row_data = df.loc[idx]

        # Populate data_dict from excel row based on field_map
        for key, mapping in settings.field_map.items():
            col_str = mapping['col']
            if col_str == '-1':
                continue
            try:
                col_index = int(col_str)
                data_dict[key] = row_data.iloc[col_index]
            except (ValueError, IndexError):
                data_dict[key] = None
                Log.Log_Error(log_file, f"Failed to read column for key '{key}' (index: {col_str})")

        serial = str(data_dict.get('key_Serial_Number', ''))
        if not serial or serial == 'nan':
            continue

        # --- DB Lookup ---
        conn, cursor = SQL.connSQL()
        if conn is None:
            Log.Log_Error(log_file, f"{serial} : Connection with Prime Failed")
            continue
        part_number, lot9 = SQL.selectSQL(cursor, serial)
        SQL.disconnSQL(conn, cursor)
        
        if not part_number or part_number == 'LDアレイ_':
            Log.Log_Info(log_file, f"Skipping {serial}: PartNumber is '{part_number}'")
            continue
            
        # --- Populate internally generated fields ---
        data_dict['key_Part_Number'] = part_number
        data_dict['key_LotNumber_9'] = lot9
        data_dict['key_TestEquipment_Dry'] = settings.tool_name

        # --- Date Handling & Validation ---
        try:
            raw_date = data_dict.get('key_Start_Date_Time')
            
            # 1. Use the custom module to get the initial (non-standard) date string
            non_standard_date_str = Convert_Date.Edit_Date(raw_date)

            # 2. Store this non-standard version specifically for the filename
            data_dict['key_Start_Date_Time_For_Filename'] = non_standard_date_str
            
            # 3. Clean and parse the non-standard string into a datetime object
            clean_datetime_str = non_standard_date_str.replace('T', ' ').replace('.', ':')
            date_obj = datetime.strptime(clean_datetime_str, '%Y-%m-%d %H:%M:%S')

            # 4. Format it back to the desired standard "YYYY-MM-DD HH:MM:SS" string
            standard_date_str = date_obj.strftime('%Y-%m-%d %H:%M:%S')

            # 5. Store the standard format for XML content and further processing
            data_dict['key_Start_Date_Time'] = standard_date_str

        except Exception as e:
            Log.Log_Error(log_file, f"{serial} : Date conversion failed. Raw date: {raw_date}. Error: {e}")
            continue

        # --- Type Checking ---
        key_to_type_map = {k: v['type'] for k, v in settings.field_map.items() if v['col'] != '-1'}
        if not data_type_check(key_to_type_map, data_dict, log_file, serial):
            continue

        # --- Calculate SORTED fields ---
        try:
            # The date string is now guaranteed to be in the correct format.
            date_obj = datetime.strptime(data_dict["key_Start_Date_Time"], "%Y-%m-%d %H:%M:%S")
            date_excel_number = int(str(date_obj - datetime(1899, 12, 30)).split()[0])
            excel_row = start_row + idx + 1
            excel_row_div = excel_row / 10**6
            data_dict["key_STARTTIME_SORTED"] = date_excel_number + excel_row_div
            data_dict["key_SORTNUMBER"] = excel_row
        except Exception as e:
            Log.Log_Error(log_file, f"{serial} : SORTED field calculation failed. Error: {e}")
            continue

        # --- Generate XML ---
        if generate_xml(settings.output_path, settings.site, settings.product_family, settings.operation, settings.test_station, data_dict, log_file):
            Log.Log_Info(log_file, f"{serial} : OK")
        else:
            Log.Log_Error(log_file, f"{serial} : XML Generation Failed")
            
    # --- Update running record ---
    next_start_row = start_row + df.shape[0] + 1
    Row_Number_Func.next_start_row_number(settings.running_rec, next_start_row)
    Log.Log_Info(log_file, f"Next start row for {settings.running_rec} set to {next_start_row}")

    if settings.backup_running_rec_path:
        try:
            backup_dir = Path(settings.backup_running_rec_path)
            backup_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy(settings.running_rec, backup_dir)
            Log.Log_Info(log_file, f"Backed up {settings.running_rec} to {backup_dir}")
        except Exception as e:
            Log.Log_Error(log_file, f"Failed to backup running_rec file. Error: {e}")


def main():
    """Main function to find and process INI files."""
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    ini_files = [f for f in os.listdir('.') if f.endswith('.ini')]
    if not ini_files:
        print("No .ini files found in the current directory.")
        return

    for ini_path in ini_files:
        log_file = None
        try:
            print(f"--- Processing config: {ini_path} ---")
            config = _read_and_parse_ini_config(ini_path)
            settings = _extract_settings_from_config(config)
            log_file = setup_logging(settings.log_path)
            Log.Log_Info(log_file, f"Start processing INI config file: {ini_path}")
            
            intermediate_path = Path(settings.intermediate_data_path)
            intermediate_path.mkdir(parents=True, exist_ok=True)

            source_files_found = False
            for input_p_str in settings.input_paths:
                input_p = Path(input_p_str)
                for pattern in settings.file_name_patterns:
                    # Find the most recently modified file that matches the pattern
                    matched_files = [p for p in input_p.glob(pattern) if not p.name.startswith('~$')]
                    if not matched_files:
                        continue
                    
                    source_files_found = True
                    latest_file = max(matched_files, key=os.path.getmtime)
                    Log.Log_Info(log_file, f"Found latest source file: {latest_file}")
                    
                    dst_path = intermediate_path / latest_file.name
                    try:
                        shutil.copy(latest_file, dst_path)
                        Log.Log_Info(log_file, f"Copied {latest_file.name} to intermediate directory.")
                        process_excel_file(dst_path, settings, log_file)
                    except Exception as e:
                        tb_str = traceback.format_exc()
                        Log.Log_Error(log_file, f"Error processing file {latest_file.name}: {e}\n{tb_str}")

            if not source_files_found:
                Log.Log_Info(log_file, "No matching source files found for this configuration.")

            Log.Log_Info(log_file, f"Finished processing INI file: {ini_path}\n")

        except Exception as e:
            tb_str = traceback.format_exc()
            error_message = f"FATAL Error processing INI file {ini_path}: {e}\n{tb_str}"
            print(error_message)
            if log_file:
                Log.Log_Error(log_file, error_message)

    print("✅ All .ini configurations have been processed.")

if __name__ == '__main__':
    main()