# -*- coding: utf-8 -*-
import os
import sys
import glob
import shutil
import logging
import numpy as np
import pandas as pd
import configparser
import traceback
from time import strftime, localtime
from datetime import date, timedelta, datetime, time
from dateutil.relativedelta import relativedelta

# --- Try importing necessary libraries ---
try:
    import xlrd # Keep for compatibility if old .xls files might exist
    import pyodbc # Keep if SQL module uses it
    import openpyxl as px # Keep if pandas uses it explicitly or for direct manipulation
except ImportError as e:
    print(f"Error importing a required library: {e}")
    print("Please ensure pandas, numpy, pyodbc, openpyxl, xlrd are installed.")
    sys.exit(1)

# -----------------------------------------------------------------------------
# Helper: Import Custom Modules
# -----------------------------------------------------------------------------
# Assume MyModule is in the parent directory relative to this script
try:
    script_dir_for_module = os.path.dirname(__file__)
    module_path = os.path.abspath(os.path.join(script_dir_for_module, '..', 'MyModule'))
    if module_path not in sys.path:
        sys.path.append(module_path)

    import Log
    import SQL
    import Check
    import Convert_Date
    import Row_Number_Func
    print(f"Successfully imported custom modules from {module_path}")
except ImportError as e:
    print(f"Error importing custom modules from {module_path}: {e}")
    print("Please ensure 'MyModule' directory exists relative to the script and contains necessary .py files.")
    sys.exit(1)
except NameError:
    print("Error: Could not determine script directory to find 'MyModule'.")
    try:
        import Log, SQL, Check, Convert_Date, Row_Number_Func
        print("Attempted direct import of custom modules.")
    except ImportError:
        print("Direct import failed. Ensure MyModule is accessible.")
        sys.exit(1)

# -----------------------------------------------------------------------------
# Configuration Loading
# -----------------------------------------------------------------------------
def load_config(config_filename='PIX_Config.ini'):
    """Loads configuration file."""
    config = configparser.ConfigParser()
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        script_dir = os.getcwd()
        print(f"Warning: Could not determine script directory via __file__, using current working directory: {script_dir}")

    config_path = os.path.join(script_dir, config_filename)
    if not os.path.exists(config_path):
        print(f"Error: Config file not found at {config_path}")
        return None, None
    try:
        config.read(config_path, encoding='utf-8')
        if not config.sections():
             print(f"Error: Config file {config_path} is empty or not parsed correctly.")
             return None, None
        print(f"Successfully loaded config from {config_path}")
        return config, script_dir
    except Exception as e:
        print(f"Error reading config file {config_path}: {e}")
        return None, None

# -----------------------------------------------------------------------------
# Logging Setup
# -----------------------------------------------------------------------------
def setup_logging(config, script_dir):
    """Sets up logging based on configuration."""
    if not config or not script_dir:
        print("Error: Cannot setup logging due to missing config or script_dir.")
        return None
    try:
        log_base_dir_rel = config.get('General', 'LogBaseDir', fallback='../Log/')
        log_file_name = config.get('General', 'LogFileName', fallback='014_PIX.log')
        log_base_dir_abs = os.path.abspath(os.path.join(script_dir, log_base_dir_rel))
    except (configparser.NoSectionError, configparser.NoOptionError) as e:
        print(f"Error reading logging configuration: {e}")
        return None

    log_folder_name = str(date.today())
    log_dir_path = os.path.join(log_base_dir_abs, log_folder_name)

    try:
        if not os.path.exists(log_dir_path):
            os.makedirs(log_dir_path)
            print(f"Created log directory: {log_dir_path}")
        log_file_path = os.path.join(log_dir_path, log_file_name)
        print(f"Logging to: {log_file_path}")
        Log.Log_Info(log_file_path, '--- Program Start ---')
        return log_file_path
    except Exception as e:
        print(f"Error setting up logging to {log_dir_path}: {e}")
        return None

# -----------------------------------------------------------------------------
# Constants and Global Variables (Initialized in main block)
# -----------------------------------------------------------------------------
# Populated after loading config
SITE = ''
PRODUCT_FAMILY = ''
OPERATION = ''
TEST_STATION = ''
DATA_SHEET_NAME = ''
OUTPUT_FILEPATH = ''
LOCAL_DATA_FILE_DIR_ABS = ''
START_ROW_FILE_BASE_DIR = '' # Optional backup directory

# Data structures used during processing
PIX_Data_List = list()
List_Index_Lot = dict()
Next_StartNumber = [0] * 3 # To store next start row for PIX1, PIX2, PIX3

# Data type definition for validation
KEY_TYPE_DEF = {
    'key_Part_Number' : str, 'key_Serial_Number' : str, 'key_LotNumber_9': str,
    'key_PIX1_Start_Date_Time' : str, 'key_PIX1_Operator' : str, 'key_PIX1_Equipment' : str,
    'key_PIX1_Step1' : float, 'key_PIX1_Step2' : float, 'key_PIX1_Step3' : float, 'key_PIX1_Step_Ave' : float, 'key_PIX1_Step_3sigma' : float,
    'key_PIX2_Start_Date_Time' : str, 'key_PIX2_Operator' : str, 'key_PIX2_Equipment' : str,
    'key_PIX2_Step1' : float, 'key_PIX2_Step2' : float, 'key_PIX2_Step3' : float, 'key_PIX2_Step_Ave' : float, 'key_PIX2_Step_3sigma' : float,
    'key_PIX3_Start_Date_Time' : str, 'key_PIX3_Operator' : str, 'key_PIX3_Equipment' : str,
    'key_PIX3_Step1' : float, 'key_PIX3_Step2' : float, 'key_PIX3_Step3' : float, 'key_PIX3_Step_Ave' : float, 'key_PIX3_Step_3sigma' : float,
    "key_STARTTIME_SORTED_PIX1" : float, "key_SORTNUMBER_PIX1" : float
}
# Indices for PIX_Data_List (assuming original structure)
# [P1_Date(0), P1_Op(1), SN(2), Part#(3), Lot9(4), P1_Eq(5), P1_S1-S5(6-10),
#  P2_Date(11), P2_Op(12), P2_Eq(13), P2_S1-S5(14-18),
#  P3_Date(19), P3_Op(20), P3_Eq(21), P3_S1-S5(22-26),
#  P1_Row(27), P2_Row(28), P3_Row(29)]
IDX_SN = 2
IDX_P1_DATE = 0
IDX_P2_DATE = 11
IDX_P3_DATE = 19
IDX_P1_ROW = 27
INVALID_MARKER = "INVALID_WRITE_ERROR" # Marker for write errors

# -----------------------------------------------------------------------------
# Helper Function: Get PIX Specific Configuration
# -----------------------------------------------------------------------------
def _get_pix_config(PIX, config, script_dir, log_file):
    """Gets PIX-specific settings from config."""
    try:
        pix_config = {
            'source_path': config.get(PIX, 'SourceFilePath'),
            'file_pattern': config.get(PIX, 'SourceFileNamePattern'),
            'start_row_file': os.path.abspath(os.path.join(script_dir, config.get(PIX, 'StartRowFile'))),
            'date_col': config.getint('Columns', 'PIX_Start_Date_Time', fallback=0),
            'op_col': config.getint('Columns', 'PIX_Operator'),
            'eq_col': config.getint('Columns', 'PIX_Equipment'),
            'sn_col': config.getint('Columns', 'PIX_Serial_Number'),
            'step_cols': [int(x.strip()) for x in config.get('Columns', f'{PIX}_Step').split(',')]
        }
        # Calculate required columns for blank check (relative to C:X slice)
        pix_config['required_cols'] = [pix_config['date_col'], pix_config['sn_col'], pix_config['op_col'], pix_config['eq_col']] + pix_config['step_cols']
        return pix_config
    except (configparser.NoSectionError, configparser.NoOptionError, ValueError) as e:
        Log.Log_Error(log_file, f"Configuration error for {PIX}: {e}")
        return None

# -----------------------------------------------------------------------------
# Helper Function: Find and Copy Excel
# -----------------------------------------------------------------------------
def _find_and_copy_excel(pix_config, log_file):
    """Finds the latest Excel file and copies it locally."""
    Log.Log_Info(log_file, f'Searching for Excel files...')
    excel_file_list = []
    try:
        search_pattern = os.path.join(pix_config['source_path'], pix_config['file_pattern'])
        found_files = glob.glob(search_pattern)
        for file in found_files:
            if not os.path.basename(file).startswith('~$') and not os.path.basename(file).startswith('.') and os.path.isfile(file):
                try:
                    dt = strftime("%Y-%m-%d %H:%M:%S", localtime(os.path.getmtime(file)))
                    excel_file_list.append([file, dt])
                except FileNotFoundError:
                     Log.Log_Error(log_file, f"File disappeared between glob and getmtime: {file}")
                     pass
    except Exception as e:
        Log.Log_Error(log_file, f"Error searching for Excel files at {search_pattern}: {e}")
        return None

    if not excel_file_list:
        Log.Log_Error(log_file, f"No valid Excel files found matching pattern: {search_pattern}")
        return None

    excel_file_list = sorted(excel_file_list, key=lambda x: x[1], reverse=True)
    source_excel_path = excel_file_list[0][0]

    try:
        if not os.path.exists(LOCAL_DATA_FILE_DIR_ABS):
            os.makedirs(LOCAL_DATA_FILE_DIR_ABS)
            Log.Log_Info(log_file, f"Created local data directory: {LOCAL_DATA_FILE_DIR_ABS}")
        local_excel_file = os.path.join(LOCAL_DATA_FILE_DIR_ABS, os.path.basename(source_excel_path))
        shutil.copy(source_excel_path, local_excel_file)
        Log.Log_Info(log_file, f"Successfully copied {source_excel_path} to {local_excel_file}")
        return local_excel_file
    except Exception as e:
        Log.Log_Error(log_file, f"Error copying Excel file from {source_excel_path} to {LOCAL_DATA_FILE_DIR_ABS}: {e}")
        return None

# -----------------------------------------------------------------------------
# Helper Function: Read Start Row
# -----------------------------------------------------------------------------
def _read_start_row(pix_config, log_file):
    """Reads the starting row number from the text file."""
    Log.Log_Info(log_file, 'Get The Starting Row Count')
    start_row_file = pix_config['start_row_file']
    start_number_raw = 0
    try:
        if not os.path.exists(start_row_file):
             Log.Log_Error(log_file, f"StartRow file not found: {start_row_file}. Creating and starting from row 0.")
             with open(start_row_file, 'w') as f: f.write('0')
        else:
            start_number_raw = Row_Number_Func.start_row_number(start_row_file)

        start_number_for_read = max(0, start_number_raw - 500) # Apply offset
        Log.Log_Info(log_file, f"Raw Start_Number: {start_number_raw}, Adjusted Start_Number for read: {start_number_for_read}")
        return start_number_raw, start_number_for_read
    except Exception as e:
        Log.Log_Error(log_file, f"Error reading or creating StartRow file {start_row_file}: {e}")
        return None, None # Indicate error

# -----------------------------------------------------------------------------
# Helper Function: Read and Prepare DataFrame
# -----------------------------------------------------------------------------
def _read_and_prepare_dataframe(local_excel_file, start_row_for_read, date_col_index, log_file):
    """Reads Excel, cleans, and filters data."""
    Log.Log_Info(log_file, f'Reading Excel: {local_excel_file} from adjusted row {start_row_for_read}')
    try:
        df = pd.read_excel(local_excel_file, header=None, sheet_name=DATA_SHEET_NAME, usecols="C:X", skiprows=start_row_for_read, dtype=str)
    except FileNotFoundError:
        Log.Log_Error(log_file, f"Excel file not found: {local_excel_file}")
        return None
    except Exception as e:
        Log.Log_Error(log_file, f"Error reading Excel file {local_excel_file}: {e}")
        return None

    df = df.dropna(how='all')
    if df.empty:
        Log.Log_Info(log_file, f"DataFrame empty after dropna(how='all').")
        return None # No data

    df.columns = range(df.shape[1]) # Rename columns 0, 1, 2...

    try:
        df[date_col_index] = pd.to_datetime(df[date_col_index], errors='coerce')
        original_rows = len(df)
        df = df.dropna(subset=[date_col_index])
        if original_rows > len(df):
             Log.Log_Error(log_file, f"Removed {original_rows - len(df)} rows with invalid date format.")
        if df.empty:
            Log.Log_Info(log_file, "DataFrame empty after date coercion/dropna.")
            return None

        months_to_keep = 2
        date_threshold = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) - relativedelta(months=months_to_keep)
        df = df[(df[date_col_index] >= date_threshold)]
        if df.empty:
             Log.Log_Info(log_file, f"DataFrame empty after filtering since {date_threshold.date()}.")
             return None

    except Exception as e:
         Log.Log_Error(log_file, f"Error processing date column: {e}")
         return None

    df = df.replace(['', 'nan', 'NaN', 'NAN'], np.nan)
    return df

# -----------------------------------------------------------------------------
# Helper Function: Initialize PIX1 Entry
# -----------------------------------------------------------------------------
def _initialize_pix1_entry(Serial_Number, original_excel_row_approx, log_file):
    """Checks Prime, initializes entry in PIX_Data_List for PIX1."""
    global PIX_Data_List, List_Index_Lot

    if Serial_Number in List_Index_Lot:
        Log.Log_Error(log_file, f"Row {original_excel_row_approx}: Duplicate SN '{Serial_Number}' during PIX1 init. Skipping.")
        return None, False # Indicate skip

    conn, cursor = SQL.connSQL()
    if conn is None:
        Log.Log_Error(log_file, f"SN {Serial_Number} (Row {original_excel_row_approx}): Connection with Prime Failed.")
        return None, True # Indicate critical failure

    Part_Number, Nine_Serial_Number = SQL.selectSQL(cursor, Serial_Number)
    SQL.disconnSQL(conn, cursor)

    if Part_Number is None or Part_Number == 'LDアレイ_':
        Log.Log_Error(log_file, f"Row {original_excel_row_approx}: SN '{Serial_Number}' - PartNumber invalid ('{Part_Number}'). Skipping.")
        return None, False # Indicate skip

    PIX_Data_List.append([0] * 30)
    index = len(PIX_Data_List) - 1
    List_Index_Lot[Serial_Number] = index
    # Pre-fill PIX1 specific static data
    PIX_Data_List[index][IDX_SN] = Serial_Number
    PIX_Data_List[index][3] = Part_Number # Assuming index 3 is Part#
    PIX_Data_List[index][4] = Nine_Serial_Number # Assuming index 4 is Lot9
    return index, False # Return index, indicate not critical failure

# -----------------------------------------------------------------------------
# Helper Function: Write Row Data to List
# -----------------------------------------------------------------------------
def _write_pix_data_to_list(PIX, index, df_row, pix_config, original_excel_row_approx, log_file):
    """Writes data from a DataFrame row to PIX_Data_List."""
    global PIX_Data_List
    try:
        # Determine column offset based on PIX type
        col_offset = 0
        if PIX == "PIX2": col_offset = 11
        elif PIX == "PIX3": col_offset = 19

        # Date and Operator
        date_obj = df_row[pix_config['date_col']]
        date_str = Convert_Date.Edit_Date(date_obj)
        if len(date_str) != 19: raise ValueError(f"Invalid date format: {date_str}")
        PIX_Data_List[index][col_offset] = date_str
        PIX_Data_List[index][col_offset + 1] = str(df_row[pix_config['op_col']]).strip()

        # Adjust offset for PIX1 specific fields (SN, Part#, Lot9 already filled)
        if PIX == "PIX1": col_offset += 3

        # Equipment (last 2 chars)
        PIX_Data_List[index][col_offset + 2] = str(df_row[pix_config['eq_col']]).strip()[-2:]

        # Step data
        for i, step_col_idx in enumerate(pix_config['step_cols']):
             step_val_str = str(df_row[step_col_idx]).strip()
             try:
                 step_val_float = float(step_val_str.replace(',', ''))
             except ValueError:
                  raise ValueError(f"Invalid numeric value for Step{i+1}: '{step_val_str}'")
             PIX_Data_List[index][col_offset + 3 + i] = step_val_float

        # Store original Excel row number
        df_index_list_idx = 27 + ["PIX1", "PIX2", "PIX3"].index(PIX)
        PIX_Data_List[index][df_index_list_idx] = original_excel_row_approx
        return True

    except (IndexError, ValueError, TypeError) as e:
        Log.Log_Error(log_file, f"Row {original_excel_row_approx} (SN:{PIX_Data_List[index][IDX_SN]}): Error writing data for {PIX}: {e}")
        # Mark PIX1 entry as invalid if write fails
        if PIX == "PIX1":
            PIX_Data_List[index][IDX_P1_DATE] = INVALID_MARKER
        return False

# -----------------------------------------------------------------------------
# Helper Function: Process DataFrame Rows
# -----------------------------------------------------------------------------
def _process_rows(PIX, df, pix_config, start_number_for_read, log_file):
    """Iterates through DataFrame rows, performs checks, and populates PIX_Data_List."""
    global List_Index_Lot # Needs modification access

    Log.Log_Info(log_file, 'Processing DataFrame rows...')
    row_end = len(df)
    df_idx = df.index.values
    Log.Log_Info(log_file, f"{PIX} Processing {row_end} rows from DataFrame index {df_idx[0]} to {df_idx[-1]}")

    processed_rows_count = 0
    skipped_blank_count = 0
    skipped_lot_error_count = 0
    skipped_prime_error_count = 0
    skipped_dict_error_count = 0
    skipped_write_error_count = 0
    last_processed_row_excel_num = -1 # Track the last successfully processed row's original number

    for df_row_idx in range(row_end):
        current_df_index = df_idx[df_row_idx]
        original_excel_row_approx = start_number_for_read + current_df_index + 1

        # Check for blanks
        try:
            if df.iloc[df_row_idx, pix_config['required_cols']].isnull().any():
                Log.Log_Error(log_file, f"Row {original_excel_row_approx}: Blank value found. Skipping.")
                skipped_blank_count += 1
                continue
        except IndexError:
             Log.Log_Error(log_file, f"Row {original_excel_row_approx}: IndexError during blank check. Skipping.")
             skipped_blank_count += 1
             continue

        # Get Serial Number
        try:
            Serial_Number = str(df.iloc[df_row_idx, pix_config['sn_col']]).strip()
            if not Serial_Number or Serial_Number.lower() == 'nan': raise ValueError("SN empty/nan")
        except (IndexError, ValueError) as e:
            Log.Log_Error(log_file, f"Row {original_excel_row_approx}: Invalid SN ({e}). Skipping.")
            skipped_lot_error_count += 1
            continue

        # Get or initialize index in PIX_Data_List
        index = -1
        critical_failure = False
        if PIX == "PIX1":
            index, critical_failure = _initialize_pix1_entry(Serial_Number, original_excel_row_approx, log_file)
            if critical_failure:
                # Stop processing this file entirely if Prime connection failed
                skipped_prime_error_count += 1 # Count failure
                return None, processed_rows_count, skipped_blank_count, skipped_lot_error_count, skipped_prime_error_count, skipped_dict_error_count, skipped_write_error_count # Indicate critical failure by returning None for next_start
            if index is None:
                skipped_prime_error_count += 1 # Count skip (duplicate or invalid part#)
                continue
        else:
            if Serial_Number not in List_Index_Lot:
                Log.Log_Error(log_file, f"Row {original_excel_row_approx}: SN '{Serial_Number}' not found from PIX1 ({PIX}). Skipping.")
                skipped_dict_error_count += 1
                continue
            index = List_Index_Lot[Serial_Number]

        # Write data to the list
        write_success = _write_pix_data_to_list(PIX, index, df.iloc[df_row_idx], pix_config, original_excel_row_approx, log_file)

        if write_success:
            processed_rows_count += 1
            last_processed_row_excel_num = original_excel_row_approx # Update last successful row
        else:
            skipped_write_error_count += 1
            # Error logged in helper, just continue
            continue

    # Determine next start row based on the last successfully processed row
    next_start = -1
    if last_processed_row_excel_num > 0:
        next_start = last_processed_row_excel_num + 1
    elif row_end > 0: # Data was read, but nothing processed successfully
         # Default to starting after the last row read in the dataframe
         next_start = start_number_for_read + df_idx[-1] + 1
    # If df was empty initially, next_start remains -1 (handled by caller)

    summary = (processed_rows_count, skipped_blank_count, skipped_lot_error_count,
               skipped_prime_error_count, skipped_dict_error_count, skipped_write_error_count)
    return next_start, summary

# -----------------------------------------------------------------------------
# Helper Function: Validate and Prepare Lot Data for XML
# -----------------------------------------------------------------------------
def _validate_and_prepare_lot_data(lot_data, index, log_file):
    """Validates completeness, dates, calculates sorted data, checks types for XML."""
    serial_num_for_log = lot_data[IDX_SN] if len(lot_data) > IDX_SN and lot_data[IDX_SN] != 0 else f"Unknown (Index {index})"

    if lot_data[IDX_P1_DATE] == INVALID_MARKER:
        Log.Log_Error(log_file, f"Skipping Lot (Index {index}, SN:{serial_num_for_log}): Marked as invalid.")
        return None, "Invalid"

    if lot_data[IDX_P1_DATE] == 0 or lot_data[IDX_P2_DATE] == 0 or lot_data[IDX_P3_DATE] == 0:
        Log.Log_Error(log_file, f"Skipping Lot (Index {index}, SN:{serial_num_for_log}): Incomplete (missing PIX date).")
        return None, "Incomplete"

    try:
        data_dict = {
            'key_Part_Number': lot_data[3], 'key_Serial_Number': lot_data[2],
            'key_PIX1_Start_Date_Time': lot_data[0], 'key_PIX1_Operator': lot_data[1], 'key_LotNumber_9': lot_data[4], 'key_PIX1_Equipment': lot_data[5],
            'key_PIX1_Step1': lot_data[6], 'key_PIX1_Step2': lot_data[7], 'key_PIX1_Step3': lot_data[8], 'key_PIX1_Step_Ave': lot_data[9], 'key_PIX1_Step_3sigma': lot_data[10],
            'key_PIX2_Start_Date_Time': lot_data[11], 'key_PIX2_Operator': lot_data[12], 'key_PIX2_Equipment': lot_data[13],
            'key_PIX2_Step1': lot_data[14], 'key_PIX2_Step2': lot_data[15], 'key_PIX2_Step3': lot_data[16], 'key_PIX2_Step_Ave': lot_data[17], 'key_PIX2_Step_3sigma': lot_data[18],
            'key_PIX3_Start_Date_Time': lot_data[19], 'key_PIX3_Operator': lot_data[20], 'key_PIX3_Equipment': lot_data[21],
            'key_PIX3_Step1': lot_data[22], 'key_PIX3_Step2': lot_data[23], 'key_PIX3_Step3': lot_data[24], 'key_PIX3_Step_Ave': lot_data[25], 'key_PIX3_Step_3sigma': lot_data[26]
        }
        pix1_excel_row_approx = lot_data[IDX_P1_ROW]
    except IndexError:
        Log.Log_Error(log_file, f"Skipping Lot (Index {index}): IndexError accessing lot_data.")
        return None, "Incomplete"

    # Validate date formats
    pix1_time_str = data_dict["key_PIX1_Start_Date_Time"]
    pix2_time_str = data_dict["key_PIX2_Start_Date_Time"]
    pix3_time_str = data_dict["key_PIX3_Start_Date_Time"]
    try:
        if not isinstance(pix1_time_str, str) or len(pix1_time_str) != 19: raise ValueError("P1 Date")
        if not isinstance(pix2_time_str, str) or len(pix2_time_str) != 19: raise ValueError("P2 Date")
        if not isinstance(pix3_time_str, str) or len(pix3_time_str) != 19: raise ValueError("P3 Date")
        dt1 = datetime.strptime(pix1_time_str, "%Y-%m-%d %H:%M:%S")
        datetime.strptime(pix2_time_str, "%Y-%m-%d %H:%M:%S")
        datetime.strptime(pix3_time_str, "%Y-%m-%d %H:%M:%S")
    except ValueError as e:
        Log.Log_Error(log_file, f"Skipping Lot (SN:{serial_num_for_log}): Date Format Error ({e}).")
        return None, "Date"

    # Calculate SORTED_DATA
    try:
        sort_number_approx = pix1_excel_row_approx
        if not isinstance(sort_number_approx, (int, float)) or sort_number_approx <= 0:
             raise ValueError(f"Invalid sort_number_approx: {sort_number_approx}")
        delta = dt1 - datetime(1899, 12, 30)
        date_excel_number = delta.days + delta.seconds / (24 * 60 * 60)
        starttime_sorted = date_excel_number + float(sort_number_approx) / 10**6
        data_dict["key_STARTTIME_SORTED_PIX1"] = starttime_sorted
        data_dict["key_SORTNUMBER_PIX1"] = float(sort_number_approx)
    except (ValueError, TypeError, KeyError) as e:
        Log.Log_Error(log_file, f"Skipping Lot (SN:{serial_num_for_log}): Error calculating SORTED_DATA: {e}")
        return None, "SortedData"

    # Data Type Validation (using Check module)
    temp_dict_for_check = data_dict.copy()
    validation_passed = True
    for k, expected_type in KEY_TYPE_DEF.items():
         if k in temp_dict_for_check and expected_type == float:
             try:
                 if not isinstance(temp_dict_for_check[k], float):
                     temp_dict_for_check[k] = float(str(temp_dict_for_check[k]).replace(',',''))
             except (ValueError, TypeError):
                  Log.Log_Error(log_file, f"Lot (SN:{serial_num_for_log}): Cannot convert {k} to float.")
                  validation_passed = False; break
    if not validation_passed: return None, "Type"

    if not Check.Data_Type(KEY_TYPE_DEF, temp_dict_for_check):
         Log.Log_Error(log_file, f"Skipping Lot (SN:{serial_num_for_log}): Check.Data_Type validation failed.")
         return None, "Type"

    return data_dict, None # Return prepared dict and no error type

# -----------------------------------------------------------------------------
# Helper Function: Create XML String
# -----------------------------------------------------------------------------
def _create_xml_string(data_dict):
    """Formats the XML content string."""
    try:
        # Use global constants for Site, Operation, etc.
        pix1_time = data_dict["key_PIX1_Start_Date_Time"]
        pix2_time = data_dict["key_PIX2_Start_Date_Time"]
        pix3_time = data_dict["key_PIX3_Start_Date_Time"]
        # Ensure all values are strings for formatting
        return f'''<?xml version="1.0" encoding="utf-8"?>
<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
       <Result startDateTime="{pix1_time}" Result="Passed">
               <Header SerialNumber="{data_dict["key_Serial_Number"]}" PartNumber="{data_dict["key_Part_Number"]}" Operation="{OPERATION}" TestStation="{TEST_STATION}" Operator="{data_dict["key_PIX1_Operator"]}" StartTime="{pix1_time}" Site="{SITE}" LotNumber="{data_dict["key_Serial_Number"]}"/>
               <HeaderMisc>
                   <Item Description="PIX1_Operator">{str(data_dict["key_PIX1_Operator"])}</Item>
                   <Item Description="PIX2_Operator">{str(data_dict["key_PIX2_Operator"])}</Item>
                   <Item Description="PIX3_Operator">{str(data_dict["key_PIX3_Operator"])}</Item>
               </HeaderMisc>
               <TestStep Name="PIX1" startDateTime="{pix1_time}" Status="Passed">
                   <Data DataType="Numeric" Name="Step1" Units="nm" Value="{str(data_dict["key_PIX1_Step1"])}"/>
                   <Data DataType="Numeric" Name="Step2" Units="nm" Value="{str(data_dict["key_PIX1_Step2"])}"/>
                   <Data DataType="Numeric" Name="Step3" Units="nm" Value="{str(data_dict["key_PIX1_Step3"])}"/>
                   <Data DataType="Numeric" Name="Step_Ave" Units="nm" Value="{str(data_dict["key_PIX1_Step_Ave"])}"/>
                   <Data DataType="Numeric" Name="Step_3sigma" Units="nm" Value="{str(data_dict["key_PIX1_Step_3sigma"])}"/>
               </TestStep>
               <TestStep Name="PIX2" startDateTime="{pix2_time}" Status="Passed">
                   <Data DataType="Numeric" Name="Step1" Units="nm" Value="{str(data_dict["key_PIX2_Step1"])}"/>
                   <Data DataType="Numeric" Name="Step2" Units="nm" Value="{str(data_dict["key_PIX2_Step2"])}"/>
                   <Data DataType="Numeric" Name="Step3" Units="nm" Value="{str(data_dict["key_PIX2_Step3"])}"/>
                   <Data DataType="Numeric" Name="Step_Ave" Units="nm" Value="{str(data_dict["key_PIX2_Step_Ave"])}"/>
                   <Data DataType="Numeric" Name="Step_3sigma" Units="nm" Value="{str(data_dict["key_PIX2_Step_3sigma"])}"/>
               </TestStep>
               <TestStep Name="PIX3" startDateTime="{pix3_time}" Status="Passed">
                   <Data DataType="Numeric" Name="Step1" Units="nm" Value="{str(data_dict["key_PIX3_Step1"])}"/>
                   <Data DataType="Numeric" Name="Step2" Units="nm" Value="{str(data_dict["key_PIX3_Step2"])}"/>
                   <Data DataType="Numeric" Name="Step3" Units="nm" Value="{str(data_dict["key_PIX3_Step3"])}"/>
                   <Data DataType="Numeric" Name="Step_Ave" Units="nm" Value="{str(data_dict["key_PIX3_Step_Ave"])}"/>
                   <Data DataType="Numeric" Name="Step_3sigma" Units="nm" Value="{str(data_dict["key_PIX3_Step_3sigma"])}"/>
               </TestStep>
               <TestStep Name="SORTED_DATA" startDateTime="{pix1_time}" Status="Passed">
                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value="{str(data_dict["key_STARTTIME_SORTED_PIX1"])}"/>
                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value="{str(data_dict["key_SORTNUMBER_PIX1"])}"/>
                   <Data DataType="String" Name="LotNumber_5" Value="{str(data_dict["key_Serial_Number"])}" CompOperation="LOG"/>
                   <Data DataType="String" Name="LotNumber_9" Value="{str(data_dict["key_LotNumber_9"])}" CompOperation="LOG"/>
               </TestStep>
               <TestEquipment>
                   <Item DeviceName="DryEtch_PIX1" DeviceSerialNumber="{str(data_dict['key_PIX1_Equipment'])}"/>
                   <Item DeviceName="DryEtch_PIX2" DeviceSerialNumber="{str(data_dict['key_PIX2_Equipment'])}"/>
                   <Item DeviceName="DryEtch_PIX3" DeviceSerialNumber="{str(data_dict['key_PIX3_Equipment'])}"/>
               </TestEquipment>
               <ErrorData/>
               <FailureData/>
               <Configuration/>
       </Result>
</Results>'''
    except KeyError as e:
        # Log error if needed, return None
        print(f"Error creating XML string: Missing key {e}") # Basic print for now
        return None

# -----------------------------------------------------------------------------
# Helper Function: Write XML File
# -----------------------------------------------------------------------------
def _write_xml_file(data_dict, xml_content, log_file):
    """Writes the XML content to a file."""
    xml_full_path = "Unknown"
    try:
        sn = data_dict["key_Serial_Number"]
        part_num = data_dict["key_Part_Number"]
        pix1_time_str = data_dict["key_PIX1_Start_Date_Time"]
        safe_testdate = pix1_time_str.replace(":", "-").replace(" ", "_")
        xml_file_name = f'Site={SITE},ProductFamily={PRODUCT_FAMILY},Operation={OPERATION},' + \
                        f'Partnumber={part_num},Serialnumber={sn},Testdate={safe_testdate}.xml'
        xml_full_path = os.path.join(OUTPUT_FILEPATH, xml_file_name)

        output_dir = os.path.dirname(xml_full_path)
        if not os.path.exists(output_dir):
             try:
                 os.makedirs(output_dir)
                 Log.Log_Info(log_file, f"Created output directory: {output_dir}")
             except OSError as e:
                  Log.Log_Error(log_file, f"Failed to create output directory {output_dir}: {e}. Cannot write XML.")
                  return False # Indicate failure

        with open(xml_full_path, 'w', encoding="utf-8") as f:
            f.write(xml_content)
        Log.Log_Info(log_file, f"Successfully wrote XML: {xml_full_path}")
        return True
    except Exception as e:
        Log.Log_Error(log_file, f"Error writing XML file {xml_full_path} for SN:{data_dict.get('key_Serial_Number', 'Unknown')}: {e}")
        return False

# -----------------------------------------------------------------------------
# Helper Function: Update and Backup Start Rows
# -----------------------------------------------------------------------------
def _update_and_backup_startrows(config, script_dir, log_file):
    """Updates StartRow files and performs backup."""
    global Next_StartNumber # Access global variable
    Log.Log_Info(log_file, 'Updating StartRow files...')
    try:
        start_row_files = [
            os.path.abspath(os.path.join(script_dir, config.get('PIX1', 'StartRowFile'))),
            os.path.abspath(os.path.join(script_dir, config.get('PIX2', 'StartRowFile'))),
            os.path.abspath(os.path.join(script_dir, config.get('PIX3', 'StartRowFile')))
        ]

        if not all(isinstance(n, (int, float)) and n >= 0 for n in Next_StartNumber):
             raise ValueError(f"Invalid values in Next_StartNumber: {Next_StartNumber}")

        for idx, file_path in enumerate(start_row_files):
            Row_Number_Func.next_start_row_number(file_path, int(Next_StartNumber[idx]))

        Log.Log_Info(log_file, f"Updated StartRow files: PIX1={int(Next_StartNumber[0])}, PIX2={int(Next_StartNumber[1])}, PIX3={int(Next_StartNumber[2])}")

        # Backup StartRow files if configured
        if START_ROW_FILE_BASE_DIR:
            if not os.path.exists(START_ROW_FILE_BASE_DIR):
                try:
                    os.makedirs(START_ROW_FILE_BASE_DIR)
                    Log.Log_Info(log_file, f"Created backup directory: {START_ROW_FILE_BASE_DIR}")
                except OSError as e:
                    Log.Log_Error(log_file, f"Error creating backup directory {START_ROW_FILE_BASE_DIR}: {e}. Skipping backup.")
                    START_ROW_FILE_BASE_DIR = None # Prevent copy attempt

            if START_ROW_FILE_BASE_DIR:
                try:
                    for file_path in start_row_files:
                        shutil.copy(file_path, START_ROW_FILE_BASE_DIR)
                    Log.Log_Info(log_file, f"Copied StartRow files to backup directory: {START_ROW_FILE_BASE_DIR}")
                except Exception as e:
                    Log.Log_Error(log_file, f"Error copying StartRow files to backup directory {START_ROW_FILE_BASE_DIR}: {e}")
        else:
             Log.Log_Error(log_file, "StartRowFileBaseDir not configured. Skipping backup.")

    except (ValueError, TypeError, KeyError, configparser.NoSectionError, configparser.NoOptionError) as e:
         Log.Log_Error(log_file, f"Error preparing to update StartRow files: {e}")
    except Exception as e:
        Log.Log_Error(log_file, f"Error updating or backing up StartRow files: {e}")


# -----------------------------------------------------------------------------
# XML Generation Orchestration Function
# -----------------------------------------------------------------------------
def generate_xml_files(config, script_dir, log_file):
    """Orchestrates the XML generation process using validated data."""
    global PIX_Data_List # Access global data list

    Log.Log_Info(log_file, "--- Starting XML Generation Process ---")
    processed_xml_count = 0
    error_summary = {"Invalid": 0, "Incomplete": 0, "Date": 0, "Type": 0, "XMLWrite": 0, "SortedData": 0}

    for i, lot_data in enumerate(PIX_Data_List):
        data_dict, error_type = _validate_and_prepare_lot_data(lot_data, i, log_file)

        if error_type:
            if error_type in error_summary:
                error_summary[error_type] += 1
            else: # Should not happen if error types are consistent
                error_summary["Unknown"] = error_summary.get("Unknown", 0) + 1
            continue # Skip to next lot

        xml_content = _create_xml_string(data_dict)
        if not xml_content:
             Log.Log_Error(log_file, f"Failed to create XML string for SN:{data_dict.get('key_Serial_Number', 'Unknown')}. Skipping.")
             error_summary["XMLWrite"] += 1
             continue

        write_success = _write_xml_file(data_dict, xml_content, log_file)
        if write_success:
            processed_xml_count += 1
        else:
            error_summary["XMLWrite"] += 1

    Log.Log_Info(log_file, f"--- Finished XML Generation Process ---")
    summary_msg = f"XML Summary: Processed={processed_xml_count}, Skipped(" + \
                  f"Invalid={error_summary['Invalid']}, Incomplete={error_summary['Incomplete']}, " + \
                  f"DateErr={error_summary['Date']}, TypeErr={error_summary['Type']}, " + \
                  f"SortedDataErr={error_summary['SortedData']}, XMLWriteErr={error_summary['XMLWrite']})"
    Log.Log_Info(log_file, summary_msg)

    # Update start rows only if XMLs were generated
    if processed_xml_count > 0:
        _update_and_backup_startrows(config, script_dir, log_file)
    else:
        Log.Log_Info(log_file, "No XML files generated, skipping StartRow update.")


# -----------------------------------------------------------------------------
# Main Data Processing Orchestration Function
# -----------------------------------------------------------------------------
def process_pix_data_type(PIX, config, script_dir, log_file):
    """Orchestrates the data collection and processing for a single PIX type."""
    global Next_StartNumber # Allow modification

    Log.Log_Info(log_file, f"--- Start processing data for {PIX} ---")
    pix_list_index = ["PIX1", "PIX2", "PIX3"].index(PIX)

    pix_config = _get_pix_config(PIX, config, script_dir, log_file)
    if not pix_config: return False # Config error

    local_excel_file = _find_and_copy_excel(pix_config, log_file)
    if not local_excel_file: return False # File copy error

    start_number_raw, start_number_for_read = _read_start_row(pix_config, log_file)
    if start_number_raw is None: return False # Start row read error

    df = _read_and_prepare_dataframe(local_excel_file, start_number_for_read, pix_config['date_col'], log_file)

    calculated_next_start = start_number_raw # Default if no data or processing fails
    if df is None or df.empty:
        Log.Log_Info(log_file, f"No processable data found for {PIX} after reading/filtering.")
        Next_StartNumber[pix_list_index] = start_number_raw # Update global state
    else:
        # Process rows
        next_start_from_process, summary = _process_rows(PIX, df, pix_config, start_number_for_read, log_file)

        # Log summary
        (proc, sk_b, sk_l, sk_p, sk_d, sk_w) = summary
        summary_msg = f"Summary for {PIX}: Processed={proc}, Skipped(Blank={sk_b}, LotErr={sk_l}, Prime/Dup={sk_p}, DictErr={sk_d}, WriteErr={sk_w})"
        Log.Log_Info(log_file, summary_msg)

        # Check for critical PIX1 failure indicated by next_start_from_process being None
        if PIX == "PIX1" and next_start_from_process is None:
            Log.Log_Error(log_file, "Critical failure during PIX1 row processing (likely Prime connection). Aborting further processing.")
            return False # Signal critical failure

        # Update Next_StartNumber based on processing result
        if next_start_from_process is not None and next_start_from_process > 0:
             calculated_next_start = next_start_from_process
        # If next_start_from_process is None (critical fail) or <= 0 (no rows processed), keep start_number_raw

        Next_StartNumber[pix_list_index] = calculated_next_start # Update global state

    Log.Log_Info(log_file, f"--- Finished processing data for {PIX} ---")
    return True # Indicate success for this PIX type (unless critical failure occurred)


# -----------------------------------------------------------------------------
# Main Execution Block
# -----------------------------------------------------------------------------
if __name__ == '__main__':

    print("Loading configuration...")
    config, script_dir = load_config()
    if not config or not script_dir: sys.exit(1)

    print("Setting up logging...")
    log_file = setup_logging(config, script_dir)

    # --- Populate Global Settings from Config ---
    try:
        SITE = config.get('General', 'Site')
        PRODUCT_FAMILY = config.get('General', 'ProductFamily')
        OPERATION = config.get('General', 'Operation')
        TEST_STATION = config.get('General', 'TestStation')
        DATA_SHEET_NAME = config.get('General', 'DataSheetName')
        OUTPUT_FILEPATH = config.get('General', 'OutputFilePath')
        local_data_file_dir_rel = config.get('General', 'LocalDataFileDir')
        LOCAL_DATA_FILE_DIR_ABS = os.path.abspath(os.path.join(script_dir, local_data_file_dir_rel))
        START_ROW_FILE_BASE_DIR = config.get('General', 'StartRowFileBaseDir', fallback=None)

        if not OUTPUT_FILEPATH: raise ValueError("OutputFilePath empty.")
        if not LOCAL_DATA_FILE_DIR_ABS: raise ValueError("LocalDataFileDir empty.")
    except (configparser.NoSectionError, configparser.NoOptionError, ValueError) as e:
        msg = f"Critical configuration error in [General] section: {e}"
        print(f"Error: {msg}")
        if log_file: Log.Log_Error(log_file, msg)
        sys.exit(1)

    # --- Main Processing Loop ---
    execution_successful = True
    try:
        pix_types_to_process = ["PIX1", "PIX2", "PIX3"]
        for pix_type in pix_types_to_process:
            success = process_pix_data_type(pix_type, config, script_dir, log_file)
            if not success and pix_type == "PIX1":
                Log.Log_Error(log_file, "Aborting due to critical failure during PIX1 processing.")
                execution_successful = False
                break # Stop processing PIX2, PIX3 if PIX1 failed critically

        # Generate XML only if all PIX types processed successfully (or PIX1 didn't fail critically)
        if execution_successful:
            generate_xml_files(config, script_dir, log_file)

    except Exception as e:
        execution_successful = False
        msg = f"An unexpected error occurred in the main execution block: {e}"
        detailed_error = traceback.format_exc()
        print(f"Error: {msg}\n{detailed_error}")
        if log_file:
            Log.Log_Error(log_file, msg)
            Log.Log_Error(log_file, detailed_error)

    # --- Program End ---
    print("\n--- Process Finished ---")
    final_msg = 'Program End' + (' with errors.' if not execution_successful else '.')
    if log_file: Log.Log_Info(log_file, final_msg)
    else: print(final_msg)