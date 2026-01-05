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

# -----------------------------------------------------------------------------
# Helper: Import Custom Modules
# -----------------------------------------------------------------------------
def import_custom_modules():
    """Imports custom modules from ../MyModule"""
    try:
        script_dir_for_module = os.path.dirname(__file__)
        module_path = os.path.abspath(os.path.join(script_dir_for_module, '..', 'MyModule'))
        if module_path not in sys.path:
            sys.path.append(module_path)

        global Log, SQL, Check, Convert_Date, Row_Number_Func
        import Log
        import SQL
        import Check
        import Convert_Date
        import Row_Number_Func
        print(f"Successfully imported custom modules from {module_path}")
        return True
    except ImportError as e:
        print(f"Error importing custom modules from {module_path}: {e}")
        print("Please ensure 'MyModule' directory exists relative to the script and contains necessary .py files.")
        return False
    except NameError:
        print("Error: Could not determine script directory to find 'MyModule'.")
        return False

# -----------------------------------------------------------------------------
# Helper: Load Configuration
# -----------------------------------------------------------------------------
def load_config(config_filename='PIX_Config.ini'):
    """Loads configuration file."""
    config = configparser.ConfigParser()
    try:
        script_dir = os.path.dirname(__file__)
    except NameError:
        print("Error: Could not determine script directory to find config file.")
        return None, None

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
# Helper: Setup Logging
# -----------------------------------------------------------------------------
def setup_logging(config, script_dir):
    """Sets up logging based on configuration."""
    if not config or not script_dir:
        print("Error: Cannot setup logging due to missing config or script_dir.")
        return None
    try:
        log_base_dir = config.get('General', 'LogBaseDir', fallback='../Log/')
        log_file_name = config.get('General', 'LogFileName', fallback='014_PIX.log')
    except configparser.NoSectionError:
        print("Error: [General] section not found in config file for logging setup.")
        return None

    log_folder_name = str(date.today())
    log_dir_path = os.path.abspath(os.path.join(script_dir, log_base_dir, log_folder_name))

    try:
        if not os.path.exists(log_dir_path):
            os.makedirs(log_dir_path)
            print(f"Created log directory: {log_dir_path}")
        log_file_path = os.path.join(log_dir_path, log_file_name)
        print(f"Logging to: {log_file_path}")
        Log.Log_Info(log_file_path, 'Program Start')
        return log_file_path
    except Exception as e:
        print(f"Error setting up logging to {log_dir_path}: {e}")
        return None

# -----------------------------------------------------------------------------
# Data Collection: Find and Copy Excel
# -----------------------------------------------------------------------------
def find_and_copy_excel(PIX, config, script_dir, log_file):
    """Finds the latest relevant Excel file and copies it locally."""
    try:
        FilePath = config[PIX]['SourceFilePath']
        FileNamePattern = config[PIX]['SourceFileNamePattern']
        LocalDataFileDir = config['General']['LocalDataFileDir']
        local_data_dir_path = os.path.abspath(os.path.join(script_dir, LocalDataFileDir))
    except (KeyError, configparser.NoSectionError) as e:
        msg = f"Configuration error finding Excel for {PIX}: {e}"
        if log_file: Log.Log_Error(log_file, msg)
        else: print(f"Error: {msg}")
        return None

    msg = f'{PIX} Searching for Excel files...'
    if log_file: Log.Log_Info(log_file, msg)
    else: print(msg)

    Excel_file_list = []
    try:
        search_path = FilePath + FileNamePattern
        found_files = glob.glob(search_path)
        for file in found_files:
            if '$' not in file and not os.path.isdir(file):
                try:
                    dt = strftime("%Y-%m-%d %H:%M:%S", localtime(os.path.getmtime(file)))
                    Excel_file_list.append([file, dt])
                except FileNotFoundError:
                     pass # Ignore files that disappear between glob and getmtime
    except Exception as e:
        msg = f"Error searching for Excel files at {FilePath}: {e}"
        if log_file: Log.Log_Error(log_file, msg)
        else: print(f"Error: {msg}")
        return None

    if not Excel_file_list:
        msg = f"No valid Excel files found matching pattern: {FileNamePattern} in {FilePath}"
        if log_file: Log.Log_Error(log_file, msg)
        else: print(f"Error: {msg}")
        return None

    Excel_file_list = sorted(Excel_file_list, key=lambda x: x[1], reverse=True)
    source_excel = Excel_file_list[0][0]

    try:
        if not os.path.exists(local_data_dir_path):
            os.makedirs(local_data_dir_path)
        Excel_File = shutil.copy(source_excel, local_data_dir_path)
        msg = f"Successfully copied {source_excel} to {Excel_File}"
        if log_file: Log.Log_Info(log_file, msg)
        else: print(msg)
        return Excel_File
    except Exception as e:
        msg = f"Error copying Excel file from {source_excel} to {local_data_dir_path}: {e}"
        if log_file: Log.Log_Error(log_file, msg)
        else: print(f"Error: {msg}")
        return None

# -----------------------------------------------------------------------------
# Data Collection: Read Start Row
# -----------------------------------------------------------------------------
def read_start_row(PIX, config, script_dir, log_file):
    """Reads the starting row number from the text file."""
    try:
        TextFile = config[PIX]['StartRowFile']
        start_row_file_path = os.path.abspath(os.path.join(script_dir, TextFile))
    except (KeyError, configparser.NoSectionError) as e:
        msg = f"Configuration error reading StartRow file for {PIX}: {e}"
        if log_file: Log.Log_Error(log_file, msg)
        else: print(f"Error: {msg}")
        return -1, -1 # Indicate error

    Start_Number_Raw = 0
    try:
        if not os.path.exists(start_row_file_path):
             msg = f"StartRow file not found: {start_row_file_path}. Creating and starting from row 0."
             if log_file: Log.Log_Warning(log_file, msg)
             else: print(f"Warning: {msg}")
             with open(start_row_file_path, 'w') as f:
                 f.write('0')
        else:
            Start_Number_Raw = Row_Number_Func.start_row_number(start_row_file_path)

        # Apply -500 offset, ensuring non-negative
        Start_Number_For_Read = max(0, Start_Number_Raw - 500)
        msg = f"{PIX} Raw Start_Number: {Start_Number_Raw}, Adjusted Start_Number for read: {Start_Number_For_Read}"
        if log_file: Log.Log_Info(log_file, msg)
        else: print(msg)
        return Start_Number_Raw, Start_Number_For_Read
    except Exception as e:
        msg = f"Error reading or creating StartRow file {start_row_file_path}: {e}"
        if log_file: Log.Log_Error(log_file, msg)
        else: print(f"Error: {msg}")
        return -1, -1 # Indicate error

# -----------------------------------------------------------------------------
# Data Collection: Read and Prepare DataFrame
# -----------------------------------------------------------------------------
def read_and_prepare_dataframe(Excel_File, Start_Number_For_Read, config, log_file):
    """Reads Excel into DataFrame and performs initial preparation."""
    try:
        Data_Sheet_Name = config['General']['DataSheetName']
    except (KeyError, configparser.NoSectionError) as e:
        msg = f"Configuration error reading DataFrame: {e}"
        if log_file: Log.Log_Error(log_file, msg)
        else: print(f"Error: {msg}")
        return None

    msg = f'Reading Excel: {Excel_File} from row {Start_Number_For_Read}'
    if log_file: Log.Log_Info(log_file, msg)
    else: print(msg)

    try:
        df = pd.read_excel(Excel_File, header=None, sheet_name=Data_Sheet_Name, usecols="C:X", skiprows=Start_Number_For_Read, dtype=str)
    except FileNotFoundError:
        msg = f"Excel file not found: {Excel_File}"
        if log_file: Log.Log_Error(log_file, msg)
        else: print(f"Error: {msg}")
        return None
    except Exception as e:
        msg = f"Error reading Excel file {Excel_File}: {e}"
        if log_file: Log.Log_Error(log_file, msg)
        else: print(f"Error: {msg}")
        return None

    df = df.dropna(how='all')
    if df.empty:
        msg = f"No new data found in {Excel_File} after row {Start_Number_For_Read}."
        if log_file: Log.Log_Info(log_file, msg)
        else: print(msg)
        return None # Return None to indicate no new data

    df.columns = range(df.shape[1]) # Rename columns 0, 1, 2...

    # Date conversion and filtering
    try:
        date_col_idx = 0 # Date is in the first column of the read DataFrame
        df[date_col_idx] = pd.to_datetime(df[date_col_idx], errors='coerce')
        original_rows = len(df)
        df = df.dropna(subset=[date_col_idx])
        if original_rows > len(df):
             msg = f"Removed {original_rows - len(df)} rows with invalid date format."
             if log_file: Log.Log_Warning(log_file, msg)
             else: print(f"Warning: {msg}")
        if df.empty: return None # No valid dates

        months_to_keep = 2
        date_threshold = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) - relativedelta(months=months_to_keep)
        df = df[(df[date_col_idx] >= date_threshold)]
        if df.empty:
             msg = f"No recent data found after filtering since {date_threshold.date()}."
             if log_file: Log.Log_Info(log_file, msg)
             else: print(msg)
             return None

    except Exception as e:
         msg = f"Error processing date column: {e}"
         if log_file: Log.Log_Error(log_file, msg)
         else: print(f"Error: {msg}")
         return None # Indicate error

    # Replace various null representations
    df = df.replace(['', 'nan', 'NaN', 'NAN'], np.nan)
    return df

# -----------------------------------------------------------------------------
# Data Collection: Process DataFrame Rows
# -----------------------------------------------------------------------------
def process_dataframe_rows(PIX, df, config, log_file, pix_data_list, list_index_lot):
    """Iterates through DataFrame rows, performs checks, and populates data list."""
    try:
        # Get column indices from config
        Col_PIX_Start_Date_Time = 0 # Date is always the first column in df
        Col_PIX_Operator = int(config['Columns']['PIX_Operator'])
        Col_PIX_Equipment = int(config['Columns']['PIX_Equipment'])
        Col_PIX_Serial_Number = int(config['Columns']['PIX_Serial_Number'])
        step_key = f'{PIX}_Step'
        Col_Step = [int(x) for x in config['Columns'][step_key].split(',')]
        required_df_cols = [Col_PIX_Start_Date_Time, Col_PIX_Serial_Number, Col_PIX_Operator, Col_PIX_Equipment] + Col_Step
    except (KeyError, configparser.NoSectionError) as e:
        msg = f"Configuration error processing DataFrame for {PIX}: {e}"
        if log_file: Log.Log_Error(log_file, msg)
        else: print(f"Error: {msg}")
        return 0, 0, 0, 0, 0, 0 # Return zero counts

    row_end = len(df)
    df_idx = df.index.values
    msg = f"{PIX} Processing {row_end} rows from DataFrame index {df_idx[0]} to {df_idx[-1]}"
    if log_file: Log.Log_Info(log_file, msg)
    else: print(msg)

    processed_rows_count = 0
    skipped_blank_count = 0
    skipped_lot_error_count = 0
    skipped_prime_error_count = 0
    skipped_dict_error_count = 0
    skipped_write_error_count = 0

    for df_row_idx in range(row_end):
        current_df_index = df_idx[df_row_idx]

        # Check for blanks in required columns
        try:
            if df.iloc[df_row_idx, required_df_cols].isnull().any():
                skipped_blank_count += 1
                continue
        except IndexError:
             skipped_blank_count += 1
             continue

        # Get Serial Number
        try:
            Serial_Number = str(df.iloc[df_row_idx, Col_PIX_Serial_Number]).strip()
            if Serial_Number == "": raise ValueError("Serial number is empty")
        except (IndexError, ValueError):
            skipped_lot_error_count += 1
            continue

        # PIX1: Connect to Prime, check Part Number, initialize list entry
        if PIX == "PIX1":
            conn, cursor = SQL.connSQL()
            if conn is None:
                msg = f"{Serial_Number} : Connection with Prime Failed. Stopping PIX1 processing."
                if log_file: Log.Log_Error(log_file, msg)
                else: print(f"Error: {msg}")
                # Return current counts, indicating PIX1 failed mid-process
                return processed_rows_count, skipped_blank_count, skipped_lot_error_count, skipped_prime_error_count, skipped_dict_error_count, skipped_write_error_count

            Part_Number, Nine_Serial_Number = SQL.selectSQL(cursor, Serial_Number)
            SQL.disconnSQL(conn, cursor)

            if Part_Number is None or Part_Number == 'LDアレイ_':
                skipped_prime_error_count += 1
                continue
            if Serial_Number in list_index_lot: # Skip duplicates for PIX1
                 continue

            pix_data_list.append([0] * 30)
            index = len(pix_data_list) - 1
            list_index_lot[Serial_Number] = index
            col_offset = 0
        # PIX2, PIX3: Find existing entry
        else:
            if Serial_Number not in list_index_lot:
                skipped_dict_error_count += 1
                continue
            index = list_index_lot[Serial_Number]
            col_offset = 11 if PIX == "PIX2" else 19

        # Write data to list
        try:
            date_obj = df.iloc[df_row_idx, Col_PIX_Start_Date_Time]
            date_str = Convert_Date.Edit_Date(date_obj)
            if len(date_str) != 19: raise ValueError(f"Invalid date format: {date_str}")
            pix_data_list[index][col_offset] = date_str
            pix_data_list[index][col_offset + 1] = str(df.iloc[df_row_idx, Col_PIX_Operator]).strip()

            if PIX == "PIX1":
                pix_data_list[index][col_offset + 2] = Serial_Number
                pix_data_list[index][col_offset + 3] = Part_Number
                pix_data_list[index][col_offset + 4] = Nine_Serial_Number
                col_offset += 3

            pix_data_list[index][col_offset + 2] = str(df.iloc[df_row_idx, Col_PIX_Equipment]).strip()[-2:]

            for i, step_col_idx in enumerate(Col_Step):
                 step_val_str = str(df.iloc[df_row_idx, step_col_idx]).strip()
                 try:
                     # Clean comma before float conversion
                     step_val_float = float(step_val_str.replace(',', ''))
                 except ValueError:
                      raise ValueError(f"Invalid numeric value for Step{i+1}: '{step_val_str}'")
                 pix_data_list[index][col_offset + 3 + i] = step_val_float

            # Store DataFrame index
            pix_data_list[index][27 + (0 if PIX == "PIX1" else 1 if PIX == "PIX2" else 2)] = current_df_index
            processed_rows_count += 1

        except (IndexError, ValueError, TypeError) as e:
            msg = f"Error processing data row for {Serial_Number} ({PIX}) at df index {current_df_index}: {e}"
            if log_file: Log.Log_Error(log_file, msg)
            else: print(f"Error: {msg}")
            skipped_write_error_count += 1
            if PIX == "PIX1" and Serial_Number in list_index_lot and list_index_lot.get(Serial_Number) == index:
                pix_data_list[index][0] = "INVALID" # Mark as invalid
                del list_index_lot[Serial_Number] # Remove from lookup
            continue

    return processed_rows_count, skipped_blank_count, skipped_lot_error_count, skipped_prime_error_count, skipped_dict_error_count, skipped_write_error_count

# -----------------------------------------------------------------------------
# Data Collection: Main Function
# -----------------------------------------------------------------------------
def collect_data_for_pix(PIX, config, script_dir, log_file, pix_data_list, list_index_lot):
    """Orchestrates data collection steps for a single PIX type."""
    msg = f"--- Start collecting data for {PIX} ---"
    if log_file: Log.Log_Info(log_file, msg)
    else: print(msg)

    next_start_row = 0 # Default return value indicates failure or no new data processed

    Excel_File = find_and_copy_excel(PIX, config, script_dir, log_file)
    if not Excel_File:
        return next_start_row # Error logged in helper

    Start_Number_Raw, Start_Number_For_Read = read_start_row(PIX, config, script_dir, log_file)
    if Start_Number_Raw < 0: # Error reading start row
        return next_start_row # Error logged in helper
    next_start_row = Start_Number_Raw # Default next start is the raw value read

    df = read_and_prepare_dataframe(Excel_File, Start_Number_For_Read, config, log_file)
    if df is None: # Error or no new data
        # If read_and_prepare returns None due to error, return 0
        # If it returns None due to no data, return Start_Number_Raw
        # We need a way to differentiate, for now assume error if None
        # Let's refine: if df is None AND Start_Number_Raw >= 0, it means no data or filter empty
        if Start_Number_Raw >= 0:
             return Start_Number_Raw
        else: # Should not happen if read_start_row worked, but safety check
             return 0


    # Calculate potential next start row based on DataFrame content
    df_idx = df.index.values
    calculated_next_start = Start_Number_For_Read + (df_idx[-1] - df_idx[0]) + 1

    # Process rows
    processed_count, skipped_b, skipped_l, skipped_p, skipped_d, skipped_w = process_dataframe_rows(
        PIX, df, config, log_file, pix_data_list, list_index_lot
    )

    msg = f"--- Finished collecting data for {PIX} ---"
    if log_file: Log.Log_Info(log_file, msg)
    else: print(msg)
    summary_msg = f"Summary for {PIX}: Processed={processed_count}, Skipped(Blank)={skipped_b}, Skipped(LotErr)={skipped_l}, Skipped(PrimeErr)={skipped_p}, Skipped(DictErr)={skipped_d}, Skipped(WriteErr)={skipped_w}"
    if log_file: Log.Log_Info(log_file, summary_msg)
    else: print(summary_msg)

    # Only update next_start_row if rows were actually processed
    # If PIX1 failed mid-process (e.g., Prime connection), process_dataframe_rows returns counts but we should signal failure (return 0)
    if PIX == "PIX1" and processed_count == 0 and (skipped_p > 0 or skipped_w > 0): # Heuristic for critical PIX1 failure
         return 0
    else:
        return calculated_next_start if processed_count > 0 else Start_Number_Raw

# -----------------------------------------------------------------------------
# XML Generation: Validate and Prepare Lot Data
# -----------------------------------------------------------------------------
def validate_and_prepare_lot(lot_data, index, log_file):
    """Validates completeness, dates, calculates sorted data, and checks types."""
    # Check for invalid marker
    if lot_data[0] == "INVALID":
        return None, "Marked as invalid during collection"

    # Check completeness (dates)
    if lot_data[0] == 0 or lot_data[11] == 0 or lot_data[19] == 0:
        serial_num = lot_data[2] if lot_data[2] != 0 else f"Unknown (Index {index})"
        return None, f"Data Incompleteness for Lot: {serial_num}"

    # Create initial dictionary
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
        pix1_df_index = lot_data[27] # Get stored DataFrame index
    except IndexError:
        return None, f"IndexError creating data_dict for list index {index}"

    # Validate date formats
    pix1_time_str = data_dict["key_PIX1_Start_Date_Time"]
    pix2_time_str = data_dict["key_PIX2_Start_Date_Time"]
    pix3_time_str = data_dict["key_PIX3_Start_Date_Time"]
    try:
        datetime.strptime(pix1_time_str, "%Y-%m-%d %H:%M:%S")
        datetime.strptime(pix2_time_str, "%Y-%m-%d %H:%M:%S")
        datetime.strptime(pix3_time_str, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None, f"{data_dict['key_Serial_Number']} : Date Format Error"

    # Calculate SORTED_DATA
    try:
        date_dt = datetime.strptime(pix1_time_str, "%Y-%m-%d %H:%M:%S")
        delta = date_dt - datetime(1899, 12, 30)
        date_excel_number = delta.days + delta.seconds / (24 * 60 * 60)
        sort_number_approx = pix1_df_index
        if not isinstance(sort_number_approx, (int, float)):
             sort_number_approx = int(sort_number_approx) # Try conversion
        starttime_sorted = date_excel_number + sort_number_approx / 10**6
        data_dict["key_STARTTIME_SORTED_PIX1"] = starttime_sorted
        data_dict["key_SORTNUMBER_PIX1"] = float(sort_number_approx)
    except (ValueError, TypeError) as e:
        return None, f"{data_dict['key_Serial_Number']}: Error calculating SORTED_DATA: {e}"

    # Check and convert data types (especially floats)
    # Define expected types for Check.Data_Type (assuming it needs this)
    full_key_type_def = {
        'key_Part_Number' : str, 'key_Serial_Number' : str, 'key_LotNumber_9': str,
        'key_PIX1_Start_Date_Time' : str, 'key_PIX1_Operator' : str, 'key_PIX1_Equipment' : str,
        'key_PIX1_Step1' : float, 'key_PIX1_Step2' : float, 'key_PIX1_Step3' : float, 'key_PIX1_Step_Ave' : float, 'key_PIX1_Step_3sigma' : float,
        'key_PIX2_Start_Date_Time' : str, 'key_PIX2_Operator' : str, 'key_PIX2_Equipment' : str,
        'key_PIX2_Step1' : float, 'key_PIX2_Step2' : float, 'key_PIX2_Step3' : float, 'key_PIX2_Step_Ave' : float, 'key_PIX2_Step_3sigma' : float,
        'key_PIX3_Start_Date_Time' : str, 'key_PIX3_Operator' : str, 'key_PIX3_Equipment' : str,
        'key_PIX3_Step1' : float, 'key_PIX3_Step2' : float, 'key_PIX3_Step3' : float, 'key_PIX3_Step_Ave' : float, 'key_PIX3_Step_3sigma' : float,
        "key_STARTTIME_SORTED_PIX1" : float, "key_SORTNUMBER_PIX1" : float
    }
    for k, expected_type in full_key_type_def.items():
        if k in data_dict:
            if expected_type == float:
                try:
                    if not isinstance(data_dict[k], float):
                        data_dict[k] = float(data_dict[k])
                except (ValueError, TypeError):
                    return None, f"{data_dict['key_Serial_Number']}: Cannot convert {k} ('{data_dict[k]}') to float"
            # Add checks for other types if needed (e.g., int, str)

    # Final check using Check.Data_Type
    # Pass both the dictionary and the type definition
    if not Check.Data_Type(full_key_type_def, data_dict):
         return None, f"{data_dict['key_Serial_Number']} : Check.Data_Type validation failed"

    return data_dict, None # Return prepared dict and no error

# -----------------------------------------------------------------------------
# XML Generation: Create XML String
# -----------------------------------------------------------------------------
def create_xml_string(data_dict, config):
    """Formats the XML content string using data_dict."""
    try:
        Site = config['General']['Site']
        ProductFamily = config['General']['ProductFamily']
        Operation = config['General']['Operation']
        TestStation = config['General']['TestStation']

        pix1_time_str = data_dict["key_PIX1_Start_Date_Time"]
        pix2_time_str = data_dict["key_PIX2_Start_Date_Time"]
        pix3_time_str = data_dict["key_PIX3_Start_Date_Time"]

        # Use f-string for cleaner formatting
        xml_content = f'''<?xml version="1.0" encoding="utf-8"?>
<Results xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
       <Result startDateTime="{pix1_time_str}" Result="Passed">
               <Header SerialNumber="{data_dict["key_Serial_Number"]}" PartNumber="{data_dict["key_Part_Number"]}" Operation="{Operation}" TestStation="{TestStation}" Operator="{data_dict["key_PIX1_Operator"]}" StartTime="{pix1_time_str}" Site="{Site}" LotNumber="{data_dict["key_Serial_Number"]}"/>
               <HeaderMisc>
                   <Item Description="PIX1_Operator">{str(data_dict["key_PIX1_Operator"])}</Item>
                   <Item Description="PIX2_Operator">{str(data_dict["key_PIX2_Operator"])}</Item>
                   <Item Description="PIX3_Operator">{str(data_dict["key_PIX3_Operator"])}</Item>
               </HeaderMisc>

               <TestStep Name="PIX1" startDateTime="{pix1_time_str}" Status="Passed">
                   <Data DataType="Numeric" Name="Step1" Units="nm" Value="{str(data_dict["key_PIX1_Step1"])}"/>
                   <Data DataType="Numeric" Name="Step2" Units="nm" Value="{str(data_dict["key_PIX1_Step2"])}"/>
                   <Data DataType="Numeric" Name="Step3" Units="nm" Value="{str(data_dict["key_PIX1_Step3"])}"/>
                   <Data DataType="Numeric" Name="Step_Ave" Units="nm" Value="{str(data_dict["key_PIX1_Step_Ave"])}"/>
                   <Data DataType="Numeric" Name="Step_3sigma" Units="nm" Value="{str(data_dict["key_PIX1_Step_3sigma"])}"/>
               </TestStep>

               <TestStep Name="PIX2" startDateTime="{pix2_time_str}" Status="Passed">
                   <Data DataType="Numeric" Name="Step1" Units="nm" Value="{str(data_dict["key_PIX2_Step1"])}"/>
                   <Data DataType="Numeric" Name="Step2" Units="nm" Value="{str(data_dict["key_PIX2_Step2"])}"/>
                   <Data DataType="Numeric" Name="Step3" Units="nm" Value="{str(data_dict["key_PIX2_Step3"])}"/>
                   <Data DataType="Numeric" Name="Step_Ave" Units="nm" Value="{str(data_dict["key_PIX2_Step_Ave"])}"/>
                   <Data DataType="Numeric" Name="Step_3sigma" Units="nm" Value="{str(data_dict["key_PIX2_Step_3sigma"])}"/>
               </TestStep>

               <TestStep Name="PIX3" startDateTime="{pix3_time_str}" Status="Passed">
                   <Data DataType="Numeric" Name="Step1" Units="nm" Value="{str(data_dict["key_PIX3_Step1"])}"/>
                   <Data DataType="Numeric" Name="Step2" Units="nm" Value="{str(data_dict["key_PIX3_Step2"])}"/>
                   <Data DataType="Numeric" Name="Step3" Units="nm" Value="{str(data_dict["key_PIX3_Step3"])}"/>
                   <Data DataType="Numeric" Name="Step_Ave" Units="nm" Value="{str(data_dict["key_PIX3_Step_Ave"])}"/>
                   <Data DataType="Numeric" Name="Step_3sigma" Units="nm" Value="{str(data_dict["key_PIX3_Step_3sigma"])}"/>
               </TestStep>

               <TestStep Name="SORTED_DATA" startDateTime="{pix1_time_str}" Status="Passed">
                   <Data DataType="Numeric" Name="STARTTIME_SORTED" Units="" Value="{str(data_dict["key_STARTTIME_SORTED_PIX1"])}"/>
                   <Data DataType="Numeric" Name="SORTNUMBER" Units="" Value="{str(data_dict["key_SORTNUMBER_PIX1"])}"/>
                   <Data DataType="String" Name="LotNumber_5" Value="{str(data_dict["key_Serial_Number"])}" CompOperation="LOG"/>
                   <Data DataType="String" Name="LotNumber_9" Value="{str(data_dict["key_LotNumber_9"])}" CompOperation="LOG"/>
               </TestStep>

               <TestEquipment>
                   <Item DeviceName="DryEtch_PIX1" DeviceSerialNumber="{str(data_dict["key_PIX1_Equipment"])}"/>
                   <Item DeviceName="DryEtch_PIX2" DeviceSerialNumber="{str(data_dict["key_PIX2_Equipment"])}"/>
                   <Item DeviceName="DryEtch_PIX3" DeviceSerialNumber="{str(data_dict["key_PIX3_Equipment"])}"/>
               </TestEquipment>

               <ErrorData/>
               <FailureData/>
               <Configuration/>
       </Result>
</Results>'''
        return xml_content
    except (KeyError, configparser.NoSectionError) as e:
        # Log error if needed, return None
        return None

# -----------------------------------------------------------------------------
# XML Generation: Write XML File
# -----------------------------------------------------------------------------
def write_xml_file(data_dict, xml_content, config, script_dir, log_file):
    """Writes the XML content to a file."""
    xml_full_path = "Unknown" # Initialize for error message
    try:
        Output_filepath_rel = config['General']['OutputFilePath']
        output_dir_path = os.path.abspath(os.path.join(script_dir, Output_filepath_rel))

        pix1_time_str = data_dict["key_PIX1_Start_Date_Time"]
        safe_testdate = pix1_time_str.replace(":", "-").replace(" ", "_")
        XML_File_Name = f'Site={config["General"]["Site"]},ProductFamily={config["General"]["ProductFamily"]},Operation={config["General"]["Operation"]},' + \
                        f'Partnumber={data_dict["key_Part_Number"]},Serialnumber={data_dict["key_Serial_Number"]},' + \
                        f'Testdate={safe_testdate}.xml'

        if not os.path.exists(output_dir_path):
            os.makedirs(output_dir_path)

        xml_full_path = os.path.join(output_dir_path, XML_File_Name)

        with open(xml_full_path, 'w', encoding="utf-8") as f:
            f.write(xml_content)
        return True, xml_full_path
    except Exception as e:
        msg = f"Error writing XML file {xml_full_path} for {data_dict.get('key_Serial_Number', 'Unknown Lot')}: {e}"
        if log_file: Log.Log_Error(log_file, msg)
        else: print(f"Error: {msg}")
        return False, None

# -----------------------------------------------------------------------------
# XML Generation: Update Start Rows
# -----------------------------------------------------------------------------
def update_and_backup_startrows(next_start_numbers, config, script_dir, log_file):
    """Updates StartRow files and performs backup."""
    msg = 'Updating StartRow files...'
    if log_file: Log.Log_Info(log_file, msg)
    else: print(msg)
    try:
        pix1_start_row_file = config['PIX1']['StartRowFile']
        pix2_start_row_file = config['PIX2']['StartRowFile']
        pix3_start_row_file = config['PIX3']['StartRowFile']
        pix1_start_row_path = os.path.abspath(os.path.join(script_dir, pix1_start_row_file))
        pix2_start_row_path = os.path.abspath(os.path.join(script_dir, pix2_start_row_file))
        pix3_start_row_path = os.path.abspath(os.path.join(script_dir, pix3_start_row_file))

        if not all(isinstance(n, (int, float)) and n >= 0 for n in next_start_numbers):
             raise ValueError(f"Invalid values in next_start_numbers: {next_start_numbers}")

        Row_Number_Func.next_start_row_number(pix1_start_row_path, int(next_start_numbers[0]))
        Row_Number_Func.next_start_row_number(pix2_start_row_path, int(next_start_numbers[1]))
        Row_Number_Func.next_start_row_number(pix3_start_row_path, int(next_start_numbers[2]))
        msg = f"Updated StartRow files: {pix1_start_row_file}={int(next_start_numbers[0])}, {pix2_start_row_file}={int(next_start_numbers[1])}, {pix3_start_row_file}={int(next_start_numbers[2])}"
        if log_file: Log.Log_Info(log_file, msg)
        else: print(msg)

        # Backup
        StartRowFileBaseDir = config.get('General', 'StartRowFileBaseDir', fallback=None)
        if StartRowFileBaseDir:
            if not os.path.exists(StartRowFileBaseDir):
                try:
                    os.makedirs(StartRowFileBaseDir)
                except OSError as e:
                    msg = f"Error creating backup directory {StartRowFileBaseDir}: {e}. Skipping backup."
                    if log_file: Log.Log_Error(log_file, msg)
                    else: print(f"Error: {msg}")
                    StartRowFileBaseDir = None # Prevent copy attempt

            if StartRowFileBaseDir:
                try:
                    shutil.copy(pix1_start_row_path, StartRowFileBaseDir)
                    shutil.copy(pix2_start_row_path, StartRowFileBaseDir)
                    shutil.copy(pix3_start_row_path, StartRowFileBaseDir)
                    msg = f"Copied StartRow files to backup directory: {StartRowFileBaseDir}"
                    if log_file: Log.Log_Info(log_file, msg)
                    else: print(msg)
                except Exception as e:
                    msg = f"Error copying StartRow files to backup directory {StartRowFileBaseDir}: {e}"
                    if log_file: Log.Log_Error(log_file, msg)
                    else: print(f"Error: {msg}")
        else:
             msg = "StartRowFileBaseDir not configured. Skipping backup."
             if log_file: Log.Log_Warning(log_file, msg)
             else: print(f"Warning: {msg}")

    except (ValueError, TypeError, KeyError, configparser.NoSectionError) as e:
         msg = f"Error preparing to update StartRow files: {e}"
         if log_file: Log.Log_Error(log_file, msg)
         else: print(f"Error: {msg}")
    except Exception as e:
        msg = f"Error updating or backing up StartRow files: {e}"
        if log_file: Log.Log_Error(log_file, msg)
        else: print(f"Error: {msg}")

# -----------------------------------------------------------------------------
# XML Generation: Main Function
# -----------------------------------------------------------------------------
def generate_xml_files_main(pix_data_list, config, script_dir, log_file, next_start_numbers):
    """Orchestrates the XML generation process."""
    msg = "--- Starting XML Generation Process ---"
    if log_file: Log.Log_Info(log_file, msg)
    else: print(msg)

    processed_count = 0
    error_counts = {"Invalid": 0, "Incomplete": 0, "Date": 0, "Type": 0, "XMLWrite": 0}

    for i, lot_data in enumerate(pix_data_list):
        data_dict, error_msg = validate_and_prepare_lot(lot_data, i, log_file)

        if error_msg:
            # Classify error for summary
            if "Invalid" in error_msg: error_counts["Invalid"] += 1
            elif "Incompleteness" in error_msg: error_counts["Incomplete"] += 1
            elif "Date Format" in error_msg: error_counts["Date"] += 1
            elif "calculating SORTED_DATA" in error_msg: error_counts["Date"] += 1
            else: error_counts["Type"] += 1 # Assume type/validation error otherwise
            # Log warning for skipped lot
            if log_file: Log.Log_Warning(log_file, f"Skipping Lot (Index {i}): {error_msg}")
            continue # Skip to next lot

        xml_content = create_xml_string(data_dict, config)
        if not xml_content:
             msg = f"Failed to create XML string for {data_dict.get('key_Serial_Number', 'Unknown Lot')}. Skipping."
             if log_file: Log.Log_Error(log_file, msg)
             else: print(f"Error: {msg}")
             error_counts["XMLWrite"] += 1 # Count as write error
             continue

        success, xml_path = write_xml_file(data_dict, xml_content, config, script_dir, log_file)
        if success:
            processed_count += 1
        else:
            error_counts["XMLWrite"] += 1

    msg = f"--- Finished XML Generation Process ---"
    if log_file: Log.Log_Info(log_file, msg)
    else: print(msg)
    summary_msg = f"XML Summary: Processed={processed_count}, Skipped(Invalid={error_counts['Invalid']}, Incomplete={error_counts['Incomplete']}, DateErr={error_counts['Date']}, TypeErr={error_counts['Type']}, XMLWriteErr={error_counts['XMLWrite']})"
    if log_file: Log.Log_Info(log_file, summary_msg)
    else: print(summary_msg)

    # Update start rows only if XMLs were generated
    if processed_count > 0:
        update_and_backup_startrows(next_start_numbers, config, script_dir, log_file)
    else:
        msg = "No XML files generated, skipping StartRow update."
        if log_file: Log.Log_Info(log_file, msg)
        else: print(msg)

# -----------------------------------------------------------------------------
# Main Execution Block
# -----------------------------------------------------------------------------
if __name__ == '__main__':

    if not import_custom_modules():
        sys.exit(1) # Stop if modules can't be imported

    print("Loading configuration...")
    config, script_dir = load_config()
    if not config or not script_dir:
        sys.exit(1) # Stop if config fails

    print("Setting up logging...")
    log_file = setup_logging(config, script_dir)
    # Continue even if logging fails, but log_file will be None

    pix_data_list = list()
    list_index_lot = dict()
    # Store next start row for PIX1, PIX2, PIX3
    next_start_numbers = [0, 0, 0]

    execution_successful = True

    try:
        # --- Data Collection Phase ---
        print("\n--- Starting Data Collection Phase ---")
        pix_types_to_process = ["PIX1", "PIX2", "PIX3"]
        collection_failed_critically = False
        for i, pix_type in enumerate(pix_types_to_process):
            start_num = collect_data_for_pix(pix_type, config, script_dir, log_file, pix_data_list, list_index_lot)
            next_start_numbers[i] = start_num
            # PIX1 failure (returns 0) is critical
            if start_num == 0 and pix_type == "PIX1":
                collection_failed_critically = True
                msg = "PIX1 data collection failed critically. Aborting."
                if log_file: Log.Log_Error(log_file, msg)
                else: print(f"Error: {msg}")
                break # Stop collecting further data
            elif start_num == 0:
                 # Failure in PIX2 or PIX3 is a warning, but continue
                 msg = f"{pix_type} data collection failed or returned no new data. Subsequent steps might be affected."
                 if log_file: Log.Log_Warning(log_file, msg)
                 else: print(f"Warning: {msg}")

        # --- XML Generation Phase ---
        if not collection_failed_critically:
            print("\n--- Starting XML Generation Phase ---")
            generate_xml_files_main(pix_data_list, config, script_dir, log_file, next_start_numbers)
        else:
            execution_successful = False # Mark as unsuccessful

    except Exception as e:
        execution_successful = False
        msg = f"An unexpected error occurred in the main execution block: {e}"
        detailed_error = traceback.format_exc()
        if log_file:
            Log.Log_Error(log_file, msg)
            Log.Log_Error(log_file, detailed_error)
        else:
            print(f"Error: {msg}")
            print(detailed_error)

    print("\n--- Process Finished ---")
    final_msg = 'Program End' + (' with errors.' if not execution_successful else '.')
    if log_file: Log.Log_Info(log_file, final_msg)
    else: print(final_msg)