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
import ast

# Mock objects for modules if they are not available for standalone execution
class MockLog:
    def Log_Info(self, file, msg): print(f"INFO: {msg}")
    def Log_Error(self, file, msg): print(f"ERROR: {msg}")
    def Log_Init(self, path): return "mock_log_file.log"
Log = MockLog()

class IniSettings:
    """Class to hold all settings read from the INI file"""
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
        self.log_path = ""
        self.running_rec = ""
        self.backup_running_rec_path = ""
        self.sheet_name = ""
        
        # Excel specific
        self.header_rows = None
        self.data_start_row = 1
        
        # Data Processing specific
        self.enable_pivot = False
        self.pivot_group_by_column = ""
        self.pivot_target_columns = []
        self.pivot_prefix_map = {}
        self.static_column_map = {}
        self.final_column_order = []


def read_ini_settings(config_path):
    """Reads settings from the specified INI file and returns an IniSettings object."""
    config = ConfigParser()
    config.read(config_path, encoding='utf-8')
    settings = IniSettings()

    # Basic Info
    basic_info = config['Basic_info']
    settings.site = basic_info.get('Site')
    settings.product_family = basic_info.get('ProductFamily')
    settings.operation = basic_info.get('Operation')
    settings.test_station = basic_info.get('TestStation')
    settings.file_name_patterns = [p.strip() for p in basic_info.get('file_name_patterns', '').split(',')]
    settings.retention_date = basic_info.getint('retention_date', 30)

    # Paths
    paths = config['Paths']
    settings.input_paths = [p.strip() for p in paths.get('input_paths', '').split(',')]
    settings.output_path = paths.get('output_path')
    settings.csv_path = paths.get('CSV_path')
    settings.log_path = paths.get('log_path')
    settings.running_rec = paths.get('running_rec')

    # Excel
    excel = config['Excel']
    settings.sheet_name = excel.get('sheet_name')
    header_rows_str = excel.get('header_rows')
    if header_rows_str:
        settings.header_rows = [int(i.strip()) for i in header_rows_str.split(',')]
    settings.data_start_row = excel.getint('data_start_row', 1)

    # DataProcessing (for special pivot logic)
    if 'DataProcessing' in config:
        processing = config['DataProcessing']
        settings.enable_pivot = processing.getboolean('enable_pivot', False)
        if settings.enable_pivot:
            settings.pivot_group_by_column = processing.get('pivot_group_by_column')
            settings.pivot_target_columns = [p.strip() for p in processing.get('pivot_target_columns').split(',')]
            
            # Use ast.literal_eval for safe parsing of dictionary-like strings
            settings.pivot_prefix_map = ast.literal_eval(processing.get('pivot_prefix_map', '{}'))
            settings.static_column_map = ast.literal_eval(processing.get('static_column_map', '{}'))
    
    # Final CSV Columns
    if 'Final_CSV_Columns' in config:
        settings.final_column_order = [
            c.strip() for c in config['Final_CSV_Columns'].get('column_order', '').split(',')
        ]

    return settings

def process_special_pivot(df, settings, log_file):
    """
    Transforms the DataFrame from long format to wide format based on pivot settings.
    """
    Log.Log_Info(log_file, "Starting special pivot processing...")
    
    # 1. Flatten multi-level header
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = ['_'.join(map(str, col)).strip().replace('Unnamed: ', 'Unnamed_').replace('_level_1', '') for col in df.columns.values]

    # 2. Identify the correct group-by column name after flattening
    group_by_col_flat = None
    for col in df.columns:
        if settings.pivot_group_by_column in col:
            group_by_col_flat = col
            break
    if not group_by_col_flat:
        raise ValueError(f"Pivot group-by column '{settings.pivot_group_by_column}' not found in DataFrame columns.")

    # 3. Filter out rows where the group-by column is NaN
    df.dropna(subset=[group_by_col_flat], inplace=True)
    if df.empty:
        Log.Log_Info(log_file, "DataFrame is empty after dropping NaNs in group-by column. No data to process.")
        return None

    Log.Log_Info(log_file, f"Grouping data by column: '{group_by_col_flat}'")

    # 4. Create the wide format DataFrame by pivoting
    df['pivot_seq'] = df.groupby(group_by_col_flat).cumcount() + 1
    
    pivot_df = df.pivot(index=group_by_col_flat, columns='pivot_seq', values=settings.pivot_target_columns)
    
    pivot_df.columns = [f"{settings.pivot_prefix_map.get(col[0], col[0])}_{col[1]}" for col in pivot_df.columns]
    pivot_df.reset_index(inplace=True)

    # 5. Get the static columns that are the same for each group
    static_cols_df = df.groupby(group_by_col_flat).first().reset_index()
    
    final_static_cols = {}
    for source_col, target_col in settings.static_column_map.items():
        if source_col in static_cols_df.columns:
            final_static_cols[target_col] = static_cols_df[source_col]
        else:
            Log.Log_Error(log_file, f"Static column '{source_col}' not found in source data. Available columns: {static_cols_df.columns.to_list()}")
            
    final_df = pd.DataFrame(final_static_cols)

    # 6. Merge pivoted data with static data
    final_df = pd.merge(final_df, pivot_df, left_on='Serial_Number', right_on=group_by_col_flat, how='left')
    
    if group_by_col_flat in final_df.columns and group_by_col_flat != 'Serial_Number':
        final_df.drop(columns=[group_by_col_flat], inplace=True)
        
    # 7. Add basic info
    final_df['Operation'] = settings.operation
    final_df['TestStation'] = settings.test_station

    # 8. Reorder columns
    if settings.final_column_order:
        existing_cols = [col for col in settings.final_column_order if col in final_df.columns]
        final_df = final_df[existing_cols]

    Log.Log_Info(log_file, f"Successfully pivoted data. Shape of final DataFrame: {final_df.shape}")
    return final_df


def process_excel_file(filepath, settings, log_file, output_csv_path):
    """Processes a single Excel file based on the provided settings."""
    try:
        Log.Log_Info(log_file, f"Reading Excel file: {filepath}, Sheet: {settings.sheet_name}")
        
        header_arg = settings.header_rows if settings.header_rows else 0
        
        df = pd.read_excel(filepath, sheet_name=settings.sheet_name, header=header_arg)
        Log.Log_Info(log_file, f"Successfully read {len(df)} rows from Excel.")
        
        final_df = None
        if settings.enable_pivot:
            final_df = process_special_pivot(df, settings, log_file)
        else:
            Log.Log_Info(log_file, "Standard processing logic would run here.")
            # Placeholder for original logic
            final_df = df 

        if final_df is not None and not final_df.empty:
            file_exists = os.path.isfile(output_csv_path)
            final_df.to_csv(output_csv_path, mode='a', header=not file_exists, index=False, encoding='utf-8-sig')
            Log.Log_Info(log_file, f"Appended {len(final_df)} rows to {output_csv_path}")

    except Exception:
        Log.Log_Error(log_file, f"Error processing Excel file {filepath}: {traceback.format_exc()}")


def generate_pointer_xml(output_path, csv_path, settings, log_file):
    """Generates the XML pointer file."""
    if not os.path.exists(csv_path):
        Log.Log_Error(log_file, f"CSV file not found, cannot generate XML pointer: {csv_path}")
        return

    try:
        root = ET.Element("TDS_INSTRUCTION_FILE")
        ET.SubElement(root, "SITE").text = settings.site
        ET.SubElement(root, "PRODUCT_FAMILY").text = settings.product_family
        ET.SubElement(root, "OPERATION").text = settings.operation
        ET.SubElement(root, "TEST_STATION").text = settings.test_station
        ET.SubElement(root, "DATA_FILE_PATH").text = csv_path
        ET.SubElement(root, "DATA_FILE_TYPE").text = "CSV"
        
        xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="   ")
        
        xml_filename = f"{settings.operation}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.xml"
        xml_filepath = os.path.join(output_path, xml_filename)
        
        with open(xml_filepath, "w", encoding="utf-8") as f:
            f.write(xml_str)
        Log.Log_Info(log_file, f"Successfully generated XML pointer file: {xml_filepath}")

    except Exception:
        Log.Log_Error(log_file, f"Error generating XML pointer file: {traceback.format_exc()}")


def main():
    """Main function to find INI files and process them."""
    ini_files = glob.glob('*.ini')
    log_file = Log.Log_Init("Main_Process")
    
    Log.Log_Info(log_file, f"Found {len(ini_files)} INI configuration files.")

    for ini_path in ini_files:
        try:
            Log.Log_Info(log_file, f"--- Starting processing for config file: {ini_path} ---")
            settings = read_ini_settings(ini_path)
            
            csv_filename = f"{settings.operation}_{datetime.now().strftime('%Y%m%d')}.csv"
            csv_filepath_for_this_ini = os.path.join(settings.csv_path, csv_filename)

            source_files_found = False
            for input_path in settings.input_paths:
                for pattern in settings.file_name_patterns:
                    search_pattern = os.path.join(input_path, pattern)
                    Log.Log_Info(log_file, f"Searching for files with pattern: {search_pattern}")
                    
                    matching_files = glob.glob(search_pattern)
                    if not matching_files:
                        continue
                    
                    source_files_found = True
                    Log.Log_Info(log_file, f"Found {len(matching_files)} matching file(s).")
                    
                    for file_path in matching_files:
                        process_excel_file(file_path, settings, log_file, csv_filepath_for_this_ini)
            
            if not source_files_found:
                Log.Log_Info(log_file, "No matching source files found for this configuration.")

            if os.path.exists(csv_filepath_for_this_ini):
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
            Log.Log_Error(log_file, error_message)

if __name__ == '__main__':
    print("Updated script 'Univeral Script_v2.py' and 'HL_Combined.ini' are ready.")
    print("Please save the INI content as 'HL_Combined.ini'.")
    print("Review the paths in the INI file, and then run the Python script.")