import os
import re
import sys
import shutil
import logging
import traceback
from datetime import datetime, date
from configparser import ConfigParser
from pathlib import Path
from typing import Dict, Any, List

import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
from xml.dom import minidom
from dateutil.relativedelta import relativedelta

# Append MyModule path
sys.path.append("../MyModule")
import Log
import SQL
import Convert_Date
import Row_Number_Func


class IniSettings:
    """Hold all settings loaded from INI"""

    def __init__(self) -> None:
        # Basic Info
        self.site: str = ""
        self.product_family: str = ""
        self.operation: str = ""
        self.test_station: str = ""
        self.retention_date: int = 30
        self.file_name_patterns: List[str] = []
        self.tool_name: str = ""

        # Paths
        self.input_paths: List[str] = []
        self.output_path: str = ""
        self.csv_path: str = ""
        self.intermediate_data_path: str = ""
        self.log_path: str = ""
        self.running_rec: str = ""

        # Excel
        self.sheet_name: List[Any] = []  # 可為 int 或 str
        self.data_columns: str = ""
        self.main_skip_rows: int = 0

        # Database
        self.db_server: str = ""
        self.db_database: str = ""
        self.db_username: str = ""
        self.db_password: str = ""
        self.db_driver: str = ""

        # DataFields
        self.field_map: Dict[str, Dict[str, str]] = {}


def setup_logging(log_dir: str, operation_name: str) -> str:
    """Set up daily rotating log"""
    log_folder = os.path.join(log_dir, str(date.today()))
    os.makedirs(log_folder, exist_ok=True)
    log_file = os.path.join(log_folder, f"{operation_name}.log")

    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)

    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    return log_file


def _read_and_parse_ini_config(config_file_path: str) -> ConfigParser:
    """Read INI"""
    config = ConfigParser()
    config.read(config_file_path, encoding="utf-8")
    return config


def _parse_fields_map_from_lines(fields_lines: List[str]) -> Dict[str, Dict[str, str]]:
    """Parse [DataFields] mapping"""
    fields = {}
    for line in fields_lines:
        if ":" in line and not line.strip().startswith("#"):
            try:
                key, col_str, dtype_str = map(str.strip, line.split(":", 2))
                fields[key] = {"col": col_str, "dtype": dtype_str}
            except ValueError:
                continue
    return fields


def _extract_settings_from_config(config: ConfigParser) -> IniSettings:
    """Extract all INI settings"""
    s = IniSettings()

    # Basic Info
    s.site = config.get("Basic_info", "Site")
    s.product_family = config.get("Basic_info", "ProductFamily")
    s.operation = config.get("Basic_info", "Operation")
    s.test_station = config.get("Basic_info", "TestStation")
    s.retention_date = config.getint("Basic_info", "Retention_date", fallback=30)
    s.file_name_patterns = [
        x.strip() for x in config.get("Basic_info", "file_name_patterns").split(",")
    ]
    s.tool_name = config.get("Basic_info", "Tool_Name")

    # Paths
    s.input_paths = [x.strip() for x in config.get("Paths", "input_paths").split(",")]
    s.output_path = config.get("Paths", "output_path")
    s.csv_path = config.get("Paths", "CSV_path")
    s.intermediate_data_path = config.get("Paths", "intermediate_data_path")
    s.log_path = config.get("Paths", "log_path")
    s.running_rec = config.get("Paths", "running_rec")

    # Excel (支援數字或名稱)
    sheet_raw = [x.strip() for x in config.get("Excel", "sheet_name").split(",")]
    sheet_list: List[Any] = []
    for x in sheet_raw:
        if x.isdigit():
            sheet_list.append(int(x))
        else:
            sheet_list.append(x)
    s.sheet_name = sheet_list
    s.data_columns = config.get("Excel", "data_columns")
    s.main_skip_rows = config.getint("Excel", "main_skip_rows")

    # Database
    s.db_server = config.get("Database", "server")
    s.db_database = config.get("Database", "database")
    s.db_username = config.get("Database", "username")
    s.db_password = config.get("Database", "password")
    s.db_driver = config.get("Database", "driver")

    # DataFields
    fields_lines = config.get("DataFields", "fields").splitlines()
    s.field_map = _parse_fields_map_from_lines(fields_lines)

    return s


def write_to_csv(csv_filepath: str, dataframe: pd.DataFrame, log_file: str) -> bool:
    """Write DataFrame to CSV"""
    try:
        file_exists = os.path.isfile(csv_filepath)
        dataframe.to_csv(
            csv_filepath,
            mode="a",
            header=not file_exists,
            index=False,
            encoding="utf-8-sig",
        )
        Log.Log_Info(log_file, f"CSV written: {csv_filepath}")
        return True
    except Exception as e:
        Log.Log_Error(log_file, f"CSV write failed: {e}")
        return False


def generate_pointer_xml(output_path: str, csv_path: str, settings: IniSettings, log_file: str) -> None:
    """Generate pointer XML file (CVD style)"""
    try:
        os.makedirs(output_path, exist_ok=True)
        now_iso = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        serial_no = Path(csv_path).stem

        xml_file_name = (
            f"Site={settings.site},ProductFamily={settings.product_family},"
            f"Operation={settings.operation},Partnumber=UNKNOWPN,"
            f"Serialnumber={serial_no},Testdate={now_iso}.xml"
        ).replace(":", ".")

        xml_file_path = os.path.join(output_path, xml_file_name)

        results = ET.Element(
            "Results",
            {
                "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
                "xmlns:xsd": "http://www.w3.org/2001/XMLSchema",
            },
        )
        result = ET.SubElement(
            results,
            "Result",
            startDateTime=now_iso,
            endDateTime=now_iso,
            Result="Passed",
        )

        ET.SubElement(
            result,
            "Header",
            SerialNumber=serial_no,
            PartNumber="UNKNOWPN",
            Operation=settings.operation,
            TestStation=settings.test_station,
            Operator="NA",
            StartTime=now_iso,
            Site=settings.site,
            LotNumber="",
        )

        test_step = ET.SubElement(
            result,
            "TestStep",
            Name=settings.operation,
            startDateTime=now_iso,
            endDateTime=now_iso,
            Status="Passed",
        )

        ET.SubElement(
            test_step,
            "Data",
            DataType="Table",
            Name=f"tbl_{settings.operation.upper()}",
            Value=str(csv_path),
            CompOperation="LOG",
        )

        xml_str = minidom.parseString(ET.tostring(results)).toprettyxml(
            indent="  ", encoding="utf-8"
        )
        with open(xml_file_path, "wb") as f:
            f.write(xml_str)
        Log.Log_Info(log_file, f"Pointer XML generated: {xml_file_path}")

    except Exception as e:
        Log.Log_Error(log_file, f"XML generation failed: {e}")


def process_excel_file(filepath_str: str, settings: IniSettings, log_file: str, csv_filepath: str) -> None:
    """Read Excel, extract N-series serials, output CSV"""
    filepath = Path(filepath_str)
    Log.Log_Info(log_file, f"Start processing {filepath.name}")

    all_data: List[pd.DataFrame] = []
    for sheet in settings.sheet_name:
        try:
            df = pd.read_excel(
                filepath,
                header=None,
                sheet_name=sheet,
                usecols=settings.data_columns,
                skiprows=settings.main_skip_rows,
            )
            df = df.dropna(how="all")
            all_data.append(df)
            Log.Log_Info(log_file, f"Read sheet {sheet}, {df.shape[0]} rows")
        except Exception as e:
            Log.Log_Error(log_file, f"Failed reading sheet {sheet}: {e}")

    if not all_data:
        Log.Log_Info(log_file, "No valid data read")
        return

    df_all = pd.concat(all_data, ignore_index=True)
    df_all.columns = range(df_all.shape[1])

    # 日期與空值過濾
    df_all = df_all.replace("nan", np.nan).dropna(subset=[0, 7])
    df_all[0] = pd.to_datetime(df_all[0], errors="coerce")
    df_all = df_all[
        df_all[0] >= datetime.now() - relativedelta(days=settings.retention_date)
    ]
    if df_all.empty:
        Log.Log_Info(log_file, "No data after date filter")
        return

    # 提取所有 N 開頭序號
    def extract_serials(cell: Any) -> List[str]:
        if isinstance(cell, str):
            return re.findall(r"N\d+", cell)
        return []

    df_all["Serial_List"] = df_all[7].apply(extract_serials)
    df_all = df_all.explode("Serial_List").reset_index(drop=True)
    df_all = df_all.rename(
        columns={
            0: "Start_Date_Time",
            2: "Operator",
            9: "Reflectance_Front",
            10: "Reflectance_Back",
        }
    )

    if df_all["Serial_List"].isna().all():
        Log.Log_Info(log_file, "No N-series serial numbers found")
        return

    # DB 查詢
    conn, cursor = None, None
    try:
        conn, cursor = SQL.connSQL()
        if conn is None:
            Log.Log_Error(log_file, "DB connection failed")
            return

        def get_db_info(serial: str) -> pd.Series:
            part_num, lot9 = SQL.selectSQL(cursor, str(serial))
            return pd.Series([part_num, lot9])

        df_all[["Part_Number", "LotNumber_9"]] = df_all["Serial_List"].apply(
            get_db_info
        )
    finally:
        if conn:
            SQL.disconnSQL(conn, cursor)

    # 計算排序欄位
    df_all["Dev"] = settings.tool_name
    df_all["SORTNUMBER"] = range(1, len(df_all) + 1)
    df_all["STARTTIME_SORTED"] = (
        df_all["SORTNUMBER"].astype(float) / 10**6
        + datetime.now().toordinal()
    )
    df_all["Serial_Number"] = df_all["Serial_List"]

    # 從 ini 動態決定輸出欄位順序
    ordered_keys = list(settings.field_map.keys())
    csv_columns = [k.replace("key_", "") for k in ordered_keys]

    # 確保所有欄位都存在
    for col in csv_columns:
        if col not in df_all.columns:
            df_all[col] = ""

    final_cols = [c for c in csv_columns if c in df_all.columns]
    df_final = df_all[final_cols]

    Log.Log_Info(log_file, f"Writing CSV with {len(df_final)} rows...")
    write_to_csv(csv_filepath, df_final, log_file)


def main() -> None:
    """Main execution"""
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    log_file = setup_logging("../Log/", "Coating_MG_Reflectance")
    Log.Log_Info(log_file, "===== Script Start =====")

    ini_files = [f for f in os.listdir(".") if f.endswith(".ini")]
    if not ini_files:
        Log.Log_Info(log_file, "No ini found, exit.")
        print("No ini found.")
        return

    for ini_path in ini_files:
        try:
            config = _read_and_parse_ini_config(ini_path)
            settings = _extract_settings_from_config(config)
            log_file = setup_logging(settings.log_path, settings.operation)

            Log.Log_Info(log_file, f"Processing INI: {ini_path}")
            Path(settings.csv_path).mkdir(parents=True, exist_ok=True)

            timestamp = datetime.now().strftime("%Y_%m_%dT%H.%M.%S")
            csv_file = Path(settings.csv_path) / f"{settings.operation}_{timestamp}.csv"

            for input_dir in settings.input_paths:
                for pattern in settings.file_name_patterns:
                    files = list(Path(input_dir).glob(pattern))
                    if not files:
                        continue
                    latest = max(files, key=os.path.getmtime)
                    dst_path = shutil.copy(latest, settings.intermediate_data_path)
                    process_excel_file(dst_path, settings, log_file, str(csv_file))

            generate_pointer_xml(settings.output_path, str(csv_file), settings, log_file)
            Log.Log_Info(log_file, f"Finished INI: {ini_path}")

        except Exception:
            error_message = f"Error in {ini_path}: {traceback.format_exc()}"
            Log.Log_Error(log_file, error_message)

    Log.Log_Info(log_file, "===== Script End =====")
    print("✅ All INI processed.")


if __name__ == "__main__":
    main()
