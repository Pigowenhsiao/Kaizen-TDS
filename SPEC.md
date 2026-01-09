# SPEC.md - E1_Qrun configuration requirements

## 1) Scope
This specification defines the configuration-driven behavior for the E1_Qrun process. All runtime variables must be sourced from INI. Python code must not hard-code values that can be expressed in INI.

## 2) Runtime requirements
- Python: 3.12.7
- Required packages: pandas, openpyxl
- Excel file types: .xlsx, .xlsm

## 3) Inputs and outputs
- Inputs: Excel files discovered by path + pattern rules in INI.
- Outputs:
  - CSV summary file (one row per source file).
  - Pointer XML file referencing the CSV.
  - Optional: dedup registry (SQLite).

## 4) Configuration policy
- Every configurable value must exist in INI.
- Python may only apply logic using INI-provided values.
- Defaults must be defined in INI, not in Python.

## 5) INI schema (sections and keys)

[Basic_info]
- output_mode: csv|xml|both
- Site
- ProductFamily
- Operation
- TestStation
- file_name_patterns (multi-line list)
- Retention_date (days)
- Tool_Name (fallback tester name)
- Part_Number (default part number)

[Paths]
- input_paths (multi-line list)
- running_rec
- output_path (XML output)
- CSV_path
- intermediate_data_path
- log_path

[FileNaming]
- filename_regex (regex for wafer_id and lot_id capture)
- exclude_substrings (comma-separated, e.g. "- Copy")
- exclude_prefixes (comma-separated, e.g. "~$")
- allowed_extensions (comma-separated, e.g. xlsx,xlsm)

[Excel]
- sheet_name
- data_columns (A1 range columns, e.g. D:KT)
- main_skip_rows
- main_nrows

[StartDateTime]
- sheet_name
- cell
- datetime_format
- fallback_mode: file_mtime|now|blank
- output_format

[TesterId]
- sheet_name
- cell
- fallback_value

[DataFields]
- fields (multi-line list): key:col:dtype
  - col >= 0: DataFrame 0-based column index
  - col = -1: assigned by Python logic
  - col = cell_<A1>: read from specific cell (if used)

[Stats]
- suffixes (comma-separated, e.g. MAX,MIN,AVG,STD)
- ddof (int)
- exclude_fields (comma-separated base names)
- fill_empty_with_zero (true|false)

[Dedup]
- enable_dedup (true|false)
- dedup_db_path
- fingerprint_mode: stat|sha256
- debug_discovery (true|false)
- debug_limit_10_files (true|false)
- db_filter_by_mtime (true|false)
- db_filter_days (int)

[Database]
- db_connection_string
- server
- database
- username
- password
- driver
- lookup_mapping (optional mapping policy for DB fields)

[WaiveLengthCategoryMapping]
- LotRule -> Waive_Leng_Cate mapping

[WaiveLengthCategory]
- missing_rule_behavior: unknown|skip_file|error
- unknown_value

[CSV]
- encoding (e.g. utf-8-sig)
- include_source_file (true|false)
- column_order (comma-separated; empty means default order)

[XML]
- file_name_template
- result_status
- header_defaults (SerialNumber, PartNumber, Operator, LotNumber)
- test_step_defaults (Status)
- data_table_name
- data_type
- comp_operation

## 6) CSV output schema
Required columns:
- Serial_Number
- Start_Date_Time
- Part_Number
- TESTER_ID
- Waive_Leng_Cate
- LotNumber_9
- Source_File (if include_source_file = true)

Optional columns:
- DB lookup fields (e.g. DB_LOOKUP_*)
- Statistics fields: {BaseName}_{MAX|MIN|AVG|STD}

## 7) XML pointer schema
- Root: Results/Result/Header/TestStep/Data
- Values and defaults are defined by [XML] in INI
- Data.Value must point to the CSV output path

## 8) File discovery rules
- Scan [Paths].input_paths
- Filter by [FileNaming].allowed_extensions
- Match [Basic_info].file_name_patterns
- Validate with [FileNaming].filename_regex
- Exclude by [FileNaming].exclude_substrings and exclude_prefixes

## 9) Excel read rules
- Read main table from [Excel].sheet_name using [Excel].data_columns
- Skip rows: [Excel].main_skip_rows
- Read rows: [Excel].main_nrows
- Read Start_Date_Time and TESTER_ID from [StartDateTime] and [TesterId]

## 10) Dedup rules
- When enabled, compute fingerprint using [Dedup].fingerprint_mode
- Persist to SQLite at [Dedup].dedup_db_path
- Optional mtime filter by [Dedup].db_filter_by_mtime and db_filter_days

## 11) DB lookup rules
- DB access uses [Database] settings
- Mapped values populate DB_LOOKUP_* fields
- Part_Number can be overridden by DB lookup when provided

## 12) Example INI snippet (template)
[FileNaming]
filename_regex = ^(N[A-Z0-9]{7})_(N[A-Z0-9]{4}).*?\.(xlsx|xlsm)$
exclude_substrings = - Copy
exclude_prefixes = ~$
allowed_extensions = xlsx,xlsm

[TesterId]
sheet_name = <sheet_name>
cell = AY23
fallback_value = <Tool_Name>

[Stats]
suffixes = MAX,MIN,AVG,STD
ddof = 1
exclude_fields = sigmaDeltaG_m0_5V
fill_empty_with_zero = true
