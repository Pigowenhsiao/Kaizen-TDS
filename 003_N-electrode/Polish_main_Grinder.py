import glob
import os
from polish_processor import PolishProcessor
from configs import configs
from datetime import date
import sys
import shutil
import openpyxl

sys.path.append('../MyModule')
import Log

# ログファイルのパスを設定
log_folder_name = str(date.today())
import configparser
log_folder_path = f"../Log/{log_folder_name}"
log_file = f"{log_folder_path}/003_N-electrode.log"

# ログフォルダが存在しない場合は作成
if not os.path.exists(log_folder_path):
    os.makedirs(log_folder_path)

# プロセッサを初期化
output_dir="C:/Users/hsi67063/Download/" # テストルート
#output_dir="//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/" # 正式ルート

processor = PolishProcessor(output_dir=output_dir)

# 讀取ini檔案
config = configparser.ConfigParser()
config.read("049 TAK_PLX/Config_TAK_PLX_9.ini", encoding='utf-8')

# 検索パスとフォルダ名を設定
base_path = "Z:/研磨/2025年/"
folder_pattern = "*年*月"

# 条件に一致するフォルダを検索
Log.Log_Info(log_file, "Searching for target folders...")
folders = glob.glob(os.path.join(base_path, folder_pattern))
if not folders:
    Log.Log_Error(log_file, "No matching folders found! Please check the path settings.")
    exit()

# 最新のフォルダを選択
folders.sort()
target_folder = folders[-1]
processed_folder = config.get('Paths', 'copy_destination_path')
#processed_folder = os.path.join(target_folder, "処理済みフォルダ")
Log.Log_Info(log_file, f"Target folder selected: {target_folder}")

# 処理後のフォルダが存在しない場合は作成
if not os.path.exists(processed_folder):
    os.makedirs(processed_folder)
    Log.Log_Info(log_file, f"Created processed folder: {processed_folder}")

# ファイル検索ルールを定義
file_pattern = "*.xls*"
Log.Log_Info(log_file, "Searching for Excel files...")
files = glob.glob(os.path.join(target_folder, file_pattern))

# 条件に一致するファイルが見つからない場合
if not files:
    Log.Log_Error(log_file, "No matching Excel files found in the target folder.")
    exit()

# すべてのファイルを処理
for filepath in files:
    Log.Log_Info(log_file, f"Processing file: {filepath}")
    
    # 複製檔案到指定路徑
    try:
        destination_path = os.path.join(processed_folder, os.path.basename(filepath))
        shutil.copy(filepath, destination_path)
        Log.Log_Info(log_file, f"File copied to: {destination_path}")
        filepath = destination_path  # 後續處理使用複製後的檔案
    except Exception as e:
        Log.Log_Error(log_file, f"Error copying file {filepath}: {e}")
        continue

    try:
        # 載入Excel檔案
        wb = openpyxl.load_workbook(filepath, data_only=True)
        
        for i in range(1, 6):
            # ファイル名に基づいて設定を選択
            if i == 1:
                config = configs["rough_polished"]
            elif i == 2:
                config = configs["wax_thickness"]
            elif i == 3:
                config = configs["mirror_polished"]
            elif i == 4:
                config = configs["etched_thickness"]
            elif i == 5:
                config = configs["initial_wafer_thickness"]
            else:
                Log.Log_Error(log_file, f"Unknown file type for {filepath}. Skipping...")
                continue
                
            # 處理模組を呼ぶ
            try:
                processor.process_file(filepath=filepath, config=config, workbook=wb)
                Log.Log_Info(log_file, f"Successfully processed config {config['operation']} for: {filepath}")
            except Exception as e:
                Log.Log_Error(log_file, f"Error processing config {config['operation']} for {filepath}: {e}")
                continue
    
        # 關閉Excel檔案
        wb.close()
    except Exception as e:
        Log.Log_Error(log_file, f"Error loading or processing Excel file {filepath}: {e}")