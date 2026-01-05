import glob
import os
from polish_processor import PolishProcessor
from configs import configs
from datetime import date
import sys
import shutil

sys.path.append('../MyModule')
import Log


# 設定日誌檔案路徑
log_folder_name = str(date.today())
log_folder_path = f"../Log/{log_folder_name}"
log_file = f"{log_folder_path}/003_N-electrode.log"

# 如果日誌資料夾不存在，則創建
if not os.path.exists(log_folder_path):
    os.makedirs(log_folder_path)

# 初始化處理器
output_dir="C:/Users/hsi67063/Downloads/処理済みフォルダ/" #test rounte
#output_dir="//li.lumentuminc.net/data/SAG/TDS/Data/Files to Insert/XML/"" #formal route
processor = PolishProcessor(output_dir=output_dir)

# 設定搜尋路徑與資料夾名稱
base_path = "Z:/研磨/2025年/"
folder_pattern = "*年*月"

# 搜尋符合條件的資料夾
Log.Log_Info(log_file, "Searching for target folders...")
folders = glob.glob(os.path.join(base_path, folder_pattern))
if not folders:
    Log.Log_Error(log_file, "No matching folders found! Please check the path settings.")
    exit()

# 選擇最新的資料夾
folders.sort()
target_folder = folders[-1]
processed_folder = "../DataFile/003_N-electrode/"
#processed_folder = os.path.join(target_folder, "処理済みフォルダ")
Log.Log_Info(log_file, f"Target folder selected: {target_folder}")

# 如果處理後的資料夾不存在，則創建
if not os.path.exists(processed_folder):
    os.makedirs(processed_folder)
    Log.Log_Info(log_file, f"Created processed folder: {processed_folder}")


# 定義檔案搜尋規則
file_pattern = "*.xls*"
Log.Log_Info(log_file, "Searching for Excel files...")
files = glob.glob(os.path.join(target_folder, file_pattern))

# 如果沒有找到符合條件的檔案
if not files:
    Log.Log_Error(log_file, "No matching Excel files found in the target folder.")
    exit()

# 處理所有檔案
for filepath in files:
    Log.Log_Info(log_file, f"Processing file: {filepath}")
    
    for i in range(1, 6):
        # 根據檔名判斷使用哪個配置
        if i == 1:
            config = configs["rough_polished"]
        elif i == 2:
            config = configs["wax_thickness"]
        elif i == 3:
            config = configs["mirror_polished"] # etched_thickness initial_wafer_thickness
        elif i == 4:
            config = configs["etched_thickness"] # etched_thickness initial_wafer_thickness
        elif i == 5:
            config = configs["initial_wafer_thickness"] # etched_thickness initial_wafer_thickness                
        else:
            Log.Log_Error(log_file, f"Unknown file type for {filepath}. Skipping...")
            continue

        # 呼叫處理模組
        try:
            processor.process_file(filepath=filepath, config=config)
            Log.Log_Info(log_file, f"Successfully processed: {filepath}")
        except Exception as e:
            Log.Log_Error(log_file, f"Error processing file {filepath}: {e}")
            continue

        # 複製處理完成的檔案到 "処理済みフォルダ"
        try:
            processed_path = os.path.join(processed_folder, os.path.basename(filepath))
            shutil.copy(filepath, processed_path)
            Log.Log_Info(log_file, f"File copied to: {processed_path}")
        except Exception as e:
            Log.Log_Error(log_file, f"Error copying file {filepath}: {e}")
