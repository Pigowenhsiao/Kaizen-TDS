import pandas as pd

# 讀取 Excel 檔案
file_path = r"Z:\MOCVD\MOCVD過去プログラム\F3炉\FE8874～.xlsx"
excel_data = pd.ExcelFile(file_path)

# 用於存儲 AL4 欄位的資料
unique_data = set()

# 迭代每個 Sheet 並讀取 AL4 欄位資料
for sheet_name in excel_data.sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="U")
    al4_data = df.iloc[2]  # AL4 欄位的資料在第 4 行 (index 為 3)
    al4_data_tuple = tuple(al4_data)  # 將 Series 轉換為 tuple
    print('sheet_name:', sheet_name,'al4_data:', al4_data_tuple)
    unique_data.add(al4_data_tuple)

# 將不重複的資料寫入 CSV 檔案
unique_df = pd.DataFrame(unique_data, columns=["AL4 Data"])
unique_df.to_csv("recipe.csv", index=False)