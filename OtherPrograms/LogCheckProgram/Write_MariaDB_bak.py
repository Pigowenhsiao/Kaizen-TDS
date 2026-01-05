########## 直近のLogフォルダ内のLogファイルを取得し、正常終了したかどうかを Maria DB に書き込む ##########

import pymysql.cursors
import glob
import os
import datetime
import time


# ----- MySQL Connection -----
# Connect to the database
connection = pymysql.connect(host='SAGAPPTDSPRD01.li.lumentuminc.net',
                             user='user01',
                             password='0claro!db',
                             db='db01',
                             cursorclass=pymysql.cursors.DictCursor)


# ----- 直近に作成したLogフォルダを取得 -----
Log_Path = "../../Log/"
Log_Folder = sorted([Folder for Folder in os.listdir(Log_Path) if os.path.isdir(os.path.join(Log_Path, Folder))], reverse=True)[0]
Path = os.path.join(Log_Path, Log_Folder)


# ----- 直近に作成したLogが今日のLogかどうかを調べる -----
Today = str(datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")).split()[0].replace("/", "-")
Log_FolderName = str(Path).split("/")[-1]
if Log_FolderName == Today:
    print("OK")
else:
    print("The Log isn't up to date.")
    exit()


# ----- プログラムの開始時間/終了時間を得る -----

# 更新日時が一番早いファイルと遅いファイルのパスを取り出す
Time_FileName = [[time.localtime(os.path.getmtime(File)), File] for File in glob.glob(Path + "/*.log")]
Time_FileName.sort()
Fast_File, Slow_File = Time_FileName[0][1], Time_FileName[-1][1]


# Fast_FileからはProram Startを見つけ、その中で一番遅い時間のものを探す
StartTime = "1900-01-01 00:00:00"
with open(Fast_File) as f:
    for line in f:
        if "Program Start" in line:
            StartTime = max(StartTime, line[10:29])


# Slow_FileからはProgram Endを見つけ、その中で一番遅い時間のものを探す
EndTime = "1900-01-01 00:00:00"
with open(Slow_File) as f:
    for line in f:
        if "Program End" in line:
            EndTime = max(EndTime, line[10:29])


# ----- エラー終了したOperationを格納 -----
Error = []


# ---- DBに書き込む情報を得る -----
Today = str(datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S"))
User = os.getlogin()


sql = "INSERT INTO kaizen_log VALUES (0, " + "\"" + str(StartTime) + "\", " + "\"" + str(EndTime) + "\", "


# ----- 時間がかかったプログラムとその時間を格納 -----
Max_Program, Max_Program_Time = "", datetime.timedelta()
Max_SQL, Max_SQL_Time = "", datetime.timedelta()



# ----- 全Logファイルの処理 -----
for File in glob.glob(Path + '/*.log'):

    # ----- 対象Operationの取得 -----
    Operation = os.path.basename(File).split('.')[0]

    # ---- Logファイルを開く -----
    with open(File) as f:
        s = [i.strip() for i in f.readlines()]

    # ----- 末尾に "Program End" が含まれていないものはエラー終了と見なす -----
    if "Program End" not in s[-1]:
        Error.append(os.path.splitext(os.path.basename(File))[0])


    # ----- 処理時間を測定する -----
    if Operation == "SQL_Program":
        Start = datetime.datetime.strptime(s[0][21:29], "%H:%M:%S")
        for i in s[1:]:
            End = datetime.datetime.strptime(i[21:29], "%H:%M:%S")
            Processing_Time = End-Start
            if Processing_Time > Max_SQL_Time:
                Max_SQL_Time = Processing_Time
                Max_SQL = i[30:-9]
            Start = End
            if i[30:-9] == "Prog":
                Max_SQL, Max_SQL_Time = "", datetime.timedelta()
                Start = datetime.datetime.strptime(i[21:29], "%H:%M:%S")


    else:
        Program_End, Program_Start = "", ""
        for i in s[::-1]:
            if "Program End" in i:
                Program_End = datetime.datetime.strptime(i[21:29], "%H:%M:%S")
            if "Program Start" in i:
                Program_Start = datetime.datetime.strptime(i[21:29], "%H:%M:%S")
                break

        if Program_Start=="" or Program_End=="": continue

        Processing_Time = Program_End - Program_Start
        if Processing_Time > Max_Program_Time:
            Max_Program_Time = Processing_Time
            Max_Program = Operation


# プログラム名 | 時間 | SQL名 | 時間
sql += "\"" + Max_Program + "\", " + "\"" + str(Max_Program_Time) + "\", \"" + Max_SQL + "\", " + "\"" + str(Max_SQL_Time) + "\", "


# ----- カンマ区切りに直しておく -----
Error = ','.join(Error)
if len(Error):
    sql += "\"" + "NG" + "\", " + "\"" + str(User) + "\", " + "\"" + Error + "\"" + ");"

else:
    sql += "\"" + "OK" + "\", " + "\"" + str(User) + "\", " + "default);"


# sqlの実行
with connection.cursor() as cursor:
    cursor.execute(sql)
    connection.commit()

connection.close()