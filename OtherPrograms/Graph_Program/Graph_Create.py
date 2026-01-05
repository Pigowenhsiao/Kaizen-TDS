import matplotlib.pyplot as plt
import seaborn as sns
import datetime as dt
import pandas as pd
import numpy as np
import glob
import sys
import os

from math import isnan



def Create_Graph(df, x, y, Alarm_min, Alarm_max, Standard_min, Standard_max, x_label, y_label, title, save_name, Log_chk, hue_name):


    # ----- グラフの土台を作成 -----
    fig = plt.figure(figsize=(16, 8), dpi=100, facecolor='w')
    ax = fig.add_subplot()
    ax.grid(True)

    Flag = 1

    # ----- グラフ描画 -----
    if hue_name == "-":
        sns.scatterplot(data=df, x=x_label, y=y_label, s=50, alpha=0.5, marker='o')
    else:
        sns.scatterplot(data=df, x=x_label, y=y_label, s=50, alpha=0.5, marker='o', hue=hue_name)
        Flag = 0


    # ----- y軸の対数表示 -----
    if Log_chk == True:
        ax.semilogy()

    # y軸範囲を求める
    y_min = np.nanmin(y)
    y_max = np.nanmax(y)


    # ----- 規格線 -----
    if isnan(Alarm_min) == False:
        ax.hlines(y=Alarm_min, xmin=min(x), xmax=max(x), color='yellow')
        y_min = min(y_min, Alarm_min)

    if isnan(Alarm_max) == False:
        ax.hlines(y=Alarm_max, xmin=min(x), xmax=max(x), color='yellow')
        y_max = max(y_max, Alarm_max)

    if isnan(Standard_min) == False:
        ax.hlines(y=Standard_min, xmin=min(x), xmax=max(x), color='red')
        y_min = min(y_min, Standard_min)

    if isnan(Standard_max) == False:
        ax.hlines(y=Standard_max, xmin=min(x), xmax=max(x), color='red')
        y_max = max(y_max, Standard_max)


    # ----- 中心線 -----
    if isnan(Standard_min) == False and isnan(Standard_max) == False:
        ax.hlines(y=(Standard_min + Standard_max) / 2, xmin=min(x), xmax=max(x), color='green', linestyle='--')

    elif isnan(Alarm_min) == False and isnan(Alarm_max) == False:
        ax.hlines(y=(Alarm_min + Alarm_max) / 2, xmin=min(x), xmax=max(x), color='green', linestyle='--')


    # ----- 凡例 ------
    if Flag:
        ax.legend(labels=[y_label], bbox_to_anchor=(1.05, 1), loc='upper left', borderaxespad=0, fontsize=12)


    # ----- グラフタイトル -----
    ax.set_title(title, fontname="MS Gothic")


    # ----- 軸名 -----
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)


    # ----- x軸範囲 -----
    try:
        plt.xlim(min(x)-dt.timedelta(days=3), max(x)+dt.timedelta(days=3))
    except:
        pass


    # ----- y軸範囲 -----
    y_min -= y_min*0.1
    y_max += y_max*0.1
    plt.ylim(y_min, y_max)


    # ----- パスを指定して保存 -----
    # path = './Graph/'
    path = "Z:/KAIZEN-TDS/Graph/"
    plt.savefig(os.path.join(path, save_name), bbox_inches="tight")


    plt.close()


########## Main ##########
if __name__ == "__main__":

    # ---- csvフォルダパスの定義 -----
    csv_Path = "Z:/KAIZEN-TDS/csv/"

    # ----- 規格等を記載したファイルの読み込み -----
    Standard_df = pd.read_csv("./Standard.csv", sep=",", header=None, skiprows=1, encoding='cp932')

    # ----- 初期設定 -----
    Standard_Row = 0

    # ----- Srandard.csvに記載した項目のGraph化 -----
    while Standard_Row < Standard_df.shape[0]:

        # ----- csvファイルを開く -----
        csv_file = os.path.join(csv_Path, Standard_df.iat[Standard_Row, 0]+'.csv')
        df = pd.read_csv(csv_file, encoding='cp932')

        # ----- dfが空であれば、次に遷移 -----
        if len(df) == 0:
            Standard_Row+=1
            continue

        # ----- 処理対象のcsvファイルの出力 -----
        print("Now:",csv_file)
        # print(df.columns)

        # ----- 品名を大分類に変換(HL13B4-BTxx -> HL13B4) -----
        df['PARTNUMBER'] = df['PARTNUMBER'].str.split('-').str[0]


        # PARTNUMBERでの抽出が必要ならば、抽出処理を行う
        if Standard_df.iat[Standard_Row, 6] != '-':

            # ----- ','で分けられている場合もあるので、リストで取得 -----
            PartNumber_List = [P.strip() for P in Standard_df.iat[Standard_Row, 6].split(',')]

            # 対象の品名を抜き出したDataFrameを作成
            df = df[df['PARTNUMBER'].isin(PartNumber_List)]

            # 行番号をリセット
            df = df.reset_index()


        # ----- なぜか[CARRIERCONCENTRATION_CONTACT]だけはstr型で取得されるのでint型に直す -----
        if Standard_df.iat[Standard_Row, 2].upper() == "CARRIERCONCENTRATION_CONTACT":
            df[Standard_df.iat[Standard_Row, 2]] = df[Standard_df.iat[Standard_Row, 2]].astype(np.float64)


        # ----- 指定した期間のデータのみを取り出す -----
        Target = Standard_df.iat[Standard_Row, 3]
        try:
            df[Target] = pd.to_datetime(df[Target])
            StartTime = Standard_df.iat[Standard_Row, 4]
            EndTime = Standard_df.iat[Standard_Row, 5]
            df = df[(df[Target] >= StartTime) & (df[Target] <= EndTime)]
        except:
            pass


        # ----- x項目を日付型に変換,　できない場合は何もしない -----
        try:
            df[Standard_df.iat[Standard_Row, 1]] = pd.to_datetime(df[Standard_df.iat[Standard_Row, 1]])
        except:
            pass

        # ----- X軸 -----
        x = df[Standard_df.iat[Standard_Row, 1]]

        # ----- Y軸 -----
        y = df[Standard_df.iat[Standard_Row, 2]]

        # ----- 規格値 -----
        Standard_min = Standard_df.iat[Standard_Row, 7]
        Standard_max = Standard_df.iat[Standard_Row, 10]

        # ----- アラーム規格値 -----
        Alarm_min = Standard_df.iat[Standard_Row, 8]
        Alarm_max = Standard_df.iat[Standard_Row, 9]

        # ----- xy軸ラベル名 -----
        x_label = Standard_df.iat[Standard_Row, 1]
        y_label = Standard_df.iat[Standard_Row, 2]

        # ----- hue -----
        hue = Standard_df.iat[Standard_Row, 11]

        # ----- グラフタイトル -----
        title = Standard_df.iat[Standard_Row, 12]

        # ----- グラフ保存名 -----
        save_name = Standard_df.iat[Standard_Row, 13] + '.png'

        # ----- Y軸のLog変換を行うかどうか -----
        Log_chk = Standard_df.iat[Standard_Row, 14]


        # ----- グラフ化 -----
        if len(x):
            print('Create Graph')
            Create_Graph(df, x, y, Alarm_min, Alarm_max, Standard_min, Standard_max, x_label, y_label, title, save_name, Log_chk, hue)

        Standard_Row += 1