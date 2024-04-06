"""
概要---
このスクリプトは、脳波データを加算処理するためのものです。特定のタイミングでの脳波データを合計し、グラフで表示する機能を持っています。

実行条件---
Python環境とPandas、matplotlib、numpy、openpyxl、scipyライブラリが必要です。
artifact_remove.py からいくつかの関数をインポートしています。
analyze_setting.txt と名付けられた設定ファイルが必要です。
加工対象のCSVファイルが、特定のディレクトリ構造内に存在する必要があります。
コマンドラインから python eeg_add.py で実行します。

出力---
加算された脳波データのグラフを表示します。
"""
import glob
import math
from matplotlib import pyplot as plt
import numpy as np
import openpyxl
from scipy import interpolate
import os
import pandas as pd

from artifact_remove import (
    get_setting_file_path,
    get_data_dir_path,
    read_dir_name_from_settings,
)

POLYMATE_SAMPLING = 0.025
INTERVAL = 3  # 刺激提示間隔
ANALYZE_START = 0.2  # 刺激提示の0.2秒前から解析する
POINT = 10000

LED_CYCLE = 0.41  # 0.2+0.21
REPEAT = 20
HIDDEN_TIME = 2.7

LED = False  # LEDを繰り返したか


def get_target_numbers(analyze_info):
    """
    Analyze_infoデータフレームから目標値(target number)のリストを取得します。

    Args:
        analyze_info (DataFrame): 分析情報が含まれるPandasのデータフレーム。

    Returns:
        List[List[float]]: 目標値のリストのリスト。各サブリストは一つの実験データに対応します。
    """
    target_rows = analyze_info[analyze_info[0].str.contains("target number")]
    target_numbers = target_rows.iloc[:, 1:].values.tolist()

    for i in range(len(target_numbers)):
        target_numbers[i] = [num for num in target_numbers[i]]

    target_numbers = [list(row) for row in zip(*target_numbers)]

    target_numbers = [
        [x for x in sublist if not math.isnan(x)] for sublist in target_numbers
    ]

    return target_numbers


def make_target_numbers(data_dir_path):
    summary_path = os.path.join(data_dir_path, "blood_excel", "summary.xlsx")
    if not os.path.isfile(summary_path):
        print("blood_excelディレクトリにsummary.xlsxがないです")
        return
    analyze_info = pd.read_excel(
        summary_path, sheet_name="analyze_info", header=None, engine="openpyxl"
    )
    target_numbers = get_target_numbers(analyze_info)  # 2次元リスト

    return target_numbers


def getOddballStartTime():
    """
    オッドボール課題が開始された時間を計算します。

    Returns:
        float: オッドボール課題の開始時間
    """
    if LED:
        oddstart_time = LED_CYCLE * REPEAT + HIDDEN_TIME
    else:
        oddstart_time = HIDDEN_TIME

    print("オドボール課題が開始された時間：", oddstart_time)

    return oddstart_time


def getSplineFunc(data):
    """
    データに対してスプライン補間を行い、補間関数を返します。

    Args:
        data (list): 補間するデータのリスト

    Returns:
        function: 補間関数
    """
    x = list(range(len(data)))
    x = [n * POLYMATE_SAMPLING for n in x]
    f = interpolate.Akima1DInterpolator(x, data)

    return f


def getTargetTime(oddstart_time, target):
    """
    特定のターゲットの時間を計算します。

    Args:
        oddstart_time (float): オッドボール課題の開始時間
        target (int): ターゲットのインデックス

    Returns:
        float: ターゲットの時間
    """
    target_time = oddstart_time + target * POLYMATE_SAMPLING

    return target_time


def readData(data_ws):
    """
    ExcelファイルからEEGデータを読み込みます。

    Args:
        data_ws (Worksheet): データが含まれるワークシート

    Returns:
        list: EEGデータのリスト
    """
    data = []
    for col in data_ws.iter_cols(
        min_row=2, min_col=2
    ):  # 1行目はTIMEとか1-REFとかヘッダが書いてあるから、2行目からみる
        for cell in col:
            data.append(float(cell.value))
    return data


def readTarget(target_ws):
    """
    ターゲットのリストをExcelファイルから読み込みます。

    Args:
        target_ws (Worksheet): ターゲットが含まれるワークシート

    Returns:
        list: ターゲットのリスト
    """
    targetlist = []
    for row in target_ws.iter_rows():
        for cell in row:
            targetlist.append(int(cell.value))
    return targetlist


def displayGraph(target_eeg_total):
    """
    処理されたEEGデータのグラフを表示します。

    Args:
        target_eeg_total (numpy.ndarray): 加算されたEEGデータ
    """
    t = np.linspace(-ANALYZE_START, 3, num=POINT)

    fig, ax = plt.subplots()
    ax.invert_yaxis()
    ax.set_xlim(-ANALYZE_START, 0.7)
    ax.set_xlabel("t[s]")
    ax.set_ylabel("eeg[μV]")
    ax.plot(t, target_eeg_total, label="eeg", color="black")
    ax.grid()

    plt.show()


def main():
    setting_file = get_setting_file_path()
    data_dir_name = read_dir_name_from_settings(setting_file)
    data_dir_path = get_data_dir_path(data_dir_name)

    csv_dir_path = os.path.join(data_dir_path, "eeg_csv", "artifact_removed")
    csv_files = glob.glob(csv_dir_path + "/*.csv")

    # target_numbersを取得
    target_numbers = make_target_numbers(data_dir_path)  # 二次元リスト

    oddstart_time = getOddballStartTime()
    target_eeg = []

    for i, file in enumerate(csv_files):
        data = pd.read_csv(file, usecols=[" 1-REF"], squeeze=True)
        data_list = data.tolist()

        for target in target_numbers[i]:
            target_time = getTargetTime(oddstart_time, target)
            data_spline_func = getSplineFunc(data_list)
            target_t = np.linspace(
                target_time - ANALYZE_START, target_time + INTERVAL, num=POINT
            )
            target_eeg.append(data_spline_func(target_t))

    target_eeg_arr = np.array(target_eeg)
    target_eeg_total = np.sum(target_eeg_arr, axis=0)

    print(np.max(target_eeg_total))
    print(np.min(target_eeg_total))

    print("ターゲットの総数：", sum(len(sublist) for sublist in target_numbers))

    displayGraph(target_eeg_total)


if __name__ == "__main__":
    main()
