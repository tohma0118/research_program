"""
概要---
このスクリプトは、脳波データのCSVファイルからアーチファクト（雑音や誤ったデータ）を除去し、結果を新しいCSVファイルに保存します。

実行条件---
Python環境が必要。
Pandasライブラリが必要。
設定ファイル（analyze_setting.txt）とデータCSVファイルが必要。
コマンドラインから python artifact_remove.py で実行。

出力---
処理後のデータは、元のファイルのディレクトリ内に artifact_removed というサブディレクトリに保存されます。
ファイル名は元の名前に _artifact_removed.csv が追加されます。
"""
import pandas as pd
import os
import glob


def get_setting_file_path():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.abspath(os.path.join(current_dir, os.pardir))
    setting_file = os.path.join(parent_dir, "analyze_setting.txt")

    return setting_file


def get_data_dir_path(data_directory_name):
    current_dir = os.path.dirname(os.path.abspath(__file__))
    grandparent_dir = os.path.abspath(os.path.join(current_dir, os.pardir, os.pardir))
    data_dir = os.path.join(grandparent_dir, "data", data_directory_name)

    return data_dir


def read_dir_name_from_settings(file_name):
    with open(file_name, "r", encoding="utf-8") as file:
        lines = file.readlines()
        for i, line in enumerate(lines):
            line = line.strip()
            if line == "directory_name" and i + 1 < len(lines):
                return lines[i + 1].strip()
    return None


def make_output_path(input_path):
    """出力ファイルのパスを作成し返す

    Args:
        input_path (str): 入力ファイル

    Return:
        output_path(str): 出力先
    """

    base_name = os.path.basename(input_path)
    file_name = base_name.rsplit(".", 1)[0]
    new_file_name = file_name + "_artifact_removed" + ".csv"
    eeg_csv_dir = os.path.dirname(input_path)

    output_dir = os.path.join(eeg_csv_dir, "artifact_removed")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    output_path = os.path.join(output_dir, new_file_name)
    print("出力先:", output_path)
    return output_path


def remove_artifacts_from_csv(file_path, output_path=None, threshold=50):
    """
    脳波データのCSVファイルからアーチファクトを除去し、クリーンなデータを新しいCSVファイルに保存します。

    パラメータ
    file_path (str): 脳波データを含むCSVファイルへのパス。
    output_path (str): クリーニングされたデータを保存するパス。
    threshold (int): アーチファクトを判定するための閾値。デフォルトは±50。
    """

    data = pd.read_csv(file_path)

    in_artifact = False
    start_index = None

    for i, value in enumerate(data[" 1-REF"]):
        # 現在値がしきい値を超えているかチェックする
        if abs(value) > threshold:
            # アーティファクトの開始
            if not in_artifact:
                in_artifact = True
                start_index = i
        else:
            # アーティファクトの終わり
            if in_artifact:
                # アーティファクトの「ピーク」または「バレー」の値をすべてゼロにする。
                data.loc[start_index:i, " 1-REF"] = 0
                in_artifact = False

    # 最後の値がアーティファクトの一部であるかどうかをチェックする
    if in_artifact:
        data.loc[start_index:, " 1-REF"] = 0

    if not output_path:
        output_path = make_output_path(file_path)

    # クリーニングしたデータを新しいCSVファイルに保存する。
    data.to_csv(output_path, index=False)


def main():
    setting_file = get_setting_file_path()
    data_dir_name = read_dir_name_from_settings(setting_file)
    data_dir_path = get_data_dir_path(data_dir_name)

    csv_dir_path = os.path.join(data_dir_path, "eeg_csv")

    # ディレクトリ内のすべての.csvファイルを取得
    csv_files = glob.glob(csv_dir_path + "/*.csv")

    # 各ファイルを一つずつ処理
    for file in csv_files:
        remove_artifacts_from_csv(file)


if __name__ == "__main__":
    main()
