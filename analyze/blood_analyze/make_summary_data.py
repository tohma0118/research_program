"""
summary.xlsxのanalyze_infoシートの情報から、データを別シートに書き込む
blood_excelディレクトリにanalyze_infoシートが完成したsummary.xlsxが存在しない場合は動作しない
"""

import os
import pandas as pd
from openpyxl.styles import Font

from create_analyze_info_check import (
    get_setting_file_path,
    get_data_dir_path,
    read_dir_name_from_settings,
)

SKIP_ROWS_NUMBER = 54


def get_csv_files_from_folder(folder_path, substring):
    """指定されたフォルダから、部分文字列.csvを含む csv ファイルのリストを返します。"""
    all_files = os.listdir(folder_path)
    return [f for f in all_files if f.endswith(".csv") and substring in f]


def read_analyze_info_from_excel(summary_path):
    """analyze_infoシートから外乱終了データを読み込んで返す"""
    try:
        analyze_info_df = pd.read_excel(
            summary_path, sheet_name="analyze_info", engine="openpyxl"
        )
        return analyze_info_df.iloc[0, 1:].dropna().values
    except FileNotFoundError:
        print(f"Error: File {summary_path} not found.")
        return []


def read_data_for_columns(csv_directory_path, csv_file, column_index, skip_rows_total):
    data = pd.read_csv(
        os.path.join(csv_directory_path, csv_file),
        encoding="Shift-JIS",
        skiprows=skip_rows_total,
        header=None,
        usecols=[column_index],
    )
    return data


def write_data_to_excel(excel_path, data_dfs):
    """Write data to the Excel file with the specified font."""
    with pd.ExcelFile(excel_path, engine="openpyxl") as xls:
        sheets = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        font = Font(name="Yu Gothic")
        for sheet_name, data in sheets.items():
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.book[sheet_name]
            for row in ws.iter_rows():
                for cell in row:
                    cell.font = font
        for ch, df in data_dfs.items():
            df.to_excel(writer, sheet_name=ch, index=False, header=False)
            ws = writer.book[ch]
            for row in ws.iter_rows():
                for cell in row:
                    cell.font = font


def main():
    setting_file = get_setting_file_path()
    data_dir_name = read_dir_name_from_settings(setting_file)
    data_dir_path = get_data_dir_path(data_dir_name)

    csv_dir_path = os.path.join(data_dir_path, "blood_csv")
    summary_path = os.path.join(data_dir_path, "blood_excel", "summary.xlsx")

    if not os.path.isfile(summary_path):
        print("blood_excelディレクトリにsummary.xlsxを追加し、analyze_infoシートを埋めてください")
        return

    disturbance_end = read_analyze_info_from_excel(summary_path)

    # disturbance_endが空の場合、エラーメッセージを表示して終了
    if len(disturbance_end) == 0:
        print("Error: analyze_infoシートのdisturbance endを埋めてください")
        return

    csv_files = get_csv_files_from_folder(csv_dir_path, "_Oxy")

    if not csv_files:
        print("指定された部分文字列を持つCSVファイルが見つかりません。")
        return

    if len(disturbance_end) != len(csv_files):
        print("CSVファイル数と開始点が一致していません。")
        return

    columns = range(7, 17)
    data_dfs = {f"CH{column}": pd.DataFrame() for column in columns}

    for i, csv_file in enumerate(csv_files):
        skip_rows_total = range(0, SKIP_ROWS_NUMBER + int(disturbance_end[i]))
        for column in columns:
            column_data = read_data_for_columns(
                csv_dir_path, csv_file, column, skip_rows_total
            )
            data_dfs[f"CH{column}"][csv_file] = column_data.squeeze()

    write_data_to_excel(summary_path, data_dfs)
    print("データを書きました")


if __name__ == "__main__":
    main()
