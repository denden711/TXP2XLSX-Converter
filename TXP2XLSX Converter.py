import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# TXPファイルを読み込む関数
def read_txp_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        data = [line.strip().split('\t') for line in lines]
        return data
    except Exception as e:
        print(f"ファイルの読み込み中にエラーが発生しました: {file_path}")
        print(e)
        return None

# データをDataFrameに変換
def convert_to_dataframe(data):
    try:
        if not data or len(data) < 2:
            raise ValueError("データが不足しています")
        df = pd.DataFrame(data[1:], columns=data[0])
        # 数値に変換できる列を数値に変換
        for col in df.columns:
            try:
                df[col] = pd.to_numeric(df[col])
            except ValueError:
                pass  # 数値に変換できない場合はそのままにする
        return df
    except Exception as e:
        print("データの変換中にエラーが発生しました。")
        print(e)
        return None

# DataFrameをExcelファイルに保存
def save_to_excel(df, output_path):
    try:
        df.to_excel(output_path, index=False)
    except Exception as e:
        print(f"Excelファイルの保存中にエラーが発生しました: {output_path}")
        print(e)

# 指定されたディレクトリ内のすべてのTXPファイルをXLSXに変換
def convert_all_txp_in_directory(directory):
    try:
        for filename in os.listdir(directory):
            if filename.endswith(".txp"):
                txp_path = os.path.join(directory, filename)
                xlsx_path = os.path.join(directory, filename.replace(".txp", ".xlsx"))
                data = read_txp_file(txp_path)
                if data is None:
                    continue
                df = convert_to_dataframe(data)
                if df is None:
                    continue
                save_to_excel(df, xlsx_path)
                print(f'{xlsx_path}にデータが保存されました。')
    except Exception as e:
        print(f"ディレクトリ内のファイル処理中にエラーが発生しました: {directory}")
        print(e)

# フォルダ選択ダイアログを開いてフォルダを選択
def select_folder_and_convert():
    try:
        root = tk.Tk()
        root.withdraw()  # メインウィンドウを表示しない
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            convert_all_txp_in_directory(folder_selected)
            messagebox.showinfo("完了", f'{folder_selected}内のすべてのTXPファイルがXLSXに変換されました。')
            print(f'{folder_selected}内のすべてのTXPファイルがXLSXに変換されました。')
        else:
            print("フォルダが選択されませんでした。")
    except Exception as e:
        print("フォルダ選択と変換中にエラーが発生しました。")
        print(e)
        messagebox.showerror("エラー", "フォルダ選択と変換中にエラーが発生しました。")

# メインの処理を実行
if __name__ == "__main__":
    select_folder_and_convert()
