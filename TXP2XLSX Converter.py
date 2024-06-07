import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def txp_to_xlsx(file_path):
    try:
        # TXPファイルの読み込み
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()

        # 必要に応じてデータを処理
        # ここではシンプルな処理を仮定
        data = [line.strip().split('\t') for line in lines]

        # データフレームに変換
        df = pd.DataFrame(data[1:], columns=data[0])

        # 出力ファイル名
        output_file = file_path.replace('.txp', '.xlsx')

        # データフレームをXLSXファイルに書き出し
        df.to_excel(output_file, index=False)
        print(f"Converted {file_path} to {output_file}")
    except Exception as e:
        print(f"Failed to convert {file_path}: {e}")

def convert_all_txp_in_directory(directory):
    for filename in os.listdir(directory):
        if filename.endswith('.txp'):
            file_path = os.path.join(directory, filename)
            txp_to_xlsx(file_path)

def select_directory():
    directory = filedialog.askdirectory()
    if directory:
        convert_all_txp_in_directory(directory)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # メインウィンドウを隠す
    select_directory()
