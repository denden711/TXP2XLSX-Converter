import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed

class TXP2XLSXConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("TXP2XLSXコンバーター")
        self.create_widgets()

        # ログの設定
        logging.basicConfig(
            filename='conversion.log', 
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            encoding='utf-8'
        )

    def create_widgets(self):
        self.select_button = tk.Button(self.root, text="ディレクトリを選択", command=self.select_directory)
        self.select_button.pack(pady=20)

    def txp_to_xlsx(self, file_path):
        try:
            # ファイルを読み込み
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()

            # タブ区切りでデータを分割
            data = [line.strip().split('\t') for line in lines]
            df = pd.DataFrame(data[1:], columns=data[0])

            # 出力ファイルのパスを決定
            output_file = file_path.replace('.txp', '.xlsx')
            df.to_excel(output_file, index=False)
            logging.info(f"{file_path} を {output_file} に変換しました")
            print(f"{file_path} を {output_file} に変換しました")
        except FileNotFoundError:
            error_message = f"ファイルが見つかりません: {file_path}"
            logging.error(error_message)
            print(error_message)
            messagebox.showerror("ファイルエラー", error_message)
        except pd.errors.EmptyDataError:
            error_message = f"ファイルが空です: {file_path}"
            logging.error(error_message)
            print(error_message)
            messagebox.showerror("データエラー", error_message)
        except Exception as e:
            error_message = f"{file_path} の変換に失敗しました: {e}"
            logging.error(error_message)
            print(error_message)
            messagebox.showerror("変換エラー", error_message)

    def convert_all_txp_in_directory(self, directory):
        try:
            with ThreadPoolExecutor(max_workers=10) as executor:
                futures = []
                for root, _, files in os.walk(directory):
                    for filename in files:
                        if filename.endswith('.txp'):
                            file_path = os.path.join(root, filename)
                            futures.append(executor.submit(self.txp_to_xlsx, file_path))
                
                for future in as_completed(futures):
                    future.result()  # ここで例外をキャッチしてログに記録する

            logging.info("すべてのファイルが正常に変換されました")
            messagebox.showinfo("成功", "すべてのファイルが正常に変換されました")
            self.show_completion_message()
        except Exception as e:
            error_message = f"ディレクトリ {directory} のファイル変換中にエラーが発生しました: {e}"
            logging.error(error_message)
            print(error_message)
            messagebox.showerror("ディレクトリエラー", error_message)

    def select_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.convert_all_txp_in_directory(directory)

    def show_completion_message(self):
        messagebox.showinfo("完了", "すべてのTXPファイルの変換が完了しました。")

if __name__ == "__main__":
    root = tk.Tk()
    app = TXP2XLSXConverter(root)
    root.mainloop()
