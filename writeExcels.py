import sys
import os
import json
import win32com.client as win32

def process_excel_files(folder_path):
    print(folder_path)
    # スクリプトの実行ディレクトリを取得
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # JSONファイルのパス
    json_file_path = os.path.join(script_dir, "data.json")

    # JSONファイルからシート名とセルの指定と書き込む値を取得
    with open(json_file_path, "r", encoding="utf-8") as json_file:
        data = json.load(json_file)
        sheet_name = data.get("sheet_name")
        cell_data = data.get("cell_data")

    # フォルダ内の全てのエクセルファイルに対して処理を行う
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path) and filename.lower().endswith((".xlsx",".xls",".xlsm")):
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False

            # ブックを開く
            workbook = excel.Workbooks.Open(file_path)

            # 指定したシート名の存在チェック
            if sheet_name in [sheet.Name for sheet in workbook.Sheets]:
                sheet = workbook.Sheets(sheet_name)

                # セルの指定と書き込む値をループして処理
                for cell in cell_data:
                    cell_value = cell_data[cell]
                    sheet.Range(cell).Value = cell_value

                # 変更を保存
                workbook.Save()

            workbook.Close()
            excel.Quit()

# コマンドライン引数からフォルダパスを取得
if len(sys.argv) < 2:
    folder_path = os.path.dirname(os.path.abspath(__file__)) # スクリプトの実行ディレクトリを使用
else:
    folder_path = sys.argv[1]

process_excel_files(folder_path)
