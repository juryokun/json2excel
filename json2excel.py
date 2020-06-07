# -*- coding: utf-8 -*-
import openpyxl
import json
import sys

def main():
    """
    メイン処理
    """
    # 設定ファイル読み込み
    with open('data.json', 'r') as f:
        try:
            json_data = json.load(f)

            book = openpyxl.load_workbook('data.xlsx')
            sheet = book['Sheet1']

        except Exception as e:
            print(e)
            sys.exit(1)

    sheet.cell(row=1,column=1).value = 'text'
    sheet.cell(row=1,column=2).value = 'size'
    sheet.cell(row=1,column=3).value = 'opacity'
    for i, data in enumerate(json_data):
        target_row = i+2
        sheet.cell(row=target_row, column=1).value = data['text']
        sheet.cell(row=target_row, column=2).value = data['size']
        sheet.cell(row=target_row, column=3).value = data['opacity']

    book.save('data.xlsx')
    sys.exit()


main()
