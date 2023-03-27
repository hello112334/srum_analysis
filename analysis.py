import openpyxl
import csv

import pandas as pd
import binascii

# Excelファイルを読み込む
wb = openpyxl.load_workbook('SRUM_DUMP_OUTPUT.xlsx')

# シート名のリストを取得
sheet_names = wb.sheetnames


def main():
    """note"""

    # 各シートの内容を取得
    for sheet_name in sheet_names:

        # 実行プログラム
        if sheet_name == "ruDbIdMapTable":
            print(f"[INFO] {sheet_name}")

            # シートデータを出力
            write2csv(sheet_name)

            # Pandasで読み込む
            # CSVファイルを読み込んでDataFrameオブジェクトとして格納
            headers = ["IdType", "IdIndex", "IdBlob"]
            df = pd.read_csv(f"{sheet_name}.csv", names=headers, header=0)

            # 空欄を0にする
            df.fillna("00", inplace=True)

            # 特定の列を1行ずつ出力する
            for index, row in df.iterrows():
                print(str(row['IdBlob']))
                res = hex_to_str(str(row['IdBlob']))
                print(f"{index}:: {df['IdBlob'][index]}....")
                df['IdBlob'][index] = res

            # df['IdBlob'] = str(df['IdBlob'])
            # df['IdBlob_new'] = df['IdBlob'].apply(hex_to_str)
            df.to_csv(f"{sheet_name}_ana.csv", index=False)

        # CPU/Battery
        if sheet_name == "5C8CF1C7-7257-":
            print(f"[INFO] {sheet_name}")

            # シートデータを出力
            write2csv(sheet_name)

            # header = [""]

# applyメソッドを使って16進数表記から文字列に変換する関数を定義
def hex_to_str(hex_str):

    try:
        hex_bytes = bytes.fromhex(hex_str)
        decoded_str = hex_bytes.decode('utf-8')
        # if type(decoded_str) == "string":
    except Exception as err1:
        # print(f"[ERROR] {err1}")
        return "-"

    return decoded_str


def write2csv(sheet_name):
    """note"""

    # CSVファイルを開く
    with open(f"{sheet_name}.csv", 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)

        # シートを選択する
        sheet = wb[sheet_name]

        # 全ての行を取得する
        rows = sheet.iter_rows(values_only=True)

        # 行ごとに処理する
        for row in rows:
            # CSVに書き込む
            writer.writerow(row)

        # シート間に空行を入れる
        writer.writerow([])

    # ファイルを閉じる
    wb.close()


if __name__ == "__main__":

    try:
        main()

    except Exception as err:
        print(f"[ERROR] {err}")
