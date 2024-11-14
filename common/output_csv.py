import pandas as pd

def run(work_file_path, sheet_num, output_csv_path):
    # 作業ファイルのCSVシートをpandasで読み込み
    df = pd.read_excel(work_file_path, sheet_name=sheet_num, index_col=None, engine='openpyxl')
    headers = df.columns
    debit = headers[13]
    credit = headers[14]
    gross = headers[15]

    # 金額がfloat型になっているのでそれぞれInt型にする
    df[debit] = df[debit].fillna(0).astype(int) 
    df[credit] = df[credit].fillna(0).astype(int)
    df[gross] = df[gross].fillna(0).astype(int)

    # 総額が0以外でフィルタ
    filtered_df = df[df.iloc[:, 15] != 0]

    # CSV出力
    filtered_df.to_csv(output_csv_path, index=False, encoding="shift-jis")

if __name__ == "__main__":
    run()