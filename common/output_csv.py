import pandas as pd
import datetime

dt_now = datetime.datetime.now()
year_str = str(dt_now.year)
month_str = str(dt_now.month - 1).zfill(2)
year_month_str = year_str + month_str

download_folder = "C:\\Users\\shinya.arai\\Downloads\\"
charge_je_csv_path = download_folder + "Edge請求仕訳" + year_month_str + ".csv"

def run(work_file_path, sheet_num):
    pd.set_option('display.unicode.east_asian_width', True)

    df = pd.read_excel(work_file_path, sheet_name=sheet_num, index_col=None, engine='openpyxl')
    headers = df.columns
    debit = headers[13]
    credit = headers[14]
    gross = headers[15]
    df[debit] = df[debit].fillna(0).astype(int) # 金額をInt型にする
    df[credit] = df[credit].fillna(0).astype(int) # 金額をInt型にする
    df[gross] = df[gross].fillna(0).astype(int) # 金額をInt型にする

    filtered_df = df[df.iloc[:, 15] != 0]

    filtered_df.to_csv(charge_je_csv_path, index=False, encoding="shift-jis")


if __name__ == "__main__":
    run()