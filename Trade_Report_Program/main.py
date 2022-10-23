"""
Purpose of code: 利用iTrade分析資料庫搜索，產生[貿易局季報]所需的數據源CSV檔案。

1. 計算出口原始資料出現的所有11碼HSCODE(非重複)。
2. 引用 TSV 產業對照表資訊，儲存資訊至 dictionary。
3. 利用產業對照表的多類HS碼，創建產業 11碼 HSCode的對照表，儲存資訊至 JSON。
4. 連接資料庫，利用 11碼資訊查詢各產業貿易額，儲存為 pandas dataframe。
   並輸出為CSV檔案: export_file/bureau_XXXX-XX-XX.csv
"""

from basic_report_cal import *
from db_connect import *
import datetime


def main():
    cnxn = create_connection()  # create a database connection

    if cnxn is not None:
        # 1. 計算出口原始資料出現的所有11碼HSCODE(非重複)。
        dict_unique_hs11 = extract_unique_hscode(cnxn, YEAR)

        # 2. 引用 TSV 產業對照表資訊，儲存資訊至 dictionary。
        dict_indus = get_indus_hscode(INDUSTRY_TABLE)

        # 3. 利用產業HSCode對照表的HS碼(2-10碼)，創建產業HSCode對照表的HS碼(11碼)為dictionary
        dict_indus_hs11 = create_indus_hs11(dict_indus, dict_unique_hs11)

        # 4.連接資料庫，利用11碼資訊查詢各產業貿易額，儲存為pandas dataframe。
        # 並輸出為CSV檔案: "export_file/bureau_XXXX-XX-XX.csv"。

        today = f"{datetime.datetime.today():%Y-%m-%d}"
        export_fname = 'export_file/bureau_{}.csv'.format(today)  # 輸出的CSV檔名
        export_all_bureau_csv(cnxn, YEAR, dict_indus_hs11, export_fname)
        cnxn.close()

    else:
        print("Error! cannot create the database connection.")


if __name__ == '__main__':

    # ============ 選取報告所需年度[可編輯] ================================
    report_year = 2022  # 輸入選取報告所需年度
    INDUSTRY_TABLE = 'import_file/Industry_hscode_20220524.tsv'  # 貿易局產業HS對照表資訊(須定期更新)。
    # ==================================================================

    YEAR = str(int(report_year) - 2)  # 資料須包含當年與過去2年內歷史資料
    main()


