"""
Purpose of code: 利用計算好的CSV數據與資料庫搜索，產生[貿易局季報Q1、Q2、Q3]所需數據源 EXCEL檔案。
"""

from stats_fcn import *
import datetime


def main():

    df = pd.read_csv(Import_CSV_file)

    ## 1. 計算特定年度之前，台灣3年內出口總額(p1)(百萬美元)
    ## 計算特定年度，特定月份區間，所有國家出口差額(p3上側文字)(千美元)
    p1 = cal_total_export(YEAR)
    p3 = cal_total_country_diff(YEAR, END_MM)

    # 2. 計算特定年度，特定月份區間，所有產業出口差額，給出產業排行(p4右側)(萬美元)
    # 計算特定年度，特定月份區間，各國出口額(p4上側)(萬美元)
    # 計算特定年度，特定月份區間，特定國家(7國)出口差額(p4下側)(萬美元)

    p4_right = cal_industry_diff(df, YEAR, START_MM, END_MM)
    p4_up = cal_total_country(YEAR, END_MM)
    p4_down = cal_total_eight_diff(df, YEAR, START_MM, END_MM)

    ## 3. 計算特定年度，特定月份區間，ICT產業各國出口差額，給出各國排行(p5右側)
    ## 計算特定年度，特定月份區間，ICT產業特定國家(7國)出口差額(p5下側)
    ## 計算特定年度，特定月份區間，ICT產業所有國家出口額(p5上側)

    p5_right = cal_itc_country_diff(df, YEAR, START_MM, END_MM)
    p5_down = cal_itc_eight_diff(df, YEAR, START_MM, END_MM)
    p5_up = cal_itc_country(df, YEAR, START_MM, END_MM)

    ## 4. 計算特定年度，特定月份區間，ICT產業出口總額&同比年增率，給出產業排行(p6-7)
    p6_p7 = cal_ict_indus_diff(df, YEAR, START_MM, END_MM)

    ## 5. 計算特定年度，特定月份區間，非ICT產業各國出口差額，給出各國排行(p8右側)
    ## 計算特定年度，特定月份區間，非ict產業特定國家(7國)出口差額(p8下側)
    ## 計算特定年度，特定月份區間，非ict產業所有國家出口額(p8上側)

    p8_right = cal_nonitc_country_diff(df, YEAR, START_MM, END_MM)
    p8_down = cal_nonitc_eight_diff(df, YEAR, START_MM, END_MM)
    p8_up = cal_nonitc_country(df, YEAR, START_MM, END_MM)

    ## 6.  計算特定年度，特定月份區間，非ICT產業出口總額&同比年增率，給出產業排行(p9-12)
    p9_p12 = cal_nonict_indus_diff(df, YEAR, START_MM, END_MM)


    # 7. 輸出所有df資料成為一Excel檔(Q1-Q3)

    lst_export = [p1, p3, p4_down, p4_up, p4_right, p5_right, p5_down, p5_up,
                  p6_p7, p8_right, p8_down, p8_up, p9_p12]
    lst_sheetname = ['(P1)台灣出口總額原始資料(百萬美元)', '(P3)文字數據(千美元)', '(P4左下)7國出口增減額(萬美元)',
                     '(P4上圖)出口至各國貿易額map(萬美元)', '(P4右側)產業出口成長衰退前五',
                     '(P5右側)ICT出口成長衰退國前五(萬美元)', '(P5左下)ICT7國貿易差額(萬美元)',
                     '(P5上側)ICT出口額map(萬美元)', '(P6-7)ICT產業出口(萬美元)',
                     '(P8右側)非ICT出口成長衰退國前五(萬美元)', '(P8左下)非ICT7國貿易差額(萬美元)',
                     '(P8上側)非ICT出口額map(萬美元)', '(p9-12)非ICT產業出口(萬美元)']
    export_to_excel(lst_export, lst_sheetname, Export_Excel_file)


if __name__ == '__main__':
    today = f"{datetime.datetime.today():%Y-%m-%d}"
    Import_CSV_file = 'export_file/bureau_{}.csv'.format(today)  # 讀取的CSV檔名
    Export_Excel_file = 'export_file/bureau_quarterly_excel_{}.xlsx'.format(today)  # 輸出的Excel檔名

    # ============ 選取報告所需年度[可編輯] ======================================
    YEAR = 2022  # 輸入選取報告所需年度
    START_MM = 1  # 數據起始月份
    END_MM = 3  # 數據結束月份
    # ==================================================================

    main()
