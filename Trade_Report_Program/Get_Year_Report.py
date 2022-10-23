"""
Purpose of code: 利用計算好的CSV數據與資料庫搜索，產生[貿易局季報Q4年度總結]所需數據源 EXCEL檔案。
"""

from stats_fcn import *
import datetime


def main():
    start_month, end_month = 1, 12  # 數據起始月份、結束月份
    df = pd.read_csv(Import_CSV_file)

    ## 1. 計算特定年度之前，台灣3年內出口總額(p2)(百萬美元)
    ## 計算特定年度，特定月份區間，所有國家出口差額(p5上側文字)(千美元)
    p2 = cal_total_export(YEAR)
    p5 = cal_total_country_diff(YEAR, end_month)

    # 2. 計算特定年度，特定月份區間，所有產業出口差額，給出產業排行(p6右側)(萬美元)
    # 計算特定年度，特定月份區間，各國出口額(p6上側)(萬美元)
    # 計算特定年度，特定月份區間，特定國家(7國)出口差額(p6下側)(萬美元)

    p6_right = cal_industry_diff(df, YEAR, start_month, end_month)
    p6_up = cal_total_country(YEAR, end_month)
    p6_down = cal_total_eight_diff(df, YEAR, start_month, end_month)

    ## 3. 計算特定年度，特定月份區間，ICT產業各國出口差額，給出各國排行(p7右側)
    ## 計算特定年度，特定月份區間，ICT產業特定國家(7國)出口差額(p7下側)
    ## 計算特定年度，特定月份區間，ICT產業所有國家出口額(p7上側)

    p7_right = cal_itc_country_diff(df, YEAR, start_month, end_month)
    p7_down = cal_itc_eight_diff(df, YEAR, start_month, end_month)
    p7_up = cal_itc_country(df, YEAR, start_month, end_month)

    ## 4. 計算單一年度，ICT產業出口總額&差額&同比年增率，給出產業排行bar(p8)(千萬美元)
    # 計算特定年度，特定月份區間，ICT產業出口總額&同比年增率，給出產業排行(p9-10)
    p8 = cal_itc_indus_rank(df, YEAR, start_month, end_month,
                            ICT_diff_width, ICT_diff_fontsize, ICT_grow_width, ICT_grow_fontsize)
    p9_p10 = cal_ict_indus_diff(df, YEAR, start_month, end_month)

    ## 5. 計算特定年度，特定月份區間，非ICT產業各國出口差額，給出各國排行(p11右側)
    ## 計算特定年度，特定月份區間，非ict產業特定國家(7國)出口差額(p11下側)
    ## 計算特定年度，特定月份區間，非ict產業所有國家出口額(p11上側)

    p11_right = cal_nonitc_country_diff(df, YEAR, start_month, end_month)
    p11_down = cal_nonitc_eight_diff(df, YEAR, start_month, end_month)
    p11_up = cal_nonitc_country(df, YEAR, start_month, end_month)

    ## 6. 計算單一年度，非ICT產業出口總額&差額&同比年增率，給出產業排行bar(p12)(千萬美元)
    # 計算特定年度，特定月份區間，非ICT產業出口總額&同比年增率，給出產業排行(p13-16)
    p12 = cal_nonitc_indus_rank(df, YEAR, start_month, end_month,
                                nonICT_diff_width, nonICT_diff_fontsize, nonICT_grow_width, nonICT_grow_fontsize)
    p13_p16 = cal_nonict_indus_diff(df, YEAR, start_month, end_month)

    ## 7. 輸出所有df資料成為一Excel檔(Q4)
    lst_export = [p2, p5, p6_down, p6_up, p6_right, p7_right, p7_down, p7_up, p8,
                  p9_p10, p11_right, p11_down, p11_up, p12, p13_p16]
    lst_sheetname = ['(P2)台灣出口總額原始資料(百萬美元)', '(P5)文字數據(千美元)', '(P6左下)7國出口增減額(萬美元)',
                     '(P6上圖)出口至各國貿易額map(萬美元)', '(P6右側)產業出口成長衰退前五',
                     '(P7右側)ICT出口成長衰退國前五(萬美元)', '(P7左下)ICT7國貿易差額(萬美元)',
                     '(P7上側)ICT出口額map(萬美元)', '(p8)ICT出口產品表現bar(千萬美元)',
                     '(p9-10)ICT產業出口(萬美元)', '(P11右側)非ICT出口成長衰退國前五(萬美元)',
                     '(P11左下)非ICT7國貿易差額(萬美元)', '(P11上側)非ICT出口額map(萬美元)',
                     '(p12)非ICT出口產品表現bar(千萬美元)', '(p13-16)非ICT產業出口(萬美元)']
    export_to_excel(lst_export, lst_sheetname, Export_Excel_file)



if __name__ == '__main__':
    today = f"{datetime.datetime.today():%Y-%m-%d}"
    Import_CSV_file = 'export_file/bureau_{}.csv'.format(today)  # 讀取的CSV檔名
    Export_Excel_file = 'export_file/bureau_yearly_excel_{}.xlsx'.format(today)  # 輸出的Excel檔名

    # ============ 選取報告所需年度[可編輯] ======================================
    YEAR = 2022  # 輸入選取報告所需年度

    # ============ 圖片調整[可編輯]  ======================================
    # 調整p8-ICT產業圖片的長條圖寬度、與數字大小調整(差額長條圖寬度, 差額數字大小, 增減率長條圖寬度, 增減率大小)
    ICT_diff_width, ICT_diff_fontsize, ICT_grow_width, ICT_grow_fontsize = 0.6, 8, 0.6, 8
    # 調整p12-非ICT產業圖片的長條圖寬度、與數字大小調整(差額長條圖寬度, 差額數字大小, 增減率長條圖寬度, 增減率大小)
    nonICT_diff_width, nonICT_diff_fontsize, nonICT_grow_width, nonICT_grow_fontsize = 0.6, 8, 0.6, 6
    # ==================================================================

    main()
