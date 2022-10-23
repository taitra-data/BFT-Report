

### A. Anaconda介紹及安裝教學 

https://medium.com/python4u/anaconda%E4%BB%8B%E7%B4%B9%E5%8F%8A%E5%AE%89%E8%A3%9D%E6%95%99%E5%AD%B8-f7dae6454ab6

### B. jupyter notebook教學

1. https://zanzan.tw/archives/14578
2. https://chienhung1519.medium.com/jupyter-notebook-%E5%9F%BA%E6%9C%AC%E6%93%8D%E4%BD%9C-6a1d624e2255

備註: Python基礎教學課程
>Python for every body 中文版
https://www.bilibili.com/video/BV16b411n7U4/

### C. 需要安裝的套件:

(Anaconda 請至 Navigator > environment 安裝套件)
(若使用Anaconda內建jupyter notebook的預設環境，以下套件都已安裝好了)
>1. fonttools 
>2. numpy
>3. pandas
>4. matplotlib
>5. pyodbc
>6. openpyxl

<br>除了pip下載以上套件......也可直接調用'requirements.txt'文件設置環境。
若是使用pipenv環境管理，可調用pipfile.

備註: Python SQL driver(pyodbc)說明文件:
https://docs.microsoft.com/en-us/sql/connect/python/python-driver-for-sql-server?view=sql-server-ver15

### D. 本程序檔案與資料夾說明: 

1. Python檔案: (同ipynb檔)
> [主要運行程序]   <br>
main.py  (計算季報原始數據主程式)    <br> 
Get_Quarter_Report.py (輸出Q1-Q3所需EXECL數據表) <br>
Get_Yearly_Report.py (輸出Q4所需EXECL數據表與圖片)          <br>


Q1-Q3季報產生運行順序: *main.py -> Get_Quarter_Report.py*  <br>
Q4季報產生運行順序:  *main.py -> Get_Yearly_Report.py* <br>
**(注意: 所有檔案年份設定需一致，程式才可正常計算)** 

[jupyter notebook運行方式] <br>
Q1-Q3季報產生運行順序: *main.ipynb -> Get_Quarter_Report.ipynb*  <br>
Q4季報產生運行順序:  *main.ipynb -> Get_Yearly_Report.ipynb* <br>
**(注意: 所有檔案年份設定需一致，程式才可正常計算)** 

> [程序引用function]     <br>
db_connect.py (連接資料庫用function)    <br> 
basic_report_cal.py (計算原始數據function) <br>
stats_fcn.py (輸出整理後數據function)     <br>
plot_fcn.py (做圖用function) 

2. 其他檔案:

> secrets.py (資料庫連接帳號密碼)    <br> 
README.md (本程序說明文件) <br>
requirements.txt (使用套件文檔)  <br>
Pipfile等檔案 (pipenv環境管理檔，可忽略) 

3. 資料夾:

> import_file (本程式引用的文件位置) <br>
> export_file (本程式csv, excel檔案輸出的文件位置)  <br>
plot (本程式圖片輸出的文件位置)  

資料夾詳細說明如下:

    a. import_file (存放輸入程式的文件): 
        Industry_hscode_20220513.tsv (季報專用產業的HSCODE對照表TSV檔)

    b. export_file (程式輸出的文件位置):
        bureau_YYYY-MM-dd.csv (原始數據計算結果)
        bureau_quarterly_excel_YYYY-MM-dd.xlsx (Q1-Q3簡報數據)
        bureau_yearly_excel_YYYY-MM-dd.xlsx (Q4簡報數據)
    
    c.plot (圖片存放位置) 
        *.jpg (Q4簡報用圖檔)
        TaipeiSansTCBeta-Regular.ttf (圖表中文字型)
