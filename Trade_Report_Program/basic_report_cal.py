"""Purpose of code: 計算貿易局季報所需的原始資料，輸出為CSV格式.

1. 計算出口原始資料出現的所有11碼HSCODE(非重複)。
2. 引用 TSV 產業對照表資訊，儲存資訊至 dictionary。
3. 利用產業對照表的多類HS碼，創建產業 11碼 HSCode的對照表，儲存資訊至 JSON。
4. 連接資料庫，利用 11碼資訊查詢各產業貿易額，儲存為 pandas dataframe。
   並輸出為CSV檔案: export_file/bureau_XXXX-XX-XX.csv
"""

from db_connect import *
import pandas as pd
import csv


def extract_unique_hscode(cnxn, year):
    """ extract unique hscode11 from selected data 讀取原始資料存在的 HS code(11碼)。
    :param cnxn: Connection object
    :param year: the start year of selected data
    :return: a dictionary with unique HS11
    """
    year = str(int(year) - 1)
    sql_select_statement = '''
          SELECT DISTINCT HSCode
          FROM [itrade_original].[dbo].[MOF_DATA_TXT]
          where year > {} and Ex_Im in (1,4)
          ORDER BY HSCode;
    '''.format(str(year))

    lst_hs11 = get_itrade_db_data(cnxn, sql_select_statement)
    dict_hs11 = {i[0]: None for i in lst_hs11}  # {key:value} = {hscode11:None}
    print('Successfully extract unique HS11 from DB export data.')
    return dict_hs11

def get_indus_hscode(indus_tsv_file):
    """ get industrial category of HSCode from a TSV file. 讀取TSV檔產業HS對照表(HS2-HS10)。
    :param indus_tsv_file: a TSV file describes industry HSCODE
    :return: a dictionary of each industrial category with a list of HSCode(HS2-HS10).
    """
    with open(indus_tsv_file, "r", encoding="utf-8") as file:
        tsv_file = csv.reader(file, delimiter="\t")
        next(tsv_file, None)  # skip the headers
        lst_indus = list(tsv_file)

    dict_indus = {}
    for i in lst_indus:
        major, detail, hscode = i[0], i[1], i[2]  # 產業對照表大項, 產業對照表細項, 產業HSCODE

        if major not in dict_indus.keys():
            dict_indus[major] = {}

        if detail not in dict_indus[major].keys():
            dict_indus[major][detail] = None

        dict_indus[major][detail] = hscode.split(',')

    print('Successfully extract industrial HS2-HS10 from a TSV file')
    return dict_indus

def create_indus_hs11(dict_indus, dict_hs11):
    """ create dict of industry categories with list of HS11. (創建產業HS2-HS10轉換成HS11的dictionary。)
    :param dict_indus: a dictionary with original industry HSCode(HS2-HS10).
    :return: a dictionary of industrial categories with HS11.
    """
    lst_hs11 = list(dict_hs11.keys())  # 原始資料有的全部HS11碼
    dict_indus_hs11 = {}

    for major in dict_indus.keys():  # 產業對照表資訊
        for detail in dict_indus[major].keys():
            lst_hs = dict_indus[major][detail]
            for hs in lst_hs:
                hs11 = [i for i in lst_hs11 if i.startswith(hs)]

                if major not in dict_indus_hs11.keys():
                    dict_indus_hs11[major] = {}
                if detail not in dict_indus_hs11[major].keys():
                    dict_indus_hs11[major][detail] = []

                dict_indus_hs11[major][detail].extend(hs11)

    for k in dict_indus_hs11.keys():
        for v in dict_indus_hs11[k].keys():
            dict_indus_hs11[k][v] = list(set(dict_indus_hs11[k][v]))  # remove duplicates
            # print(k, v, dict_indus_hs11[k][v])

    print('Successfully extract industrial info of HS11 from current data')
    return dict_indus_hs11

def calc_indus_value(cnxn, year, insertion):
    """ Input a list of HS11, get the SUM of USD value from iTrade Database
    :param cnxn: Connection object
    :param year: the start year of selected data
    :param insertion: a list of hs11
    :return: a pandas dataframe with columns = ('Year', 'Month', 'country', 'USD_value')
    """
    year = str(int(year) - 1)

    if len(insertion) == 1:
        insertion = '(\'{}\')'.format(insertion[0])
    else:
        insertion = tuple(insertion)

    sql_select_statement = '''
        SELECT Year, Month, Zh_Name, CAST(SUM(Value_Month_USD) AS float)
        FROM [itrade_original].[dbo].[MOF_DATA_TXT]	
            JOIN [itrade_original].[dbo].[MOF_COUNTRY] on 	
                ([MOF_DATA_TXT].Country = [itrade_original].[dbo].[MOF_COUNTRY].Country)
        where Year > {} and Ex_Im in (1,4) and HSCode in {}
        group by Year, Month, Zh_Name;
    '''.format(year, insertion)

    lst_data = get_itrade_db_data(cnxn, sql_select_statement)  # get data from SQL Server
    df = pd.DataFrame(lst_data, columns=['Year', 'Month', 'Country', 'Value_USD'])
    return df

def export_all_bureau_csv(cnxn, year, indus_hs11, fname):
    """ Calculate all Industrial TradeValue from iTrade DB, the info of HS11 of all industries
    is from "INDUSTRY_HS11".
    :param cnxn: Connection object
    :param year: the start year of selected data.
    :param indus_hs11: a dictionary containing industrial HS11 info.
    :param fname: The file name of csv data to be exported.
    :return: None
    """

    print("\nStart to get all industrial export data from DB")

    first = True
    for major in indus_hs11.keys():
        for minor in indus_hs11[major].keys():
            lst_hs11 = indus_hs11[major][minor]
            df = calc_indus_value(cnxn, year, lst_hs11)  # call DB to get sum TradeValue from list of HS11
            df['major'], df['minor'] = major, minor
            df['date'] = df['Year'] + '/' + df['Month'] + '/01'
            print(major, minor, df.shape)

            if first:
                df_join = df
                first = False
            else:
                df_join = pd.concat([df_join, df])

    print('-' * 20)
    print('Total data size: ', df_join.shape)
    df_join.to_csv(fname, index=False, encoding="utf_8_sig")
    print('Successfully export csv file: ', fname)

def main():
    pass

if __name__ == '__main__':
    pass
    # YEAR = '2020'  # set the start year of data
    # INDUSTRY_TABLE = 'import_file/Industry_hscode_20220513.tsv'  # 貿易局產業HS對照表資訊(須定期更新)。
    # main()
