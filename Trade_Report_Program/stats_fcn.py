# -*- coding: utf-8 -*-
"""
Purpose of code: calculate industrial statistics for 貿易局簡報 2022 Q1-Q4.

"""

from db_connect import *
from plot_fcn import *
from decimal import Decimal
import pandas as pd
import numpy as np


# 輸出資料成Excel檔
def export_to_excel(lst_export, lst_name, excel_file):
    with pd.ExcelWriter(excel_file) as writer:
        for df, sheet_name in zip(lst_export, lst_name):
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print('Successfully export data to Excel file: ', excel_file)

# 計算特定年度之前，台灣3年內出口總額(百萬美元)
def cal_total_export(year):
    print('計算特定年度之前，台灣3年內出口總額(百萬美元)\n')

    sql_select_statement = """
        SELECT Year, Month, CAST(SUM(Value_Month_USD) AS float)
        FROM [itrade_original].[dbo].[MOF_DATA_TXT]	
        where Year > {} and Year < {} and Ex_Im in (1,4)
        group by Year, Month
        order by Year, Month;
    """.format(year - 4, year + 1)

    cnxn = create_connection()  # create a database connection

    if cnxn is not None:
        result = get_itrade_db_data(cnxn, sql_select_statement)
        df = pd.DataFrame(result, columns=['Year', 'Month', 'USD_value'])
        cnxn.close()

    else:
        print("Error! cannot create the database connection.")


    # data calculation
    df['pre_export'] = [None for i in range(12)] + df.USD_value.tolist()[0:-12]  # 同月前一年出口額
    df['diff'] = df['USD_value'] - df['pre_export']
    df_cut = df[df['Year'].astype(int) > year-3].copy()
    df_cut['per_change'] = df_cut['diff'].div(df_cut['pre_export']).mul(100)
    # export data to a proper format
    df_result = df_cut[['Year', 'Month', 'USD_value', 'per_change']].copy()
    df_result['USD_value'] = df_result['USD_value'].div(1000)  # 單位千美元轉換至單位百萬美元
    df_result['per_change'] = df_result['per_change'].map(
        lambda x: Decimal(x).quantize(Decimal('.01'), rounding='ROUND_HALF_UP'))
    df_result = df_result.rename(columns={"Year": "年", "Month": "月", "USD_value": "出口總值", "per_change": "增減率"})
    # print(df_result.to_string())
    return df_result

# 計算特定年度，特定月份區間，所有國家出口差額(千美元)
def cal_total_country_diff(year, end_month):
    print('計算特定年度，特定月份區間，所有國家出口差額(千美元)\n')

    sql_select_statement = """
        SELECT Year, Zh_Name, CAST(SUM(Value_Month_USD) AS float)
        FROM [itrade_original].[dbo].[MOF_DATA_TXT]	
            JOIN [itrade_original].[dbo].[MOF_COUNTRY] on 	
                ([MOF_DATA_TXT].Country = [itrade_original].[dbo].[MOF_COUNTRY].Country)
        where Ex_Im in (1,4) and Year > {} and Year < {} and month < {} and Zh_Name not like '其他%'
        group by Year, Zh_Name;
     """.format(year - 2, year + 1, end_month + 1)

    cnxn = create_connection()
    if cnxn is not None:
        result = get_itrade_db_data(cnxn, sql_select_statement)
        df = pd.DataFrame(result, columns=['Year', 'country', 'USD_value'])
        cnxn.close()
    else:
        print("Error! cannot create the database connection.")

    # data calculation
    pd.set_option('display.float_format', lambda x: '%.5f' % x)  # not show float data in scientific notation
    df_reshape = df.pivot(index='country', columns='Year')['USD_value'].reset_index()
    df_reshape.columns.name = None
    df_reshape[str(year)] = df_reshape[str(year)].fillna(0)
    df_reshape[str(year - 1)] = df_reshape[str(year - 1)].fillna(0)
    df_reshape['diff'] = df_reshape[str(year)] - df_reshape[str(year - 1)]
    df_reshape['per_change'] = df_reshape['diff'].div(df_reshape[str(year - 1)]).mul(100)
    # export data to a proper format
    df_result = df_reshape.replace([np.inf, -np.inf], np.nan)
    df_result['per_change'] = df_result['per_change'].map(
        lambda x: Decimal(x).quantize(Decimal('.01'), rounding='ROUND_HALF_UP'))
    df_result = df_result.sort_values(by=['diff'], ascending=False)
    df_result = df_result.rename(columns={"country": "國家", "diff": "差額(千美元)", "per_change": "2021增減率"})
    # print(df_result.to_string())
    return df_result

# 計算特定年度，特定月份區間，所有產業出口差額，給出產業排行(萬美元)
def cal_industry_diff(df, year, start_month, end_month):
    print('計算特定年度，特定月份區間，所有產業出口差額，給出產業排行(萬美元)\n')
    select = (df['Year'] >= year - 1) & (df['Year'] <= year) & (df['Month'] >= start_month) \
             & (df['Month'] <= end_month) & (df['major'] != 'ICT產業') & (df['major'] != '非ICT產業')
    df_sh = df[select]
    df_sh = df_sh.groupby(['Year', 'major'], as_index=False).sum()[['Year', 'major', 'Value_USD']]
    df_reshape = df_sh.pivot(index='major', columns='Year')['Value_USD'].reset_index()
    df_reshape.columns.name = None
    df_reshape[year] = df_reshape[year].fillna(0)
    df_reshape[year - 1] = df_reshape[year - 1].fillna(0)
    df_reshape['difference'] = df_reshape[year] - df_reshape[year - 1]
    df_reshape['per_change'] = df_reshape['difference'].div(df_reshape[year - 1]).mul(100)
    # calculation is done, now export data to a proper format
    df_cut = df_reshape.iloc[:, [1, 2, 3]].div(10)
    df_cut = df_cut.applymap(lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    df_result = pd.concat([df_reshape.iloc[:, 0], df_cut], axis=1)
    df_result['per_change'] = df_reshape['per_change'].map(
        lambda x: Decimal(x).quantize(Decimal('.01'), rounding='ROUND_HALF_UP'))
    df_result = df_result.sort_values(by=['difference'], ascending=False)
    lst_colname = ['產業', year - 1, year, '差額(萬美元)', '年增率']
    df_result = df_result.rename(columns=dict(zip(df_result.columns, lst_colname)))
    # print(df_result.head(7))
    return df_result

# 計算特定年度，特定月份區間，各國出口額(萬美元)
def cal_total_country(year, end_month):
    print('計算特定年度，特定月份區間，各國出口額(萬美元)\n')

    sql_select_statement = """
        SELECT Year, Zh_Name, CAST(SUM(Value_Month_USD) AS float)
        FROM [itrade_original].[dbo].[MOF_DATA_TXT]	
            JOIN [itrade_original].[dbo].[MOF_COUNTRY] on 	
                ([MOF_DATA_TXT].Country = [itrade_original].[dbo].[MOF_COUNTRY].Country)
        where Ex_Im in (1,4) and Year > {} and Year < {} and month < {} and Zh_Name not like '其他%'
        group by Year, Zh_Name;
     """.format(year - 1, year + 1, end_month + 1)

    cnxn = create_connection()
    if cnxn is not None:
        result = get_itrade_db_data(cnxn, sql_select_statement)
        df = pd.DataFrame(result, columns=['Year', 'country', 'Value_USD'])
        cnxn.close()
    else:
        print("Error! cannot create the database connection.")

    # data calculation
    pd.set_option('display.float_format', lambda x: '%.5f' % x)  # not show float data in scientific notation
    df_reshape = df.pivot(index='country', columns='Year')['Value_USD'].reset_index()
    df_reshape.columns.name = None
    df_reshape[str(year)] = df_reshape[str(year)].fillna(0)
    # export data to a proper format
    df_result = df_reshape.sort_values(by=[str(year)], ascending=False)
    df_result[str(year)] = df_result[str(year)].div(10)
    df_result[str(year)] = df_result[str(year)].map(
        lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    lst_colname = ['國家', str(year) + '(萬美元)']
    df_result = df_result.rename(columns=dict(zip(df_result.columns, lst_colname)))
    # print(df_result.to_string())
    return df_result

# 計算特定年度，特定月份區間，特定國家(7國)出口差額(萬美元)
def cal_total_eight_diff(df, year, start_month, end_month):
    print('計算特定年度，特定月份區間，特定國家(7國)出口差額(萬美元)\n')

    sql_select_statement = """
        SELECT Year, Zh_Name, CAST(SUM(Value_Month_USD) AS float)
        FROM [itrade_original].[dbo].[MOF_DATA_TXT]	
            JOIN [itrade_original].[dbo].[MOF_COUNTRY] on 	
                ([MOF_DATA_TXT].Country = [itrade_original].[dbo].[MOF_COUNTRY].Country)
        where Ex_Im in (1,4) and Year > {} and Year < {} and month < {}
            and Zh_Name in ('中國大陸', '美國', '日本', '南韓', '香港', '新加坡', '德國')
        group by Year, Zh_Name;
     """.format(year - 3, year + 1, end_month + 1)

    cnxn = create_connection()
    if cnxn is not None:
        result = get_itrade_db_data(cnxn, sql_select_statement)
        df = pd.DataFrame(result, columns=['Year', 'country', 'USD_value'])
        cnxn.close()
    else:
        print("Error! cannot create the database connection.")

    # data calculation
    df_sh = df.groupby(['Year', 'country'], as_index=False).sum()[['Year', 'country', 'USD_value']]
    df_reshape = df_sh.pivot(index='country', columns='Year')['USD_value'].reset_index()
    df_reshape.columns.name = None
    df_reshape[str(year)] = df_reshape[str(year)].fillna(0)
    df_reshape[str(year - 1)] = df_reshape[str(year - 1)].fillna(0)
    df_reshape['diff_year-1'] = df_reshape[str(year - 1)] - df_reshape[str(year - 2)]
    df_reshape['diff_year'] = df_reshape[str(year)] - df_reshape[str(year - 1)]
    # export data to a proper format
    df_cut = df_reshape.iloc[:, [1, 2, 3, 4, 5]].div(10)
    df_cut = df_cut.applymap(lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    df_result = pd.concat([df_reshape.iloc[:, 0], df_cut], axis=1)
    lst_colname = ['國家', year - 2, year - 1, year, str(year - 1) + '差額', str(year) + '差額']
    df_result = df_result.rename(columns=dict(zip(df_result.columns, lst_colname)))
    # sort country by specific order
    label_dict = {'中國大陸': 0, '南韓': 3, '德國': 6, '新加坡': 5, '日本': 2, '美國': 1, '香港': 4}
    df_result['label'] = df_result['國家'].apply(lambda x: label_dict[x])
    df_result = df_result.sort_values('label', ascending=True).drop(columns=['label'])
    # print(df_result)
    return df_result

# 計算特定年度，特定月份區間，ICT產業各國出口差額，給出各國排行(萬美元)
def cal_itc_country_diff(df, year, start_month, end_month):
    print('計算特定年度，特定月份區間，ICT產業各國出口差額，給出各國排行(萬美元)\n')
    select = (df['Year'] >= year - 1) & (df['Year'] <= year) & (df['Month'] >= start_month) \
             & (df['Month'] <= end_month) & (df['major'] == 'ICT產業')
    df_sh = df[select]
    df_sh = df_sh.groupby(['Year', 'Country'], as_index=False).sum()[['Year', 'Country', 'Value_USD']]
    df_reshape = df_sh.pivot(index='Country', columns='Year')['Value_USD'].reset_index()
    df_reshape.columns.name = None
    df_reshape[year] = df_reshape[year].fillna(0)
    df_reshape[year - 1] = df_reshape[year - 1].fillna(0)
    df_reshape['difference'] = df_reshape[year] - df_reshape[year - 1]
    df_reshape['per_change'] = df_reshape['difference'].div(df_reshape[year - 1]).mul(100)
    df_reshape = df_reshape.replace([np.inf, -np.inf], np.nan)
    # calculation is done, now export data to a proper format
    df_cut = df_reshape.iloc[:, [1, 2, 3]].div(10)
    df_cut = df_cut.applymap(lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    df_result = pd.concat([df_reshape.iloc[:, 0], df_cut], axis=1)
    df_result['per_change'] = df_reshape['per_change'].map(
        lambda x: Decimal(x).quantize(Decimal('.01'), rounding='ROUND_HALF_UP'))
    df_result = df_result.sort_values(by=['difference'], ascending=False)
    lst_colname = ['國家', year - 1, year, '差額', '年增率']
    df_result = df_result.rename(columns=dict(zip(df_result.columns, lst_colname)))
    # print(df_result.head(10))
    return df_result

# 計算特定年度，特定月份區間，ICT產業特定國家(7國)出口差額(萬美元)
def cal_itc_eight_diff(df, year, start_month, end_month):
    print('計算特定年度，特定月份區間，ict產業特定國家(7國)出口差額(萬美元)\n')
    select = (df['Year'] >= year - 2) & (df['Year'] <= year) & (df['Month'] >= start_month) \
             & (df['Month'] <= end_month) & (df['major'] == 'ICT產業')
    df_sh = df[select]
    df_sh = df_sh[df_sh['Country'].isin(['中國大陸', '美國', '日本', '南韓', '香港', '新加坡', '德國'])]
    df_sh = df_sh.groupby(['Year', 'Country'], as_index=False).sum()[['Year', 'Country', 'Value_USD']]
    df_reshape = df_sh.pivot(index='Country', columns='Year')['Value_USD'].reset_index()
    df_reshape.columns.name = None
    df_reshape[year] = df_reshape[year].fillna(0)
    df_reshape[year - 1] = df_reshape[year - 1].fillna(0)
    df_reshape['diff_year-1'] = df_reshape[year - 1] - df_reshape[year - 2]
    df_reshape['diff_year'] = df_reshape[year] - df_reshape[year - 1]
    # calculation is done, now export data to a proper format
    df_cut = df_reshape.iloc[:, [1, 2, 3, 4, 5]].div(10)
    df_cut = df_cut.applymap(lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    df_result = pd.concat([df_reshape.iloc[:, 0], df_cut], axis=1)
    df_result = df_result.sort_values(by=['diff_year'], ascending=False)
    lst_colname = ['國家', year - 2, year - 1, year, str(year - 1) + '差額', str(year) + '差額']
    df_result = df_result.rename(columns=dict(zip(df_result.columns, lst_colname)))
    # sort country by specific order
    label_dict = {'中國大陸': 0, '南韓': 3, '德國': 6, '新加坡': 5, '日本': 2, '美國': 1, '香港': 4}
    df_result['label'] = df_result['國家'].apply(lambda x: label_dict[x])
    df_result = df_result.sort_values('label', ascending=True).drop(columns=['label'])
    # print(df_result)
    return df_result

# 計算特定年度，特定月份區間，ICT產業所有國家出口額(萬美元)
def cal_itc_country(df, year, start_month, end_month):
    print('計算特定年度，特定月份區間，ict產業所有國家出口額(萬美元)\n')
    select = (df['Year'] == year) & (df['Month'] >= start_month) \
             & (df['Month'] <= end_month) & (df['major'] == 'ICT產業')
    df_sh = df[select]
    # df_sh = df_sh[~df_sh["country"].str.contains('其他')]
    df_sh = df_sh.groupby(['Year', 'Country'], as_index=False).sum()[['Year', 'Country', 'Value_USD']]
    df_reshape = df_sh.pivot(index='Country', columns='Year')['Value_USD'].reset_index()
    df_reshape.columns.name = None
    df_reshape[year] = df_reshape[year].fillna(0).div(10)
    # calculation is done, now export data to a proper format
    df_result = df_reshape.sort_values(by=[year], ascending=False)
    df_result[year] = df_result[year].map(lambda x: Decimal(x).quantize(Decimal('0'),
                                                                        rounding='ROUND_HALF_UP'))
    lst_colname = ['國家', str(year) + '(萬美元)']
    df_result = df_result.rename(columns=dict(zip(df_result.columns, lst_colname)))
    # print(df_result.to_string())
    return df_result

# 計算單一年度，ICT產業出口總額&差額&同比年增率，給出產業排行bar(千萬美元)
def cal_itc_indus_rank(df, year, start_month, end_month,
                       diff_width=0.6, diff_fontsize=8, grow_width=0.6, grow_fontsize=8):
    print('計算單一年度，ICT產業出口總額&差額&同比年增率，給出產業排行bar(千萬美元)\n')
    select = (df['Year'] >= year - 1) & (df['Year'] <= year) & (df['major'] != 'ICT產業') \
             & (df['major'] != '非ICT產業') & (df['minor'] == 'ICT產業') & (df['Month'] >= start_month) \
             & (df['Month'] <= end_month)
    df_sh = df[select]
    df_sh = df_sh.groupby(['Year', 'major'], as_index=False).sum()[['Year', 'major', 'Value_USD']]
    df_reshape = df_sh.pivot(index='major', columns='Year')['Value_USD'].reset_index()
    df_reshape.columns.name = None
    df_reshape[year] = df_reshape[year].fillna(0)
    df_reshape['difference'] = df_reshape[year] - df_reshape[year - 1]
    df_reshape['per_change'] = df_reshape['difference'].div(df_reshape[year - 1]).mul(100)
    df_reshape = df_reshape.replace([np.inf, -np.inf], np.nan)
    df_reshape = df_reshape.sort_values(by=['difference'], ascending=False)
    # calculation is done, now export data to a proper format
    df_cut = df_reshape.iloc[:, [1, 2, 3]].div(10000)  # 單位千美元轉換至單位千萬美元
    df_cut = df_cut.applymap(lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    df_result = pd.concat([df_reshape.iloc[:, 0], df_cut], axis=1)
    df_result['per_change'] = df_reshape['per_change'].map(
        lambda x: Decimal(str(x)).quantize(Decimal('.0'), rounding='ROUND_HALF_UP'))  # 利用Decimal取正確四捨五入
    df_result['per_change'] = df_result['per_change'].map(lambda x: str(x))  # Decimal轉回str
    lst_colname = ['產業', year - 1, year, '增減額', '年增率']
    df_result = df_result.rename(columns=dict(zip(df_result.columns, lst_colname)))
    # print(df_result.to_string())
    # plot the bar charts
    data1 = dict(zip(df_result['產業'], df_result['增減額']))
    data2 = dict(zip(df_result['產業'], df_result['年增率']))
    plot_bar(data1, fname='ICT_diff.jpg', digit=False, width=diff_width, fontsize=diff_fontsize)
    plot_bar(data2, fname='ICT_growth.jpg', width=grow_width, fontsize=grow_fontsize)
    return df_result

# 計算特定年度，特定月份區間，ICT產業出口總額&同比年增率，給出產業排行(萬美元)
def cal_ict_indus_diff(df, year, start_month, end_month):
    print('計算特定年度，特定月份區間，ICT產業出口總額&同比年增率，給出產業排行(萬美元)\n')
    select = (df['Year'] >= year - 2) & (df['Year'] <= year) & (df['Month'] >= start_month) \
             & (df['Month'] <= end_month) & (df['major'] != 'ICT產業') & (df['major'] != '非ICT產業') \
             & (df['minor'] == 'ICT產業')
    df_sh = df[select]
    df_sh = df_sh.groupby(['Year', 'major'], as_index=False).sum()[['Year', 'major', 'Value_USD']]
    df_reshape = df_sh.pivot(index='major', columns='Year')['Value_USD'].reset_index()
    df_reshape.columns.name = None
    df_reshape[year] = df_reshape[year].fillna(0)
    df_reshape[year - 1] = df_reshape[year - 1].fillna(0)
    df_reshape['difference_1'] = df_reshape[year - 1] - df_reshape[year - 2]
    df_reshape['difference_2'] = df_reshape[year] - df_reshape[year - 1]
    df_reshape['per_change_1'] = df_reshape['difference_1'].div(df_reshape[year - 2]).mul(100)
    df_reshape['per_change_2'] = df_reshape['difference_2'].div(df_reshape[year - 1]).mul(100)
    df_reshape = df_reshape.replace([np.inf, -np.inf], np.nan)
    df_reshape = df_reshape.sort_values(by=['per_change_2'], ascending=False)
    # calculation is done, now export data to a proper format
    df_cut = df_reshape.iloc[:, [1, 2, 3]].div(10)
    df_cut = df_cut.applymap(lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    df_result = pd.concat([df_reshape.iloc[:, 0], df_cut], axis=1)
    df_result['per_change_1'] = df_reshape['per_change_1'].map(
        lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    df_result['per_change_2'] = df_reshape['per_change_2'].map(
        lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    lst_colname = ['產業', year - 2, year - 1, year, str(year - 1) + '年增率', str(year) + '年增率']
    df_result = df_result.rename(columns=dict(zip(df_result.columns, lst_colname)))
    # print(df_result.to_string())
    return df_result

# 計算特定年度，特定月份區間，非ICT產業各國出口差額，給出各國排行(萬美元)
def cal_nonitc_country_diff(df, year, start_month, end_month):
    print('計算特定年度，特定月份區間，非ICT產業各國出口差額，給出各國排行(萬美元)\n')
    select = (df['Year'] >= year - 1) & (df['Year'] <= year) & (df['Month'] >= start_month) \
             & (df['Month'] <= end_month) & (df['major'] == '非ICT產業')
    df_sh = df[select]
    df_sh = df_sh.groupby(['Year', 'Country'], as_index=False).sum()[['Year', 'Country', 'Value_USD']]
    df_reshape = df_sh.pivot(index='Country', columns='Year')['Value_USD'].reset_index()
    df_reshape.columns.name = None
    df_reshape[year] = df_reshape[year].fillna(0)
    df_reshape[year - 1] = df_reshape[year - 1].fillna(0)
    df_reshape['difference'] = df_reshape[year] - df_reshape[year - 1]
    df_reshape['per_change'] = df_reshape['difference'].div(df_reshape[year - 1]).mul(100)
    df_reshape = df_reshape.replace([np.inf, -np.inf], np.nan)
    # calculation is done, now export data to a proper format
    df_cut = df_reshape.iloc[:, [1, 2, 3]].div(10)
    df_cut = df_cut.applymap(lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    df_result = pd.concat([df_reshape.iloc[:, 0], df_cut], axis=1)
    df_result['per_change'] = df_reshape['per_change'].map(
        lambda x: Decimal(x).quantize(Decimal('.01'), rounding='ROUND_HALF_UP'))
    df_result = df_result.sort_values(by=['difference'], ascending=False)
    lst_colname = ['國家', year - 1, year, '差額', str(year) + '年增率']
    df_result = df_result.rename(columns=dict(zip(df_result.columns, lst_colname)))
    # print(df_result.head(10))
    return df_result

# 計算特定年度，特定月份區間，非ICT產業特定國家(7國)出口差額(萬美元)
def cal_nonitc_eight_diff(df, year, start_month, end_month):
    print('計算特定年度，特定月份區間，非ict產業特定國家(7國)出口差額(萬美元)\n')
    select = (df['Year'] >= year - 2) & (df['Year'] <= year) & (df['Month'] >= start_month) \
             & (df['Month'] <= end_month) & (df['major'] == '非ICT產業')
    df_sh = df[select]
    df_sh = df_sh[df_sh['Country'].isin(['中國大陸', '美國', '日本', '南韓', '香港', '新加坡', '德國'])]
    df_sh = df_sh.groupby(['Year', 'Country'], as_index=False).sum()[['Year', 'Country', 'Value_USD']]
    df_reshape = df_sh.pivot(index='Country', columns='Year')['Value_USD'].reset_index()
    df_reshape.columns.name = None
    df_reshape[year] = df_reshape[year].fillna(0)
    df_reshape[year - 1] = df_reshape[year - 1].fillna(0)
    df_reshape['diff_year-1'] = df_reshape[year - 1] - df_reshape[year - 2]
    df_reshape['diff_year'] = df_reshape[year] - df_reshape[year - 1]
    # calculation is done, now export data to a proper format
    df_cut = df_reshape.iloc[:, [1, 2, 3, 4, 5]].div(10)
    df_cut = df_cut.applymap(lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    df_result = pd.concat([df_reshape.iloc[:, 0], df_cut], axis=1)
    lst_colname = ['國家', year - 2, year - 1, year, str(year - 1) + '差額', str(year) + '差額']
    df_result = df_result.rename(columns=dict(zip(df_result.columns, lst_colname)))
    # sort country by specific order
    label_dict = {'中國大陸': 0, '南韓': 3, '德國': 6, '新加坡': 5, '日本': 2, '美國': 1, '香港': 4}
    df_result['label'] = df_result['國家'].apply(lambda x: label_dict[x])
    df_result = df_result.sort_values('label', ascending=True).drop(columns=['label'])
    # print(df_result)
    return df_result

# 計算特定年度，特定月份區間，非ICT產業所有國家出口額(萬美元)
def cal_nonitc_country(df, year, start_month, end_month):
    print('計算特定年度，特定月份區間，非ict產業所有國家出口額(萬美元)\n')
    select = (df['Year'] == year) & (df['Month'] >= start_month) \
             & (df['Month'] <= end_month) & (df['major'] == '非ICT產業')
    df_sh = df[select]
    # df_sh = df_sh[~df_sh["country"].str.contains('其他')]
    df_sh = df_sh.groupby(['Year', 'Country'], as_index=False).sum()[['Year', 'Country', 'Value_USD']]
    df_reshape = df_sh.pivot(index='Country', columns='Year')['Value_USD'].reset_index()
    df_reshape.columns.name = None
    df_reshape[year] = df_reshape[year].fillna(0).div(10)
    # calculation is done, now export data to a proper format
    df_result = df_reshape.sort_values(by=[year], ascending=False)
    df_result[year] = df_result[year].map(lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    lst_colname = ['國家', str(year) + '(萬美元)']
    df_result = df_result.rename(columns=dict(zip(df_result.columns, lst_colname)))
    # print(df_result.to_string())
    return df_result

# 計算單一年度，非ICT產業出口總額&差額&同比年增率，給出產業排行(千萬美元)
def cal_nonitc_indus_rank(df, year, start_month, end_month,
                          diff_width=0.6, diff_fontsize=8, grow_width=0.6, grow_fontsize=8):
    print('計算單一年度，非ICT產業出口總額&差額&同比年增率，給出產業排行(千萬美元)\n')
    select = (df['Year'] >= year - 1) & (df['Year'] <= year) & (df['major'] != 'ICT產業') \
             & (df['major'] != '非ICT產業') & (df['minor'] != 'ICT產業') & (df['Month'] >= start_month) \
             & (df['Month'] <= end_month)
    df_sh = df[select]
    df_sh = df_sh.groupby(['Year', 'major'], as_index=False).sum()[['Year', 'major', 'Value_USD']]
    df_reshape = df_sh.pivot(index='major', columns='Year')['Value_USD'].reset_index()
    df_reshape.columns.name = None
    df_reshape[year] = df_reshape[year].fillna(0)
    df_reshape['difference'] = df_reshape[year] - df_reshape[year - 1]
    df_reshape['per_change'] = df_reshape['difference'].div(df_reshape[year - 1]).mul(100)
    df_reshape = df_reshape.replace([np.inf, -np.inf], np.nan)
    df_reshape = df_reshape.sort_values(by=['difference'], ascending=False)
    # calculation is done, now export data to a proper format
    df_cut = df_reshape.iloc[:, [1, 2, 3]].div(10000)  # 單位千美元轉換至單位千萬美元
    df_cut = df_cut.applymap(lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    df_result = pd.concat([df_reshape.iloc[:, 0], df_cut], axis=1)
    df_result['per_change'] = df_reshape['per_change'].map(
        lambda x: Decimal(str(x)).quantize(Decimal('.0'), rounding='ROUND_HALF_UP'))  # 利用Decimal取正確四捨五入
    df_result['per_change'] = df_result['per_change'].map(lambda x: str(x))  # Decimal轉回str
    lst_colname = ['產業', year - 1, year, '增減額', '年增率']
    df_result = df_result.rename(columns=dict(zip(df_result.columns, lst_colname)))
    df_result = df_result.replace({"產業": {"電動車及相關零組件": "電動車", "照明": "照明光電"}})  # 修正產業命名
    df_result.drop(df_result[df_result['產業'] == '光電'].index, inplace=True)  # 刪除重複產業'光電'
    # print(df_result.to_string())
    # plot the bar charts
    data1 = dict(zip(df_result['產業'], df_result['增減額']))
    data2 = dict(zip(df_result['產業'], df_result['年增率']))
    plot_bar(data1, fname='nonICT_diff.jpg', digit=False, width=diff_width, fontsize=diff_fontsize)
    plot_bar(data2, fname='nonICT_growth.jpg', width=grow_width, fontsize=grow_fontsize)
    return df_result

# 計算特定年度，特定月份區間，非ICT產業出口總額&同比年增率，給出產業排行(萬美元)
def cal_nonict_indus_diff(df, year, start_month, end_month):
    print('計算特定年度，特定月份區間，非ICT產業出口總額&同比年增率，給出產業排行(萬美元)\n')
    select = (df['Year'] >= year - 2) & (df['Year'] <= year) & (df['Month'] >= start_month) \
             & (df['Month'] <= end_month) & (df['major'] != 'ICT產業') & (df['major'] != '非ICT產業') \
             & (df['minor'] != 'ICT產業')
    df_sh = df[select]
    df_sh = df_sh.groupby(['Year', 'major'], as_index=False).sum()[['Year', 'major', 'Value_USD']]
    df_reshape = df_sh.pivot(index='major', columns='Year')['Value_USD'].reset_index()
    df_reshape.columns.name = None
    df_reshape[year] = df_reshape[year].fillna(0)
    df_reshape[year - 1] = df_reshape[year - 1].fillna(0)
    df_reshape['difference_1'] = df_reshape[year - 1] - df_reshape[year - 2]
    df_reshape['difference_2'] = df_reshape[year] - df_reshape[year - 1]
    df_reshape['per_change_1'] = df_reshape['difference_1'].div(df_reshape[year - 2]).mul(100)
    df_reshape['per_change_2'] = df_reshape['difference_2'].div(df_reshape[year - 1]).mul(100)
    df_reshape = df_reshape.replace([np.inf, -np.inf], np.nan)
    df_reshape = df_reshape.sort_values(by=['per_change_2'], ascending=False)
    # calculation is done, now export data to a proper format
    df_cut = df_reshape.iloc[:, [1, 2, 3]].div(10)
    df_cut = df_cut.applymap(lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    df_result = pd.concat([df_reshape.iloc[:, 0], df_cut], axis=1)
    df_result['per_change_1'] = df_reshape['per_change_1'].map(
        lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    df_result['per_change_2'] = df_reshape['per_change_2'].map(
        lambda x: Decimal(x).quantize(Decimal('0'), rounding='ROUND_HALF_UP'))
    lst_colname = ['產業', year - 2, year - 1, year, str(year - 1) + '年增率', str(year) + '年增率']
    df_result = df_result.rename(columns=dict(zip(df_result.columns, lst_colname)))
    # print(df_result.to_string())
    return df_result


# def main():
#     df = pd.read_csv(CSV_file)


if __name__ == '__main__':
    pass
    # CSV_file = 'export_file/bureau_20220412.csv'  # 讀取的CSV檔名
    # Export_Excel_file = 'export_file/bureau_excel_20220412.xlsx'  # 輸出的Excel檔名
    # main()


