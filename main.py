import pandas as pd
import numpy as np


def sheet_price_rating(weight, sheet):
    sheet.loc[:, 'point'] = 0
    for i in range(len(sheet.columns[:-1])):
        for bank in sheet.index:
            if sheet.loc[bank, sheet.columns[i]] == np.nan:
                sheet.loc[bank, 'point'] -= 1 * weight[i]
            elif sheet.loc[bank, sheet.columns[i]] < sheet[sheet.columns[i]].mean():
                j = 1
                while sheet.loc[bank, sheet.columns[i]] < sheet[sheet.columns[i]].mean() - j * sheet[
                    sheet.columns[i]].std():
                    j += 1
                sheet.loc[bank, 'point'] += j * weight[i] * 2
    point_normalizing()


def point_normalizing():
    if 'point' in sheet2.columns:
        sheet2.loc[:, 'point'] *= 0.6
    if 'point' in sheet1.columns:
        sheet1.loc[:, 'point'] *= 0.85


def fiz_sec_rating(weight):
    sheet5.loc[:, 'pointSec'] = 0
    for i in range(2, 10):
        for bank in sheet5.index:
            sheet5.loc[bank, 'pointSec'] += sheet5.loc[bank, sheet5.columns[i]] * weight[i - 3] / 0.7
    for bank in sheet5.index:
        j = 0
        while sheet5.loc[bank, 'Народный рейтинг'] > sheet5['Народный рейтинг'].mean() + j * sheet5['Народный рейтинг'].std():
            j += 1
        sheet5.loc[bank, 'pointSec'] += j


def ur_sec_rating(weight):
    sheet6.loc[:, 'pointSec'] = 0
    for i in range(2, len(sheet6.columns[:7])):
        for bank in sheet6.index:
            sheet6.loc[bank, 'pointSec'] += sheet6.loc[bank, sheet6.columns[i]] * int(weight[i]) / 5 * 1.5


def fiz_inter_rating(weight):
    sheet5.loc[:, 'pointInter'] = 0
    for i in range(2):
        for bank in sheet5.index:
            sheet5.loc[bank, 'pointInter'] += sheet5.loc[bank, sheet5.columns[i]] * weight[i] / 2


def ur_inter_rating(weight):
    sheet6.loc[:, 'pointInter'] = 0
    for i in range(2):
        for bank in sheet6.index:
            sheet6.loc[bank, 'pointInter'] += sheet6.loc[bank, sheet6.columns[len(sheet6.columns) - 3 - i]] * weight[i]


def fiz_func_rating(weight):
    sheet7.loc[:, 'pointFiz'] = 0
    for i in range(len(sheet7.columns[:3])):
        for bank in sheet7.index:
            sheet7.loc[bank, 'pointFiz'] += sheet7.loc[bank, sheet7.columns[i]] * weight[i] * 0.7


def ur_func_rating(weight):
    sheet7.loc[:, 'pointUr'] = 0
    for i in range(len(sheet7.columns)):
        for bank in sheet7.index:
            sheet7.loc[bank, 'pointUr'] += sheet7.loc[bank, sheet7.columns[i]] * weight[i] * 0.09


sheet1 = pd.read_excel('Banki.xlsx', 'Лист1')
sheet2 = pd.read_excel('Banki.xlsx', 'Лист2')
sheet3 = pd.read_excel('Banki.xlsx', 'Лист3')
sheet5 = pd.read_excel('Banki.xlsx', 'Лист5')
sheet6 = pd.read_excel('Banki.xlsx', 'Лист6')
sheet7 = pd.read_excel('Banki.xlsx', 'Лист7')

# Лист1
sheet1 = sheet1.drop(columns=['Unnamed: 7', 'Unnamed: 8',
                              'Unnamed: 9', 'Unnamed: 10', 'Открытие вклада(100 000 рублей, 1год)'])
sheet1['SMS-банкинг'].loc[sheet1['SMS-банкинг'] == 'нет'] = np.nan
frees = ['SMS-банкинг', 'Интернет-банкинг (подключение)',
         'Интернет-банкинг (использование)', 'Мобильный-банкинг (подключение)', 'Мобильный-банкинг (использование)']
for x in frees:
    sheet1.loc[sheet1[x] == 'бесплатно', x] = 0
sheet1.index = sheet1['Списки банков']
sheet1 = sheet1.drop(columns=['Списки банков'])
sheet1.loc['Саровбизнесбанк', 'SMS-банкинг'] = 60

# Лист2
for x in sheet2.columns:
    sheet2[x].loc[sheet2[x] == 'бесплатно'] = 0
    sheet2[x].loc[sheet2[x] == 'нет'] = np.nan
sheet2.index = sheet2['Списки/Стоим.услуг для малого и среднего бизнеса']
sheet2 = sheet2.drop(columns=['Списки/Стоим.услуг для малого и среднего бизнеса'])

# Лист3
for x in sheet3.columns:
    sheet3[x].loc[sheet3[x] == 'бесплатно'] = 0
    sheet3[x].loc[sheet3[x] == 'нет'] = np.nan
sheet3.index = sheet3['Списки/Стоим.услуг для корпораций:']
sheet3 = sheet3.drop(columns=['Списки/Стоим.услуг для корпораций:'])

# Лист5
sheet5 = sheet5.drop(columns=['Unnamed: 1'])
for x in sheet5.columns:
    sheet5[x].loc[sheet5[x] == 'есть'] = 1
    sheet5[x].loc[sheet5[x] == 'нет'] = 0
sheet5.columns = sheet5.iloc[0]
sheet5 = sheet5.drop(0)
sheet5.index = sheet5['Списки/Оценка преимуществ и недостатков различных услуг(физические лица):']
sheet5 = sheet5.drop(columns=['Списки/Оценка преимуществ и недостатков различных услуг(физические лица):'])

# Лист6
sheet6.columns = sheet6.iloc[0]
sheet6 = sheet6.drop(0)
sheet6.index = sheet6['Списки/Оценка преимуществ и недостатков услуг(корпорации)']
sheet6 = sheet6.drop(columns=['Списки/Оценка преимуществ и недостатков услуг(корпорации)'])

# Лист7
sheet7.columns = sheet7.iloc[0]
sheet7 = sheet7.drop(0)
sheet7.index = sheet7['Списки/Оценка преимуществ и недостатков услуг(общее)']
sheet7 = sheet7.drop(columns=['Списки/Оценка преимуществ и недостатков услуг(общее)', 'Информационные сервисы'])
sheet7 = sheet7.fillna(0)

sheet_price_rating([1, 1, 1, 1, 1], sheet1)
sheet_price_rating([1, 1, 1, 1, 1, 1], sheet2)
sheet_price_rating([1, 1, 1, 1], sheet3)
fiz_inter_rating([1, 1])
fiz_sec_rating([1, 1, 1, 1, 1, 1, 1, 1])
ur_sec_rating([1, 1, 1, 1, 1, 1, 1])
ur_inter_rating([1, 1])
fiz_func_rating([1, 1, 1])
ur_func_rating(np.ones(21))

# Костыль для GUI, чтобы работала функция detailed
sheet1['point1'] = np.nan
sheet2['point1'] = np.nan
sheet3['point1'] = np.nan

# запись в отдельный файл для GUI
with pd.ExcelWriter('output.xlsx') as writer:
    sheet1.to_excel(writer, sheet_name='Лист1')
    sheet2.to_excel(writer, sheet_name='Лист2')
    sheet3.to_excel(writer, sheet_name='Лист3')
    sheet5.to_excel(writer, sheet_name='Лист5')
    sheet6.to_excel(writer, sheet_name='Лист6')
    sheet7.to_excel(writer, sheet_name='Лист7')

