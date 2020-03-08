from tkinter import *
import pandas as pd
import numpy as np


sheet1 = pd.read_excel('Banki.xlsx', 'Лист1')
sheet1 = sheet1.drop(columns=['Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10'])
sheet1['SMS-банкинг'].loc[sheet1['SMS-банкинг'] == 'нет'] = np.nan
sheet1['Открытие вклада(100 000 рублей, 1год)'].loc[sheet1['Открытие вклада(100 000 рублей, 1год)'] == 'нет'] = np.nan
frees = ['SMS-банкинг', 'Интернет-банкинг(подкл)', 'Инт-б(использование)', 'Мобильный-банкинг(подкл)', 'Моб-б(исп)']
for x in frees:
    sheet1.loc[sheet1[x] == 'бесплатно', x] = 0
sheet1.index = sheet1['Списки банков']
sheet1 = sheet1.drop(columns=['Списки банков'])

sheet2 = pd.read_excel('Banki.xlsx', 'Лист2')
for x in sheet2.columns:
    sheet2[x].loc[sheet2[x] == 'бесплатно'] = 0
    sheet2[x].loc[sheet2[x] == 'нет'] = np.nan
sheet2.index = sheet2['Списки/Стоим.услуг для малого и среднего бизнеса']
sheet2 = sheet2.drop(columns=['Списки/Стоим.услуг для малого и среднего бизнеса'])

sheet3 = pd.read_excel('Banki.xlsx', 'Лист3')
for x in sheet3.columns:
    sheet3[x].loc[sheet3[x] == 'бесплатно'] = 0
    sheet3[x].loc[sheet3[x] == 'нет'] = np.nan
sheet3.index = sheet3['Списки/Стоим.услуг для корпораций:']
sheet3 = sheet3.drop(columns=['Списки/Стоим.услуг для корпораций:'])

sheet5 = pd.read_excel('Banki.xlsx', 'Лист5')
sheet5 = sheet5.drop(columns=['Unnamed: 1'])
for x in sheet5.columns:
    sheet5[x].loc[sheet5[x] == 'есть'] = 1
    sheet5[x].loc[sheet5[x] == 'нет'] = 0
sheet5.columns = sheet5.iloc[0]
sheet5 = sheet5.drop(0)
sheet5.index = sheet5['Списки/Оценка преимуществ и недостатков различных услуг(физические лица):']
sheet5 = sheet5.drop(columns=['Списки/Оценка преимуществ и недостатков различных услуг(физические лица):'])
"""
sheet5: первые 2 колонки- интерфейс, остальные - безопасность, последняя - доверие
"""

sheet6 = pd.read_excel('Banki.xlsx', 'Лист6')
sheet6.columns = sheet6.iloc[0]
sheet6 = sheet6.drop(0)
sheet6.index = sheet6['Списки/Оценка преимуществ и недостатков услуг(корпорации)']
sheet6 = sheet6.drop(columns=['Списки/Оценка преимуществ и недостатков услуг(корпорации)'])
"""
sheet6: последние 2- интерфейс, остальные- безопасность
"""

sheet7 = pd.read_excel('Banki.xlsx', 'Лист7')
sheet7.columns = sheet7.iloc[0]
sheet7 = sheet7.drop(0)
sheet7.index = sheet7['Списки/Оценка преимуществ и недостатков услуг(общее)']
sheet7 = sheet7.drop(columns=['Списки/Оценка преимуществ и недостатков услуг(общее)'])

"""
sheet7: первые 3- ФЮ, 4- Ф, остальные- Ю
"""



root = Tk()
root.title('Брокер ДБО')
root.resizable(False, False)
f_top = Frame(root)
f_bot = Frame(root)
b = Button(f_top, text="Физическое лицо", width=20, height=3, font=('CenturyGothic', 12))
b1 = Button(f_top, text="Юридическое лицо", width=20, height=3, font=('CenturyGothic', 12))
b2 = Button(f_top, text="Корпорацию", width=20, height=3, font=('CenturyGothic', 12))
l1 = Label(bg='white', fg='black', width=60, font=('CenturyGothic', 16))
l1['text'] = 'Вы представляете'

b['activebackground'] = 'green'
b1['activebackground'] = 'green'
b2['activebackground'] = 'green'


def fiz(event):
    fizwindow = Toplevel()
    fizwindow.title('Брокер ДБО Физическое лицо')
    fizwindow.geometry("600x400")
    fizwindow.resizable(False, False)

def ur(event):
    urwindow = Toplevel()
    urwindow.title('Брокер ДБО Юридическое лицо')
    urwindow.geometry("600x400")
    urwindow.resizable(False, False)

def korp(event):
    korpwindow = Toplevel()
    korpwindow.title('Брокер ДБО Корпорация')
    korpwindow.geometry("600x400")
    korpwindow.resizable(False, False)


b.bind('<Button-1>', fiz)
b1.bind('<Button-1>', ur)
b2.bind('<Button-1>', korp)

l1.pack(side='top', pady=55)
f_top.pack()
f_bot.pack(pady= 55)
b.pack(side='left', padx= 5)
b1.pack(side='left', padx= 5)
b2.pack(side='left', padx= 5)
root.geometry("600x400")
root.mainloop()
