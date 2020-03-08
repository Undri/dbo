import tkinter as tk
from tkinter import ttk
import pandas as pd


def make_result_label(sheet, point, name):
    ans = sheet.sort_values(by=point, ascending=False)[name].tolist()
    text = '1)' + ' ' + ans[0] + '\n' + '2)' + ' ' + ans[1] + '\n' + '3)' + ' ' + ans[2]
    out = tk.Label(text=text)
    out.pack()


def kostil():
    widget_list = root.winfo_children()
    # ATTENTION!! DO NOT TRY THIS AT HOME!! КОСТЫЛЬ!!!
    # Нужен для того, чтобы заменять прошлый результат поиска банка(виджет Label) новым
    if len(widget_list) > 7:
        for w in widget_list[7:]:
            w.destroy()


def show():
    global clients
    global detailed_pressed
    global checks
    if not detailed_pressed:
        current_client = (choose_type.get(), choose_category.get())
        kostil()
        sheet = pd.read_excel('output.xlsx', clients[current_client][0])
        make_result_label(sheet, clients[current_client][1], clients[current_client][2])
        return
    else:
        print("detailed pressed")
        print(len(checks))


def detailed():
    kostil()
    global detailed_pressed
    global checks
    checks = list()
    detailed_pressed = True
    if str(choose_category.get()) == '' and str(choose_type.get()) == '':
        error_label = tk.Label(text='Недостаточно информации')
        error_label.pack()
    else:
        current_client = (choose_type.get(), choose_category.get())
        sheet = pd.read_excel('output.xlsx', clients[current_client][0])
        for col in sheet.columns[1:-2]:
            var = tk.BooleanVar()
            check = tk.Checkbutton(text=col, variable=var)
            check.pack()
            checks.append(check)


clients = {('Физическое лицо', 'Удобство'):
               ('Лист7', 'pointFiz', 'Списки/Оценка преимуществ и недостатков услуг(общее)'),
           ('Физическое лицо', 'Низкая стоимость услуг'): ('Лист1', 'point', 'Списки банков'),
           ('Физическое лицо', 'Безопасность'):
               ('Лист5', 'pointSec', 'Списки/Оценка преимуществ и недостатков различных услуг(физические лица):'),
           ('Юридическое лицо', 'Удобство'):
               ('Лист7', 'pointUr', 'Списки/Оценка преимуществ и недостатков услуг(общее)'),
           ('Юридическое лицо', 'Низкая стоимость услуг'):
               ('Лист2', 'point', 'Списки/Стоим.услуг для малого и среднего бизнеса'),
           ('Юридическое лицо', 'Безопасность'):
               ('Лист5', 'pointSec', 'Списки/Оценка преимуществ и недостатков различных услуг(физические лица):'),
           ('Корпорация', 'Удобство'):
               ('Лист6', 'pointInter', 'Списки/Оценка преимуществ и недостатков услуг(корпорации)'),
           ('Корпорация', 'Низкая стоимость услуг'): ('Лист3', 'point', 'Списки/Стоим.услуг для корпораций:'),
           ('Корпорация', 'Безопасность'):
               ('Лист6', 'pointSec', 'Списки/Оценка преимуществ и недостатков услуг(корпорации)')}


root = tk.Tk()
root.title("DBO")
root.geometry("550x1050")
root.resizable(False, False)

detailed_pressed = False
checks = list()

choose_type_label = tk.Label(text='Тип субъекта')
choose_type_label.pack()

choose_type = ttk.Combobox(values=['Физическое лицо', 'Юридическое лицо', 'Корпорация'])
choose_type.pack()

choose_cat_label = tk.Label(text='Наиболее значимая характеристика')
choose_cat_label.pack()

choose_category = ttk.Combobox(values=['Удобство', 'Низкая стоимость услуг', 'Безопасность'])
choose_category.pack()

detailed_but = tk.Button(text="Подробнее..", command=detailed)
detailed_but.pack()

result_but = tk.Button(text="Подобрать банк", command=show)
result_but.pack()

exit_but = tk.Button(text='Выход', command=root.destroy)
exit_but.pack()

root.mainloop()
