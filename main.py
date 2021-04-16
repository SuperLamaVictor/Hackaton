import os
import pprint
import pandas as pd
import xlrd

book = xlrd.open_workbook('Accenture_Датасет.xlsx')
df1 = pd.read_excel(os.path.join('Accenture_Датасет.xlsx'), engine='openpyxl')
things = {}
for thing, kg, buy, sell in zip(df1["Код товара"], df1["Продажи в кг"], df1["Сумма в ценах закупки"], df1["Сумма в ценах продажи"]):
    if thing not in things:
        things[thing] = [1, kg, buy, sell]
    else:
        things[thing][0] += 1       # количество
        things[thing][1] += kg      # кг
        things[thing][2] += buy     # цена покупки
        things[thing][3] += sell    # цена продажи
for i in things.items():
    a = str()
    for j in range(len(i[1])):
        if j == 0:
            a += "количество " + str(i[1][j])
        elif j == 1:
            a += " кг " + str(i[1][j])
        elif j == 2:
            a += " цена покупки " + str(i[1][j])
        elif j == 3:
            a += " цена продажи " + str(i[1][j])
    print(a)
    print()
