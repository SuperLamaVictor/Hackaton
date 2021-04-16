import os
import pprint
import pandas as pd
import matplotlib.pyplot as plt
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

statik_first = {}
for date, count, money in zip(df1["Дата"][:3857], df1["Продажи в кг"][:3857], df1["Сумма в ценах продажи"][:3857]):
    if date not in things:
        statik_first[date] = [count, money]
    else:
        statik_first[date][0] += count      # количество в кг
        statik_first[date][1] += money         # продажа

statik_second = {}
for date, count, money in zip(df1["Дата"][3858:], df1["Продажи в кг"][3858:], df1["Сумма в ценах продажи"][3858:]):
    if date not in statik_second:
        if date not in things:
            statik_second[date] = [count, money]
        else:
            statik_second[date][0] += count  # количество в кг
            statik_second[date][1] += money  # продажа

date_1, summa_kg_1, summa_1 = [], [], []
date_2, summa_kg_2, summa_2 = [], [], []
print(list(statik_first.values()))
for i in sorted(zip(statik_first.keys(), list(statik_first.values())), key=lambda x: (x[0].split('.')[1], x[0].split('.')[0])):
    date_1.append(i[0])
    summa_kg_1.append(i[1][0])
    summa_1.append(i[1][1])
print(date_1)

for i in sorted(zip(statik_second.keys(), list(statik_second.values())), key=lambda x: (x[0].split('.')[1], x[0].split('.')[0])):
    date_1.append(i[0])
    summa_kg_1.append(i[1][0])
    summa_1.append(i[1][1])

print("Рост продажи товаров (в кг) составил: " + str(int(sum(summa_2) / sum(summa_1) * 100)) + "%")
'''
plt.style.use('dark_background')  # чёерный фон
plt.plot(date, summa)
plt.show()
'''
