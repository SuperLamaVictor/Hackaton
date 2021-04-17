import os
import xlrd
import pprint
import xlsxwriter
import pandas as pd
import matplotlib.pyplot as plt

book = xlrd.open_workbook('Accenture_Датасет.xlsx')
df1 = pd.read_excel(os.path.join('Accenture_Датасет.xlsx'), engine='openpyxl')
things = {}

for thing, kg, buy, sell in zip(df1["Код товара"][:3857], df1["Продажи в кг"][:3857], df1["Сумма в ценах закупки"][:3857], df1["Сумма в ценах продажи"][:3857]):
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
            a += "; кг " + str(i[1][j])
        elif j == 2:
            a += "; сумма покупок " + str(int(i[1][j]))
        elif j == 3:
            a += "; суииа продаж " + str(int(i[1][j]))
            a += "; чистая прибль товара " + str(int(i[1][j] - i[1][j - 1]))
    print(a)
print()
things_2 = {}

for thing, kg, buy, sell in zip(df1["Код товара"][3857:], df1["Продажи в кг"][3857:], df1["Сумма в ценах закупки"][3857:], df1["Сумма в ценах продажи"][3857:]):
    if thing not in things_2:
        things_2[thing] = [1, kg, buy, sell]
    else:
        things_2[thing][0] += 1       # количество
        things_2[thing][1] += kg      # кг
        things_2[thing][2] += buy     # цена покупки
        things_2[thing][3] += sell    # цена продажи

for i in things_2.items():
    a = str()
    for j in range(len(i[1])):
        if j == 0:
            a += "количество " + str(i[1][j])
        elif j == 1:
            a += "; кг " + str(i[1][j])
        elif j == 2:
            a += "; сумма покупок " + str(int(i[1][j]))
        elif j == 3:
            a += "; суииа продаж " + str(int(i[1][j]))
            a += "; чистая прибль товара " + str(int(i[1][j] - i[1][j - 1]))
    print(a)

print()

statik_first = {}
for date, count, money in zip(df1["Дата"][:3857], df1["Продажи в кг"][:3857], df1["Сумма в ценах продажи"][:3857]):
    if date not in statik_first:
        statik_first[date] = [count, money]
    else:
        statik_first[date][0] += count      # количество в кг
        statik_first[date][1] += money      # продажа

statik_second = {}
for date, count, money in zip(df1["Дата"][3858:], df1["Продажи в кг"][3858:], df1["Сумма в ценах продажи"][3858:]):
    if date not in statik_second:
        statik_second[date] = [count, money]
    else:
        statik_second[date][0] += count  # количество в кг
        statik_second[date][1] += money  # продажа

date_1, summa_kg_1, summa_1 = [], [], []
date_2, summa_kg_2, summa_2 = [], [], []

for i in sorted(zip(statik_first.keys(), list(statik_first.values())), key=lambda x: (x[0].split('.')[1], x[0].split('.')[0])):
    date_1.append(i[0])
    summa_kg_1.append(i[1][0])
    summa_1.append(i[1][1])

for i in sorted(zip(statik_second.keys(), list(statik_second.values())), key=lambda x: (x[0].split('.')[1], x[0].split('.')[0])):
    date_2.append(i[0])
    summa_kg_2.append(i[1][0])
    summa_2.append(i[1][1])

print("Изменения продажи товаров (в кг) составил:  \t" + str(int((sum(summa_kg_2) - sum(summa_kg_1)) / sum(summa_kg_1) * 10000) / 100) + "%  \tс 2015 по 2016")
print("Изменения продажи выручки товаров составил: \t" + str(int((sum(summa_2)- sum(summa_1)) / sum(summa_1) * 10000) / 100) + "%  \tс 2015 по 2016")
print()
for i in range(3):
    f = list(things.values())[i][3] / list(things.values())[i][1]
    s = list(things_2.values())[i][3] / list(things_2.values())[i][1]
    print("Изменения цен товара " + str(list(things_2.keys())[i]) + " за кг составили: \t" + str(int((s - f) / f * 10000) / 100) + "%  \tс 2015 по 2016")
print()

for i in range(3):
    f = list(things.values())[i][3] - list(things.values())[i][2]
    s = list(things_2.values())[i][3] - list(things_2.values())[i][2]
    print("Изменения прибыли товара " + str(list(things_2.keys())[i]) + " составили: \t\t" + str(int((s - f) / f * 10000) / 100) + "%  \tс 2015 по 2016")
    if f < s:
        f_p = list(things.values())[i][3] / list(things.values())[i][1]
        s_p = list(things_2.values())[i][3] / list(things_2.values())[i][1]

        f_kg = list(things.values())[i][1]
        s_kg = list(things_2.values())[i][1]

        f_buy = list(things.values())[i][3]
        s_buy = list(things_2.values())[i][3]

        k_kg = round((s_kg - f_kg) / f_kg * 100 / 2, 2)
        k_buy = round((s_buy - f_buy) / f_buy * 100 / 2, 2)

        print("Рекомендуемая цена на следующий год: \t\t\t" + str(round(s_p + (s_p-f_p) / f_p / 2 * 100, 2)) + '\t\tс 2016 по 2017')
        print("Примерная доход: \t\t\t\t\t\t\t\t" + str(int(s_buy + s_buy * (k_buy / 100))) + '\t\tс 2016 по 2017')
        print("Примерный объём закупок: \t\t\t\t\t\t" + str(int(s_kg + s_kg * (k_kg / 100)))+ '\t\tс 2016 по 2017')
        print()

    elif f >= s:
        f_p = list(things.values())[i][3] / list(things.values())[i][1]
        s_p = list(things_2.values())[i][3] / list(things_2.values())[i][1]

        f_kg = list(things.values())[i][1]
        s_kg = list(things_2.values())[i][1]

        f_buy = list(things.values())[i][3]
        s_buy = list(things_2.values())[i][3]

        k_kg = abs(round((s_kg - f_kg) / f_kg * 100, 2))
        k_buy = abs(round((s_buy - f_buy) / f_buy * 100, 2))

        print("Рекомендуемая цена на следующий год: \t\t\t" + str(round(s_p - (s_p - f_p) / f_p * 100, 2)) + '\t\tс 2017 по 2018')
        print("Примерная доход: \t\t\t\t\t\t\t\t" + str(int(s_buy + s_buy * (k_buy / 100))) + '\t\tс 2017 по 2018')
        print("Примерный объём закупок: \t\t\t\t\t\t" + str(int(s_kg + s_kg * (k_kg / 100))) + '\t\tс 2017 по 2018')
        print()
print()

for i in range(3):
    f = list(things.values())[i][1]
    s = list(things_2.values())[i][1]
    print("Изменения спроса на товар " + str(list(things_2.keys())[i]) + " составили: \t\t" + str(int((s - f) / f * 10000) / 100) + "%   \tс 2015 по 2016")
print()

print('Изменения спроса при рекомендуемых параметрах')
for i in range(3):
    f_p = list(things.values())[i][3] / list(things.values())[i][1]
    s_p = list(things_2.values())[i][3] / list(things_2.values())[i][1]

    f_kg = list(things.values())[i][1]
    s_kg = list(things_2.values())[i][1]

    f_buy = list(things.values())[i][3]
    s_buy = list(things_2.values())[i][3]

    k_kg = abs(round((s_kg - f_kg) / f_kg * 100 / 2, 2))
    k_buy = abs(round((s_buy - f_buy) / f_buy * 100, 2))
    print("Товар " + str(list(things_2.keys())[i]) + " составят: \t\t" + str(k_kg) + "%   \tс 2017 по 2018")

df = pd.DataFrame({'Месяц': [1, 2, 3],
                    'Код товара': [list(things.keys())[0], list(things.keys())[1], list(things.keys())[2]],
                    'Продажи в 2015': [list(things.values())[0][0], list(things.values())[1][0], list(things.values())[2][0]],
                    'Продажи в 2016': [list(things_2.values())[0][0], list(things_2.values())[1][0], list(things_2.values())[2][0]]})

workbook = xlsxwriter.Workbook('диаграммы.xlsx')
worksheet = workbook.add_worksheet()

# Данные
data = [list(things.values())[0][0], list(things.values())[1][0], list(things.values())[2][0]]
worksheet.write_column('A1', data)

# Тип диаграммы
chart = workbook.add_chart({'type': 'pie'})

# Строим по нашим данным
chart.add_series({'values': '=Sheet1!A1:A3'})
chart.add_series({'categories': "=Sheet1!A1:A3",
                  'values': "=Sheet1!A1:A3",
                  'name': 'Продажи'})
worksheet.insert_chart('B2 x M2', chart)
workbook.close()

df.to_excel('./teams.xlsx')
df.head()
# os.startfile(r'teams.xlsx')
'''
plt.style.use('dark_background')  # чёерный фон
plt.plot(date, summa)
plt.show()
'''
