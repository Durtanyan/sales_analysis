# -*- coding: utf-8 -*-
"""
Created on Sun Oct 11 12:52:49 2020

@author: lukin
"""

"""
Задача: распарсить журнал логов и построить отчет по самым 
популярным браузерам и самым продаваемым товарам.
Дано:
Файл логов: logs.xlsx
Файл-отчет: report.xlsx

Программа должна анализировать данные из файла logs.xlsx 
и результаты вычислений записывать в файл report.xlsx.

Необходимо вычислить:

1. 7 самых популярных браузеров. 
Названия браузеров заполнить в ячейках A5-A11. 
Ячейки “Количество посещений” заполнить количеством 
посещений для каждого браузера по месяцам.

2. 7 самых популярных товаров. 
Названия товаров заполнить в ячейках A5-A11.
Ячейки “Количество продаж” заполнить количеством продаж каждого товара 
с учетом того, что 1 посетитель купил 1 единицу товара.

3. Заполнить раздел “Предпочтения”, 
вычислив самые популярные и самые не востребованные товары 
среди мужчин и женщин. 
Самый популярный товар - товар с наибольшим количеством продаж. 
Самый невостребованный - с наименьшим.
"""
import openpyxl as op
import pandas

# Читаем файл эксель и результат передаем в переменную excel_data
# Переменная excel_data имеет тип <class 'pandas.core.frame.DataFrame'>
excel_data = pandas.read_excel('logs.xlsx', sheet_name='log')

# Преобразуем переменную excel_data в словарь с помощью метода to_dict()
# Результат передаем в переменную excel_data_dict
excel_data_dict = excel_data.to_dict()
#print(excel_data_dict)

#получаем колонку браузеров
dict_brauser = excel_data_dict['Браузер']
#dict_data = excel_data_dict['Дата посещения']
#print(len(dict_brauser), len(dict_data))
#создаем set для получения задействованных в аналитике браузеров
#и поместим в в этот сет значения dict_brauser.values()
#таким образом мы избежим дублирования
set_brauzer = set()
for value in dict_brauser.values():
    set_brauzer.add(value)
#print(set_brauzer)

# Преобразуем переменную excel_data в словарь с помощью метода to_dict()
# Результат передаем в переменную excel_data_dict
excel_data_dict = excel_data.to_dict(orient='records')
list_brauzer_and_months = list()
for element_dict in excel_data_dict:
    element_data = element_dict['Дата посещения']
    element_data = element_data.to_pydatetime()
    element_data = element_data.date()
    element_data = str(element_data)
    list_brauzer_and_months.append([element_dict['Браузер'], element_data[5:7]])
    print(list_brauzer_and_months)
print(len(list_brauzer_and_months))





#wb = op.load_workbook(filename='report.xlsx')