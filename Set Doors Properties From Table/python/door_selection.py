# -*- coding: utf-8 -*-
"""
 door_selection.py
 идея - вызывать этот скрипт из setDoorPropertiesFromTable.cpp
 получая словарь csv, где на каждую характеристику двери есть список
 из наименований (ОВ ДУ СС ЭОМ ПУИ тамбур квартира)
 и категорий помещений (Жилая нежилая технические МОП)
 и таблицу cpp из GUID дверей (номер) и помещений и категорий,
 здесь мы получаем словарь ключей в виде уникальных списков характеристик дверей
 со значениями в виде списков GUID дверей.
 Возвращая это таблицей мы можем определить количество типов для проекта
 и исходя из таблицы пронумеровать типы скриптом setDoorPropertiesFromTable.cpp?
 вернув в него таблицу из GUID дверей и их категориями дверей.
"""

import csv

Корпус = 'В'
filecats = '../ЗОНЫ И ДВЕРИ EXPORT.csv'
filedoor = '../src/КОРПУС_'+Корпус+'.csv'
имяфайла = '../IDдвери_тип_'+Корпус+'.csv'

def read_and_adapt(filename):
    with open(filename, newline='\n') as csvfile:
        reader = csv.reader(csvfile, delimiter=';', quotechar="|")
        return reader

def соотнесение(категории, двери):
    for d in двери:
        d0 = d[0]
        for c in категории:
            if d in
    return

def вывод(filename, data):
    # write new csv
    with open(filename, 'w', newline='') as csvfile:
        spamwriter = csv.writer(csvfile, delimiter=';',
                                quotechar='|', quoting=csv.QUOTE_MINIMAL)
        for group in data:
            for line in group:
                spamwriter.writerow(line)
    return

if __name__ == "__main__":
    категории = read_and_adapt(filecats)
    двери = read_and_adapt(filedoor)
    вывод = соотнесение(категории, двери)
    вывести(имяфайла, вывод)
