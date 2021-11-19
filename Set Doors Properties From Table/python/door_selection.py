# -*- coding: utf-8 -*-

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
