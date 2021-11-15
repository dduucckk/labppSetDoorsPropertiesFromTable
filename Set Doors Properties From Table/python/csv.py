# -*- coding: utf-8 -*-


# import xml.dom.minidom
# import urllib.request
import csv


with open('../ЗОНЫ И ДВЕРИ.csv', newline='') as csvfile:
    spamreader = csv.reader(csvfile, delimiter=';', quotechar='"')
    for row in spamreader:
        print(', '.join(row))

'''
with open('../ЗОНЫ И ДВЕРИ_вывод.csv', 'w', newline='') as csvfile:
    spamwriter = csv.writer(csvfile, delimiter=' ',
                            quotechar='|', quoting=csv.QUOTE_MINIMAL)
    spamwriter.writerow(['Spam'] * 5 + ['Baked Beans'])
    spamwriter.writerow(['Spam', 'Lovely Spam', 'Wonderful Spam'])
'''