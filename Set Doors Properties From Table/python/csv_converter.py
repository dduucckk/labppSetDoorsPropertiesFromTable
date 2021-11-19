# -*- coding: utf-8 -*-

# import xml.dom.minidom
# import urllib.request
import csv
#from combinatorics import *
#from itertools import combinations


filename = '../ЗОНЫ И ДВЕРИ.csv'
filenameoutput = '../ЗОНЫ И ДВЕРИ_new.csv'

def read_and_adapt(filename):
    with open(filename, newline='\n') as csvfile:
        reader = csv.reader(csvfile, delimiter=';', quotechar="'")
        newcsv = ''
        for row in reader:
            newcsv += ';'.join(row)+'\n'
        newcsv = newcsv.replace('\n";',';')
        newcsv = newcsv.replace('"','')
        #print(newcsv)
        # 1. заменить все \n--------\n на символ $
        newcsv_ = newcsv.replace('\n--------\n','$')
        #print(newcsv_)
        # 2. разбить по переносам строк
        newcsvlist = newcsv_.split('\n')
        #print(newcsvlist)
        # 3. разбить по ;
        newcsvsublist = []
        for i in newcsvlist:
            newcsvsublist.append(i.split(';'))
        #print(newcsvsublist)
        # 4. разбить по $
        finalcsvlist = []
        for i in newcsvsublist:
            temp = []
            for k in i:
                if '$' in k:
                    temp.append(k.split('$'))
                else:
                    temp.append([k])
            finalcsvlist.append(temp)
        print(finalcsvlist)
        return finalcsvlist[:-1]


def combinatorics(data):
    # cross fit
    # для каждой двери каждого свойства
    # итерируются другие свойства. Это ужас
    # ll == line length, для каждой ячейки её длина
    # далее идёт говнокод
    adapted = []
    def lldef(ll,adapted_line): # не пригодилось
        if ll > 1:
            ll -= 1
        return ll
    def assignvalues(ll,i):
        out = []
        #print('проверка ',i,ll)
        for k in range(9):
            a = '"'+str(i[k][ll[k]-1])+'"'
            out.append(a)
        return out
    for i in data: # строки
        # по длине каждой ячейки проходим все остальные
        # сегодня без рекурсии, задрало
        # а как без рекурсии?
        #v0, v1, v2, v3, v4, v5, v6, v7, v8 = '','','','','','','','',''
        adapted_line = []
        ll = [len(k) for k in i] # длины каждой ячейки
        # так делать нельзя:
        '''while sum(ll) > 0:
            for r,t in enumerate(ll):
                if t>0:
                    ll[r] = t-1
                    print(ll)
                    adapted_line.append(assignvalues(ll,i))'''
        # главное не запутаться в этом дерьме
        for t0 in range(ll[0]):
            ll0 = ll.copy()
            ll0[0] = t0
            for t1 in range(ll[1]):
                ll1 = ll0.copy()
                ll1[1] = t1
                for t2 in range(ll[2]):
                    ll2 = ll1.copy()
                    ll2[2] = t2
                    for t3 in range(ll[3]):
                        ll3 = ll2.copy()
                        ll3[3] = t3
                        for t4 in range(ll[4]):
                            ll4 = ll3.copy()
                            ll4[4] = t4
                            for t5 in range(ll[5]):
                                ll5 = ll4.copy()
                                ll5[5] = t5
                                for t6 in range(ll[6]):
                                    ll6 = ll5.copy()
                                    ll6[6] = t6
                                    for t7 in range(ll[7]):
                                        ll7 = ll6.copy()
                                        ll7[7] = t7
                                        for t8 in range(ll[8]):
                                            ll8 = ll7.copy()
                                            ll8[8] = t8
                                            #print(ll8)
                                            adapted_line.append(assignvalues(ll8,i))
            #print(adapted_line)
        adapted.append(adapted_line)
    return adapted


def writefile(filename, data):
    # write new csv
    with open(filename, 'w', newline='') as csvfile:
        spamwriter = csv.writer(csvfile, delimiter=';',
                                quotechar='"', quoting=csv.QUOTE_MINIMAL)
        for group in data:
            for line in group:
                spamwriter.writerow(line)
    return

if __name__ == '__main__':
    data = read_and_adapt(filename)
    #for i in data:
    #    print(i[4],i)
    adapted = combinatorics(data)
    #print(len(adapted))
    #for i in adapted[2]:
    #    print(*i)
    writefile(filenameoutput, adapted)
