# -*- coding: utf-8 -*-

import csv


filename = '../ЗОНЫ И ДВЕРИ_new.csv'
filedoor = '../src/КОРПУС_В.csv'

def read_and_adapt(filename):
    with open(filename, newline='\n') as csvfile:
        reader = csv.reader(csvfile, delimiter=';', quotechar="'")
        doors = []

        for row in reader:

        return

if __name__ == "__main__":
    read_and_adapt(filename)
