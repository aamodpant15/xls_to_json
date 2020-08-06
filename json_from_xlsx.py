#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import openpyxl as xl

#open workbook
print("\n")
print("-"*10)
name = input("Name of file in <fileName.xls> or <fileName.xlsx> format\n")
workbook = xl.load_workbook(name)


print("\n")
print("-"*10)
#get sheet
name = input("Name of sheet inside file\n")
sheet = workbook[name]

#member count
memCount= 0

#number of rows
X = sheet.max_row

#number of columns
Y = sheet.max_column

#column titles
attributes= []

for col in range (1, Y+1):
    title = sheet.cell(row = 1, column= col).value
    attributes.append(title)

#members
membas= [] 

#keep count of which attribute from the array is being added
attributeCount= 0

# print(len(attributes))
idCount=1

for rows in range(2, X+1):
    attributeCount= 0
    membas.append({})
    memba = membas[memCount]
    memba["id"] = idCount
    idCount+=1
    for cols in range(1, Y+1):
        attribute= attributes[attributeCount]
        data = str(sheet.cell(row=rows, column=cols).value)
        memba[attribute]=data
        attributeCount+= 1
    memCount+= 1



json.dump(membas, open("result.json", "w"))