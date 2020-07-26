#!/usr/bin/python3

import openpyxl 
out=openpyxl.load_workbook('C://Users//Pps//Desktop//PROJECT Mod. M.D//value.xlsx',data_only=True)
sheet=out.active

d2=sheet['D2']
print(d2.value)

d3=sheet['D3']
print(d3.value)

d4=sheet['D4']
print(d4.value)

d5=sheet['D5']
print(d5.value)

d6=sheet['D6']
print(d6.value)

d7=sheet['D7']
print(d7.value)

d8=sheet['D8']
print(d8.value)

d9=sheet['D9']
print(d9.value)

d10=sheet['D10']
print(d10.value)

d11=sheet['D11']
print(d11.value)

d12=sheet['D12']
print(d12.value)

d13=sheet['D13']
print(d13.value)

d14=sheet['D14']
print(d14.value)

d15=sheet['D15']
print(d15.value)

d16=sheet['D16']
print(d16.value)

