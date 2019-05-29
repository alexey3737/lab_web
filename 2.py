# -*- coding: utf-8 -*-
import xlrd
import csv
import datetime as dt
import matplotlib.pyplot as plt
import numpy as np

# Open the file
wb = xlrd.open_workbook('Analytics.xlsx')

# Get the list of the sheets name
sheet_list = wb.sheet_names()
print(sheet_list)

# Select one sheet and get its size
s = wb.sheet_by_name(sheet_list[0])
print(s.nrows, s.ncols)

x1,y1 = [],[]
x2,y2 = [],[]
x3,y3 = [],[]
i=0
first=1

#Считавыние данных с файла и запись в массив
while (i < s.nrows - 4):
    number1 = s.cell(first,0).value
    value1 = s.cell(first, 3).value
    x1.append(number1)
    y1.append(value1)
    number2 = s.cell(first+1, 0).value
    value2 = s.cell(first+1, 3).value
    x2.append(number2)
    y2.append(value2)
    number3 = s.cell(first+2, 0).value
    value3 = s.cell(first+2, 3).value
    x3.append(number3)
    y3.append(value3)
    i=i+3
    first = first + 3

#Вывод графиков с данными
plt.subplot(211)
plt.plot(x1,y1)
plt.plot(x2,y2)
plt.plot(x3,y3)
plt.legend(['Все пользователи','Новые пользователи','Трафик с мобильных телефонов'])
plt.show()


#Нормализация данных
normal = y1
normal -= np.min(normal, axis=0)
normal /= np.ptp(normal, axis=0)
#print(normal)
norma2 = y2
norma2 -= np.min(norma2, axis=0)
norma2 /= np.ptp(norma2, axis=0)
#print(norma2)
norma3 = y3
norma3 -= np.min(norma3, axis=0)
norma3 /= np.ptp(norma3, axis=0)
#print(norma3)


#Вывод нормализированого графика
plt.subplot(211)
plt.plot(x1,normal)
plt.plot(x2,norma2)
plt.plot(x3,norma3)
plt.legend(['Все пользователи','Новые пользователи','Трафик с мобильных телефонов'])
plt.show()

korr1 = []
zero = 0
one = 1
i=0
korelac1 = normal
korelac2 = norma2
korelac3 = norma3

#Прывышение первого критерия относительно двух остальных
result_1_2 = normal - norma2
result_1_2 = np.absolute(result_1_2)
aver1_2 = np.average(result_1_2)
result1_2 = np.where(result_1_2 > aver1_2, result_1_2, 0)
result_1_3 = normal - norma3
result_1_3 = np.absolute(result_1_3)
aver1_3 = np.average(result_1_3)
result1_3 = np.where(result_1_3 > aver1_3, result_1_2, 0)

#Вывод нормализированого графика по первому
plt.subplot(111)
plt.plot(x2,result1_2)
plt.plot(x3,result1_3)
plt.legend(['Новые пользователи','Трафик с мобильных телефонов'])
plt.show()

#Прывышение первого критерия относительно двух остальных
result_2_1 = norma2 - normal
result_2_1 = np.absolute(result_2_1)
aver2_1 = np.average(result_2_1)
result_2_1 = np.where(result_2_1 > aver2_1, result_2_1, 0)
result_2_3 = norma2 - norma3
result_2_3 = np.absolute(result_2_3)
aver2_3 = np.average(result_2_3)
result_2_3 = np.where(result_2_3 > aver2_3, result_2_3, 0)

#Вывод нормализированого графика по первому
plt.subplot(111)
plt.plot(x2,result_2_1)
plt.plot(x3,result_2_3)
plt.legend(['Все пользователи','Трафик с мобильных телефонов'])
plt.show()

#Прывышение первого критерия относительно двух остальных
result_3_1 = norma3 - normal
result_3_1 = np.absolute(result_3_1)
aver3_1 = np.average(result_3_1)
result_3_1 = np.where(result_3_1 > aver3_1, result_3_1, 0)
result_3_2 = norma3 - norma2
result_3_2 = np.absolute(result_3_2)
aver3_2 = np.average(result_3_2)
result_3_2 = np.where(result_3_2 > aver3_2, result_3_2, 0)

#Вывод нормализированого графика по первому
plt.subplot(111)
plt.plot(x2,result_3_1)
plt.plot(x3,result_3_2)
plt.legend(['Все пользователи','Новые пользователи'])
plt.show()

#Доверенный интервал 1 графика
aver1_2 = np.average(normal)
normal = np.where(normal > aver1_2*2, normal, 0)
print(aver1_2)#0.28
print(normal)
plt.subplot(111)
plt.plot(x1,normal)
plt.legend(['Все пользователи'])
plt.show()

#Доверенный интервал 2 графика
aver1_2 = np.average(norma2)
norma2 = np.where(norma2 > aver1_2*2, norma2, 0)
print(aver1_2)#0.27
print(norma2)
plt.subplot(111)
plt.plot(x2,norma2)
plt.legend(['Новые пользователи'])
plt.show()

#Доверенный интервал 3 графика
aver1_2 = np.average(norma3)
norma3 = np.where(norma3 > aver1_2*2, norma3, 0)
print(aver1_2) #0.3
print(norma3)
plt.subplot(111)
plt.plot(x3,norma3)
plt.legend(['Трафик с мобильных телефонов'])
plt.show()