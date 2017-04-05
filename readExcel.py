#-*- coding: utf8 -*-
import xlrd
# import csv
import pandas as pd
import json
import sys

list_id = ['6.92364424042e+12','6.92364426892e+12', '6.92364426891e+12', '6.92364426492e+12', '6.92364426495e+12', '6.92364425148e+12', '6.92364427904e+12', '6.92364427263e+12']
# list_id = ['6923644240424','6923644268916','6923644268909','6923644264918','6923644264949','6923644251475','6923644279035','6923644272630']
json_data = {}
json_data_sku = {}

fname = "D:\mengniu\\1_3.xlsx"
# fname = "aaa.xlsx"
data = xlrd.open_workbook(fname)
table = data.sheets()[0]

for i in range(1, table.nrows):
    row = table.row_values(i)
    # try:
    if str(row[7]) in list_id:
        month = xlrd.xldate_as_tuple(row[2], 0)[1]
        city = row[1].split('/')[1]
        # location = row[1].split('/')
        # city = location[1]
        count = row[9]
        money = float(row[9]) * float(row[11])

        key_city_count = 'count_' + str(month) + '_' + city
        key_city_sku_count = 'count_' + str(month) + '_' + city + '_' + str(list_id.index(str(row[7])))
        key_city_money = 'money_' + str(month) + '_' + city
        key_city_sku_money = 'money_' + str(month) + '_' + city + '_' + str(list_id.index(str(row[7])))

        if json_data. has_key(key_city_count):
            json_data[key_city_count] = int(json_data[key_city_count]) + int(count)
            json_data[key_city_money] = float(json_data[key_city_money]) + float(money)
        else:
            json_data[key_city_count] = count
            json_data[key_city_money] = money

        if json_data_sku. has_key(key_city_sku_count):
            json_data_sku[key_city_sku_count] = int(json_data_sku[key_city_sku_count]) + int(count)
            json_data_sku[key_city_sku_money] = float(json_data_sku[key_city_sku_money]) + float(money)
        else:
            json_data_sku[key_city_sku_count] = count
            json_data_sku[key_city_sku_money] = money
    # except:
    #     print i

with open('D:\mengniu\\result\\1to3_city_dump.json', 'w') as f:
    f.write(json.dumps(json_data))
with open('D:\mengniu\\result\\1to3_sku_dump.json', 'w') as f:
    f.write(json.dumps(json_data_sku))

# 城市柱状图文件二次处理
list_city = list()
list_money = list()
list_count = [0] * int(len(json_data)/2)

for k in json_data:
    keys = k.split('_')
    if keys[1] != '1':
        continue
    if 'money' == keys[0]:
        # list_money.append(json_data[k])
        # list_city.append(keys[2])
        if len(list_money) == 0:
            list_money.append(round(json_data[k], 2))
            list_city.append(keys[2])
        else:
            for money_dump in list_money:
                if json_data[k] < list_money[len(list_money) - 1]:
                    list_money.insert(len(list_money), round(json_data[k], 2))
                    list_city.insert(len(list_city), keys[2])
                    break
                elif json_data[k] >= money_dump:
                    list_city.insert(list_money.index(money_dump), keys[2])
                    list_money.insert(list_money.index(money_dump), round(json_data[k], 2))
                    break

for k in json_data:
    keys = k.split('_')
    if keys[1] != '1':
        continue
    if 'count' == keys[0]:
        city = keys[2]
        index = list_city.index(city)
        list_count[index] = json_data[k]
# print len(list_city),len(list_count),len(list_money)
# print list_city
# print list_count
# print list_money
json_need_city = {'city': list_city, 'count': list_count, 'money': list_money}

with open('D:\mengniu\\result\\1_city.json', 'w') as f:
    f.write(json.dumps(json_need_city))

# SKU饼状图数据二次处理
json_sku = {}
for k in json_data_sku:
    keys = k.split('_')
    if keys[1] != '1':
        continue
    if 'money' == keys[0]:
        city = keys[2]
        if json_sku. has_key(city):
            list_sku = json_sku[city]
            list_sku[int(keys[3])] = round(json_data_sku[k], 2)
            json_sku[city] = list_sku
        else:
            list_sku = [0] * 8
            list_sku[int(keys[3])] = round(json_data_sku[k], 2)
            json_sku[city] = list_sku

with open('D:\mengniu\\result\\1_sku.json', 'w') as f:
        f.write(json.dumps(json_sku))


# 城市柱状图文件二次处理
list_city = list()
list_money = list()
list_count = [0] * int(len(json_data)/2)

for k in json_data:
    keys = k.split('_')
    if keys[1] != '2':
        continue
    if 'money' == keys[0]:
        # list_money.append(json_data[k])
        # list_city.append(keys[2])
        if len(list_money) == 0:
            list_money.append(round(json_data[k], 2))
            list_city.append(keys[2])
        else:
            for money_dump in list_money:
                if json_data[k] < list_money[len(list_money) - 1]:
                    list_money.insert(len(list_money), round(json_data[k], 2))
                    list_city.insert(len(list_city), keys[2])
                    break
                elif json_data[k] >= money_dump:
                    list_city.insert(list_money.index(money_dump), keys[2])
                    list_money.insert(list_money.index(money_dump), round(json_data[k], 2))
                    break

for k in json_data:
    keys = k.split('_')
    if keys[1] != '2':
        continue
    if 'count' == keys[0]:
        city = keys[2]
        index = list_city.index(city)
        list_count[index] = json_data[k]
# print len(list_city),len(list_count),len(list_money)
# print list_city
# print list_count
# print list_money
json_need_city = {'city': list_city, 'count': list_count, 'money': list_money}

with open('D:\mengniu\\result\\2_city.json', 'w') as f:
    f.write(json.dumps(json_need_city))

# SKU饼状图数据二次处理
json_sku = {}
for k in json_data_sku:
    keys = k.split('_')
    if keys[1] != '2':
        continue
    if 'money' == keys[0]:
        city = keys[2]
        if json_sku. has_key(city):
            list_sku = json_sku[city]
            list_sku[int(keys[3])] = round(json_data_sku[k], 2)
            json_sku[city] = list_sku
        else:
            list_sku = [0] * 8
            list_sku[int(keys[3])] = round(json_data_sku[k], 2)
            json_sku[city] = list_sku

with open('D:\mengniu\\result\\2_sku.json', 'w') as f:
        f.write(json.dumps(json_sku))



# 城市柱状图文件二次处理
list_city = list()
list_money = list()
list_count = [0] * int(len(json_data)/2)

for k in json_data:
    keys = k.split('_')
    if keys[1] != '3':
        continue
    if 'money' == keys[0]:
        # list_money.append(json_data[k])
        # list_city.append(keys[2])
        if len(list_money) == 0:
            list_money.append(round(json_data[k], 2))
            list_city.append(keys[2])
        else:
            for money_dump in list_money:
                if json_data[k] < list_money[len(list_money) - 1]:
                    list_money.insert(len(list_money), round(json_data[k], 2))
                    list_city.insert(len(list_city), keys[2])
                    break
                elif json_data[k] >= money_dump:
                    list_city.insert(list_money.index(money_dump), keys[2])
                    list_money.insert(list_money.index(money_dump), round(json_data[k], 2))
                    break

for k in json_data:
    keys = k.split('_')
    if keys[1] != '3':
        continue
    if 'count' == keys[0]:
        city = keys[2]
        index = list_city.index(city)
        list_count[index] = json_data[k]
# print len(list_city),len(list_count),len(list_money)
# print list_city
# print list_count
# print list_money
json_need_city = {'city': list_city, 'count': list_count, 'money': list_money}

with open('D:\mengniu\\result\\3_city.json', 'w') as f:
    f.write(json.dumps(json_need_city))

# SKU饼状图数据二次处理
json_sku = {}
for k in json_data_sku:
    keys = k.split('_')
    if keys[1] != '3':
        continue
    if 'money' == keys[0]:
        city = keys[2]
        if json_sku. has_key(city):
            list_sku = json_sku[city]
            list_sku[int(keys[3])] = round(json_data_sku[k], 2)
            json_sku[city] = list_sku
        else:
            list_sku = [0] * 8
            list_sku[int(keys[3])] = round(json_data_sku[k], 2)
            json_sku[city] = list_sku

with open('D:\mengniu\\result\\3_sku.json', 'w') as f:
        f.write(json.dumps(json_sku))
