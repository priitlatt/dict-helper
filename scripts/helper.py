# -*- coding: utf-8 -*-

import xlrd
import xlwt3 as xlwt
import locale

locale.setlocale(locale.LC_ALL, 'Estonian_Estonia')

def split_rows(array):
    new_array = []
    for row in array:
        new_array.append(row.split(', '))
    return new_array

def add_codes(array):
    new_array = []
    array2 = split_rows(array)
    
    for i in range(len(array2)):
        for word in array2[i]:
            word_ = word
            if word_.startswith('('):
                w_ = word_.split(')')
                word_ = w_[1]
                if word_.startswith(' '):
                    word_ = word_[1:]
            new_row = [word_]
            for k in range(len(array2)):
                for w in array2[k]:
                    if word == w:
                        new_row.append(codes[k])
            new_array.append(new_row)
    return new_array

def make_string_array(array):
    array_codes_strings = []
    for row in array:
        row_string = row[0] + ' ' + row[1]
        if len(row) > 2:
            for i in range(2, len(row)):
                row_string += ', ' + row[i]
        array_codes_strings.append(row_string)
    return array_codes_strings

filepath = "C:/Users/Priit/Desktop/splchk_for_editing2.xls"
book = xlrd.open_workbook(filepath, encoding_override="cp1252")
sheet = book.sheets()[0]
codes = sheet.col_values(0)
est = sheet.col_values(1)
eng = sheet.col_values(2)
rus = sheet.col_values(3)

eng_codes = make_string_array(add_codes(eng))
rus_codes = make_string_array(add_codes(rus))
eng_codes_sorted = sorted(set(eng_codes), key=locale.strxfrm)
rus_codes_sorted = sorted(set(rus_codes), key=locale.strxfrm)

wb = xlwt.Workbook()
ws = wb.add_sheet('Translations codes')

for i in range(len(eng_codes_sorted)):
    ws.write(i, 0, eng_codes_sorted[i])
for j in range(len(rus_codes_sorted)):
    ws.write(j, 1, rus_codes_sorted[j])

wb.save("C:/Users/Priit/Desktop/test.xls")
print('valmis')