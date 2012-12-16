# -*- coding: utf-8 -*-

import xlrd
import xlwt3 as xlwt
import locale

class XlsProcessor:
    
    def __init__(self):
        self.filepath = ""
        self.message = "Processing"
        self.everything_ok = False
    
    def open_book(self, filepath):
        self.filepath = filepath
        try:
            self.book = xlrd.open_workbook(self.filepath, encoding_override="cp1252")
            self.sheet = self.book.sheets()[0]
            self.codes = self.sheet.col_values(0)
            self.est = self.sheet.col_values(1)
            self.eng = self.sheet.col_values(2)
            self.rus = self.sheet.col_values(3)
        except Exception as e:
            print(e)
        finally:
            return True
        
    def process_xls(self):
        self.everything_ok = True
        self.message = "Finding codes"
        
        try:
            eng_codes = self.make_string_array(self.add_codes(self.eng))
            rus_codes = self.make_string_array(self.add_codes(self.rus))
            
            self.message = "Sorting lists"
            eng_codes_sorted = sorted(set(eng_codes), key=locale.strxfrm)
            rus_codes_sorted = sorted(set(rus_codes), key=locale.strxfrm)
            
            wb = xlwt.Workbook()
            ws = wb.add_sheet('Translations codes')
            
            self.message = "Writing document"
            
            for i in range(len(eng_codes_sorted)):
                ws.write(i, 0, eng_codes_sorted[i])
            for j in range(len(rus_codes_sorted)):
                ws.write(j, 1, rus_codes_sorted[j])
            
            self.save_file(wb)
            
        except Exception as e:
            print(e)
            self.everything_ok = False
            self.message = "Error: File is not compatible"
        
    def save_file(self, wb):
        file_exists = True
        count = 1
        path_array = self.filepath.split('/')
        filename = path_array[-1]
        if filename.endswith('x'):
            filename = filename[:-1]
        path_array.pop()
        path = '/'.join(path_array) + '/'
        while file_exists:
            try:
                if count == 1:
                    filename_modification = "processed_"
                else:
                    filename_modification = "processed" + str(count) + "_"
                with open(path+filename_modification+filename):
                    pass
            except Exception as e:
                print(e)
                file_exists = False
            finally:
                count += 1
        try:
            wb.save(path+filename_modification+filename)
        except Exception as e:
            print(e)
        self.message = "Finished"

    def split_rows(self, array):
        new_array = []
        for row in array:
            new_array.append(row.split(', '))
        return new_array
    
    def add_codes(self, array):
        new_array = []
        array2 = self.split_rows(array)
        
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
                            new_row.append(self.codes[k])
                new_array.append(new_row)
        return new_array
    
    def make_string_array(self, array):
        array_codes_strings = []
        for row in array:
            row_string = row[0] + ' ' + row[1]
            if len(row) > 2:
                for i in range(2, len(row)):
                    row_string += ', ' + row[i]
            array_codes_strings.append(row_string)
        return array_codes_strings
