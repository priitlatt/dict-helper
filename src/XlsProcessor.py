# -*- coding: utf-8 -*-

import xlrd
import time
class XlsProcessor:
    
    def __init__(self):
        self.filepath = ""
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
        
    def do_smth(self):
        try:
            time.sleep(1)
        except Exception as e:
            print(e)
            return
        #self.save_file()
        self.everything_ok = True
        return
        
    def save_file(self):
        file_exists = True
        count = 1
        path_array = self.filepath.split('/')
        filename = path_array[-1]
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
            with open(path+filename_modification+filename, 'w'):
                pass
        except Exception as e:
            print(e)