# -*- coding: utf-8 -*-

import xlrd
import xlwt3 as xlwt
import locale

#import sys, traceback

class XlsProcessor:
    
    def __init__(self):
        self.filepath = ""
        self.message = "Processing"
        self.show_rus_warning = False
        self.show_eng_warning = False
        self.everything_ok = False
        self.all_codes_present = True
    
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
            self.message = "Something went wrong\n" + str(e)
            print(e)
        finally:
            self.codes_trimmed = [code for code in self.codes if code != ""] 
            if len(self.codes_trimmed) < len(self.est):
                self.all_codes_present = False
            return True

    def process_xls(self):
        self.everything_ok = True
        
        try:
            self.message = "Finding codes for english words"
            eng_with_codes = self.make_sortable_array(self.add_codes(self.eng))
            self.message = "Finding codes for russian words"
            rus_with_codes = self.make_sortable_array(self.add_codes(self.rus))
            
            self.message = "Sorting lists"

            eng_sorted = sorted(eng_with_codes, key= lambda k: locale.strxfrm(k[0]))
            rus_sorted = sorted(rus_with_codes, key= lambda k: locale.strxfrm(k[0]))
            
            wb = xlwt.Workbook()
            ws = wb.add_sheet('Translations codes')
            
            self.message = "Writing document"
            
            for i in range(len(eng_sorted)):
                ws.write(i, 0, eng_sorted[i][0] + ' ' + eng_sorted[i][1])
            for j in range(len(rus_sorted)):
                ws.write(j, 1, rus_sorted[j][0] + ' ' + rus_sorted[j][1])
            
            self.set_warnings(eng_sorted[-1][0], rus_sorted[-1][0])
            
            self.save_file(wb)
            
        except Exception as e:
            print(e)
            self.everything_ok = False
            self.message = "Error: " + str(e)
            #traceback.print_exc(file=sys.stdout)
            
    def set_warnings(self, eng, rus):
        if ord(eng[0]) > 500:
            self.show_eng_warning = True
        if ord(rus[0]) > 1500:
            self.show_rus_warning = True
        
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
            except IOError:
                print("File doesn't exist")
                file_exists = False
            except Exception as e:
                print(e)
                self.message = "Something went wrong\n" + str(e)
                file_exists = False
            finally:
                count += 1
        try:
            wb.save(path+filename_modification+filename)
        except Exception as e:
            self.message = "Something went wrong\n" + str(e)
            print(e)
        self.message = "Finished. " + "Saved file\n" + str(filename_modification+filename)

    def split_rows(self, array, splitter = ', '):
        new_array = []
        for row in array:
            new_array.append(row.split(splitter))
        return new_array
    
    def add_codes(self, array):
        """
        takes in array with translations like
        array = ['individual, item, object', 'integrate', 'complex', 'curve, line']
        and returns new array that has words attached codes in arrays like
        new_array = [['individual', 'T5', 'A8'], ['item', 'F4'], ['object', 'D6'], 
        ['integrate', 'G98'], ['complex', 'C45', 'C72', 'C80'],
        ['curve', 'K100'], ['line', 'J50']]
        """
        new_array = []
        splitted_array = self.split_rows(array)
        
        for i in range(len(splitted_array)):
            for word in splitted_array[i]:
                word_ = word
                if word_.startswith('('):
                    w_ = word_.split(')')
                    word_ = w_[1]
                    if word_.startswith(' '):
                        word_ = word_[1:]
                new_row = [word_]
                for k in range(len(splitted_array)):
                    for w in splitted_array[k]:
                        if word == w:
                            new_row.append(self.codes[k])
                new_array.append(new_row)
        return new_array

    def make_sortable_array(self, array):
        """
        takes in array that has words attached codes in arrays like
        array = [['individual', 'T5', 'A8'], ['item', 'F4'], 
        ['integrate', 'G98'], ['complex', 'C45', 'C72', 'C80']] and 
        returns new array that has words as one string and all codes as another string
        like [['individual', 'T5, A8'], ['item', 'F4'], 
        ['integrate', 'G98'], ['complex', 'C45, C72, C80']]
        """
        array_codes_strings = []
        for row in array:
            row_string = row[0] + '*@*' + row[1]
            if len(row) > 2:
                for i in range(2, len(row)):
                    row_string += ', ' + row[i]
            array_codes_strings.append(row_string)
        array_codes_strings = list(set(array_codes_strings))
        return self.split_rows(array_codes_strings, '*@*')