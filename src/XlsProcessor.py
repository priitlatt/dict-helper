# -*- coding: utf-8 -*-

import xlrd
import xlwt3 as xlwt
import locale

import sys, traceback

class XlsProcessor:
    
    def __init__(self):
        self.filepath = ""
        self.message = "Processing"
        self.show_rus_warning = False
        self.show_eng_warning = False
        self.everything_ok = False
        self.all_codes_present = True
        
        self.codes = None
        self.est = None
        self.eng = None
        self.rus = None
    
    def open_book(self, filepath):
        self.filepath = filepath
        try:
            book = xlrd.open_workbook(self.filepath, encoding_override="cp1252")
            sheet = book.sheets()[0]
            self.codes = sheet.col_values(0)
            self.est = sheet.col_values(1)
            self.eng = sheet.col_values(2)
            self.rus = sheet.col_values(3)
        except Exception as e:
            self.message = "Something went wrong\n" + str(e)
            print(e)
        finally:
            self.codes_trimmed = [code for code in self.codes if code != ""] 
            if len(self.codes_trimmed) < len(self.est):
                self.all_codes_present = False
            return True

    def fix_original(self):
        """
        Util method that calls remove_empty_lines and remove_whitespaces
        functions on itself to get rid of such entries in the table that doesn' 
        have estonian keywords and removes whitespaces from left and right of 
        all rows and columns in table.
        Finally this method saves the fixed table into new file.
        """
        
        self.remove_empty_lines()
        self.codes = self.remove_whitespaces(self.codes)
        self.est = self.remove_whitespaces(self.est)
        self.eng = self.remove_whitespaces(self.eng)
        self.rus = self.remove_whitespaces(self.rus)
        
        self.save_fixed_original()

    def remove_empty_lines(self):
        """
        Util method that actually removes 
        lines that are not needed in the table.
        """
        new_codes = []
        new_est = []
        new_eng = []
        new_rus = []
        for i in range(len(self.est)):
            if str(self.est[i]).lstrip().rstrip() != "":
                new_codes.append(self.codes[i])
                new_est.append(self.est[i])
                new_eng.append(self.eng[i])
                new_rus.append(self.rus[i])
        self.codes = new_codes
        self.est = new_est
        self.eng = new_eng
        self.rus = new_rus
        
    def remove_whitespaces(self, original_array):
        """
        Method that takes in table column as staing array and then 
        removes whitespaces from left and right of the column elements
        and then returns the trimmed array.
        Original array is not affected. 
        """
        words = []
        for word in original_array:
            word = str(word).lstrip().rstrip()
            if len(word) > 0 and word[-1] == ',':
                word = word[:-1]
            words.append(word)
        return words
    
    def save_fixed_original(self):
        """
        Method that prepares fixed original table to be saved into new file
        as xlwt table and then calls save_file to actually save the table.
        """
        wb = xlwt.Workbook()
        ws = wb.add_sheet('Page1')
        for i in range(len(self.codes)):
            ws.write(i, 0, self.codes[i])
            ws.write(i, 1, self.est[i])
            ws.write(i, 2, self.eng[i])
            ws.write(i, 3, self.rus[i])
        self.save_file(wb, original=True)
        
    def process_xls(self):
        self.everything_ok = True
        
        try:
            self.message = "Verifying and fixing original file"
            self.fix_original()
            
            self.message = "Finding codes for english words"
            eng_with_codes = self.make_sortable_array(self.add_codes(self.eng))
            self.message = "Finding codes for russian words"
            rus_with_codes = self.make_sortable_array(self.add_codes(self.rus))
            
            self.message = "Sorting lists"

            eng_sorted = sorted(eng_with_codes, key= lambda k: locale.strxfrm(k[0]))
            rus_sorted = sorted(rus_with_codes, key= lambda k: locale.strxfrm(k[0]))
            
            self.message = "Adding letters"
            eng_sorted_letters = self.add_letters(eng_sorted)
            rus_sorted_letters = self.add_letters(rus_sorted)
            
            wb = xlwt.Workbook()
            ws = wb.add_sheet('Translations codes')
            
            self.message = "Writing document"
            
            for i in range(len(eng_sorted_letters)):
                ws.write(i, 0, eng_sorted_letters[i])
            for j in range(len(rus_sorted_letters)):
                ws.write(j, 1, rus_sorted_letters[j])
            
            self.set_warnings(eng_sorted_letters[-1], rus_sorted_letters[-1])
            
            self.save_file(wb)
            
        except Exception as e:
            print(e)
            self.everything_ok = False
            self.message = "Error: " + str(e)
            traceback.print_exc(file=sys.stdout)
            
    def add_letters(self, array):
        """
        Takes in a sorted array of strings array and returns new array where
        letter is inserted between words with different first letters and
        individual subarrays are merged into one string.
        input: [..., [s1, k1], [s2, k2], ..., [sn, kn], [t1, kn+1], [t2, kn+2], ... ]
        output: [..., 's1 k1', 's2 k2', ..., 'sn, kn', '', 'T', '', 't1, kn+1', t2, kn+2', ... ]
        """
        new_array = []
        letter = str(array[0][0])[0].upper()
        new_array.append(letter)
        new_array.append('')
        for row in array:
            if str(row[0])[0].upper() != letter:
                letter = str(row[0])[0].upper()
                new_array.append('')
                new_array.append(letter)
                new_array.append('')
            new_array.append(row[0] + ' ' + row[1])
        return new_array
    
    def set_warnings(self, eng, rus):
        if ord(eng[0]) > 500:
            self.show_eng_warning = True
        if ord(rus[0]) > 1500:
            self.show_rus_warning = True
        
    def save_file(self, wb, original=False):
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
                    if original:
                        filename_modification = "fixed_"
                    else:
                        filename_modification = "processed_"
                else:
                    if original:
                        filename_modification = "fixed" + str(count) + "_"
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
        if not original:
            self.message = "Finished. " + "Saved file\n" + str(filename_modification+filename)

    def split_rows(self, array, splitter = ', '):
        new_array = []
        for row in array:
            new_array.append(row.split(splitter))
        return new_array
    
    def add_codes(self, array):
        """
        Takes in array with translations like
        array = ['individual, item, object', 'integrate', 'complex', 'curve, line']
        and returns new array that has words attached codes in arrays like
        new_array = [['individual', 'T5', 'A8'], ['item', 'F4'], ['object', 'D6'], 
        ['integrate', 'G98'], ['complex', 'C45', 'C72', 'C80'],
        ['curve', 'K100'], ['line', 'J50']].
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
        Takes in array that has words attached codes in arrays like
        array = [['individual', 'T5', 'A8'], ['item', 'F4'], 
        ['integrate', 'G98'], ['complex', 'C45', 'C72', 'C80']] and 
        returns new array that has words as one string and all codes as another string
        like [['individual', 'T5, A8'], ['item', 'F4'], 
        ['integrate', 'G98'], ['complex', 'C45, C72, C80']].
        """
        array_codes_strings = []
        for row in array:
            if str(row[0]) != '':
                row_string = row[0].lstrip().rstrip() + '*@*' + row[1]
                if len(row) > 2:
                    for i in range(2, len(row)):
                        row_string += ', ' + row[i]
                array_codes_strings.append(row_string)
        array_codes_strings = list(set(array_codes_strings))
        return self.split_rows(array_codes_strings, '*@*')
