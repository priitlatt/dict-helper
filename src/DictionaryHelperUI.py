# -*- coding: utf-8 -*-

import tkinter as tk
import tkinter.filedialog as FileDialog

class DictHelperUI:
    
    def __init__(self, parent):
        BTN_WIDTH = 12
        PARENT_HEIGHT = 220 
        PARENT_WIDTH = 250
        
        self.my_parent = parent
        self.my_parent.title("Dictionary Helper")
        self.my_parent.geometry(str(PARENT_HEIGHT)+"x"+str(PARENT_WIDTH))
        self.my_parent.resizable(width=False, height=False)
        self.my_container = tk.Frame(parent)
        self.my_container.pack(expand=tk.YES, fill=tk.BOTH)
        
        self.buttons_frame = tk.Frame(self.my_container)
        self.buttons_frame.pack(side=tk.TOP, expand=tk.NO, fill=tk.Y, ipadx=5, ipady=10)
        
        self.text_frame = tk.Frame(self.my_container)
        self.text_frame.pack(side=tk.BOTTOM, expand=tk.YES, fill=tk.BOTH, ipadx=5, ipady=10)
        
        self.select_file_btn = tk.Button(self.buttons_frame, text="Select file", width=BTN_WIDTH)
        self.select_file_btn.pack(side=tk.LEFT)
        self.select_file_btn.bind("<Return>", self.select_file_btn_click)
        self.select_file_btn.bind("<ButtonRelease-1>", self.select_file_btn_click)
        
        self.close_btn = tk.Button(self.buttons_frame, text="Close", width=BTN_WIDTH)
        self.close_btn.pack(side=tk.RIGHT)
        self.close_btn.bind("<Button-1>", self.close_btn_click)
        self.close_btn.bind("<Return>", self.close_btn_click)
        
        self.my_message="Please select a file"
        self.label = tk.Label(self.text_frame, text=self.my_message, justify=tk.LEFT)#.pack(side=BOTTOM, anchor=W)
        self.label.pack()
        
    def select_file_btn_click(self, event):
        #filename = self.filePicker()
        #self.label.configure(text=filename)
        self.readfile()
        
    def close_btn_click(self, event):
        self.my_parent.destroy()
        
    def file_picker(self):
        filename = FileDialog.askopenfilename()
        # print(filename)
        return filename
        
    def readfile(self):
        filename = self.file_picker()
        
        if not filename.endswith('.txt') and not filename.endswith('.csv') and not filename.endswith('.py'):
            if filename != "":
                lbl_text = "File '" + str(filename) + "' is not compatible, please select 'txt' or 'csv' file!"
                self.label.configure(text=lbl_text)
            
        else:
            with open(filename) as f:
                file_text = ""
                for line in f:
                    #print(line)
                    file_text += line + "\n"
                self.label.configure(text=file_text)
