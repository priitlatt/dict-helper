# -*- coding: utf-8 -*-

import tkinter as tk
import tkinter.filedialog as FileDialog
import XlsProcessor as XlsP
import threading
import time

class DictHelperUI:
    
    def __init__(self, parent):
        
        BTN_WIDTH = 12
        PARENT_HEIGHT = 100 
        PARENT_WIDTH = 220
        RESIZABLE = False
        PROGRAM_TITLE = "Dictionary Helper"
        
        self.counter = 0
        self.thread_runnable = True
        
        self.my_parent = parent
        self.my_parent.title(PROGRAM_TITLE)
        self.my_parent.geometry(str(PARENT_WIDTH)+"x"+str(PARENT_HEIGHT))
        self.my_parent.resizable(width=RESIZABLE, height=RESIZABLE)
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
        self.close_btn.bind("<ButtonRelease-1>", self.close_btn_click)
        self.close_btn.bind("<Return>", self.close_btn_click)
        
        self.lbl_text=tk.StringVar()
        self.label = tk.Label(self.text_frame, textvariable=self.lbl_text,
                              wraplength=PARENT_WIDTH-40, anchor=tk.W, justify=tk.CENTER)
        self.lbl_text.set("Please select a file")
        self.label.pack()
        
    def select_file_btn_click(self, event):
        self.thread_runnable = False
        self.readfile()
        
    def close_btn_click(self, event):
        self.my_parent.destroy()
        
    def file_picker(self):
        return FileDialog.askopenfilename()
        
    def readfile(self):
        filepath = self.file_picker()
        
        if not filepath.endswith('.xls') and not filepath.endswith('.xlsx'):
            if filepath != "":
                filename = filepath.split('/')[-1]
                if len(filename) > 20:
                    filename = filename[:20] + "..."
                self.lbl_text.set("File '" + str(filename) + "' is not compatible.\nPlease select 'xls' or 'csv' file!")
        else:
            self.start_process(filepath)
            
    def start_process(self, filepath):
        start = time.time()
        
        xls_processor = XlsP.XlsProcessor()
        self.thread_runnable = True
        self.lbl_text.set("Opening file")
        xls_processor.open_book(filepath)
        t = threading.Thread(target = xls_processor.process_xls)
        t.start()
        while t.is_alive():
            if self.thread_runnable:
                self.counter += 1
                self.lbl_text.set(xls_processor.message + (self.counter % 4)*".")
                self.my_parent.update()
                time.sleep(0.5)
            else:
                t._stop()
        t.join()
        stop = time.time()
        if xls_processor.everything_ok and self.thread_runnable:
            self.lbl_text.set(xls_processor.message + 
                              '\nTotal time ' + str(int(stop-start)) + " seconds")
        else:
            self.lbl_text.set("Something went wrong")
