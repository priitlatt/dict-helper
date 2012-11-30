# -*- coding: utf-8 -*-

from tkinter import *
import tkinter.filedialog as FileDialog

class DictHelperUI:
    def __init__(self, parent):
        btn_width = 12
        
        self.myParent = parent
        self.myParent.title("Dictionary Helper")
        self.myParent.geometry("220x250")
        self.myParent.resizable(width=False, height=False)
        self.myContainer = Frame(parent)
        self.myContainer.pack(expand=YES, fill=BOTH)

        self.buttons_frame = Frame(self.myContainer)
        self.buttons_frame.pack(side=TOP, expand=NO, fill=Y, ipadx=5, ipady=5)
        
        self.text_frame = Frame(self.myContainer)
        self.text_frame.pack(side=BOTTOM, expand=YES, fill=BOTH, ipadx=5, ipady=5)
        
        self.selectFileBtn = Button(self.buttons_frame, text="Vali fail", width=btn_width)
        self.selectFileBtn.pack(side=LEFT)
        self.selectFileBtn.bind("<Button-1>", self.selectFileBtnClick)
        self.selectFileBtn.bind("<Return>", self.selectFileBtnClick)
        
        self.cancelBtn = Button(self.buttons_frame, text="Cancel", width=btn_width)
        self.cancelBtn.pack(side=RIGHT)
        self.cancelBtn.bind("<Button-1>", self.cancelBtnClick)
        self.cancelBtn.bind("<Return>", self.cancelBtnClick)  ### (2)
        
        self.myMessage="Palun valige fail"
        self.label = Label(self.text_frame, text=self.myMessage, justify=LEFT)#.pack(side=BOTTOM, anchor=W)
        self.label.pack()

    def selectFileBtnClick(self, event):
        #filename = self.filePicker()
        #self.label.configure(text=filename)
        self.readfile()
        
    def cancelBtnClick(self, event):
        self.myParent.destroy()
        
    def filePicker(self):
        fileName = FileDialog.askopenfilename()
        # print(fileName)
        return fileName

    def readfile(self):
        fname = self.filePicker()
        
        if not fname.endswith('.txt') and not fname.endswith('.csv') and not fname.endswith('.py'):
            lbl_text = "Fail '" + str(fname) + "' ei ole sobiv, palun valige 'txt' või 'csv' fail!"
            self.label.configure(text=lbl_text)
        
        else:
            with open(fname) as f:
                file_text = ""
                for line in f:
                    #print(line)
                    file_text += line + "\n"
                self.label.configure(text=file_text)
        
root = Tk()

myapp = DictHelperUI(root)
root.mainloop()