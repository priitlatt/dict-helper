# -*- coding: utf-8 -*-

import tkinter as tk
import DictionaryHelperUI as dhUI
import locale

locale.setlocale(locale.LC_ALL, 'Estonian_Estonia')

root_window = tk.Tk()
dict_helper_UI = dhUI.DictHelperUI(root_window)

root_window.mainloop()