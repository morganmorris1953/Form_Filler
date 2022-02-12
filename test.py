import tkinter as tk
from tkinter import font as tkFont
import tkinter
from tkinter import OptionMenu, Tk, Variable, mainloop, TOP, ttk
from tkinter.ttk import Button
from typing import ValuesView
import openpyxl
import os
import tkinter as tk
from tkinter import font as tkFont
root = tk.Tk()

helv36 = tkFont.Font(family='Helvetica', size=36)
forms = ('Form 910 (AB - TSgt EPR)', 'Form 911 (MSgt - SMSgt EPR)', 'Form 4 (Reenlistment)')
optionVar = tk.StringVar(root, value='Choose form')

# optionVar = tk.StringVar()

formOptionMenu = tk.OptionMenu(root, optionVar, *forms)
formOptionMenu.config(font=helv36) # set the button font
menu = root.nametowidget(formOptionMenu.menuname)
menu.config(font=helv36)  # Set the dropdown menu's font
formOptionMenu.grid(row=0, column=0, sticky='nsew')

root.mainloop()



# helv36 = tkFont.Font(family='Helvetica', size=36)
# optionsList = 'eggs spam toast'.split()
# selectedOption = tk.StringVar(root, value=optionsList[0])

# chooseTest = tk.OptionMenu(root, selectedOption, *optionsList)
# chooseTest.config(font=helv36) # set the button font

# helv20 = tkFont.Font(family='Helvetica', size=20)
# menu = root.nametowidget(chooseTest.menuname)
# menu.config(font=helv20)  # Set the dropdown menu's font
# chooseTest.grid(row=0, column=0, sticky='nsew')

# root.mainloop()