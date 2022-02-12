import tkinter
import tkinter as tk
from tkinter import ttk
from ctypes import pythonapi
from sys import setswitchinterval
import pyautogui, time
import os
import openpyxl
import datetime
from datetime import datetime
pyautogui.FAILSAFE = True

class App(tk.Tk):
    


    def __init__(self):
        super().__init__()
        self.geometry("500x200")
        self.title('Automated Form Filler')

        # initialize data
        self.forms = ('Form 910 (AB - TSgt EPR)', 'Form 911 (MSgt - SMSgt)', 'Form 4 (Reenlistment)')
        self.menuHeader = ('Choose form')
        # set up variable
        self.option_var = tk.StringVar(self)

        # create widget
        self.create_wigets()

    def runProgram():
        os.system('start Form_4_reenlistment_v1.py')     

    def create_wigets(self):
        root = tkinter.Tk()
        # padding for widgets using the grid layout
        paddings = {'padx': 5, 'pady': 5}

        # label
        label = ttk.Label(self,  text='Select your form to automate:', border = 10)
        label.grid(column=0, row=0, sticky=tk.W, **paddings)

        # option menu
        option_menu = ttk.OptionMenu(
            root,
            self,
            self.option_var,
            self.menuHeader,
            *self.forms,
            command=self.applySelections)

        option_menu.grid(column=1, row=0, sticky=tk.W, padx=10, pady=20)

        # output label
        self.output_label = ttk.Label(self, foreground='red')
        self.output_label.grid(column=1, row=2, sticky=tk.W, **paddings)

        #Buttons
        buttonVariable = tkinter.IntVar()
        button1 = tkinter.Radiobutton(
            root,
            text = "Select by rank",
            variable=buttonVariable,
            value=1
            )


        button2 = tkinter.Radiobutton(
            root,
            text = "Select by name",
            variable=buttonVariable,
            value=2
            )

        runButton = tkinter.Button(
        root,
        text = "OK",
        # command=lambda: runProgram()
        )
        
        # runButton = ttk.Button(
        # # root,
        # text = "OK",
        # command=lambda: runProgram(processingLabel)
        # )

        button1.grid(row=1, column=0, sticky=tk.W, **paddings)
        button2.grid(row=2, column=0, sticky=tk.W, **paddings)
        runButton.grid(row=3, column=0, padx=30, pady=10)

    def applySelections(self, *args):
        
        self.output_label['text'] = f'You selected: {self.option_var.get()}'

    # def runProgram():
    #     os.system('start Form_4_reenlistment_v1.py')     


if __name__ == "__main__":
    app = App()
    app.mainloop()