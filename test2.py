import tkinter
from tkinter import OptionMenu, Tk, Variable, mainloop, TOP, ttk
from tkinter.ttk import Button
from typing import ValuesView
import openpyxl
import os
import tkinter as tk
from ctypes import pythonapi
from sys import setswitchinterval
import pyautogui, time
from tkinter import ttk
import datetime
from datetime import datetime
from tkinter import *
def Display_Text():
    print("This is from the test2 file")

pyautogui.FAILSAFE = True


def create_GUI():
    root = tkinter.Tk()
    root.geometry('450x250')
    root.title("Automated Form Filler")
    root.config(bg='#F2B90C')
    return root
root = create_GUI()

def create_left_frame(root):
    leftSideFrameVariable = tkinter.Frame(root)
    leftSideFrameVariable.grid(row = 0, column = 0)
    # leftSideFrameVariable.config(bg='#F2B90C')
    return leftSideFrameVariable
leftSideFrameVariable = create_left_frame(root)

def create_right_frame(root):
    rightSideFrameVariable = tkinter.Frame(root)
    rightSideFrameVariable.grid(row = 0, column = 1)
    scrollbarVariable = tkinter.Scrollbar(
    rightSideFrameVariable,
    orient = tkinter.VERTICAL)
    listBoxVariable = tkinter.Listbox(
    rightSideFrameVariable,
    width = 25,
    yscrollcommand = scrollbarVariable.set,
    selectmode = tkinter.EXTENDED)
    scrollbarVariable.config(command = listBoxVariable.yview)
    scrollbarVariable.pack(side = tkinter.RIGHT, fill = tkinter.Y)
    listBoxVariable.pack()
    return listBoxVariable
# create_right_frame(root)
listBoxVariable = create_right_frame(root)

buttonVariable = tkinter.IntVar()
optionVar = tkinter.StringVar()
forms = ('Form 910 (AB - TSgt EPR)', 'Form 911 (MSgt - SMSgt EPR)', 'Form 4 (Reenlistment)')
ranks = ['AB', 'A1C', 'SrA', 'SSgt', 'TSgt', 'MSgt', 'SMSgt']
nameList = []
rankList = []


def determine_program_to_run():
    # pass
    optionMenuValue = optionVar.get()
    if optionMenuValue == 'Form 910 (AB - TSgt EPR)':
        programToStart = 'AF_form_910_v3.py'
    elif optionMenuValue == 'Form 911 (MSgt - SMSgt EPR)':
        programToStart = 'AF_form_911_v2.py'
    elif optionMenuValue == 'Form 4 (Reenlistment)':
        programToStart = 'Form_4_reenlistment_v1.py'
    else:
        # error_message_popup("You must select a form")
        # textDisplayed = 'You must select a form'
        # processingLabel.configure(text = textDisplayed)
        programToStart = ''
    # print(programToStart)
    return programToStart

formOptionMenu = ttk.OptionMenu(
    leftSideFrameVariable, 
    optionVar, 
    'Choose form', 
    *forms)
    # ,
    # command = print_value_of_dropdown_menu) 
buttonLabel = tkinter.Label(
    leftSideFrameVariable,
    text = 'Choose:')
