import tkinter
from tkinter import Variable, ttk
from typing import Literal


root = tkinter.Tk()

buttonVariable = tkinter.IntVar()

firstOptions = ttk.Menubutton(
    root,
    text = "choose",
    # variable = form910,
    # value = 1
)

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

button3 = tkinter.Radiobutton(
    root,
    text = "Form 4 (Reenlistment)",
    variable=buttonVariable,
    value=3
)

runButton = ttk.Button(
    root,
    text = "OK",
    command=lambda: runProgram(processingLabel)
)

runButton2 = tkinter.Button(
    root,
    text = "OK",
    command=lambda: runProgram(processingLabel)
)

processingLabel = tkinter.Label(
    root,
    text=""
)
firstOptions.grid(row=0, column=0)

# button1.grid(row=0, column=0)
button2.grid(row=1, column=0)
button3.grid(row=2, column=0)
runButton.grid(row=3, column=0, padx=300, pady=100)
runButton2.grid(row=4, column=0, padx=100, pady=30)
processingLabel.grid(row=0, column=2)

def runProgram(processingLabel):
    if buttonVariable.get() == 1:
        textDisplayed = "Processing all items"
    else:
        textDisplayed = "Processing selected items"
    processingLabel.configure(text = textDisplayed)

root.mainloop()     #this keeps the window on top