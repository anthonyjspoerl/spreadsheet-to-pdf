import os
import re
import win32com.client as win32
from tkinter import *
import tkinter.simpledialog as simpledialog

#----------------------------------------------------------

def test():
    def printContents():
        print ( textwidget1.get(), textwidget2.get() ) 

    window = Tk()
    #frame1 = Frame(window)
    label1 = Label(window, text = "Label 1: ")
    label1.grid(row = 1, column = 0, pady = 10)
    label2 = Label(window, text = "Label 2: ")
    label2.grid(row = 2, column = 0, pady = 10)
    label3 = Label(window, text = "Label 3: ")
    label3.grid(row = 3, column = 0, pady = 10)
    label4 = Label(window, text = "Label 4: ")
    label4.grid(row = 4, column = 0, pady = 10)
    label5 = Label(window, text = "Label 5: ")
    label5.grid(row = 5, column = 0, pady = 10)
    textwidget1 = Entry(window, width = 20)
    textwidget1.grid(row = 1, column = 1)
    textwidget2 = Entry(window, width = 20)
    textwidget2.grid(row = 2, column = 1)
    textwidget3 = Entry(window, width = 20)
    textwidget3.grid(row = 3, column = 1)
    textwidget4 = Entry(window, width = 20)
    textwidget4.grid(row = 4, column = 1)
    textwidget5 = Entry(window, width = 20)
    textwidget5.grid(row = 5, column = 1)
    buttonwidget = Button(window, text = "Submit", command = printContents)
    buttonwidget.grid(row = 6, column = 2, sticky = S, pady = 10)
    #checkbutton = Checkbutton(window, text = "Checkbox", selectcolor = 'pink')
    #checkbutton.grid(row = 2, column = 2
    exitbuttonwidget = Button(window, text = "Cancel", command = window.quit, bg = "red")
    exitbuttonwidget.grid(row = 6, column = 3, sticky = E, columnspan = 2, pady = 10, padx = 10)
    window.mainloop()


if __name__ == "__main__":
    test()
