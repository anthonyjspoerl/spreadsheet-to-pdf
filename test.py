import os
import re
import win32com.client as win32
from tkinter import *
import tkinter.simpledialog as simpledialog

TEMPLATE_PATH = os.getcwd() + '/templates/'
RANGE = range(3, 8)
ADD_REGEX = re.compile('.*(add).*', re.IGNORECASE)
MULTIPLY_REGEX = re.compile('.*(multiply).*', re.IGNORECASE)

excel = 0
ss = 0
word = 0
doc = 0

# This will contain all enums/module constants from coms bound with EnsureDispatch
COM_CONSTANTS = win32.constants

#----------------------------------------------------------------------

def excelToWord():
    setup()
    spreadsheet = openExcel()
    openWordTemplate(spreadsheet, 'Tribals.docx')
    cleanup()

def setup():
    global word, excel
    word = win32.gencache.EnsureDispatch('Word.Application')
    excel = win32.gencache.EnsureDispatch('Excel.Application')

def cleanup():
    ss.Close(False)
    excel.Application.Quit()
    doc.SaveAs( os.getcwd() + '/test.docx' )
    doc.ExportAsFixedFormat(os.getcwd() + '/test.pdf', COM_CONSTANTS.wdExportFormatPDF)
    doc.Close(False)
    word.Application.Quit()

def openExcel():
    global excel, ss
    ss = excel.Workbooks.Open( os.getcwd() + '/test_sheet')
    sh = ss.ActiveSheet

    excel.Visible = False

    return sh

def openWordTemplate(spreadsheet, templateName):
    global word, doc
    doc = word.Documents.Open(TEMPLATE_PATH + templateName)
    selection = word.Selection
    word.Visible = False
 
    rng = doc.Range(0,0)

    index = 2
    val = spreadsheet.Cells(index,1).Value

    sum = 0
    product = 1
    while val:
        numVal = spreadsheet.Cells(index,2).Value
        if( ADD_REGEX.match(val) != None ):
            sum += numVal
        elif( MULTIPLY_REGEX.match(val) != None ):
            product = product * numVal

        index += 1
        val = spreadsheet.Cells(index,1).Value

    selection.Find.Execute('%amount%')
    selection.Text = fillWithWhitespace(str(sum), len(selection.Text))

def fillWithWhitespace(str, expectedSize):
    #TODO This will need to use the largest amount (num of digits) as expected
    difference = expectedSize - len(str)
    if(difference <= 0):
        return str
    else:
        return (' ' * difference) + str

def test():
    def printContents():
        print ( textwidget.get() )

    window = Tk()
    textwidget = Entry(window, width = 20)
    textwidget.pack()
    buttonwidget = Button(window, text = "Button", command = printContents)
    buttonwidget.pack()
    exitbuttonwidget = Button(window, text = "Exit", command = window.quit, bg = "red")
    exitbuttonwidget.pack()
    window.mainloop()


if __name__ == "__main__":
    excelToWord()
