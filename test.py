import os
import re
import win32com.client as win32
from tkinter import *

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

def excelToWord(invoiceNum):
    try:
        setup()
        spreadsheet = openExcel()
        openWordTemplate(spreadsheet, 'Tribals.docx', invoiceNum)
        saveDocs('test')
        cleanup()
    except:
        print('Encountered an error: %s', sys.exc_info()[0])
        cleanup()

def setup():
    global word, excel
    word = win32.gencache.EnsureDispatch('Word.Application')
    excel = win32.gencache.EnsureDispatch('Excel.Application')

def saveDocs(filename):
    doc.SaveAs( os.getcwd() + '/' + filename )
    doc.ExportAsFixedFormat(os.getcwd() + '/' + filename, COM_CONSTANTS.wdExportFormatPDF)

def cleanup():
    ss.Close(False)
    excel.Application.Quit()
    doc.Close(False)
    word.Application.Quit()

def openExcel():
    global excel, ss
    ss = excel.Workbooks.Open( os.getcwd() + '/test_sheet')
    sh = ss.ActiveSheet

    excel.Visible = False

    return sh

def openWordTemplate(spreadsheet, templateName, invoiceNum):
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

    selection.Find.Execute('_invoice_num_')
    selection.Text = invoiceNum
    selection.WholeStory()
    selection.Find.Execute('_amount_')
    selection.Text = fillWithWhitespace(str(sum), len(selection.Text))

def fillWithWhitespace(str, expectedSize):
    #TODO This will need to use the largest amount (num of digits) as expected
    difference = expectedSize - len(str)
    if(difference <= 0):
        return str
    else:
        return (' ' * difference) + str

def getInputs():
    def submit():
        excelToWord( invoiceEntry.get() )
        window.quit()

    window = Tk()
    invoiceLabel = Label(window, text = 'Invoice #:')
    invoiceLabel.pack()
    invoiceEntry = Entry(window, width = 20)
    invoiceEntry.pack()
    submitButton = Button(window, text = "Submit", command = submit)
    submitButton.pack()
    exitButtonWidget = Button(window, text = "Exit", command = window.quit, bg = "red")
    exitButtonWidget.pack()
    window.mainloop()


if __name__ == "__main__":
    getInputs()