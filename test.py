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

def excelToWord(invoiceNum, subdivision, referenceNum, mps, location, county, state):
    try:
        setup()
        spreadsheet = openExcel()
        testSum = openWordTemplate(spreadsheet, 'Tribals.docx')
        replaceEntryFields(invoiceNum, subdivision, referenceNum, mps, location, county, state, testSum)
        saveDocs('test')
        cleanup()
    except Exception as e:
        print('Encountered an error: ', e)
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

def openWordTemplate(spreadsheet, templateName):
    global word, doc
    doc = word.Documents.Open(TEMPLATE_PATH + templateName)
    word.Visible = False
 
    rng = doc.Range(0,0)

    index = 2
    val = spreadsheet.Cells(index,1).Value

    testSum = 0
    product = 1
    while val:
        numVal = spreadsheet.Cells(index,2).Value
        if( ADD_REGEX.match(val) != None ):
            testSum += numVal

        index += 1
        val = spreadsheet.Cells(index,1).Value

    return testSum
        

def replaceEntryFields(invoiceNum, subdivision, referenceNum, mps, location, county, state, testSum):
    selection = word.Selection

    selection.Find.Execute('_invoice_num_')
    selection.Text = invoiceNum
    selection.WholeStory()

    selection.Find.Execute('_subdivision_')
    selection.Text = subdivision
    selection.WholeStory()

    selection.Find.Execute('_reference_num_')
    selection.Text = referenceNum
    selection.WholeStory()

    selection.Find.Execute('_mps_')
    selection.Text = mps
    selection.WholeStory()

    selection.Find.Execute('_location_')
    selection.Text = location
    selection.WholeStory()

    selection.Find.Execute('_county_')
    selection.Text = county
    selection.WholeStory()

    selection.Find.Execute('_state_')
    selection.Text = state
    selection.WholeStory()

    selection.Find.Execute('_amount_')
    selection.Text = fillWithWhitespace(str(testSum), len(selection.Text))
    

def fillWithWhitespace(str, expectedSize):
    #TODO This will need to use the largest amount (num of digits) as expected
    difference = expectedSize - len(str)
    if(difference <= 0):
        return str
    else:
        return (' ' * difference) + str

def getInputs():
    def submit():
        excelToWord( invoiceEntry.get(), subdivisionEntry.get(), referenceEntry.get(), mpEntry.get(), locationEntry.get(), countyEntry.get(), stateEntry.get() )
        window.quit()

    window = Tk()
    
    titleLabel = Label(window, text = "Enter data: ", font = "-weight bold")
    titleLabel.grid(row = 0, column = 0)

    invoiceLabel = Label(window, text = "Invoice #: ")
    invoiceLabel.grid(row = 1, column = 0, pady = 10)
    invoiceEntry = Entry(window, width = 20)
    invoiceEntry.grid(row = 1, column = 1)

    subdivisionLabel = Label(window, text = "Subdivision: ")
    subdivisionLabel.grid(row = 2, column = 0, pady = 10)
    subdivisionEntry = Entry(window, width = 20)
    subdivisionEntry.grid(row = 2, column = 1)
    
    referenceLabel = Label(window, text = "Reference #: ")
    referenceLabel.grid(row = 3, column = 0, pady = 10)
    referenceEntry = Entry(window, width = 20)
    referenceEntry.grid(row = 3, column = 1)
    
    mpLabel = Label(window, text = "MP(s): ")
    mpLabel.grid(row = 4, column = 0, pady = 10)
    mpEntry = Entry(window, width = 20)
    mpEntry.grid(row = 4, column = 1)
    
    locationLabel = Label(window, text = "Location (site): ")
    locationLabel.grid(row = 5, column = 0, pady = 10)
    locationEntry = Entry(window, width = 20)
    locationEntry.grid(row = 5, column = 1, padx = (0,10))

    countyLabel = Label(window, text = "County : ")
    countyLabel.grid(row = 5, column = 2, pady = 10)
    countyEntry = Entry(window, width = 20)
    countyEntry.grid(row = 5, column = 3, padx = (0,10))

    stateLabel = Label(window, text = "State : ")
    stateLabel.grid(row = 5, column = 4, pady = 10)
    stateEntry = Entry(window, width = 5)
    stateEntry.grid(row = 5, column = 5)
    
    submitButton = Button(window, text = "Submit", command = submit)
    submitButton.grid(row = 8, column = 4, sticky = S, pady = 10)
    exitButtonWidget = Button(window, text = "Exit", command = window.quit, bg = "red")
    exitButtonWidget.grid(row = 8, column = 5, sticky = E, columnspan = 2, pady = 10, padx = 10)
    
    window.mainloop()


if __name__ == "__main__":
    getInputs()