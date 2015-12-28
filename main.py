import os
import re
import win32com.client as win32
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

APPLICATION_NAME = 'Spreadsheet Too PDF'
TEMPLATE_PATH = os.getcwd() + '/templates/'
RANGE = range(3, 8)
ADD_REGEX = re.compile('.*(add).*', re.IGNORECASE)
MULTIPLY_REGEX = re.compile('.*(multiply).*', re.IGNORECASE)
INPUT_FILETYPES = [('Excel', '*.xlsx;*.xls;*.xlsm'),('All', '*.*')]

# Tribe list consts
TRIBE_LIST_FILE = TEMPLATE_PATH + 'TribeList.xlsx'
SAGE_TRIBE_COLUMN = 0
GSS_TRIBE_COLUMN = 1
FEE_COLUMN = 3

excel = 0
ss = 0
word = 0
doc = 0

# This will contain all enums/module constants from coms bound with EnsureDispatch
COM_CONSTANTS = win32.constants

#----------------------------------------------------------------------

def excelToWord(spreadsheetName, invoiceNum, subdivision, referenceNum, mps, location, county, state):
    try:
        setup()
        spreadsheet = openExcel(spreadsheetName)
        saveTribals(spreadsheet, invoiceNum, subdivision, referenceNum, mps, location, county, state)
        cleanup()
    except Exception as e:
        messagebox.showerror("Error", str(e))
        cleanup()

def setup():
    global word, excel
    word = win32.gencache.EnsureDispatch('Word.Application')
    excel = win32.gencache.EnsureDispatch('Excel.Application')

def saveTribals(spreadsheet, invoiceNum, subdivision, referenceNum, mps, location, county, state):
    testSum = openWordTemplate(spreadsheet, 'Tribals.docx')
    replaceEntryFields(invoiceNum, subdivision, referenceNum, mps, location, county, state, testSum)
    saveDoc('Tribals_out')

def saveDoc(filename):
    doc.SaveAs( os.getcwd() + '/' + filename )
    doc.ExportAsFixedFormat(os.getcwd() + '/' + filename, COM_CONSTANTS.wdExportFormatPDF)

def cleanup():
    if(ss != 0):
        ss.Close(False)
    if(excel != 0):
        excel.Application.Quit()
    if(doc != 0):
        doc.Close(False)
    if(word != 0):
        word.Application.Quit()

def openExcel(spreadsheetName):
    global excel, ss
    ss = excel.Workbooks.Open(spreadsheetName)
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
    def findAndReplace(searchTerm, replacement):
        selection.Find.Execute(searchTerm)
        selection.Text = replacement
        selection.WholeStory()

    selection = word.Selection

    findAndReplace('_invoice_num_', invoiceNum)
    findAndReplace('_subdivision_', subdivision)
    findAndReplace('_reference_num_', referenceNum)
    findAndReplace('_mps_', mps)
    findAndReplace('_location_', location)
    findAndReplace('_county_', county)
    findAndReplace('_state_', state)

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
    def getSpreadsheetName():
        filename = filedialog.askopenfilename(defaultextension = '.xlsx', filetypes = INPUT_FILETYPES)
        if(filename != ''):
            fileEntry.delete(0, END)
            fileEntry.insert(0, filename)
            fileEntry.xview_moveto(1)

    def submit(event = None):
        excelToWord( fileEntry.get(), invoiceEntry.get(), subdivisionEntry.get(), referenceEntry.get(), mpEntry.get(), locationEntry.get(), countyEntry.get(), stateEntry.get() )
        window.quit()

    window = Tk()
    window.wm_title(APPLICATION_NAME)
    
    fileFrame = Frame(window)
    fileFrame.grid(row = 1, column = 0, sticky=W)

    titleLabel = Label(fileFrame, text = "Enter data: ", font = "-weight bold")
    titleLabel.grid(row = 0, column = 0)

    window.bind("<Return>", submit)

    fileLabel = Label(fileFrame, text = "Spreadsheet: ")
    fileLabel.grid(row = 1, column = 0, pady = 10, sticky = W)
    fileEntry = Entry(fileFrame, width = 60)
    fileEntry.grid(row = 1, column = 1, sticky = W)
    openFileButton = Button(fileFrame, text = 'Open...', command = getSpreadsheetName)
    openFileButton.grid(row = 1, column = 2)

    entryFrame = Frame(window)
    entryFrame.grid(row = 2, column = 0, sticky = W)
    invoiceLabel = Label(entryFrame, text = "Invoice #: ")
    invoiceLabel.grid(row = 2, column = 0, pady = 10, sticky = W)
    invoiceEntry = Entry(entryFrame, width = 20)
    invoiceEntry.grid(row = 2, column = 1, sticky = W)

    subdivisionLabel = Label(entryFrame, text = "Subdivision: ")
    subdivisionLabel.grid(row = 3, column = 0, pady = 10, sticky = W)
    subdivisionEntry = Entry(entryFrame, width = 20)
    subdivisionEntry.grid(row = 3, column = 1, stick = W)
    
    referenceLabel = Label(entryFrame, text = "Reference #: ")
    referenceLabel.grid(row = 4, column = 0, pady = 10, sticky = W)
    referenceEntry = Entry(entryFrame, width = 20)
    referenceEntry.grid(row = 4, column = 1, sticky = W)
    
    mpLabel = Label(entryFrame, text = "MP(s): ")
    mpLabel.grid(row = 5, column = 0, pady = 10, sticky = W)
    mpEntry = Entry(entryFrame, width = 20)
    mpEntry.grid(row = 5, column = 1, sticky = W)
    
    locationLabel = Label(entryFrame, text = "Location (site): ")
    locationLabel.grid(row = 6, column = 0, pady = 10, sticky = W)
    locationEntry = Entry(entryFrame, width = 20)
    locationEntry.grid(row = 6, column = 1, padx = (0,10), sticky = W)

    countyLabel = Label(entryFrame, text = "County : ")
    countyLabel.grid(row = 6, column = 2, pady = 10, sticky = W)
    countyEntry = Entry(entryFrame, width = 20)
    countyEntry.grid(row = 6, column = 3, padx = (0,10), sticky = W)

    stateLabel = Label(entryFrame, text = "State : ")
    stateLabel.grid(row = 6, column = 4, pady = 10)
    stateEntry = Entry(entryFrame, width = 5)
    stateEntry.grid(row = 6, column = 5)
    
    submitButton = Button(entryFrame, text = "Submit", command = submit)
    submitButton.grid(row = 7, column = 4, sticky = S, pady = 10)
    exitButtonWidget = Button(entryFrame, text = "Exit", command = window.quit, bg = "red")
    exitButtonWidget.grid(row = 7, column = 5, sticky = E, columnspan = 2, pady = 10, padx = 10)
    
    window.mainloop()


if __name__ == "__main__":
    getInputs()

#GSS Admin Fee: # of tribes * 40
#Sante Sioux: markup %15 of cost
#Eat My Shorts