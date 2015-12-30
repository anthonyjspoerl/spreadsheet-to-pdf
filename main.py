import os
import re
import win32com.client as win32
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

APPLICATION_NAME = 'Spreadsheet Too PDF'
TEMPLATE_PATH = os.getcwd() + '/templates/'
INPUT_FILETYPES = [('Excel', '*.xlsx;*.xls;*.xlsm'),('All', '*.*')]

# Tribe list consts
TRIBE_LIST_FILE = TEMPLATE_PATH + 'TribeList.xlsx'
LIST_START_INDEX = 4
LIST_END = 'END'
SAGE_TRIBE_COLUMN = 1
GSS_TRIBE_COLUMN = 2
FEE_COLUMN = 4

TRIBAL_FEE_DICTIONARY = {}
excel = 0
ss = 0
word = 0
doc = 0

# This will contain all enums/module constants from coms bound with EnsureDispatch
COM_CONSTANTS = win32.constants

#----------------------------------------------------------------------

def setupTribalsDictionary():
    spreadsheet = openExcel(TRIBE_LIST_FILE)
    loadFees(spreadsheet)
    closeExcel()

def loadFees(spreadsheet):
    global TRIBAL_FEE_DICTIONARY
    sageTribe = ''
    index = LIST_START_INDEX
    while sageTribe != LIST_END:
        sageTribe = spreadsheet.Cells(index, SAGE_TRIBE_COLUMN).Value
        if(sageTribe != None and sageTribe.strip() != ''):
            tribe = spreadsheet.Cells(index, GSS_TRIBE_COLUMN).Value
            fee = spreadsheet.Cells(index, FEE_COLUMN).Value
            TRIBAL_FEE_DICTIONARY[sageTribe] = [tribe, fee]
        index += 1
    # Could fail if there is no 'END' signifier, maybe add a timeout to be sure

def excelToWord(spreadsheetName, invoiceNum, subdivision, referenceNum, mps, location, county, state):
    try:
        spreadsheet = openExcel(spreadsheetName)
        saveTribals(spreadsheet, invoiceNum, subdivision, referenceNum, mps, location, county, state)
    except Exception as e:
        messagebox.showerror("Error", str(e))
        cleanup()

def setup():
    global word, excel
    word = win32.gencache.EnsureDispatch('Word.Application')
    excel = win32.gencache.EnsureDispatch('Excel.Application')

def saveTribals(spreadsheet, invoiceNum, subdivision, referenceNum, mps, location, county, state):
    openWordTemplate(spreadsheet, 'Tribals.docx')
    replaceEntryFields(invoiceNum, subdivision, referenceNum, mps, location, county, state)
    descriptions = getDescriptionsInSpreadsheet(spreadsheet)
    insertTribalFees( filterTribes(descriptions) )
    saveDoc('Tribals_out')

def saveDoc(filename):
    doc.SaveAs( os.getcwd() + '/' + filename )
    doc.ExportAsFixedFormat(os.getcwd() + '/' + filename, COM_CONSTANTS.wdExportFormatPDF)

def cleanup():
    closeExcel()
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

def closeExcel():
    if(ss != 0):
        ss.Close(False)
    if(excel != 0):
        excel.Application.Quit()

def openWordTemplate(spreadsheet, templateName):
    global word, doc
    doc = word.Documents.Open(TEMPLATE_PATH + templateName)
    word.Visible = False

        
def replaceEntryFields(invoiceNum, subdivision, referenceNum, mps, location, county, state):
    findAndReplace('_invoice_num_', invoiceNum)
    findAndReplace('_subdivision_', subdivision)
    findAndReplace('_reference_num_', referenceNum)
    findAndReplace('_mps_', mps)
    findAndReplace('_location_', location)
    findAndReplace('_county_', county)
    findAndReplace('_state_', state)

def getDescriptionsInSpreadsheet(spreadsheet):
    index = 2
    descriptions = []
    description = spreadsheet.Cells(index,1).Value

    while description:
        descriptions.append(description)
        index += 1
        description = spreadsheet.Cells(index,1).Value

    return descriptions

def filterTribes(descriptions):
    tribes = []
    for index in range(0, len(descriptions)):
        tribe = descriptions[index]
        if tribe in TRIBAL_FEE_DICTIONARY:
            tribes.append(tribe)
    return tribes

def insertTribalFees(tribes):
    setCopyText(len(tribes))
    for tribe in tribes:
        if tribe in TRIBAL_FEE_DICTIONARY:
            findAndReplace('_tribe_', TRIBAL_FEE_DICTIONARY[tribe][0])
            findAndReplace('_amount_', TRIBAL_FEE_DICTIONARY[tribe][1])

def setCopyText(numTribes):
    selection = word.Selection

    selection.Find.Execute('_tribe_')
    selection.Expand(COM_CONSTANTS.wdLine)
    selection.Copy()
    for index in range(0,numTribes):
        selection.Paste()
    selection.WholeStory()

def fillWithWhitespace(str, expectedSize):
    #TODO This will need to use the largest amount (num of digits) as expected
    difference = expectedSize - len(str)
    if(difference <= 0):
        return str
    else:
        return (' ' * difference) + str

def findAndReplace(searchTerm, replacement):
    selection = word.Selection

    selection.Find.Execute(searchTerm)
    selection.Text = replacement
    selection.WholeStory()

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
    setup()
    setupTribalsDictionary()
    getInputs()
    cleanup()

#GSS Admin Fee: # of tribes * 40
#Sante Sioux: markup %15 of cost
#Ponca Tribe: PTC vs Non PTC (special case) ## use PTC by default