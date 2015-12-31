import os
import re
import win32com.client as win32
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

APPLICATION_NAME = 'Spreadsheet Too PDF'
TEMPLATE_PATH = os.getcwd() + '/templates/'
DEFAULT_OUTPUT_FOLDER = os.path.expanduser('~') + '/Documents/'
INPUT_FILETYPES = [('Excel', '*.xlsx;*.xls;*.xlsm'),('All', '*.*')]
EMERGENCY_EXIT_THRESHOLD = 100
GET_DESCRIPTION_ERROR = 'Search ran too long in Sage spreadsheet without finding "Report". See help for more details.'

PER_TRIBE_GSS_FEE = 40
TCNS_REGEX = re.compile('.*(TCNS).*', re.IGNORECASE)

# Tribe list consts
TRIBE_LIST_FILE = TEMPLATE_PATH + 'TribeList.xlsx'
LIST_START_INDEX = 4
LIST_END = 'END'
SAGE_TRIBE_COLUMN = 1
GSS_TRIBE_COLUMN = 2
FEE_COLUMN = 4

# Sage spreadsheet consts
JOB_COLUMN = 1
DESCRIPION_COLUMN = 7
TCNS_COLUMN = 8
SAGE_END_DELIMETER = 'Report'

TRIBAL_FEE_DICTIONARY = {}
excel = 0
ss = 0
word = 0
doc = 0

# This will contain all enums/module constants from coms bound with EnsureDispatch
COM_CONSTANTS = win32.constants

#----------------------------------------------------------------------

def menuHelp():
    print("Help")

def menuAbout():
    print("About")    

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
    doc.SaveAs(DEFAULT_OUTPUT_FOLDER + filename)
    doc.ExportAsFixedFormat(DEFAULT_OUTPUT_FOLDER + filename, COM_CONSTANTS.wdExportFormatPDF)

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
    doc.ActiveWindow.View.Type = COM_CONSTANTS.wdPrintView
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
    tcnsNumber = ''
    index = 2
    emergencyExitCounter = 0
    delimeter = spreadsheet.Cells(index,JOB_COLUMN).Value
    descriptions = []
    description = spreadsheet.Cells(index,DESCRIPION_COLUMN).Value
    while delimeter != SAGE_END_DELIMETER and emergencyExitCounter < EMERGENCY_EXIT_THRESHOLD:
        if description:
            descriptions.append(description)
        index += 1
        description = spreadsheet.Cells(index,DESCRIPION_COLUMN).Value
        
        delimeter = spreadsheet.Cells(index,JOB_COLUMN).Value
        if delimeter == None:
            emergencyExitCounter += 1
        else:
            emergencyExitCounter = 0

        tempTcns = spreadsheet.Cells(index,TCNS_COLUMN).Value
        if tempTcns != None and tcnsNumber == '' and TCNS_REGEX.match(tempTcns) != None:
            tcnsNumber = tempTcns

    if(emergencyExitCounter >= 100):
        raise Exception(GET_DESCRIPTION_ERROR)

    findAndReplace('_trans_ref_num_', tcnsNumber)
    return descriptions

def filterTribes(descriptions):
    tribes = {}
    for index in range(0, len(descriptions)):
        tribe = descriptions[index].split('-')[0].strip()
        if tribe in TRIBAL_FEE_DICTIONARY:
            if tribe in tribes:
                tribes[tribe] += 1
            else:
                tribes[tribe] = 1
    return tribes

def insertTribalFees(tribes):
    setCopyText(len(tribes))

    tribeCount = 0
    total = 0
    for tribe in tribes:
        if tribe in TRIBAL_FEE_DICTIONARY:
            tribeName = TRIBAL_FEE_DICTIONARY[tribe][0]
            fee = TRIBAL_FEE_DICTIONARY[tribe][1]

            tribeCount += tribes[tribe]
            total += fee

            findAndReplace('_tribe_', tribeName)
            findAndReplace('_amount_', fee)
    adminFee = tribeCount * PER_TRIBE_GSS_FEE
    findAndReplace('_admin_fee_', adminFee)
    total += adminFee
    findAndReplace('_total_',total)

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

    menubar = Menu(window)
    filemenu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label = "Help", menu = filemenu)
    filemenu.add_command(label = "Help", command = menuHelp)
    filemenu.add_command(label = "About", command = menuAbout)
    window.config(menu = menubar)

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

#Sante Sioux: markup %15 of cost
#Ponca Tribe: PTC vs Non PTC (special case) ## use PTC by default