import os, re, time, traceback, subprocess
import win32com.client as win32
import webbrowser
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

ABOUT_TEXT = 'Product developed by Anthony Spoerl and Zach Garlinghouse for GSS Inc.\n\nIf you have any questions or comments please contact anthonyjspoerl@gmail.com'

APPLICATION_NAME = 'Spreadsheet To PDF'
PROP_FILE = '.prop'
TEMPLATE_PATH = os.getcwd() + '/templates/'
HELP_PATH = os.getcwd() + '/help.html'
DEFAULT_OUTPUT_FOLDER = os.path.expanduser('~') + '/Documents/'
INPUT_FILETYPES = [('Excel', '*.xlsx;*.xls;*.xlsm'),('All', '*.*')]
EMERGENCY_EXIT_THRESHOLD = 100
GET_DESCRIPTION_ERROR = 'Search ran too long in Sage spreadsheet without finding "Report". See help for more details.'
DATE_TUPLE = 0
TRIBE_TUPLE = 1
FEE_TUPLE = 2
DATETIME_FORMAT = '%m/%d/%y'

PER_TRIBE_GSS_FEE = 40
TCNS_REGEX = re.compile('.*(TCNS).*', re.IGNORECASE)

# Tribe list consts
TRIBE_LIST_FILE = TEMPLATE_PATH + 'TribeList.xlsx'
LIST_START_INDEX = 4
LIST_END = 'END'
SAGE_TRIBE_COLUMN = 1
GSS_TRIBE_COLUMN = 2
 
# Sage spreadsheet consts
JOB_COLUMN = 1
DATE_COLUMN = 6
DESCRIPION_COLUMN = 7
TCNS_COLUMN = 8
FEE_COLUMN = 9
JOB_COLUMN_DELIMETER = 'Job ID'
DATE_COLUMN_DELIMETER = 'Trx Date'
DESCRIPION_COLUMN_DELIMETER = 'Trans Description'
TCNS_COLUMN_DELIMETER = 'Trans Ref'
FEE_COLUMN_DELIMETER = 'Amount'

SAGE_END_DELIMETER = 'Report'

TRIBAL_FEE_DICTIONARY = {}
MAPPING_LIST = ['ar', 'bm']
FIELD_SURVEY_HPI_IDENTIFIER = 'hpi'
RECORD_SEARCH_FEE_IDENTIFIER = 'rsf'
DEFAULT_OPEN_PATH = ''
DEFAULT_SAVETO_PATH = ''
tcnsNumberSet = set()
dates = set()
savePath = ''
excel = 0
ss = 0
word = 0
doc = 0

# This will contain all enums/module constants from coms bound with EnsureDispatch
COM_CONSTANTS = win32.constants

#----------------------------------------------------------------------
def menuHelp():
    webbrowser.open(HELP_PATH)

def menuAbout():
    top = Toplevel()
    top.title("About Info")
    msg = Message(top, text = ABOUT_TEXT, width = 1000, justify = CENTER)
    msg.pack()

def setupTribalsDictionary():
    spreadsheet = openExcel(TRIBE_LIST_FILE)
    loadTribes(spreadsheet)
    closeExcel()

def loadTribes(spreadsheet):
    global TRIBAL_FEE_DICTIONARY
    sageTribe = ''
    index = LIST_START_INDEX
    sageTribe = spreadsheet.Cells(index, SAGE_TRIBE_COLUMN).Value
    while sageTribe != LIST_END:
        if(sageTribe != None and sageTribe.strip() != ''):
            sageTribe = sageTribe[:30] # Sage only uses 30 chars in description
            tribe = spreadsheet.Cells(index, GSS_TRIBE_COLUMN).Value
            TRIBAL_FEE_DICTIONARY[sageTribe] = tribe
        index += 1
        sageTribe = spreadsheet.Cells(index, SAGE_TRIBE_COLUMN).Value
    # Could fail if there is no 'END' signifier, maybe add a timeout to be sure

def excelToWord(spreadsheetName, invoiceNum, subdivision, referenceNum, mps, location, county, state):
    spreadsheet = openExcel(spreadsheetName)
    descriptions = getDescriptionsInSpreadsheet(spreadsheet)
    tribes = filterTribes(descriptions)
    saveTribals(tribes, invoiceNum, subdivision, referenceNum, mps, location, county, state)
    fieldSurveyFees = filterFieldSurvey(descriptions)
    saveFieldSurveyFee(fieldSurveyFees, invoiceNum, subdivision, referenceNum, mps, location, county, state)
    # mappings = filterMappings(descriptions)
    # saveMappings(mappings, invoiceNum, subdivision, referenceNum, mps, location, county, state)

def filterFieldSurvey(descriptions):
    fieldSurveyFees = []
    fieldSurveyFees.append(0) # fee
    fieldSurveyFees.append(0) # admin fee
    for index in range(0, len(descriptions)):
        description = descriptions[index][TRIBE_TUPLE]
        if description.find('Item: ' + FIELD_SURVEY_HPI_IDENTIFIER) != -1:
            fieldSurveyFees[0] += float(descriptions[index][FEE_TUPLE])
            fieldSurveyFees[1] += 800.0
        elif description.find('Item: ' + RECORD_SEARCH_FEE_IDENTIFIER) != -1:
            fieldSurveyFees[0] += float(descriptions[index][FEE_TUPLE])
            fieldSurveyFees[1] += float(descriptions[index][FEE_TUPLE]) * 0.5

    return fieldSurveyFees

def saveFieldSurveyFee(fee, invoiceNum, subdivision, referenceNum, mps, location, county, state):
    if fee:
        openWordTemplate('Field Survey.docx')
        replaceFieldSurveyEntryFields(fee, invoiceNum, subdivision, referenceNum, mps, location, county, state)
        saveName = savePath + ' Field Survey'
        saveDoc(saveName)

def setup():
    global word, excel
    word = win32.gencache.EnsureDispatch('Word.Application')
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    loadPropertyFile()

def loadPropertyFile():
    if os.path.isfile(PROP_FILE):
        global DEFAULT_OPEN_PATH, DEFAULT_SAVETO_PATH, TRIBE_LIST_FILE
        propFile = open(PROP_FILE, 'r')
        propFile.readline() # [open]
        DEFAULT_OPEN_PATH = propFile.readline().strip('\n')
        propFile.readline() # [saveTo]
        DEFAULT_SAVETO_PATH = propFile.readline().strip('\n')
        propFile.readline() # [tribalsFile]
        tribalsFile = propFile.readline().strip('\n')
        if os.path.isfile(tribalsFile):
            TRIBE_LIST_FILE = tribalsFile

def isTexasInvoiceNumber(invoiceNum):
    return invoiceNum[0] == 'D' or invoiceNum[0] == 'd'

def saveTribals(tribes, invoiceNum, subdivision, referenceNum, mps, location, county, state):
    if tribes:
        openWordTemplate('Tribals_TX.docx' if isTexasInvoiceNumber(invoiceNum) else 'Tribals.docx')
        replaceEntryFields(invoiceNum, subdivision, referenceNum, mps, location, county, state, tcnsNumberSet)
        insertTribalFees( tribes )
        saveName = savePath + ' Tribals'
        saveDoc(saveName)
        openPDF(saveName)

def saveMappings(mappings, invoiceNum, subdivision, referenceNum, mps, location, county, state):
    if mappings:
        openWordTemplate('Mapping_TX.docx' if isTexasInvoiceNumber(invoiceNum) else 'Mapping.docx')
        replaceEntryFields(invoiceNum, subdivision, referenceNum, mps, location, county, state, tcnsNumberSet)
        saveDoc(savePath + ' Mapping')

def saveDoc(saveName):
    doc.SaveAs(saveName)
    doc.ExportAsFixedFormat(saveName, COM_CONSTANTS.wdExportFormatPDF)

def openPDF(saveName):
    subprocess.Popen(saveName + '.pdf',shell=True)

def cleanup():
    closeExcel()
    closeWord()

def openExcel(spreadsheetName):
    global excel, ss
    ss = excel.Workbooks.Open(spreadsheetName)
    sh = ss.ActiveSheet

    excel.Visible = False
    return sh

def closeExcel():
    global ss
    if(ss != 0):
        ss.Close(False)
        ss = 0
    if(excel != 0):
        excel.Application.Quit()

def closeWord():
    global doc
    if(doc != 0):
        doc.Close(False)
        doc = 0
    if(word != 0):
        word.Application.Quit()

def openWordTemplate(templateName):
    global word, doc
    doc = word.Documents.Open(TEMPLATE_PATH + templateName)
    doc.ActiveWindow.View.Type = COM_CONSTANTS.wdPrintView
    word.Visible = False
        
def replaceEntryFields(invoiceNum, subdivision, referenceNum, mps, location, county, state, tcnsNumbers = None):
    findAndReplace('_invoice_num_', invoiceNum)
    findAndReplace('_subdivision_', subdivision)
    findAndReplace('_reference_num_', referenceNum)
    if mps:
        findAndReplace('_mps_', mps)
    else:
        findAndReplace('_mps_', 'NA')
    findAndReplace('_location_', location)
    findAndReplace('_county_', county)
    findAndReplace('_state_', state)
    multipleFindAndReplace('_date_paid_', dates)
    if tcnsNumbers:
        multipleFindAndReplace('_trans_ref_num_', tcnsNumbers)

def replaceFieldSurveyEntryFields(fee, invoiceNum, subdivision, referenceNum, mps, location, county, state, tcnsNumbers = None):
    findAndReplace('_invoice_num_', invoiceNum)
    findAndReplace('_subdivision_', subdivision)
    findAndReplace('_reference_num_', referenceNum)
    if mps:
        findAndReplace('_mps_', mps)
    else:
        findAndReplace('_mps_', 'NA')
    findAndReplace('_location_', location)
    findAndReplace('_county_', county)
    findAndReplace('_state_', state)
    multipleFindAndReplace('_dates_paid_', dates)
    findAndReplace('_fee_', fee[0])
    findAndReplace('_admin_fee_', fee[1])
    findAndReplace('_total_', fee[0] + fee[1])


def findColumnHeaderIndices(spreadsheet):
    global JOB_COLUMN, DATE_COLUMN, DESCRIPION_COLUMN, TCNS_COLUMN, FEE_COLUMN
    JOB_COLUMN = spreadsheet.Cells.Find(JOB_COLUMN_DELIMETER).Column
    DATE_COLUMN = spreadsheet.Cells.Find(DATE_COLUMN_DELIMETER).Column
    DESCRIPION_COLUMN = spreadsheet.Cells.Find(DESCRIPION_COLUMN_DELIMETER).Column
    TCNS_COLUMN = spreadsheet.Cells.Find(TCNS_COLUMN_DELIMETER).Column
    FEE_COLUMN = spreadsheet.Cells.Find(FEE_COLUMN_DELIMETER).Column


def getDescriptionsInSpreadsheet(spreadsheet):
    global tcnsNumberSet
    findColumnHeaderIndices(spreadsheet)
    index = 2
    emergencyExitCounter = 0
    delimeter = spreadsheet.Cells(index,JOB_COLUMN).Value
    descriptions = []
    description = spreadsheet.Cells(index,DESCRIPION_COLUMN).Value
    while delimeter != SAGE_END_DELIMETER and emergencyExitCounter < EMERGENCY_EXIT_THRESHOLD:
        if description:
            date = spreadsheet.Cells(index, DATE_COLUMN).Value
            fee = spreadsheet.Cells(index, FEE_COLUMN).Value # Only used for tribe, bother checking for None?
            descriptions.append( (date, description, fee) )
        index += 1
        description = spreadsheet.Cells(index,DESCRIPION_COLUMN).Value

        delimeter = spreadsheet.Cells(index,JOB_COLUMN).Value
        if delimeter == None:
            emergencyExitCounter += 1
        else:
            emergencyExitCounter = 0

        tempTcns = spreadsheet.Cells(index,TCNS_COLUMN).Value
        if tempTcns != None and TCNS_REGEX.match(tempTcns) != None:
            tcnsNumberSet.add(tempTcns[5:]) # slice out tcns at beggining

    if(emergencyExitCounter >= 100):
        raise Exception(GET_DESCRIPTION_ERROR)

    return descriptions

def filterTribes(descriptions):
    global dates
    tribes = {}
    unmatchedTribes = []
    for index in range(0, len(descriptions)):
        tribe = descriptions[index][TRIBE_TUPLE].split('- Item:')[0].strip()
        if tribe in TRIBAL_FEE_DICTIONARY:
            fee = descriptions[index][FEE_TUPLE]
            if tribe in tribes:
                tribes[tribe][0] += 1
                tribes[tribe][1] += fee
            else:
                tribes[tribe] = []
                tribes[tribe].append(1) # index 0
                tribes[tribe].append(fee) # index 1
            dates.add( descriptions[index][DATE_TUPLE].strftime(DATETIME_FORMAT) )
        else:
            unmatchedTribes.append(tribe)
    print('These tribes were not found in the tribal list')
    print(unmatchedTribes)
    return tribes

def insertTribalFees(tribes):
    setCopyText(len(tribes))

    tribeCount = 0
    total = 0
    print('inserting tribes:')
    print(tribes)
    for tribe in tribes:
        if tribe in TRIBAL_FEE_DICTIONARY:
            tribeName = TRIBAL_FEE_DICTIONARY[tribe]
            fee = tribes[tribe][1]

            tribeCount += tribes[tribe][0]
            print('Tribal fee:')
            print(fee)
            total += fee

            findAndReplace('_tribe_', tribeName)
            findAndReplace('_amount_', "{:,.2f}".format(fee))
    adminFee = tribeCount * PER_TRIBE_GSS_FEE
    findAndReplace('_admin_fee_', "{:,.2f}".format(adminFee))
    total += adminFee
    findAndReplace('_total_',"{:,.2f}".format(total))

def filterMappings(descriptions):
    mappingsFee = 0
    for index in range(0, len(descriptions)):
        print(descriptions[index][TRIBE_TUPLE])
        for identifier in MAPPING_LIST:
            print(descriptions[index][TRIBE_TUPLE].find('Item: ' + identifier) != -1)
            if descriptions[index][TRIBE_TUPLE].find('Item: ' + identifier) != -1:
                mappingsFee += descriptions[index][FEE_TUPLE]
                break

    print('mapping fee: ' + str(mappingsFee))
    return {}

def multipleFindAndReplace(placeholder, itemSet):
    replacementText = ''
    for item in itemSet:
        replacementText += item + ' & '

    findAndReplace(placeholder, replacementText[:-3]) # Trim hanging ampersand

def setCopyText(numTribes):
    selection = word.Selection

    selection.Find.Execute('_tribe_')
    selection.Expand(COM_CONSTANTS.wdLine)
    selection.Copy()
    for index in range(0,numTribes):
        selection.Paste()
    selection.WholeStory()

def findAndReplace(searchTerm, replacement):
    selection = word.Selection

    selection.Find.Execute(searchTerm)
    if type(replacement) is float or type(replacement) is int:
        selection.Text = "{:.2f}".format(replacement)
    else:
        selection.Text = str(replacement)

    selection.WholeStory()

def getInputs():
    def getSpreadsheetName():
        filename = filedialog.askopenfilename(initialdir = fileEntry.get(), defaultextension = '.xlsx', filetypes = INPUT_FILETYPES)
        if(filename != ''):
            fileEntry.delete(0, END)
            fileEntry.insert(0, filename)
            fileEntry.xview_moveto(1)

    def getSaveFilePath():
        filename = filedialog.asksaveasfilename(initialdir = saveFileEntry.get())
        if(filename != ''):
            saveFileEntry.delete(0, END)
            saveFileEntry.insert(0, filename)
            saveFileEntry.xview_moveto(1)

    def submit(event = None):
        try:
            global savePath
            if invoiceEntry.get() == '':
                messagebox.showinfo('Missing Invoice Number', 'Please provide an invoice number.')
            else:
                savePath = os.path.abspath(saveFileEntry.get().replace("\\", "\\\\"))
                excelToWord( fileEntry.get(), invoiceEntry.get(), subdivisionEntry.get(), referenceEntry.get(), mpEntry.get(), locationEntry.get(), countyEntry.get(), stateEntry.get() )
                window.quit()
        except:
            messagebox.showerror("Error", "An error has occured. For more information, see errors.log in your Sage to PDF folder.")

            errorLog = open('errors.log', 'a')
            errorLog.write(time.strftime("\n%m/%d/%y %H:%M:%S\n"))
            errorLog.write(traceback.format_exc())
            errorLog.close()

            cleanup()

    def openPreferencesDialog():
        def savePreferences():
            DEFAULT_OPEN_PATH = openEntry.get()
            DEFAULT_SAVETO_PATH = saveToEntry.get()
            fileEntry.insert(0,DEFAULT_OPEN_PATH)
            saveFileEntry.insert(0,DEFAULT_SAVETO_PATH)

            propFile = open(PROP_FILE, 'r')

            data = propFile.readlines()
            data[1] = DEFAULT_OPEN_PATH + '\n'
            data[3] = DEFAULT_SAVETO_PATH + '\n'

            propFile = open(PROP_FILE, 'w')
            propFile.writelines(data)

            preferencesWindow.destroy()

        preferencesWindow = Tk()
        preferencesWindow.wm_title('Preferences')

        openLabel = Label(preferencesWindow, text = "Default open path: ")
        openLabel.grid(row = 0, column = 0)
        openEntry = Entry(preferencesWindow, width = 80)
        openEntry.insert(0, DEFAULT_OPEN_PATH)
        openEntry.grid(row = 0, column = 1)
        saveToLabel = Label(preferencesWindow, text = "Default save path and name: ")
        saveToLabel.grid(row = 1, column = 0)
        saveToEntry = Entry(preferencesWindow, width = 80)
        saveToEntry.insert(0, DEFAULT_SAVETO_PATH)
        saveToEntry.grid(row = 1, column = 1)
        tribalsListLabel = Label(preferencesWindow, text = "Tribe list spreadsheet: ")
        tribalsListLabel.grid(row = 2, column = 0)
        tribalsListEntry = Entry(preferencesWindow, width = 80)
        tribalsListEntry.insert(0, TRIBE_LIST_FILE)
        tribalsListEntry.grid(row = 2, column = 1)
        preferencesFrame = Frame(preferencesWindow)
        preferencesFrame.grid(row = 3, column = 1, sticky = E)
        okButton = Button(preferencesFrame, text = "Ok", command = savePreferences)
        okButton.grid(row = 0, column = 0, sticky = E, padx = 10)
        cancelButtonWidget = Button(preferencesFrame, text = "Cancel", command = preferencesWindow.destroy, bg = "red")
        cancelButtonWidget.grid(row = 0, column = 1, sticky = E)

        preferencesWindow.mainloop()

    window = Tk()
    window.wm_title(APPLICATION_NAME)
    
    fileFrame = Frame(window)
    fileFrame.grid(row = 1, column = 0, sticky=W)

    titleLabel = Label(fileFrame, text = "Enter data: ", font = "-weight bold")
    titleLabel.grid(row = 0, column = 0)

    window.bind("<Return>", submit)

    menubar = Menu(window)
    filemenu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label = "Preferences", command = openPreferencesDialog)
    menubar.add_cascade(label = "Help", menu = filemenu)
    filemenu.add_command(label = "Help", command = menuHelp)
    filemenu.add_command(label = "About", command = menuAbout)
    window.config(menu = menubar)

    fileLabel = Label(fileFrame, text = "Spreadsheet: ")
    fileLabel.grid(row = 1, column = 0, pady = 10, sticky = W)
    fileEntry = Entry(fileFrame, width = 60)
    fileEntry.insert(0,DEFAULT_OPEN_PATH)
    fileEntry.grid(row = 1, column = 1, sticky = W)
    openFileButton = Button(fileFrame, text = 'Open...', command = getSpreadsheetName)
    openFileButton.grid(row = 1, column = 2)

    saveFileLabel = Label(fileFrame, text = "Save as: ")
    saveFileLabel.grid(row = 2, column = 0, pady = 10, sticky = W)
    saveFileEntry = Entry(fileFrame, width = 60)
    saveFileEntry.insert(0,DEFAULT_SAVETO_PATH)
    saveFileEntry.grid(row = 2, column = 1, sticky = W)
    openFileButton = Button(fileFrame, text = 'Open...', command = getSaveFilePath)
    openFileButton.grid(row = 2, column = 2)

    entryFrame = Frame(window)
    entryFrame.grid(row = 3, column = 0, sticky = W)
    invoiceLabel = Label(entryFrame, text = "Invoice #: ")
    invoiceLabel.grid(row = 3, column = 0, pady = 10, sticky = W)
    invoiceEntry = Entry(entryFrame, width = 20)
    invoiceEntry.grid(row = 3, column = 1, sticky = W)

    subdivisionLabel = Label(entryFrame, text = "Subdivision: ")
    subdivisionLabel.grid(row = 4, column = 0, pady = 10, sticky = W)
    subdivisionEntry = Entry(entryFrame, width = 20)
    subdivisionEntry.grid(row = 4, column = 1, stick = W)
    
    referenceLabel = Label(entryFrame, text = "Reference #: ")
    referenceLabel.grid(row = 5, column = 0, pady = 10, sticky = W)
    referenceEntry = Entry(entryFrame, width = 20)
    referenceEntry.grid(row = 5, column = 1, sticky = W)
    
    mpLabel = Label(entryFrame, text = "MP(s): ")
    mpLabel.grid(row = 6, column = 0, pady = 10, sticky = W)
    mpEntry = Entry(entryFrame, width = 20)
    mpEntry.grid(row = 6, column = 1, sticky = W)
    
    locationLabel = Label(entryFrame, text = "Location (site): ")
    locationLabel.grid(row = 7, column = 0, pady = 10, sticky = W)
    locationEntry = Entry(entryFrame, width = 20)
    locationEntry.grid(row = 7, column = 1, padx = (0,10), sticky = W)

    countyLabel = Label(entryFrame, text = "County : ")
    countyLabel.grid(row = 7, column = 2, pady = 10, sticky = W)
    countyEntry = Entry(entryFrame, width = 20)
    countyEntry.grid(row = 7, column = 3, padx = (0,10), sticky = W)

    stateLabel = Label(entryFrame, text = "State : ")
    stateLabel.grid(row = 7, column = 4, pady = 10)
    stateEntry = Entry(entryFrame, width = 5)
    stateEntry.grid(row = 7, column = 5)
    
    submitButton = Button(entryFrame, text = "Submit", command = submit)
    submitButton.grid(row = 8, column = 4, sticky = S, pady = 10)
    exitButtonWidget = Button(entryFrame, text = "Exit", command = window.quit, bg = "red")
    exitButtonWidget.grid(row = 8, column = 5, sticky = E, columnspan = 2, pady = 10, padx = 10)
    
    window.mainloop()


if __name__ == "__main__":
    try:
        setup()
        setupTribalsDictionary()
        getInputs()
        cleanup()
    except:
        messagebox.showerror("Error", "An error has occured. For more information, see errors.log in your Sage to PDF folder.")
        
        errorLog = open('errors.log', 'a')
        errorLog.write(time.strftime("\n%m/%d/%y %H:%M:%S\n"))
        errorLog.write(traceback.format_exc())
        errorLog.close()
        
        cleanup()


#Sante Sioux: markup %15 of cost
#Ponca Tribe: PTC vs Non PTC (special case) ## use PTC by default

### Remove Enter to submit ###