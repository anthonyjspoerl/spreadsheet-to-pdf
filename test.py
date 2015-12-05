import os
import win32com.client as win32

RANGE = range(3, 8)

excel = 0
ss = 0
word = 0
doc = 0

#----------------------------------------------------------------------

def excelToWord():
    spreadsheet = openExcel()
    openWord(spreadsheet)
    cleanup()

def cleanup():
    ss.Close(False)
    excel.Application.Quit()
    doc.SaveAs( os.getcwd() + '/test' )
    doc.Close(False)
    word.Application.Quit()

def openExcel():
    global excel, ss
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    #ss = excel.Workbooks.Add()
    ss = excel.Workbooks.Open( os.getcwd() + '/test_sheet')
    sh = ss.ActiveSheet

    excel.Visible = False

    #for i in range(2,8):
    #    sh.Cells(i,1).Value = 'Line %i' % i

    return sh

def openWord(spreadsheet):
    global word, doc
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Add()
    word.Visible = True
 
    rng = doc.Range(0,0)

    index = 2
    val = spreadsheet.Cells(index,1).Value

    sum = 0
    product = 1
    while val:
        numVal = spreadsheet.Cells(index,2).Value
        if(val == 'add this'):
            sum += numVal
        elif(val == 'multiply this'):
            product = product * numVal

        index += 1
        val = spreadsheet.Cells(index,1).Value

    rng.InsertAfter('sum: ')
    rng.InsertAfter(sum)
    rng.InsertAfter('\n')
    rng.InsertAfter('product: ')
    rng.InsertAfter(product)



if __name__ == "__main__":
    excelToWord()
