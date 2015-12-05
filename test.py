import time
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
    doc.SaveAs('test1')
    doc.Close(False)
    word.Application.Quit()

def openExcel():
    """"""
    global excel, ss
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    ss = excel.Workbooks.Add()
    sh = ss.ActiveSheet

    excel.Visible = True
    time.sleep(1)

    sh.Cells(1,1).Value = 'Hacking Excel with Python Demo'

    time.sleep(1)
    for i in range(2,8):
        sh.Cells(i,1).Value = 'Line %i' % i
        time.sleep(1)

    return sh

def openWord(spreadsheet):
    global word, doc
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Add()
    word.Visible = True
    time.sleep(1)
 
    rng = doc.Range(0,0)
    rng.InsertAfter(spreadsheet.Cells(1,1).Value)
    time.sleep(1)
    for i in RANGE:
        rng.InsertAfter('Line %d\r\n' % i)
        time.sleep(1)
    rng.InsertAfter("\r\nPython rules!\r\n")
 


if __name__ == "__main__":
    excelToWord()
