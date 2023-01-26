import win32com.client

localPath = 'C:/Projects/Python/data-analysis/teste.xlsx'

def refresh_excel():
    xlapp = win32com.client.DispatchEx('Excel.Application')
    wb = xlapp.Workbooks.Open(localPath)
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    xlapp.DisplayAlerts = False
    wb.Save()
    wb.Close()

from tkinter import *

windown = Tk()
windown.title('REFRESH EXCEL')

orientation_text = Label(windown, text='Click to refresh Excel')
orientation_text.grid(column=0, row=0)

btn = Button(windown, text='Refresh', command=refresh_excel )
btn.grid(column=0, row=1)

windown.mainloop()
