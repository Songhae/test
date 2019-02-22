import win32com.client

# xptmxm

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('C:/Users/04869/Downloads/Pandas Study/data/korea_pop_2012.xls')
ws = wb.ActiveSheet
print(ws.Cells(3,3).Value)
ws.cells(1,1).value = ws.Cells(3,3).Value + ws.Cells(4,4).Value
# excel.Quit()