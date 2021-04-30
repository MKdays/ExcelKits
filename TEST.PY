import win32com.client
excel = win32com.client.GetActiveObject("Excel.Application")
wb = excel.activeWorkbook
ws = wb.ActiveSheet

# colored_value = []
# colored_value_code = []
# for i in ws.UsedRange:
#     if i.Interior.ColorIndex != -4142: #기본값이 아니라면
#         colored_value.append([i.value2, None])
#         colored_value_code.append([i.Interior.ColorIndex, None])

# cnt = len(colored_value)
# ws.Range(excel.Cells(2, 1),excel.Cells(1+cnt, 1)).value2 = colored_value
# ws.Range(excel.Cells(2, 2),excel.Cells(1+cnt, 2)).value2 = colored_value_code

ws.UsedRange

A = len(ws.UsedRange)

for i in A:
    print()