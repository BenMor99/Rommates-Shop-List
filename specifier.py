# this module will make  a table which specify the debters and creditors
#and simply specifies who should pay and how much should be paid

import Code

sum1 = []

wb = openpyxl.load_workbook('Expenses.xlsx',data_only=True)
ws = wb.active

for c in range(1,19):
    r = Code.row_finder(c)
    if r <= 2:
        sum1.append(0)
    else:
        sum1.append(ws.cell(row=r,column=c).value)

sum2 = []

sum2.append(sum1[4] + sum1[3]/3) # BtoA = sum2[0]
sum2.append(sum1[7] + sum1[6]/3) # BtoH = sum2[1]
sum2.append(sum1[1] + sum1[0]/3) # AtoB = sum2[2]
sum2.append(sum1[8] + sum1[6]/3) # AtoH = sum2[3]
sum2.append(sum1[2] + sum1[0]/3) # HtoB = sum2[4]
sum2.append(sum1[5] + sum1[3]/3) # HtoA = sum2[5]

