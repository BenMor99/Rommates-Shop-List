# this module will make  a table which specify the debters and creditors
#and simply specifies who should pay and how much should be paid

import Code
import openpyxl
def spec():
    sum1 = []

    wb = openpyxl.load_workbook('Expenses.xlsx',data_only=True)
    ws = wb.active

    for c in range(1,18,2):
        r = Code.row_finder(c)
        if r <= 2:
            sum1.append(0)
        else:
            val = 0
            for row in range(2,r):
               val += ws.cell(row = row,column = c+1).value
            sum1.append(val)

    del ws
    del wb

    sum2 = []

    sum2.append(sum1[4] + sum1[3]/3) # BtoA = sum2[0]
    sum2.append(sum1[7] + sum1[6]/3) # BtoH = sum2[1]
    sum2.append(sum1[1] + sum1[0]/3) # AtoB = sum2[2]
    sum2.append(sum1[8] + sum1[6]/3) # AtoH = sum2[3]
    sum2.append(sum1[2] + sum1[0]/3) # HtoB = sum2[4]
    sum2.append(sum1[5] + sum1[3]/3) # HtoA = sum2[5]

    class acount:
        def __init__(self,debt,credit):
            self.debt = debt
            self.credit = credit

    B = acount(sum2[0]+sum2[1],sum2[2]+sum2[4])
    b = B.credit - B.debt
    A = acount(sum2[2]+sum2[3],sum2[0]+sum2[5])
    a = A.credit - A.debt
    H = acount(sum2[4]+sum2[5],sum2[1]+sum2[3])
    h = H.credit - H.debt

    f = open('message.txt','w')

    if a == 0 and b == 0 and h == 0:
        f.write('all roommates are even!')

    elif a==0:
        if b > 0:
            f.write('Hamed should pay %d to Behnam' %b)
        else:
            f.write('Behnam should pay %d to Hamed' %h)

    elif b==0:
        if a > 0:
            f.write('Hamed should pay %d to Alireza' %a)
        else:
            f.write('Alireza should pay %d to Hamed' %h)

    elif h==0:
        if b > 0:
            f.write('alireza should pay %d to Behnam' %b)
        else:
            f.write('Behnam should pay %d to Alireza' %a)

    elif (a>=0 and b>=0):
        f.write('Hamed should pay %d to Behnam and %d to Alireza!' %(b , a))

    elif (a>=0 and h>=0):
        f.write('Behnam should pay %d to Hamed and %d to Alireza!' %(h , a))

    elif (b>=0 and h>=0):
        f.write('Alireza should pay %d to Behnam and %d to Hamed!' %(b , h))

    elif b>=0:
        f.write('Hamed should pay %d and Alireza should pay %d to Behnam' %(abs(h) , abs(a)))

    elif a>=0:
        f.write('Hamed should pay %d and Behnam should pay %d to Alireza' %(abs(h) , abs(b)))

    elif h>=0:
        f.write('Alireza should pay %d and Behnam should pay %d to Hamed' %(abs(a) , abs(b)))

    else:
        f.write('Check out the code!\nSomething must have gone wrong!!!')

    f.close()
