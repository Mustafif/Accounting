from openpyxl import load_workbook
from openpyxl.styles import Font


def transaction():
    # Set Account Columns
    workbook = load_workbook(filename="Accounting.xlsx")
    sheet = workbook["Journal"]

    count = sheet.max_row + 2

    Date = "A"

    Day = "B"
    Acc = "C"
    Dr = "E"
    Cr = "F"

    print("Please enter DD-MM-YYYY: ")
    date = str(input())

    day = date[0] + date[1]
    month_year = f'{date[2] + date[3]}, {date[4] + date[5] + date[6] + date[7]}'

    sheet[Date + str(count)] = month_year
    count += 1
    sheet[Day + str(count)] = day

    dr = int(input("How many Debit accounts: "))
    debit_acc = []
    i = 0
    print("Enter Debit Account: ")
    while i < dr:
        acc = str(input())
        debit_acc.append(acc)
        sheet[Acc + str(count)] = debit_acc[i]
        amount = float(input("Enter amount: "))
        sheet[Dr + str(count)] = amount
        count += 1
        i += 1
        if dr > 1 and i < dr:
            print("Enter Debit Account: ")

    cr = int(input("How many Credit accounts: "))
    credit_acc = []
    k = 0
    print("Enter Credit Account: ")
    while k < cr:
        acc = str(input())
        credit_acc.append(acc)
        sheet[Acc + str(count)] = credit_acc[k]
        sheet[Acc + str(count)].font = Font(name='Arial', size=10, bold=True)
        amount = float(input("Enter amount: "))
        sheet[Cr + str(count)] = amount
        count += 1
        k += 1
        if cr > 1 and k < cr:
            print("Enter Credit Account: ")
    sheet[Acc + str(count)] = str(input("Enter Transaction Description: "))
    sheet[Acc + str(count)].font = Font(name='Arial', size=10, italic=True)

    workbook.save(filename="Accounting.xlsx")
