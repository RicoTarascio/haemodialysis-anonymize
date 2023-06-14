from openpyxl import Workbook


def test():
    wb = Workbook()

    wb.save("balances.xlsx")
