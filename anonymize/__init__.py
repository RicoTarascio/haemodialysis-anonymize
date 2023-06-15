from openpyxl import Workbook


def test():
    wb = Workbook()

    wb.save("tests/balances.xlsx")
