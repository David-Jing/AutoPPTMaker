import xlrd

loc = "PPTConfigurator.xlsx"


def getPPTData():
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    print(sheet.cell_value(3, 4))


if __name__ == '__main__':
    # something something
