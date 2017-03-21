import xlrd

def main():
    workbook = xlrd.open_workbook("NCEI_test01.xlsx")
    for sheet in sheets:
        parse(sheet)

def get_all_sheets(workbook):
    sheets = []
    for n in range(workbook.nsheets):
        sheets.append(workbook.sheet_by_index(n))
    return sheets

def get_sheet_data(sheet):
    sheetArray = []
    for n in range(sheet.ncols):
        print n

if __name__ == '__main__':
    main()
