import xlrd

def main():
    workbook = xlrd.open_workbook("NCEI_test01.xlsx")

def get_all_sheets(workbook):
    sheets = []
    for n in range(workbook.nsheets):
        sheets.append(workbook.sheet_by_index(n))
    return sheets

def get_sheet_data(sheet):
    sheetArray = []
    for n in range(sheet.ncols):
        print n

def get_all_people(workbook):
    pass

def get_all_variables():
    pass

def get_funding_agencies():
    pass

def get_all_projects():
    pass

def get_dates():
    pass

def get_boundaries():
    pass

def get_ships():
    pass

def get_sea_areas():
    pass

def get_package_descriptions():
    pass

if __name__ == '__main__':
    main()
