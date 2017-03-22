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
    '''
    This will return a dictionary of all the people listed in the workbook.

    Columns 0-4 are each tied to one person per row.
    Columns 1 and 4 are not required (these are the middle name and email fields).

    `cell` objects accept row first and then column such as `sheet.cell(row,col).value`
    '''
    sheet = workbook.sheet_by_index(0)
    people = []
    for row in range(2,sheet.nrows):
        people.append(sheet.row_values(row,end_colx=6))
    return people

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
