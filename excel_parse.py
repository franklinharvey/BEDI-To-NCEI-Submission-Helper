import xlrd

def main():
    workbook = xlrd.open_workbook("NCEI_test01.xlsx")
    print get_funding_agencies(workbook)
    print get_all_projects(workbook)
    print get_all_people(workbook)
    print get_all_variables(workbook)

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
    This will return an array of all the people listed in the workbook.

    Columns 0-5 are each tied to one person per row.
    Columns 1, 4, 5 are not required (these are the middle name and email fields).
    '''
    sheet = workbook.sheet_by_index(0)
    people = []
    for row in range(2,sheet.nrows):
        people.append(sheet.row_values(row,end_colx=6))
    return people

def get_all_variables(workbook):
    sheet = workbook.sheet_by_index(2)
    variables = []
    for row in range(2,sheet.nrows):
        variables.append(sheet.row_values(row))
    return variables

def get_funding_agencies(workbook):
    sheet = workbook.sheet_by_index(0)
    agencies = []
    for row in range(2,sheet.nrows):
        agencies.append(sheet.cell(row,6).value)
    return agencies

def get_all_projects(workbook):
    sheet = workbook.sheet_by_index(0)
    projects = []
    for row in range(2,sheet.nrows):
        projects.append(sheet.cell(row,7).value)
    return projects

def get_dates(workbook):
    sheet = workbook.sheet_by_index(1)
    return sheet.row_values(2,end_colx=2)

def get_boundaries(workbook):
    sheet = workbook.sheet_by_index(1)
    return sheet.row_values(2,start_colx=2,end_colx=6)

def get_ships(workbook):
    sheet = workbook.sheet_by_index(1)
    ships = []
    for row in range(2,sheet.nrows):
        ships.append(sheet.cell(row,6).value)
    return ships

def get_sea_areas(workbook):
    sheet = workbook.sheet_by_index(1)
    sea_areas = []
    for row in range(2,sheet.nrows):
        sea_areas.append(sheet.cell(row,7).value)
    return sea_areas

def get_package_descriptions(workbook):
    sheet = workbook.sheet_by_index(3)
    return sheet.row_values(2)

if __name__ == '__main__':
    main()
