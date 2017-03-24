import xlrd

def main():
    wb = xlrd.open_workbook("NCEI_test01.xlsx")
    # print get_all_people(workbook)
    # print get_all_variables(workbook)
    # print get_funding_agencies(workbook)
    # print get_all_projects(workbook)
    # print get_dates(workbook)
    # print get_boundaries(workbook)
    # print get_ships(workbook)
    # print get_sea_areas(workbook)
    # print get_package_descriptions(workbook)

    va = get_all_variables(wb)
    print filter_list(va)

def get_all_people(workbook):
    '''
    Input: workbook
    Output: nested array of people

    Columns 0-5 are each tied to one person per row.
    Columns 1, 4, 5 are not required (these are the middle name, email, and institution fields).
    '''
    sheet = workbook.sheet_by_index(0)
    people = []
    for row in range(2,sheet.nrows):
        people.append(sheet.row_values(row, end_colx=6))
    return people

def get_all_variables(workbook):
    sheet = workbook.sheet_by_index(2)
    variables = []
    for row in range(2,sheet.nrows):
        variables.append(sheet.row_values(row))
    return filter_list(variables)

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

def get_headers(sheet):
    return sheet.row_values(1)

def get_explanations(sheet):
    return sheet.row_values(0)

def filter_list(itemList):
    for i, item in enumerate(itemList):
        if itemList[i]:
            pass
        else:
            print "HERE"
    return itemList


if __name__ == '__main__':
    main()
