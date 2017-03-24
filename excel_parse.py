import xlrd

def main():
    log = ""
    workbook = xlrd.open_workbook("NCEI_test01.xlsx")
    log += "{}\n".format(filter_list(get_all_people(workbook)))
    log += "{}\n".format(filter_list(get_funding_agencies(workbook)))
    log += "{}\n".format(filter_list(get_all_projects(workbook)))
    log += "{}\n".format(filter_list(get_all_variables(workbook)))
    log += "{}\n".format(get_dates(workbook))
    log += "{}\n".format(get_boundaries(workbook))
    log += "{}\n".format(filter_list(get_ships(workbook)))
    log += "{}\n".format(filter_list(get_sea_areas(workbook)))
    log += "{}\n".format(filter_list(get_package_descriptions(workbook)))

    with open("log.txt", 'w+') as logFile:
        logFile.write(log)

def get_all_people(workbook):
    '''
    Input: workbook
    Output: array of people (nested if multiple rows)

    Columns 0-5 are each tied to one person per row.
    Columns 1, 4, 5 are not required.
    '''
    sheet = workbook.sheet_by_index(0)
    people = []
    for row in range(2,sheet.nrows):
        people.append(sheet.row_values(row, end_colx=6))
    return people

def get_all_variables(workbook):
    '''
    Input: workbook
    Output: array of variables (nested if multiple rows)
    '''
    sheet = workbook.sheet_by_index(2)
    variables = []
    for row in range(2,sheet.nrows):
        variables.append(sheet.row_values(row))
    return variables

def get_funding_agencies(workbook):
    '''
    Input: workbook
    Output: array of funding agencies
    '''
    sheet = workbook.sheet_by_index(0)
    agencies = []
    for row in range(2,sheet.nrows):
        agencies.append(sheet.cell(row,6).value)
    return agencies

def get_all_projects(workbook):
    '''
    Input: workbook
    Output: array of projects
    '''
    sheet = workbook.sheet_by_index(0)
    projects = []
    for row in range(2,sheet.nrows):
        projects.append(sheet.cell(row,7).value)
    return projects

def get_dates(workbook):
    '''DOES NOT WORK'''
    sheet = workbook.sheet_by_index(1)
    return sheet.row_values(2,end_colx=2)

def get_boundaries(workbook):
    '''
    Input: workbook
    Output: array of boundaries
    '''
    sheet = workbook.sheet_by_index(1)
    return sheet.row_values(2,start_colx=2,end_colx=6)

def get_ships(workbook):
    '''
    Input: workbook
    Output: array of ships
    '''
    sheet = workbook.sheet_by_index(1)
    ships = []
    for row in range(2,sheet.nrows):
        ships.append(sheet.cell(row,6).value)
    return ships

def get_sea_areas(workbook):
    '''
    Input: workbook
    Output: array of sea areas
    '''
    sheet = workbook.sheet_by_index(1)
    sea_areas = []
    for row in range(2,sheet.nrows):
        sea_areas.append(sheet.cell(row,7).value)
    return sea_areas

def get_package_descriptions(workbook):
    '''
    Input: workbook
    Output: array of package description elements
    '''
    sheet = workbook.sheet_by_index(3)
    return sheet.row_values(2)

def get_headers(sheet):
    '''
    Input: sheet
    Output: array of headers
    '''
    return sheet.row_values(1)

def get_explanations(sheet):
    '''
    Input: sheet
    Output: array of header explanations
    '''
    return sheet.row_values(0)

def filter_list(itemList, replace=False, test=False):
    '''
    Input: list
    Output: list with empty elements removed
    '''
    nested = any(isinstance(i, list) for i in itemList)
    if test:
        print "Input list: {}".format(itemList)
        print "Is nested: {}".format(nested)
    if nested: # if nested
        for subList in itemList:
            for item in subList:
                if item:
                    pass
                else:
                    subList.remove(item)
    else: # if not
        itemList[:] = [x for x in itemList if not determine(x)]

    if test:
        print "Output list: {}".format(itemList)
    return itemList

def determine(element):
    '''
    Input: element
    Output: Boolean
    '''
    if element: #checking to see if element needs to be deleted
        return False
    else:
        return True

if __name__ == '__main__':
    main()
