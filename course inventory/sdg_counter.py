import excelrd
import xlsxwriter
import numpy as np


def write_sheet(loc, name, title_list, title_width):
    """
    Function that writes all the rows that found a keyword into a new spreadsheet
    :param loc: Name of spreadsheet you want to search through
    :param name: What you want to name your new sorted spreadsheet
    :param title_list: The column labels for the original spreadsheet
    :param title_width: The column widths for the original spreadsheet
    :return: None. Creates a new spreadsheet of the results of the keyword search
    """
    wb = excelrd.open_workbook(loc)  # reading workbook
    sheet = wb.sheet_by_index(0)  # sheet index
    row_list = []  # init list of rows

    # adding each row of the spreadsheet minus the headers to our list
    for i in range(1, sheet.nrows):
        row_list.append(sheet.row_values(i))

    # loop that removes all extra whitespaces
    for i in range(len(row_list)):
        for j in range(len(row_list[i])):
            if type(row_list[i][j]) is str:  # check if the element is a string
                row_list[i][j] = ' '.join(row_list[i][j].split())  # join the string by single spaces

    workbook = xlsxwriter.Workbook(name)  # writing new workbook
    worksheet = workbook.add_worksheet()  # adding spreadsheet
    bold = workbook.add_format({'bold': True})  # bold formatting

    formatting(worksheet, title_list, title_width, bold)

    counts = count_sdg(row_list)
    print(counts)

    workbook.close()


def formatting(worksheet, title_list, title_width, bold):
    """
    Function that formats the new spreadsheet the same as the original one
    :param worksheet: New sheet to be created
    :param title_list: List of the column titles for the original spreadsheet
    :param title_width: List of the column widths for the original spreadsheet
    :param bold: Bold formatting
    :return: None
    """

    # setting headers and cell widths
    for i in range(len(title_list)):
        worksheet.set_column(i, i, title_width[i])
        worksheet.write(0, i, title_list[i], bold)

def count_sdg(rows):
    count_arr = np.zeros(16, dtype=int)
    for i in rows:
        sdgs = i[-1].split(', ')
        for j in sdgs:
            for k in range(16):
                if j.upper() == 'SDG' + str(k+1) or j.upper() == 'SDG ' + str(k+1) or j.upper() == 'SDG  ' + str(k+1):
                    count_arr[k] += 1
    return count_arr





sheet_path = "Course Inventory 2022-23.xlsx"  # name of full spreadsheet
#sheet_path = 'Classeur3.xlsx'
wb_name = 'TEST count SDGs.xlsx'  # new workbook name

# headings
titles = ['Course Code', 'Course Title', 'Course Description', 'Division', 'Unit', 'Keyword(s)', 'SDG(s)']
# width of headings
col_width = [20, 40, 80, 20, 20, 40, 20]

write_sheet(sheet_path, wb_name, titles, col_width)
