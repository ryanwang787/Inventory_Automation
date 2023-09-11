import excelrd
import xlsxwriter


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
                row_list[i][j] = row_list[i][j].replace('\t', '')
    print(row_list[1][8].encode('utf-8'))
    print(row_list[1][8].encode('utf-8', 'ignore').decode('utf-8'.replace('\xc2', '')))

    workbook = xlsxwriter.Workbook(name)  # writing new workbook
    worksheet = workbook.add_worksheet()  # adding spreadsheet
    bold = workbook.add_format({'bold': True})  # bold formatting
    cell_format = workbook.add_format()  # text warp format
    cell_format.set_text_wrap()

    formatting(worksheet, title_list, title_width, bold)

    for i in range(len(row_list)):
        for j in range(sheet.ncols):
            if j == sheet.ncols - 18 or j == sheet.ncols - 17:  # if its the 'keywords' or 'SDGs covered' column
                worksheet.write(i + 1, j, row_list[i][j], cell_format)  # warp the text
            else:
                worksheet.write(i + 1, j, row_list[i][j])


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

    # adding the SDG group columns
    title_list.append('Keyword(s)')
    title_width.append(18)
    for i in range(1, 17):
        title_list.append('SDG' + str(i))
        title_width.append(8)
    # setting headers and cell widths
    for i in range(len(title_list)):
        worksheet.set_column(i, i, title_width[i])
        worksheet.write(0, i, title_list[i], bold)

sheet_path = "sgs programs - keep FINAL.xlsx"  # name of full spreadsheet
wb_name = 'sgs programs - keep FINAL FINAL.xlsx'  # new workbook name

# headings
titles = ['Program Name', 'Graduate Unit', 'Department Website', 'Calendar URL', 'SGS URL', 'Faculty Affiliation',
          'Campus', 'Degree(s)',
          'Program Description']
# list of column titles
col_width = [40, 40, 20, 20, 20, 40, 20, 40, 80]  # list of column widths

write_sheet(sheet_path, wb_name, titles, col_width)