import excelrd
import xlsxwriter


def read_write(loc, name, title_list, title_width):
    """
    Function that reads the original spreadsheet, merges the duplicates, and writes to a new spreadsheet
    :param loc: File name of the original spreadsheet
    :param name: File name of the new spreadsheet
    :param title_list: List of the column titles for the original spreadsheet (not including SDG groups)
    :param title_width: List of column widths for the original spreadsheet
    :return: None. Creates a new spreadsheet with duplicates removed
    """
    wb = excelrd.open_workbook(loc)  # opening file
    sheet = wb.sheet_by_index(0)  # reading first sheet
    row_list = []  # init list of rows

    # adding each row to list
    for i in range(1, sheet.nrows):
        row_list.append(sheet.row_values(i))

    workbook = xlsxwriter.Workbook(name)  # writing new workbook
    worksheet = workbook.add_worksheet()  # adding spreadsheet

    bold = workbook.add_format({'bold': True})  # bold formatting
    cell_format = workbook.add_format()  # text warp format
    cell_format.set_text_wrap()

    formatting(worksheet, title_list, title_width, bold)

    newList = merge(row_list, sheet)

    for i in range(50):  # calling keepLast 15 times (if >15 dups, run it more times)
        keep_last(newList)

    SDGcovered(newList, sheet)

    # adding the newList into a spreadsheet
    for i in range(len(newList)):
        for j in range(sheet.ncols + 1):
            if j == sheet.ncols - 17 or j == sheet.ncols - 16:  # if its the 'keywords' or 'SDGs covered' column
                worksheet.write(i + 1, j, newList[i][j], cell_format)  # warp the text
            else:
                worksheet.write(i + 1, j, newList[i][j])

    workbook.close()


def formatting(worksheet, title_list, title_width, bold):
    """
    Function that formats the new spreadsheet according to original spreadsheet
    :param worksheet: New sheet to be created
    :param title_list: List of column titles of original spreadsheet (not including SDG groups)
    :param title_width: List of column widths of original spreadsheet
    :param bold: Bold formatting
    :return: None
    """
    # adding SDGs covered column
    title_list.append('SDGs Covered')
    title_width.append(18)
    # adding the SDG group columns
    for i in range(1, 17):
        title_list.append('SDG' + str(i))
        title_width.append(8)

    # setting headers and cell widths
    for i in range(len(title_list)):
        worksheet.set_column(i, i, title_width[i])
        worksheet.write(0, i, title_list[i], bold)


def merge(list, sheet):
    """
    Function that merges all the keywords for duplicate rows into the last duplicate
    :param list: List of rows of original spreadsheet
    :param sheet: Original sheet
    :return: newList - A list of all the rows with keywords merged
    """
    newList = []
    # checking if there are duplicates and merging keywords if there are any
    for i in range(len(list)):
        list[i].insert(sheet.ncols - 16, '')  # adding SDGs covered column
        if list[i][2] == list[i - 1][2]:  # if the one before it is a duplicate
            # if the SDG word was not already included
            if list[i][sheet.ncols - 17] not in list[i - 1][sheet.ncols - 17]:
                # add the SDG to the keywords list
                list[i][sheet.ncols - 17] = list[i][sheet.ncols - 17] + ', ' + list[i - 1][sheet.ncols - 17]
            '''else:  # if the keyword is already included
                list[i][sheet.ncols - 17] = list[i - 1][sheet.ncols - 17]  # keep the keywords as they are'''
            for j in range(16):  # for checking the SDGs covered
                if list[i][j + sheet.ncols - 15] == '':  # if its empty
                    # add whichever SDG group the previous keyowrd was part of
                    list[i][j + sheet.ncols - 15] = list[i][j + sheet.ncols - 15] + list[i - 1][j + sheet.ncols - 15]
            '''newList.append(list[i])  # add that row to a new list
        else:  # if its not a duplicate, no need to change anything
            newList.append(list[i])  # add that row to a new list'''
        newList.append(list[i])
    return newList


def keep_last(newList):
    """
    Function that removes all the duplicates except for the last one, which contains all of the keywords
    :param newList: Merged keywords list
    :return: newList - A list with all the duplicates removed
    """
    for i in range(len(newList)):
        if i < len(newList) - 1:
            if newList[i][2] == newList[i + 1][2]:  # if the row after it is the same
                newList.remove(newList[i])  # take the current row out
    return newList


def SDGcovered(newList, sheet):
    """
    Function that keeps track of which keyword groups were covered
    :param newList: Duplicates removed list
    :param sheet: Original sheet
    :return: newList - A list with an extra column after the "Keyword(s)" column that tracks SDG groups covered
    """
    for i in range(len(newList)):
        covered = ''  # init covered string
        for j in range(16):  # going through each of the SDG group columns
            if newList[i][j + sheet.ncols - 15] != '':  # if the column is not empty
                if covered == '':  # if its the first SDG group we find
                    covered = 'SDG' + str(j + 1)  # add it to the covered string
                else:  # if its not the first SDG group we find
                    covered = covered + ', SDG' + str(j + 1)  # add it to the covered string
            newList[i][sheet.ncols - 16] = covered  # set the string to the SDG covered column
    return newList


sheet_path = 'Test - Doctoral Keywords New.xlsx'  # name of SORTED spreadsheet to be highlighted
wb_name = 'Test - Doctoral Removed Duplicates New.xlsx'  # new workbook name
# headings
titles = ['Author', 'Advisor', 'Title', 'Department', 'Date issued', 'Abstract', 'Degree', 'Permanent URL', 'Subject',
          'Keyword(s)']
# width of headings
col_width = [43, 51, 38, 51, 11, 80, 24, 32, 80, 18]

read_write(sheet_path, wb_name, titles, col_width)
