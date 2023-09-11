import xlsxwriter


def write_sheet(file_name_list, wb_name_list, title_list, title_width):
    """
    Function that writes list(s) to spreadsheet(s)
    :param file_name_list: The list of file names you want to write to spreadsheet
    :param wb_name_list: Name of spreadsheets you want to create
    :param title_list: Column titles for the spreadsheets you want to create (not including SDG groups)
    :param title_width: Column widths for the spreadsheets you want to create (not including SDG groups)
    :return: None
    """
    workbook = []  # init workbook list
    worksheet = []  # init worksheet list

    for i in range(len(wb_name_list)):  # adding workbook and worksheet for each spreadsheet to their respective lists
        workbook.append(xlsxwriter.Workbook(wb_name_list[i]))  # writing new workbook
        worksheet.append(workbook[i].add_worksheet())  # adding spreadsheet

    formatting(file_name_list, worksheet, workbook, title_list, title_width)  # calling formatting function

    newList = read_file(file_name_list)  # calling read_file function
    # newList is now a 3D list with the first layer being keep, remove, or saved, each of those containing rows
    # and each row containing columns

    for i in range(len(worksheet)):  # going through each worksheet
        cell_format = workbook[i].add_format()  # text warp format
        cell_format.set_text_wrap()

        if i == 1:  # if its the remove worksheet
            for j in range(len(newList[1])):  # goes through the removed rows of the newList
                # adds the columns up to 'Keyword(s)' and the removal reason to a new list
                myList = newList[1][j][:9] + newList[1][j][-1:]
                for k in range(len(myList)):  # runs through each element of that new list
                    if k == 7 or k == 8:  # if its the keywords or removal reason column
                        worksheet[i].write(j + 1, k, myList[k], cell_format)  # write it in with wrapped text
                    else:
                        worksheet[i].write(j + 1, k, myList[k])  # write the columns in the spreadsheet
        else:  # if its not the remove worksheet i.e. its the keep or saved worksheet
            for j in range(len(newList[i])):  # goes through the rows of the lists
                for k in range(len(newList[i][j])):  # goes through the columns of the rows
                    if k == 7 or k == 8:  # if its the keywords or SDG covered column
                        worksheet[i].write(j + 1, k, newList[i][j][k], cell_format)  # write it in with wrapped text
                    else:
                        worksheet[i].write(j + 1, k, newList[i][j][k])  # write the columns in the spreadsheet

        workbook[i].close()


def file_len(file_name):
    """
    Function that finds the number of rows in a file
    :param file_name: File name
    :return: Int
    """
    f = open(file_name, encoding='utf-8')  # opening the file
    index = 0  # init index variable
    for index, line in enumerate(f):  # goes through every index and line in f
        pass
    f.close()
    return int(
        (index + 1) / 2)  # returns the index + 1 to account for starting the index at 0, /2 to account for empty lines


def read_file(file_name_list):
    """
    Function that reads the files from a list and adds all of them to a 3D list
    :param file_name_list: List of the file names to be read
    :return: List
    """
    file_list = []  # init file list list
    line_count = []  # init line count list
    for i in range(len(file_name_list)):  # for each file
        file_list.append(open(file_name_list[i], 'r', encoding='utf-8'))  # read the file
        line_count.append(file_len(file_name_list[i]))  # count the number of lines in that file

    myList = []  # init temporary list

    for i in range(len(line_count)):  # goes through the number of rows in each file
        myList.append(file_list[i].read().split('@%&'))  # adds the rows of each file split at the '@%&' symbol, which
        # divides each column of each row
        file_list[i].close()

    for i in range(len(myList)):  # going through the temp list and replacing every '\n' with an empty string
        for j in range(len(myList[i])):
            myList[i][j] = myList[i][j].replace('\n', '')

    newList = []  # temp list to remove the empty lists inside of myList
    for i in range(len(myList)):
        if myList[i] != ['']:  # removing empty lists from myList
            newList.append(myList[i])

    final_list = [[], [], []]  # final 3D list to be returned
    # newList is currently formatted as a 2D list, with newList[0] being a string of the rows of the keep file,
    # newList[1] being a string of the rows of the remove file, and newList[2] being a string of the rows of the
    # saved file
    # This loop will go through each of those and separate the rows into their individual lists of columns, and each of
    # the rows will be added to a list of keep, remove, or saved rows
    for i in range(len(newList)):  # going through newList
        for j in range(line_count[i]):  # going through each row of each file
            final_list[i].append(
                newList[i][int(26 * j):int(26 * (j + 1))])  # adding the rows of each file to the final list

    return final_list  # returning the final 3D list


def formatting(file_name_list, worksheet_list, workbook_list, title_list, title_width):
    """
    Function that formats new spreadsheets
    :param file_name_list: List of file names
    :param worksheet_list: List of worksheets
    :param workbook_list: List of workbooks
    :param title_list: List of column titles (minus SDG groups)
    :param title_width: List of column widths (minus SDG groups)
    :return: None
    """
    # adding SDGs covered column
    title_list.append('SDGs Covered')
    title_width.append(18)
    # adding the SDG group columns
    for j in range(1, 17):
        title_list.append('SDG' + str(j))
        title_width.append(8)
    # setting headers and cell widths
    for i in range(len(file_name_list)):
        bold = workbook_list[i].add_format({'bold': True})  # bold formatting

        if i == 1:  # if its the remove spreadsheet
            for j in range(len(title_list) - 17):  # not including the SDG groups
                worksheet_list[i].set_column(j, j, title_width[j])
                worksheet_list[i].write(0, j, title_list[j], bold)
            # adding extra column for removal reason
            worksheet_list[i].set_column(len(title_list) - 17, len(title_list) - 17, 40)
            worksheet_list[i].write(0, len(title_list) - 17, 'Removal Reason', bold)
        else:
            for j in range(len(title_list)):  # setting column title and widths
                worksheet_list[i].set_column(j, j, title_width[j])
                worksheet_list[i].write(0, j, title_list[j], bold)


file_paths = ['keep_list.txt', 'del_list.txt', 'save_list.txt']  # list of file names
wb_names = ['test missed courses - keep.xlsx', 'test missed courses - deleted.xlsx', 'test missed courses - saved.xlsx']  # list of spreadsheet names

# headings
titles = ['Course Code', 'Course Title', 'Course Description', 'Division', 'Unit', 'Campus', 'Credit Value', 'Keyword(s)']
# width of headings
col_width = [20, 40, 80, 40, 40, 20, 20, 30]

write_sheet(file_paths, wb_names, titles, col_width)
