import excelrd
import xlsxwriter


def write_sheet(loc1, loc2, name, title_list, title_width):
    wb_course_finder = excelrd.open_workbook(loc1)
    sheet_course_finder = wb_course_finder.sheet_by_index(0)
    wb_rosi = excelrd.open_workbook(loc2)  # reading workbook
    sheet_rosi = wb_rosi.sheet_by_index(0)  # sheet index
    course_finder_list = []
    rosi_list = []

    # adding each row of the spreadsheet minus the headers to our list
    for i in range(1, sheet_course_finder.nrows):
        course_finder_list.append(sheet_course_finder.row_values(i))

    for i in range(1, sheet_rosi.nrows):
        rosi_list.append(sheet_rosi.row_values(i))


    workbook = xlsxwriter.Workbook(name)  # writing new workbook
    worksheet_sesh = workbook.add_worksheet()  # adding spreadsheet
    worksheet_missed = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})  # bold formatting

    final_list = merge_sesh(course_finder_list, rosi_list)
    missed_list = missed(course_finder_list, rosi_list)

    for i in range(len(final_list)):
        for j in range(len(final_list[i])):
            worksheet_sesh.write(i + 1, j, final_list[i][j])
    for i in range(len(missed_list)):
        for j in range(len(missed_list[i])):
            worksheet_missed.write(i + 1, j, missed_list[i][j])

    formatting(worksheet_sesh, title_list, title_width, bold)
    formatting(worksheet_missed, title_list[:-1], title_width[:-1], bold)

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


def merge_sesh(course_finder, rosi):
    newList = []
    for i in course_finder:
        for j in rosi:
            if i[0] == j[2]:
                i.append(j[6])
                newList.append(i)

    return newList

def missed(cm, rosi):
    newList = []
    for i in cm:
        include = False
        for j in rosi:
            if i[0] == j[2]:
                include = True
        if not include:
            newList.append(i)

    return newList

sheet_path_rosi = "remove dups ROSI courses.xlsx"  # name of full spreadsheet
sheet_path_course_finder = "CECCS-2021-09-10.xlsx"
wb_name = 'TEST - Course Inv with Sessions.xlsx'  # new workbook name

# headings
titles = ['Course Code', 'Course Title', 'Course Description', 'Division', 'Unit', 'Campus', 'Credit Value', 'Offered Term']
# width of headings
col_width = [20, 40, 80, 40, 40, 20, 20, 20]


write_sheet(sheet_path_course_finder, sheet_path_rosi, wb_name, titles, col_width)
