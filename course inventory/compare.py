import excelrd
import xlsxwriter


def write_sheet(loc1, loc2, loc3, rosi_name, cm_name):
    wb_rosi = excelrd.open_workbook(loc1)  # reading workbook
    sheet_rosi = wb_rosi.sheet_by_index(0)  # sheet index
    wb_cm = excelrd.open_workbook(loc2)
    sheet_cm = wb_cm.sheet_by_index(0)
    wb_combined = excelrd.open_workbook(loc3)
    sheet_combined = wb_combined.sheet_by_index(0)
    rosi_list = []
    cm_list = []
    combined_list = []

    # adding each row of the spreadsheet minus the headers to our list
    for i in range(1, sheet_rosi.nrows):
        rosi_list.append(sheet_rosi.row_values(i))

    for i in range(1, sheet_cm.nrows):
        cm_list.append(sheet_cm.row_values(i))

    for i in range(1, sheet_combined.nrows):
        combined_list.append(sheet_combined.row_values(i))

    # loop that removes all extra whitespaces
    for i in range(len(rosi_list)):
        for j in range(len(rosi_list[i])):
            if type(rosi_list[i][j]) is str:  # check if the element is a string
                rosi_list[i][j] = ' '.join(rosi_list[i][j].split())  # join the string by single spaces

    for i in range(len(cm_list)):
        for j in range(len(cm_list[i])):
            if type(cm_list[i][j]) is str:  # check if the element is a string
                cm_list[i][j] = ' '.join(cm_list[i][j].split())  # join the string by single spaces

    for i in range(len(combined_list)):
        for j in range(len(combined_list[i])):
            if type(combined_list[i][j]) is str:  # check if the element is a string
                combined_list[i][j] = ' '.join(combined_list[i][j].split())  # join the string by single spaces

    workbook_rosi = xlsxwriter.Workbook(rosi_name)  # writing new workbook
    worksheet_rosi = workbook_rosi.add_worksheet()  # adding spreadsheet
    workbook_cm = xlsxwriter.Workbook(cm_name)  # writing new workbook
    worksheet_cm = workbook_cm.add_worksheet()  # adding spreadsheet
    # bold = workbook.add_format({'bold': True})  # bold formatting

    final_list = compare(combined_list, rosi_list, cm_list)

    for i in range(len(final_list[0])):
        for j in range(len(final_list[0][i])):
            worksheet_rosi.write(i + 1, j, final_list[0][i][j])

    for i in range(len(final_list[1])):
        for j in range(len(final_list[1][i])):
            worksheet_cm.write(i + 1, j, final_list[1][i][j])

    # formatting(worksheet, title_list, title_width, bold)

    workbook_rosi.close()
    workbook_cm.close()


'''def formatting(worksheet, title_list, title_width, bold):
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
        worksheet.write(0, i, title_list[i], bold)'''


def compare(combined, rosi, cm):
    rosi_list = []
    cm_list = []

    in_rosi = False
    in_cm = False
    for i in combined:
        for j in rosi:
            if i[2] == j[2]:
                rosi.remove(j)
        for j in cm:
            if i[2] == j[0]:
                cm.remove(j)
        '''if not in_rosi:
            rosi_list.append(i)
        if not in_cm:
            cm_list.append(i)
        in_rosi = False
        in_cm = False'''
    print(len(rosi), len(cm))
    return [rosi, cm]


sheet_path_rosi = 'remove dups ROSI courses.xlsx'  # name of full spreadsheet
sheet_path_cm = 'remove dups CM courses.xlsx'
sheet_path_combined = 'Combined sheets.xlsx'
wb_name_rosi = 'extra rosi.xlsx'  # new workbook name
wb_name_cm = 'extra cm.xlsx'

# headings
# titles = ['Author', 'Advisor', 'Title', 'Department', 'Date issued', 'Abstract', 'Degree', 'Permanent URL', 'Subject']
# width of headings
# col_width = [43, 51, 38, 51, 11, 80, 24, 32, 80]


write_sheet(sheet_path_rosi, sheet_path_cm, sheet_path_combined, wb_name_rosi, wb_name_cm)
