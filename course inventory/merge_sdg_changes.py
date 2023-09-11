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

    workbook = xlsxwriter.Workbook(name)  # writing new workbook
    worksheet = workbook.add_worksheet()  # adding spreadsheet
    bold = workbook.add_format({'bold': True})  # bold formatting

    formatting(worksheet, title_list, title_width, bold)

    sdgs = combine_sdg(row_list)

    # write in new changes to new sheet
    for i in range(len(row_list)):
        for j in range(len(row_list[i]) - 3):
            worksheet.write(i + 1, j, row_list[i][j])
        worksheet.write(i + 1, len(row_list[i]) - 3, sdgs[i])

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

def combine_sdg(rows):
    sdg_list = []
    for i in rows:
        temp_sdg = ''
        for j in range(len(i)):
            if j == 8: # SDGs
                if len(str(i[j])) == 1:
                    i[j] = 'SDG' + str(i[j])
                temp_sdg += str(i[j]).upper()
            elif j == 9: # removed sdg
                rem_sdg = str(i[j]).upper().split(',') # splitting sdg into list
                rem_sdg = [k.replace('.0', '') for k in rem_sdg] # removing any floats
                for k in range(len(rem_sdg)): # checking if its in single number format and correcting it
                    if rem_sdg[k] != '' and len(rem_sdg[k]) == 1:
                        rem_sdg[k] = 'SDG' + rem_sdg[k]
                print(temp_sdg)
                print(rem_sdg)
                print(i)
                for k in rem_sdg:

                    temp_sdg = temp_sdg.replace(k, '') # replacing removed sdg with nothing
                    temp_sdg = temp_sdg.replace(', , ', ', ') # replacing all empty comma spaces left behind
                    if temp_sdg != '':
                        if temp_sdg[0] == ',': # if comma space is in front
                            temp_sdg = temp_sdg[2:]
                        if temp_sdg[-1] == ',': # if comma space is at the back
                            temp_sdg = temp_sdg[:-1]
                #print(temp_sdg)
            elif j == 10: # added sdg
                keep_sdg = str(i[j]).upper().split(',')
                keep_sdg = [k.replace('.0', '') for k in keep_sdg]
                for k in range(len(keep_sdg)):
                    if i[j] != '' and len(keep_sdg[k]) == 1:
                        keep_sdg[k] = 'SDG' + keep_sdg[k]
                    if i[j] !='':
                        temp_sdg += ', '
                        temp_sdg += keep_sdg[k].upper()

        sdg_list.append(temp_sdg)

    return sdg_list




sheet_path = "keyword_search_result_to_manually_review (1).xlsx"  # name of full spreadsheet
#sheet_path = 'Classeur3.xlsx'
wb_name = 'TEST combined SDGs.xlsx'  # new workbook name

# headings
titles = ['Course Code', 'Course Title', 'Course Description', 'Division', 'Unit', 'Keep (K) or Remove (R)', 'Removal Reason', 'Keyword(s)', 'SDG(s)']
# width of headings
col_width = [20, 40, 80, 40, 40, 20, 20, 40, 20]

write_sheet(sheet_path, wb_name, titles, col_width)
