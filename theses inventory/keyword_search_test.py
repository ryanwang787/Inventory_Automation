import excelrd
import xlsxwriter


def write_sheet(loc, name, title_list, title_width, kl, gl):
    """
    Function that writes all the rows that found a keyword into a new spreadsheet
    :param loc: Name of spreadsheet you want to search through
    :param name: What you want to name your new sorted spreadsheet
    :param title_list: The column labels for the original spreadsheet
    :param title_width: The column widths for the original spreadsheet
    :param kl: Keyword list you want to search through
    :param gl: The SDG groups with each keyword
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
    search(row_list, sheet, worksheet, kl, gl)

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

    # adding the keywords column
    title_list.append('Keyword(s)')
    title_width.append(18)
    # adding the SDG groups columns
    title_list.append('SDGs Covered')
    title_width.append(18)
    # setting headers and cell widths
    for i in range(len(title_list)):
        worksheet.set_column(i, i, title_width[i])
        worksheet.write(0, i, title_list[i], bold)


def search(list, sheet, worksheet, keyword, SDGs):
    """
    Function that searches through the cells of the original spreadsheet for the keywords
    :param list: Rows of the original spreadsheet
    :param sheet: Original sheet
    :param worksheet: New spreadsheet
    :param keyword: Keyword list
    :param SDGs: SDG group dictionary
    :return: None. Writes the rows that found a keyword to the new sheet, including which keyword was found
    """

    row = 1  # init row
    for i in list:
        SDG_covered = []
        keywords = []
        write = False
        for j in range(len(keyword)):
            if keyword[j] in str(i[2]).lower() or keyword[j] in str(i[5]).lower() or keyword[j] in str(i[8]).lower():
                write = True

                if keyword[j] not in keywords:  # checking for no duplicate words
                    keywords.append(keyword[j])  # adding the found keyword to a keyword list
                for m in range(sheet.ncols):  # for every column in that row
                    if m == 4:  # if its the date
                        try:  # try formatting the date to excel date format
                            datetime = excelrd.xldate_as_datetime(i[m],
                                                                  0)  # turns the excel date into a datetime object
                            date_obj = datetime.date()  # datetime object into a date object (year, month, day)
                            str_date = date_obj.isoformat()  # writes date object to string in the form of YYYY-MM-DD
                            worksheet.write(row, m, str_date)  # write in the date
                        except:  # if it doesn't wanna be formatted (only YYYY-MM), keep it as it is
                            worksheet.write(row, m, i[m])
                    else:
                        worksheet.write(row, m, i[m])  # write in the copied info for each cell

                for n in range(1, 17):  # adding SDG groups that the found keyword appears in
                    if keyword[j] in SDGs['SDG' + str(n)] and 'SDG' + str(n) not in SDG_covered:
                        SDG_covered.append('SDG' + str(n))
        if write:
            worksheet.write(row, sheet.ncols, ', '.join(keywords))  # add the keywords that were used
            worksheet.write(row, sheet.ncols + 1, ', '.join(SDG_covered))  # add the SDG groups that were covered
            row += 1


sheet_path = "master theses.xlsx"  # name of full spreadsheet
wb_name = 'Test - Masters Keywords.xlsx'  # new workbook name

# headings
titles = ['Author', 'Advisor', 'Title', 'Department', 'Date issued', 'Abstract', 'Degree', 'Permanent URL', 'Subject']
# width of headings
col_width = [43, 51, 38, 51, 11, 80, 24, 32, 80]

# list of SDG keywords
keyword_list = ['poverty', 'income distribution', 'wealth distribution', 'socioeconomic', 'socio-economic',
                'socio economic',
                'agricultur', 'food', 'nutrition', 'health', 'well being', 'wellbeing', 'well-being', 'educat',
                'inclusiv',
                'equitable', 'gender', 'women', 'equality', 'girl', 'queer', 'water', 'sanita', 'energy',
                'renewabl',
                'wind', 'solar', 'geothermal', 'hydroelectric', 'employment', 'economic growth',
                'sustainable development',
                'labour', 'labor', 'worker', 'wage', 'infrastructure', 'innovat', 'industr', 'buildings', 'trade',
                'inequality',
                'financial market', 'taxation', 'cities', 'urban', 'resilien', 'rural', 'consum', 'production', 'waste',
                'natural resource', 'recycl', 'industrial ecology', 'sustainable design', 'climate', 'greenhouse gas',
                'environment', 'global warming', 'weather', 'ocean', 'marine', 'water', 'pollut', 'conserv', 'fish',
                'forest', 'biodivers', 'ecolog', 'pollut', 'conserv', 'land use', 'institut', 'justice',
                'governance', 'peace', 'rights']

# dict of all SDG keywords and which SDG group they belong to
SDG_groups = {
    'SDG1': ['poverty', 'income distribution', 'wealth distribution', 'socioeconomic', 'socio-economic',
             'socio economic'],
    'SDG2': ['agricultur', 'food', 'nutrition'],
    'SDG3': ['health', 'well being', 'wellbeing', 'well-being'],
    'SDG4': ['educat', 'inclusiv', 'equitable'],
    'SDG5': ['gender', 'women', 'equality', 'girl', 'queer'],
    'SDG6': ['water', 'sanita'],
    'SDG7': ['energy', 'renewabl', 'wind', 'solar', 'geothermal', 'hydroelectric'],
    'SDG8': ['employment', 'economic growth', 'sustainable development', 'labour', 'labor', 'worker', 'wage'],
    'SDG9': ['infrastructure', 'innovat', 'industr', 'buildings'],
    'SDG10': ['trade', 'inequality', 'financial market', 'taxation'],
    'SDG11': ['cities', 'urban', 'resilien', 'rural'],
    'SDG12': ['consum', 'production', 'waste', 'natural resource', 'recycl', 'industrial ecology',
              'sustainable design'],
    'SDG13': ['climate', 'greenhouse gas', 'environment', 'global warming', 'weather'],
    'SDG14': ['ocean', 'marine', 'water', 'pollut', 'conserv', 'fish'],
    'SDG15': ['forest', 'biodivers', 'ecolog', 'pollut', 'conserv', 'land use'],
    'SDG16': ['institut', 'justice', 'governance', 'peace', 'rights']
}

write_sheet(sheet_path, wb_name, titles, col_width, keyword_list, SDG_groups)
