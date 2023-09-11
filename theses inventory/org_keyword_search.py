import xlrd
import xlsxwriter


# import pandas as pd


def write_sheet(loc, wb_name, titles, widths, kl, gl):
    """
    :param loc: Name of spreadsheet you want to search through
    :param wb_name: What you want to name your new sorted spreadsheet
    :param titles: The column labels for the original spreadsheet
    :param widths: The column widths for the original spreadsheet
    :param kl: Keyword list you want to search through
    :param gl: The SDG groups with each keyword
    :return: None. Creates a new spreadsheet of the results of the keyword search
    """

    wb = xlrd.open_workbook(loc)  # reading workbook
    sheet = wb.sheet_by_index(0)  # sheet index
    list = []  # init list of rows

    # adding each row of the spreadsheet minus the headers to our list
    for i in range(1, sheet.nrows):
        list.append(sheet.row_values(i))

    workbook = xlsxwriter.Workbook(wb_name)  # writing new workbook
    worksheet = workbook.add_worksheet()  # adding spreadsheet

    bold = workbook.add_format({'bold': True})  # bold formatting

    formatting(titles, widths, sheet, worksheet, bold)
    search(list, sheet, worksheet, kl, gl)

    workbook.close()


def formatting(title_list, title_width, sheet, worksheet, bold):
    """ Formatting the new sheet """

    # adding the SDG group columns
    title_list.append('Keyword(s)')
    title_width.append(18)
    for i in range(1, 17):
        title_list.append('SDG' + str(i))
        title_width.append(8)
    # setting headers and cell widths
    for i in range(sheet.ncols + 17):
        worksheet.set_column(i, i, title_width[i])
        worksheet.write(0, i, title_list[i], bold)


def search(list, sheet, worksheet, keyword, SDGs):
    """ Searches through a given list for a given set of keywords and writes it to a new worksheet """

    row = 1  # init row
    # run through the list of keywords and check if the keywords appear in our rows
    for i in range(len(keyword)):
        for j in list:  # taking each row
            write = False  # writing boolean

            # checking if keyword is in title, abstract, or subject
            '''if keyword[i] in str(j[2]).lower() or keyword[i] in str(j[5]).lower() or keyword[i] in str(j[8]).lower():
                write = True'''

            for k in j:  # taking each cell in that row
                if keyword[i] in str(k).lower():  # checking if the keyword appears in the cell
                    write = True  # if it does, set write to True

            # if the keyword appears in that row, we want to copy that row into our new spreadsheet
            if write:
                for m in range(sheet.ncols):  # for every column in that row
                    if m == 4:  # if its the date
                        try:  # try formatting the date to excel date format
                            datetime_date = xlrd.xldate_as_datetime(j[m], 0)
                            date_object = datetime_date.date()
                            string_date = date_object.isoformat()
                            worksheet.write(row, m, string_date)  # write in the date
                        except:  # if it doesn't wanna be formatted, keep it as it is
                            worksheet.write(row, m, j[m])
                    else:
                        worksheet.write(row, m, j[m])  # write in the copied info for each cell
                worksheet.write(row, sheet.ncols, keyword[i])  # add the keyword that was used

                for n in range(1, 17):  # adding the 'X' marking for which SDG group the keyword belonged to
                    if keyword[i] in SDGs['SDG' + str(n)]:
                        worksheet.write(row, sheet.ncols + n, "X")

                row += 1  # moving to next row


sheet_path = "Test - Orgs.xlsx"  # name of full spreadsheet
wb = 'Test - Orgs Keywords.xlsx'  # new workbook name

# headings
titles = ['Club Name', 'Club Description', 'URL']
# width of headings
col_width = [40, 160, 40]

# list of all SDG keywords
keyword_list = ['poverty', 'income distribution', 'wealth distribution', 'socioeconomic', 'socio-economic',
                'socio economic',
                'agriculture', 'food', 'nutrition', 'health', 'well being', 'wellbeing', 'well-being', 'educat',
                'inclusiv',
                'equitable', 'gender', 'women', 'equality', 'girl', 'queer', 'water', 'sanitation', 'energy',
                'renewable',
                'wind', 'solar', 'geothermal', 'hydroelectric', 'employment', 'economic growth',
                'sustainable development',
                'labour', 'worker', 'wage', 'infrastructure', 'innovation', 'industr', 'buildings', 'trade',
                'inequality',
                'financial market', 'taxation', 'cities', 'urban', 'resilien', 'rural', 'consum', 'production', 'waste',
                'natural resources', 'recycl', 'industrial ecology', 'sustainable design', 'climate', 'greenhouse gas',
                'environment', 'global warming', 'weather', 'ocean', 'marine', 'water', 'pollut', 'conserv', 'fish',
                'forest', 'biodiversity', 'ecology', 'pollut', 'conserv', 'land use', 'institution', 'justice',
                'governance', 'peace', 'rights']

# dict of all SDG keywords and which SDG group they belong to
SDG_groups = {
    'SDG1': ['poverty', 'income distribution', 'wealth distribution', 'socioeconomic', 'socio-economic',
             'socio economic'],
    'SDG2': ['agriculture', 'food', 'nutrition'],
    'SDG3': ['health', 'well being', 'wellbeing', 'well-being'],
    'SDG4': ['educat', 'inclusiv', 'equitable'],
    'SDG5': ['gender', 'women', 'equality', 'girl', 'queer'],
    'SDG6': ['water', 'sanitation'],
    'SDG7': ['energy', 'renewable', 'wind', 'solar', 'geothermal', 'hydroelectric'],
    'SDG8': ['employment', 'economic growth', 'sustainable development', 'labour', 'worker', 'wage'],
    'SDG9': ['infrastructure', 'innovation', 'industr', 'buildings'],
    'SDG10': ['trade', 'inequality', 'financial market', 'taxation'],
    'SDG11': ['cities', 'urban', 'resilien', 'rural'],
    'SDG12': ['consum', 'production', 'waste', 'natural resources', 'recycl', 'industrial ecology',
              'sustainable design'],
    'SDG13': ['climate', 'greenhouse gas', 'environment', 'global warming', 'weather'],
    'SDG14': ['ocean', 'marine', 'water', 'pollut', 'conserv', 'fish'],
    'SDG15': ['forest', 'biodiversity', 'ecology', 'pollut', 'conserv', 'land use'],
    'SDG16': ['institution', 'justice', 'governance', 'peace', 'rights']
}

write_sheet(sheet_path, wb, titles, col_width, keyword_list, SDG_groups)
