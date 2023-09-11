# coding = utf-8
from bs4 import BeautifulSoup
import xlsxwriter


def get_programs_info(soup_list, url):
    """
    Function that gets info of each program from a list of soups
    :param soup_list: List of HTMLs
    :param url: List of program urls
    :return: 2D list of title and description for each program
    """

    info = []  # init info list
    print(len(soup_list), len(url))
    for i in range(len(soup_list)):  # going through the HTML list
        soup_list[i] = soup_list[i] + '</html>'  # adding </html> at the end so bs4 can parse it
        soup_list[i] = BeautifulSoup(soup_list[i], 'html.parser')  # replacing each HTML with its parsed version

        program_name = str(soup_list[i].find(class_='col display-4 page-title'))[37:-5]  # find the program name, which is the display
        grad_unit_start = str(soup_list[i].find(class_='col-sm-4 order-1')).find('<h3>')
        grad_unit_end = str(soup_list[i].find(class_='col-sm-4 order-1')).find('</h3>')
        grad_unit = str(soup_list[i].find(class_='col-sm-4 order-1').h3)[4:-5].encode('ascii', 'ignore').decode('ascii').replace('\\xe2\\x80\\x8b', '')
                    #.encode('ascii', 'ignore').decode('utf-8')  # find the graduate unit, the first h3 of the side column
        dep_site = str(soup_list[i].find(class_='list-unstyled').findAll('a')[1].get(
            'href'))  # find the department site, which is the second url in the side column list
        fac_aff = str(soup_list[i].find(class_='col-sm-4 order-1').find('p'))[
                  3:-4]  # find faculty affiliation, which is the first paragraph in the side column

        campus = ''
        print(program_name, grad_unit)
        #print(soup_list[13], '\n')
        for j in range(1, len(soup_list[i].find(class_='col-sm-4 order-1').findAll('p'))):
            #print(i)
            if str(soup_list[i].find(class_='col-sm-4 order-1').findAll('p')[-j]).find('University of Toronto') != -1:  # find where university of toronto appears in the last paragraph, which is usually the address for the program
                #print(j)
                campus_pos1 = str(soup_list[i].find(class_='col-sm-4 order-1').findAll('p')[-j]).find(
                    'University of Toronto')
                campus_pos2 = str(soup_list[i].find(class_='col-sm-4 order-1').findAll('p')[-j])[campus_pos1:].find(
                    'br') + campus_pos1 - 1  # find where the br appears in the final paragraph, signalling the end
                # of the campus
                campus = str(soup_list[i].find(class_='col-sm-4 order-1').findAll('p')[-j])[
                         campus_pos1:campus_pos2]  # slicing the final paragraph to just grab the campus
                break

        # degree
        degrees_list = []  # init degree list
        degrees = soup_list[i].find(class_='col-sm-8 order-2').findAll('h3')  # finding all degree headers
        for j in range(len(degrees)):
            degrees_list.append(str(degrees[j])[4:-5])  # adding the degrees to a degree list
        # print(degrees_list)

        # program description
        degree_index = []  # init index for degree
        index = 0  # init dummy counter
        for j in range(len(degrees_list)):
            # adding the start index of each separate degree, with the dummy counter in place to skip past the
            # sections we already searched through including the h tags to avoid finding the degree name in the
            # paragraphs
            degree_index.append(
                str(soup_list[i].find(class_='col-sm-8 order-2'))[index:].find('<h3>' + degrees_list[j]) + index)
            index = degree_index[j]

        # grabbing the program overview by finding where the quick facts table starts
        qf_index = str(soup_list[i].find(class_='col-sm-8 order-2'))[0:degree_index[0]].find('<h2>Quick Facts</h2>')
        program_desc_text = str(soup_list[i].find(class_='col-sm-8 order-2'))[35:qf_index]

        degree_info_list = []  # init list of information for each degree
        degree_desc_index = []  # init list of indices for the program description
        degree_req_index = []  # init list of indices for the program requirements
        for j in range(len(degrees_list)):
            if j < len(degrees_list) - 1:
                # adding each degree section to an info list
                degree_info_list.append(
                    str(soup_list[i].find(class_='col-sm-8 order-2'))[degree_index[j]:degree_index[j + 1]])
            else:
                degree_info_list.append(str(soup_list[i].find(class_='col-sm-8 order-2'))[degree_index[j]:-1])

            # finding where the program description starts for each degree
            degree_desc_index.append(degree_info_list[j].find('<h4>Program Description</h4>'))
            # finding where the program requirements starts for each degree
            degree_req_index.append(degree_info_list[j].find('Requirements</h4'))
            # adding the sliced program description section to the program description text
            program_desc_text += degree_info_list[j][degree_desc_index[j]:degree_req_index[j]]

        # print(degree_index)
        # print(program_desc_index)
        # print(program_req_index)

        # changing the program description text into a soup object and then back to string with the get_text function
        # to get rid of the html stuffs
        program_desc = BeautifulSoup(program_desc_text, 'html.parser').get_text()
        print(type(program_desc))

        # finding the SGS calendar link for the program
        calendar_site = str(soup_list[i].find(class_='list-unstyled').find('a').get('href'))

        # adding the collected info to info
        info.append(
            [program_name, grad_unit, dep_site, url[i], calendar_site, fac_aff, campus, ', '.join(degrees_list),
             program_desc])

    return info


def write_sheet(name, title_list, title_width, info):
    """
    Function that writes stuff into Excel spreadsheet
    :param name: Name of spreadsheet to be created
    :param title_list: List of column titles
    :param title_width: List of column widths
    :param info: Info to be written
    :return: None
    """
    workbook = xlsxwriter.Workbook(name)  # creating workbook
    worksheet = workbook.add_worksheet()  # adding a sheet

    bold = workbook.add_format({'bold': True})  # bold formatting

    formatting(title_list, title_width, worksheet, bold)  # calling formatting

    for i in range(len(info)):  # going through info for each row
        for j in range(len(info[i])):  # going through info for each column of each row
            worksheet.write(i + 1, j, str(info[i][j]))  # writing it

    workbook.close()


def formatting(title_list, title_width, worksheet, bold):
    """
    Function that formats the spreadsheet with column titles and widths
    :param title_list: List of column titles
    :param title_width: List of column widths
    :param worksheet: Worksheet
    :param bold: Bold formatting
    :return: None
    """
    # setting headers and cell widths
    for i in range(len(title_list)):
        worksheet.set_column(i, i, title_width[i])
        worksheet.write(0, i, title_list[i], bold)


programs_list_file = open('SGS_programs.txt', 'r', encoding='utf-8')  # opening the cached HTMLs of each organization
programs_list = programs_list_file.read().split('</html>')  # reading it and splitting at the end of each HTML
programs_list.pop()  # removing the last element as it is just </html>
programs_list_file.close()

programs_cache_file = open('SGS_programs_cache.txt',
                           'r', encoding='utf-8')  # opening the url cache file for each organization, used to get the org url
url_list = programs_cache_file.read().splitlines()  # reading it and splitting the lines
programs_cache_file.close()
programs_info = get_programs_info(programs_list, url_list)  # getting org info

wb_name = 'Test - SGS Programs.xlsx'  # workbook name
titles = ['Program Name', 'Graduate Unit', 'Department Website', 'Calendar URL', 'SGS URL', 'Faculty Affiliation',
          'Campus', 'Degree(s)',
          'Program Description']  # list of column titles
col_width = [40, 40, 20, 20, 20, 40, 20, 40, 80]  # list of column widths
write_sheet(wb_name, titles, col_width, programs_info)  # writing sheet
