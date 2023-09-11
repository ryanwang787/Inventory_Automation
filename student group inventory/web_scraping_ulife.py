from bs4 import BeautifulSoup
import xlsxwriter


def get_org_info(soup_list, url):
    """
    Function that gets info of each organization from a list of soups
    :param soup_list: List of HTMLs
    :param url: List of organization urls
    :return: 2D list of title and description for each organization
    """

    info = []  # init info list
    #print(len(soup_list), len(url))

    for i in range(len(soup_list)):  # going through the HTML list
        print('Scraping: ', url[i])
        soup_list[i] = soup_list[i] + '</html>'  # adding </html> at the end so bs4 can parse it
        soup_list[i] = BeautifulSoup(soup_list[i], 'html.parser')  # replacing each HTML with its parsed version
        #print(soup_list[0])
        # all club names are stored under the "detailHeading" class
        # all descriptions can be found with the paragraph attribute
        club_name = str(soup_list[i].find(class_='container mt-6 md:mt-12 text-5xl font-light flex flex-col md:flex-row items-start md:items-center gap-6'))[122:-542]  # grabbing the club name
        #print(club_name)
        desc_soup = soup_list[i].find_all('p') # club description paragraphs as soup objects
        desc_list = [i.get_text() for i in desc_soup] # converting to text
        #print(desc_soup)
        club_desc = ' '.join(desc_list[:-1])  # joining as a single string
        #print(club_desc)
        campus = str(soup_list[i].find(class_='inline-flex items-center text-sm bg-secondary font-bold px-2 py-2 gap-2 rounded text-white uppercase'))[513:-7]
        #print(campus)
        primary_contact = str(soup_list[i].find(class_='flex gap-4 border bg-white hover:bg-slate-50 transition-colors duration-300 rounded border-slate-200 p-2 items-center'))[463:-86]
        #print(primary_contact)
        interest_class_soup = soup_list[i].find(class_='list-none flex flex-wrap gap-4')
        interest_list_soup = interest_class_soup.find_all('li')
        interest_list = [i.get_text() for i in interest_list_soup]
        interest = ' '.join(interest_list)
        #interest_soup = str(soup_list[i].find(class_='list-none flex flex-wrap gap-4'))
        #print(interest, '\n\n')

        socials_soup = soup_list[i].find(class_='flex gap-4 mb-4')
        #print(socials_soup)
        facebook = ''
        twitter = ''
        instagram = ''
        if socials_soup is not None:
            #print(socials_soup)
            socials_list_soup = socials_soup.find_all('a', href=True)
            #print(socials_list_soup)
            socials_list = [i['href'] for i in socials_list_soup]

            for j in socials_list:
                if str(j)[12:13] == 'f':
                    facebook = j
                elif str(j)[12:13] == 't':
                    twitter = j
                elif str(j)[12:13] == 'i':
                    instagram = j
            #print(socials)

        email_soup = soup_list[i].find_all('a', href=True)
        email_list_soup = [i['href'] for i in email_soup]
        email = 'None'
        for j in email_list_soup:
            if str(j)[:6] == 'mailto':
                email = j

        #print(email)
        socials = 0
        # adding the collected info to info
        info.append([club_name, club_desc, url[i], primary_contact, email, campus, interest, instagram, facebook, twitter])
        print('Done scraping: ', url[i], '\n')

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


org_list_file = open('orgs_html_cache.txt', 'r', encoding='utf-8')  # opening the cached HTMLs of each organization
org_list = org_list_file.read().split('</html>')  # reading it and splitting at the end of each HTML
org_list.pop()  # removing the last element as it is just </html>
org_list_file.close()

org_cache_file = open('orgs_urls_cache.txt','r', encoding='utf-8')  # opening the url cache file for each organization, used to get the org url
url_list = org_cache_file.read().splitlines()  # reading it and splitting the lines
org_cache_file.close()

print(len(org_list), len(url_list))

org_info = get_org_info(org_list, url_list)  # getting org info

wb_name = 'Test - Org Info Part 4.xlsx'  # workbook name
titles = ['Club Name', 'Club Description', 'URL', 'Primary Contact', 'Club Email', 'Campus Association',
          'Areas of Interest', 'Instagram', 'Facebook', 'Twitter']  # list of column titles
col_width = [40, 160, 40, 40, 40, 40, 120, 40, 40, 40]  # list of column widths
write_sheet(wb_name, titles, col_width, org_info)  # writing sheet
