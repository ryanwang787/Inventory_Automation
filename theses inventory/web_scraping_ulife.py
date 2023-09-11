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
    print(len(soup_list), len(url))
    for i in range(len(soup_list)):  # going through the HTML list
        soup_list[i] = soup_list[i] + '</html>'  # adding </html> at the end so bs4 can parse it
        soup_list[i] = BeautifulSoup(soup_list[i], 'html.parser')  # replacing each HTML with its parsed version

        # all club names are stored under the "detailHeading" class
        # all descriptions can be found with the paragraph attribute
        club_name = str(soup_list[i].find(class_='detailHeading'))[26:-5]  # grabbing the club name
        desc_soup = soup_list[i].find_all('p')[:-2] # club description paragraphs as soup objects
        desc_list = [i.get_text() for i in desc_soup] # converting to text
        club_desc = ', '.join(desc_list)  # joining as a single string

        dd_list = soup_list[i].find_all('dd')  # searching for all dd attributes and adding them to list
        dl_list = soup_list[i].find_all(class_='infoDl')  # searching for the infoDl class and adding them to list

        dd_list = [str(j) for j in dd_list]  # turning the dd list into strings

        dl_list = [str(j) for j in dl_list]  # turning the dt list into strings

        # In this section we check if each of these categories exist on the organization's site
        # If they do exist, we set their count to 1. If not, their counts are set to 0.
        # This will help us index the dd list later when we are trying to grab info from it
        if dl_list[0].find('Primary Contact') == -1:
            primary_count = 0
        else:
            primary_count = 1
        if dl_list[0].find('Secondary Contact') == -1:
            secondary_count = 0
        else:
            secondary_count = 1
        if dl_list[0].find('Telephone') == -1:
            telephone_count = 0
        else:
            telephone_count = 1
        if dl_list[0].find('Group Email Address') == -1:
            email_count = 0
        else:
            email_count = 1
        if dl_list[0].find('Website') == -1:
            website_count = 0
        else:
            website_count = 1
        if dl_list[0].find('Mailing Address') == -1:
            mailing_count = 0
        else:
            mailing_count = 1

        if dl_list[1].find('Organization Type') == -1:
            org_type_count = 0
        else:
            org_type_count = 1

        # renewal date
        renewal_count = dl_list[1].count('Renewal Date')

        # This part does the same has the areas of interest except with socials
        social = dl_list[1].find('<dt>Social')
        social_count = dl_list[1][social:-1].count('<dd>')

        # campus association
        campus_association = dl_list[1].find('Campus Association')
        interest = dl_list[1].find('Areas of Interest')
        campus_count = dl_list[1][campus_association:interest].count('<dd>')
        print(campus_count)
        # This part counts how many areas of interests are listed by counting the number of <dd> between Areas of
        # Interest and Renewal Date
        date = dl_list[1].find('Renewal Date')
        if date == -1:
            interest_count = dl_list[1][interest:social].count('<dd>')
        else:
            interest_count = dl_list[1][interest:date].count('<dd>')


        if primary_count == 1:
            club_prim_contact = dd_list[0][4:-5]  # primary club contact
        else:
            club_prim_contact = ''
        if email_count == 1:
            club_email = dd_list[primary_count + secondary_count + telephone_count][13:-21]  # club email
        else:
            club_email = ''

        if website_count == 0:  # if there is no website
            club_website = ''
        else:  # if website exists
            club_website = dd_list[primary_count + secondary_count + telephone_count + email_count][13:-33]  # club website
        print(dd_list)
        if campus_count == 0:
            club_campus = ''
        else:
            club_campus = dd_list[primary_count + secondary_count + telephone_count + email_count + website_count + mailing_count + org_type_count][
                      4:-5]  # campus association

        club_interests = []  # init club interest list
        print(club_name)
        for j in range(interest_count):  # going through how many interests there are
            club_interests.append(
                dd_list[primary_count + secondary_count + telephone_count + email_count + website_count + mailing_count + org_type_count + campus_count + j][
                4:-5])  # adding those interests

        club_socials = ['', '', '']  # init socials list, in the order twitter, facebook, instagram
        # social index so we don't need to keep typing all of this
        social_index_start = primary_count + secondary_count + telephone_count + email_count + website_count + mailing_count + org_type_count + campus_count + interest_count + renewal_count
        for j in range(social_count):  # going through how many socials they have listed
            if 'Twitter' in dd_list[social_index_start + j]:  # if it's twitter
                club_socials[0] = dd_list[social_index_start + j][13:-21]
            elif 'Facebook' in dd_list[social_index_start + j]:  # if it's facebook
                club_socials[1] = dd_list[social_index_start + j][13:-22]
            elif 'Instagram' in dd_list[social_index_start + j]:  # if it's instagram
                club_socials[2] = dd_list[social_index_start + j][13:-23]

        # adding the collected info to info
        info.append(
            [club_name, club_desc, url[i], club_prim_contact, club_email, club_website, club_campus,
             ', '.join(club_interests),
             club_socials[0], club_socials[1], club_socials[2]])

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


org_list_file = open('org_link_list_ulife2.txt', 'r', encoding='utf-8')  # opening the cached HTMLs of each organization
org_list = org_list_file.read().split('</html>')  # reading it and splitting at the end of each HTML
org_list.pop()  # removing the last element as it is just </html>
org_list_file.close()

org_cache_file = open('url_org_cache_ulife2.txt',
                      'r', encoding='utf-8')  # opening the url cache file for each organization, used to get the org url
url_list = org_cache_file.read().splitlines()  # reading it and splitting the lines
org_cache_file.close()

org_info = get_org_info(org_list, url_list)  # getting org info

wb_name = 'Test - Orgs.xlsx'  # workbook name
titles = ['Club Name', 'Club Description', 'URL', 'Primary Contact', 'Club Email', 'Club Website', 'Campus Association',
          'Areas of Interest', 'Twitter', 'Facebook', 'Instagram']  # list of column titles
col_width = [40, 160, 40, 40, 40, 40, 40, 120, 40, 40, 40]  # list of column widths
write_sheet(wb_name, titles, col_width, org_info)  # writing sheet
