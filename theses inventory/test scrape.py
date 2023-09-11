from bs4 import BeautifulSoup

test_file = open("test_scrape.txt", 'r')

html = test_file.read()
soup = BeautifulSoup(html, 'html.parser')

dd_list = soup.find_all('dd')
dt_list = soup.find_all(class_='infoDl')

"""for i in range(len(dd_list)):
    dd_list[i] = str(dd_list[i])"""

dd_list = [str(i) for i in dd_list]
print(dd_list)
for i in range(len(dt_list)):
    dt_list[i] = str(dt_list[i])
print(dt_list)
'''secondary_count = 0
telephone_count = 0
website_count = 0
mailing_count = 0'''

if dt_list[0].find('Secondary Contact') == -1:
    secondary_count = 0
else:
    secondary_count = 1
if dt_list[0].find('Telephone') == -1:
    telephone_count = 0
else:
    telephone_count = 1
if dt_list[0].find('Website') == -1:
    website_count = 0
else:
    website_count = 1
if dt_list[0].find('Mailing Address') == -1:
    mailing_count = 0
else:
    mailing_count = 1



social = dt_list[1].find('<dt>Social')
social_count = dt_list[1][social:-1].count('<dd>')

interest = dt_list[1].find('Areas of Interest')
date = dt_list[1].find('Renewal Date')
if date == -1:
    interest_count = dt_list[1][interest:social].count('<dd>')
else:
    interest_count = dt_list[1][interest:date].count('<dd>')

club_prim_contact = dd_list[0][4:-5]
club_email = dd_list[1 + secondary_count + telephone_count][13:-21]
club_website = ''
if website_count == 0:
    club_website = ''
else:
    club_website = dd_list[2 + secondary_count + telephone_count][13:-33]

club_campus = dd_list[3 + secondary_count + telephone_count + website_count + mailing_count][4:-5]

club_interests = []
for i in range(interest_count):
    club_interests.append(dd_list[4 + secondary_count + website_count + telephone_count + mailing_count + i][4:-5])

club_socials = ['','','']
#club_socials = ''
social_index_start = 5 + secondary_count + website_count + telephone_count + mailing_count + interest_count
for i in range(social_count):
    if 'Twitter' in dd_list[social_index_start + i]:
        club_socials[0] = dd_list[social_index_start + i][13:-19]
        # club_socials += dd_list[social_index_start + i][13:-19] + ', '
    elif 'Facebook' in dd_list[social_index_start + i]:
        club_socials[1] = dd_list[social_index_start + i][13:-22]
        # club_socials += dd_list[social_index_start + i][13:-22] + ', '
    elif 'Instagram' in dd_list[social_index_start + i]:
        club_socials[2] = dd_list[social_index_start + i][13:-23]
        # club_socials += dd_list[social_index_start + i][13:-23] + ', '
# club_socials = club_socials[:-2]
print(dd_list)
print(club_prim_contact)
print(club_email)
print(club_website)
print(club_campus)
print(club_interests)
print(club_socials)
