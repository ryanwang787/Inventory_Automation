from bs4 import BeautifulSoup
import re

file = open('SGS_programs3.txt', 'r', encoding='utf-8')
html = file.read()
file.close()
soup = BeautifulSoup(html, 'html.parser')

program_name = str(soup.find(class_='col display-4 page-title'))[37:-5] # find the program name, which is the display
grad_unit = str(soup.find(class_='col-sm-4 order-1').h3)[4:-5].encode('ascii', 'ignore').decode('utf-8') # find the graduate unit, the first h3 of the side column
print(grad_unit.encode('ascii', 'ignore').decode('utf-8'))
dep_site = str(soup.find(class_='list-unstyled').findAll('a')[1].get('href')) # find the department site, which is the second url in the side column list
fac_aff = str(soup.find(class_='col-sm-4 order-1').find('p'))[3:-4] # find faculty affiliation, which is the first paragraph in the side column

campus_pos1 = -1
campus_pos2 = -1
campus = ''
for i in range(1, 10):
    if str(soup.find(class_='col-sm-4 order-1').findAll('p')[-i]).find('University of Toronto') != -1: # find where university of toronto appears in the last paragraph, which is usually the address for the program
        print(i)
        campus_pos1 = str(soup.find(class_='col-sm-4 order-1').findAll('p')[-i]).find('University of Toronto')
        campus_pos2 = str(soup.find(class_='col-sm-4 order-1').findAll('p')[-i])[campus_pos1:].find(
            'br') + campus_pos1 - 1  # find where the br appears in the final paragraph, signalling the end of the campus
        campus = str(soup.find(class_='col-sm-4 order-1').findAll('p')[-i])[
                 campus_pos1:campus_pos2]  # slicing the final paragraph to just grab the campus
        break
print(campus_pos1, campus_pos2)
print(str(soup.find(class_='col-sm-4 order-1').findAll('p')[-1]).find('University of Toronto'))


# degree
degrees_list = []
degrees = soup.find(class_='col-sm-8 order-2').findAll('h3')
for i in range(len(degrees)):
    degrees_list.append(str(degrees[i])[4:-5])
#print(degrees_list)

# program description
program_index = []
index = 0
for i in range(len(degrees_list)):
    program_index.append(str(soup.find(class_='col-sm-8 order-2'))[index:].find('<h3>' + degrees_list[i]) + index)
    index = program_index[i]

qf_index = str(soup.find(class_='col-sm-8 order-2'))[0:program_index[0]].find('<h2>Quick Facts</h2>')
program_desc_text = str(soup.find(class_='col-sm-8 order-2'))[35:qf_index]

#print(program_desc_text)

program_info_list = []
program_desc_index = []
program_req_index = []
for i in range(len(degrees_list)):
    if i < len(degrees_list) - 1:
        program_info_list.append(str(soup.find(class_='col-sm-8 order-2'))[program_index[i]:program_index[i+1]])
    else:
        program_info_list.append(str(soup.find(class_='col-sm-8 order-2'))[program_index[i]:-1])

    #if i == 0:
        #program_desc_text += program_info_list[i][]
    program_desc_index.append(program_info_list[i].find('<h4>Program Description</h4>'))
    program_req_index.append(program_info_list[i].find('Requirements</h4'))

    program_desc_text += program_info_list[i][program_desc_index[i]:program_req_index[i]]


#print(program_desc_index)
#print(program_req_index)

program_desc = BeautifulSoup(program_desc_text, 'html.parser').get_text()

#sgs_site = url[i]
calendar_site = str(soup.find(class_='list-unstyled').find('a').get('href'))

print(program_name)
print(grad_unit)
print(dep_site)
print(fac_aff)
print(campus)
#print(program_desc)
print(calendar_site)