''' This program is not really needed, the remove duplicates one makes this one redundant '''

import xlrd
import xlsxwriter

loc = 'Test - Masters Keywords.xlsx' # name of SORTED spreadsheet to be highlighted

wb = xlrd.open_workbook(loc) # opening file
sheet = wb.sheet_by_index(0) # reading first sheet
list = [] # init list of rows

# adding each row to list
for i in range(1, sheet.nrows):
    list.append(sheet.row_values(i))

wb_name = 'Test - Masters Duplicates.xlsx'  # new workbook name
workbook = xlsxwriter.Workbook(wb_name)  # writing new workbook
worksheet = workbook.add_worksheet()  # adding spreadsheet

bold = workbook.add_format({'bold': True})  # bold formatting

row = 1  # init row
col = 0  # init column

# headings
titles = ['Author', 'Advisor', 'Title', 'Department', 'Date issued', 'Abstract', 'Degree', 'Permanent URL', 'Subject',
          'Keyword(s)']
# width of headings
col_width = [43, 51, 38, 51, 11, 80, 24, 32, 80, 18]

# adding the SDG group columns
for i in range(1, 17):
    titles.append('SDG' + str(i))
    col_width.append(8)

# setting headers and cell widths
for i in range(sheet.ncols):
    worksheet.set_column(i, i, col_width[i])
    worksheet.write(0, i, titles[i], bold)

# yellow highlight format
format_yellow = workbook.add_format({'bg_color': '#ffeb99',
                                     'font_color': '#997300'})
# green highlight format
format_green = workbook.add_format({'bg_color': '#b3ffb3',
                                    'font_color': '#006600'})

count = 0 # iterating count to change colours
colour = [format_green, format_yellow] # colour list

for i in range (len(list)):  # taking each row
    change = False # changing colour bool
    for j in range(sheet.ncols):  # for every column in that row
        if i < len(list) - 1: # for all but the last entry, otherwise we'd go out of range
            if list[i][2] == list[i + 1][2]: # if its the same author as the next entry
                worksheet.write(row, j, list[i][j], colour[count]) # colour code it into our new sheet

            elif list[i][2] != list[i + 1][2] and list[i][2] == list[i-1][2]: # if its not the same as the next but was the same as the one before
                worksheet.write(row, j, list[i][j], colour[count]) # colour code it into our new sheet
                change = True # change the colour bool so we know to switch colours

            else: # if its not a duplicate
                worksheet.write(row, j, list[i][j])  # write in the copied info for each cell
        else: # for the last case
            if list[i][2] == list[i - 1][2]: # if its the same as the one before
                worksheet.write(row, j, list[i][j], colour[count]) # add it with the same colour
            else: # if not
                worksheet.write(row, j, list[i][j])  # write in the copied info for each cell

    row += 1  # moving to next row
    if change: # if we need to change our colour
        count = (count + 1) % 2 # iterates between 0 and 1
        change = False # stop changing colour

workbook.close()
