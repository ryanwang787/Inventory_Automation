import xlrd
import xlsxwriter

loc = 'Test - Orgs Keywords.xlsx' # name of SORTED spreadsheet to be highlighted
wb = xlrd.open_workbook(loc) # opening file
sheet = wb.sheet_by_index(0) # reading first sheet
list = [] # init list of rows

# adding each row to list
for i in range(1, sheet.nrows):
    list.append(sheet.row_values(i))

wb_name = 'Test - Orgs Removed Duplicates.xlsx'  # new workbook name
workbook = xlsxwriter.Workbook(wb_name)  # writing new workbook
worksheet = workbook.add_worksheet()  # adding spreadsheet

bold = workbook.add_format({'bold': True})  # bold formatting
cell_format = workbook.add_format() # text warp format
cell_format.set_text_wrap()

# headings
titles = ['Club Name', 'Club Description', 'URL', 'Keyword(s)']
# width of headings
col_width = [40, 160, 40, 18]

# adding SDGs covered column
titles.append('SDGs Covered')
col_width.append(18)

# adding the SDG group columns
for i in range(1, 17):
    titles.append('SDG' + str(i))
    col_width.append(8)

# setting headers and cell widths
for i in range(len(titles)):
    worksheet.set_column(i, i, col_width[i])
    worksheet.write(0, i, titles[i], bold)

newList = []
row = 1  # init row

# checking if there are duplicates and merging keywords if there are any
for i in range(len(list)):
    list[i].insert(sheet.ncols - 16, '') # adding SDGs covered column
    if list[i][0] == list[i - 1][0]: # if the one before it is a duplicate
        if list[i][sheet.ncols - 17] not in list[i-1][sheet.ncols - 17]: # if the SDG word was not already included
            list[i][sheet.ncols - 17] = list[i][sheet.ncols - 17] + ', ' + list[i - 1][sheet.ncols - 17] # add the SDG to the keywrods list
        else: # if the keyword is already included
            list[i][sheet.ncols - 17] = list[i-1][sheet.ncols - 17] # keep the keywords as they are
        for j in range(16): # for checking the SDGs covered
            if list[i][j + sheet.ncols - 15] == '': # if its empty
                list[i][j + sheet.ncols - 15] = list[i][j + sheet.ncols - 15] + list[i-1][j + sheet.ncols - 15] # add whichever SDG group the previous keyowrd was part of
        newList.append(list[i]) # add that row to a new list
    else: # if its not a duplicate, no need to change anything
        newList.append(list[i]) # add that row to a new list

def keepLast(myList):
    """ Function that removes all but the last row of a group of duplicates """
    for i in range(len(myList)):
        if i < len(newList) - 1:
            if myList[i][0] == myList[i + 1][0]: # if the row after it is the same
                myList.remove(myList[i]) # take the current row out

for i in range(15): # calling keepLast 15 times (if >15 dups, run it more times)
    keepLast(newList)

def SDGcovered(myList):
    """ Function that keeps track of the SDG groups covered """
    for i in range(len(myList)):
        covered = '' # init covered string
        for j in range(16): # going through each of the SDG group columns
            if myList[i][j + sheet.ncols - 15] != '': # if the column is not empty
                if covered == '': # if its the first SDG group we find
                    covered = 'SDG' + str(j + 1) # add it to the covered string
                else: # if its not the first SDG group we find
                    covered = covered + ', SDG' + str(j + 1) # add it to the covered string
            myList[i][sheet.ncols - 16] = covered # set the string to the SDG covered column

SDGcovered(newList)

# adding the newList into a spreadsheet
for i in range(len(newList)):
    for j in range(sheet.ncols + 1):
        if j == sheet.ncols - 17 or j == sheet.ncols - 16: # if its the 'keywords' or 'SDGs covered' column
            worksheet.write(row, j, newList[i][j], cell_format) # warp the text
        else:
            worksheet.write(row, j, newList[i][j])
    row += 1

workbook.close()