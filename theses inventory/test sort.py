import excelrd
import sys
from os import path


# The following code was from the answer to "Python - Returning a break statement" on Stackoverflow thanks to
# Jun Zhang (https://stackoverflow.com/users/7424826/jun-zhang) at the following url:
# https://stackoverflow.com/questions/42291272/python-returning-a-break-statement
# The snippet of code is used on lines 10-11, 68, 189, and 194-195
class BreakException(Exception):
    pass


def wipe(data_list):
    """
    Function that wipes the data from all lists
    :param data_list: List of data lists
    :return: None
    """
    # asking user if they would like to wipe the data
    print("Would you like to wipe the Rows List Data? [y/n]")
    inp = sys.stdin.readline()

    if inp.strip().lower() == 'y':  # if yes
        print("Are you sure? [y/n]")
        confirm = sys.stdin.readline()
        if confirm.strip().lower() == 'y':
            for i in range(len(data_list)):  # goes through each data list
                file_wipe_list = open(data_list[i], 'w')
                file_wipe_list.write('')  # replaces all content with an empty string
                file_wipe_list.close()
            print("The data was erased")
        elif confirm.strip().lower() == 'n':
            wipe(data_list)
        else:
            print("Not a valid input, please try again")
            wipe(data_list)
    elif inp.strip().lower() == 'n':  # if no, does nothing
        print("No data was erased")
    else:  # catching invalid inputs
        print("Invalid input. Please try again")
        wipe(data_list)


def check_file_exist(file):
    """
    Function that checks if the file exists, and creates it if it doesn't
    :param file: The file name you want to check
    :return: None
    """
    if not path.exists(file):  # if the file doesn't exist
        temp = open(file, 'x')  # create it
        temp.close()


def display(row):
    """
    Function that displays the Title and keywords of each row individually and allows for user to choose what to do
    with that row
    :param row: Row you want to check
    :return: None
    """
    print('Title: ' + row[2], '\nKeywords: ' + row[9])  # displays title and keywords
    print('Keep [Q]\tRemove [W]\tSave [E]\tSee more Info [R]\tStop [P]')  # user options
    user_in = sys.stdin.readline()  # user input
    if user_in.strip().lower() == 'q':  # if they decide to keep the row
        keep(row)  # call the keep function
    elif user_in.strip().lower() == 'w':  # if they decide to remove the row
        not_keep(row)  # call the not_keep function
    elif user_in.strip().lower() == 'e':  # if they want to see abstract
        save(row)  # call the readmore function
    elif user_in.strip().lower() == 'r':  # if they want to save it for later
        readmore(row)  # call the save_for_later function
    elif user_in.strip().lower() == 'p':  # if they want to stop the program
        print("Stopping Process")
        raise BreakException()  # raise the BreakException class, which returns a pass that breaks the for loop
    else:  # catching invalid inputs
        print('Not a valid input. Please try again')
        display(row)


def readmore(row):
    """
    Function that displays more information, in this case abstract and subjects
    :param row: Row you want to check
    :return: None
    """
    print("Abstract: ", row[5])  # displaying abstract
    print("Subject(s): ", row[8])  # displaying subjects
    print("Keep [Q]\tRemove [W]\tSave [E]\tStop [P]")  # displaying user options
    user_in = sys.stdin.readline()  # user input
    if user_in.strip().lower() == 'q':  # if they want to keep
        keep(row)  # call keep function
    elif user_in.strip().lower() == 'w':  # if they want to remove
        not_keep(row)  # call not_keep function
    elif user_in.strip().lower() == 'e':  # if they want to save it for later
        save(row)  # call save_for_later function
    elif user_in.strip().lower() == 'p':  # if they want to stop the program
        print("Stopping Process")
        raise BreakException()  # raise the BreakException class, which returns a pass that breaks the for loop
    else:  # catching invalid inputs
        print('Not a valid input. Please try again.')
        readmore(row)


def keep(row):
    """
    Function that adds the row to a keep list
    :param row: Row you want to keep
    :return: None
    """
    print("Are you sure you want to keep?\tYes [Q]\tNo [R]")
    confirm = sys.stdin.readline()
    if confirm.strip().lower() == 'q':
        print("Keeping " + row[2] + '\n\n')
        keep_list.append(row)  # adding row to keep list
        checked_list.append(row)  # adding row to checked list
    elif confirm.strip().lower() == 'r':
        display(row)
    else:
        print("Please enter a valid option")
        keep(row)


def not_keep(row):
    """
    Function that adds the row to a remove list
    :param row: Row you want to remove
    :return: None
    """
    # prompting user for a reason for removal
    print("Which line/What reason?\tSee more Info [R]\t(Enter ['Back'] to go back to previous menu)")
    reason = sys.stdin.readline()  # user input
    if reason.strip().lower() == 'back':  # if the user wants to go back
        display(row)  # goes back to display menu
    elif reason.strip().lower() == 'r':  # if the user wants to see more information
        readmore(row)  # calls readmore function
    else:
        print("Your reason for removal is: " + reason)  # confirming the user's reason
        print("Is this correct and would you like to remove this row?\tYes [W]\tNo[R]\t(Enter ['Back'] to go back to "
              "previous menu)")
        confirm = sys.stdin.readline()  # user input
        if confirm.strip().lower() == 'w':  # if the user confirms
            print("Deleting: " + row[2] + '\n\n')
            del_list.append(row)  # add the row to the delete list
            reason_list.append(reason)  # add the reason for removal to the reason list
            checked_list.append(row)  # add the row to checked list
        elif confirm.strip().lower() == 'r':  # if the user declines
            print('The row was not removed. Please enter a new reason.')  # asks for a new reason
            not_keep(row)  # calls function again
        elif confirm.strip().lower() == 'back':  # if user wnats to go back
            display(row)  # calls display function
        else:  # invalid input catch
            print("Please enter a valid option")  # asks for valid input
            not_keep(row)  # runs function again


def save(row):
    """
    Function that adds the row to a save for later list
    :param row: Row you want to save for later
    :return: None
    """
    print("Are you sure you want to save?\tYes [E]\tNo [R]")
    confirm = sys.stdin.readline()
    if confirm.strip().lower() == 'e':
        print("Saving: " + row[2] + '\n\n')
        saved_list.append(row)  # adding row to save for later list
        checked_list.append(row)  # adding row to checked list
    elif confirm.strip().lower() == 'r':
        display(row)
    else:
        print("Please enter a valid option")
        save(row)


def write(file, file_list):
    """
    Function that writes each list to the corresponding file
    :param file: File you are writing to
    :param file_list: List you want to write in to the file
    :return: None
    """
    # The file list is formatted in [row, row, row, etc] with each row having multiple columns of strings
    # goes through each column of each row of the file list and writes it in the file
    for i in range(len(file_list)):
        for j in range(len(file_list[i])):
                file.write(file_list[i][j] + '@%&')  # adding a '@%&' at the end of each element to separate the columns
        file.write('\n')  # adds a new line at the end of each row


sheet_path = 'test book.xlsx'  # spreadsheet we are sorting through
wb_name = excelrd.open_workbook(sheet_path)  # reading workbook
sheet = wb_name.sheet_by_index(0)  # sheet index
row_list = []  # init list of rows

# adding each row of the spreadsheet minus the headers to our list
for i in range(1, sheet.nrows):
    row_list.append(sheet.row_values(i))

# replacing all new lines with a single space
for i in range(len(row_list)):
    for j in range(len(row_list[i])):
        row_list[i][j] = row_list[i][j].replace('\n', ' ')

list_list = ['k_list.txt', 'd_list.txt', 's_list.txt', 'c_rows.txt']  # list of lists

for i in list_list:  # checking if each of the lists exist
    check_file_exist(i)

wipe(list_list)  # giving the user the option to wipe the lists and start over

checked_file = open('c_rows.txt', 'r', encoding='utf-8')  # reading which rows we've already sorted

# init lists
keep_list = []
del_list = []
reason_list = []
saved_list = []
checked_list = checked_file.readlines()  # this list will be the rows we've already checked

checked_file.close()

check = ''  # init string
for i in range(len(checked_list)):  # going through every element of the checked list
    checked_list[i] = checked_list[i].replace('\n', '')  # removing the new line from the end of each element
    check += checked_list[i]  # adding it to our checked string

try:  # if no BreakException was raised
    for i in row_list:  # checking if the rows we are going for is not in our checked string
        if i[2] not in check:
            display(i)  # if we haven't checked them yet, display them for the user to sort through

except BreakException:  # if the BreakException is raised in the display function
    pass  # don't run the for loop, i.e. stop displaying

# opening the files to write to
# we want to append to keep, remove, and save for later files so we keep our previous decisions
# while adding on our new ones
# while we write to the checked files since our checked list contains all checked rows, old and new
keep_file = open('k_list.txt', 'a', encoding='utf-8')
del_file = open('d_list.txt', 'a', encoding='utf-8')
saved_file = open('s_list.txt', 'a', encoding='utf-8')
checked_file = open('c_rows.txt', 'w', encoding='utf-8')

for i in range(len(keep_list)):  # adding a new line after each row
    keep_list[i].append('\n')

for i in range(len(del_list)):  # adding the reason why we removed the row
    del_list[i].append(reason_list[i])  # the reason already has '\n' after it so no need to manually add

for i in range(len(saved_list)):  # adding a new line after each row
    saved_list[i].append('\n')

# calling write function to write each list to their respective files
write(keep_file, keep_list)
write(del_file, del_list)
write(saved_file, saved_list)
for i in checked_list:
    for j in i:
        checked_file.write(j)
    checked_file.write('\n')

keep_file.close()
del_file.close()
saved_file.close()
checked_file.close()
