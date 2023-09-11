import requests
import sys
from os import path
from time import sleep
from bs4 import BeautifulSoup


def check_file_exist(file):
    """
    Function that checks if the file exists, and creates it if it doesn't
    :param file: The file name you want to check
    :return: None
    """
    if not path.exists(file):  # if the file doesn't exist
        temp = open(file, 'x')  # create it
        temp.close()


def wipe(files):
    """
    Function that wipes the given files
    :param files: List of files to wipe
    :return: None
    """
    print("Would you like to wipe the Programs Data? y/n")
    inp = sys.stdin.readline()  # user input

    if inp.strip() == 'y':
        for url in files:  # going through each file
            file_wipe_list = open(url, 'w')  # opening it in write mode
            file_wipe_list.write('')  # replacing all content with an empty string
            file_wipe_list.close()
        print("The data was erased")
    elif inp.strip() == 'n':
        print("No data was erased")
    else:
        print("Invalid input. Please try again")
        wipe(files)


def read_cache(file):
    """
    Function that reads the cache to see which urls were already scraped
    :param file: cache file
    :return: List of cached urls
    """
    url_cache_file = open(file, 'r', encoding='utf-8')  # opening file
    url_cache = url_cache_file.readlines()  # reading the lines of the file
    url_cache = [i.strip('\n') for i in url_cache]  # removing the '\n' at the end of each element
    url_cache_file.close()  # closing file
    return url_cache


def get_html_from_web(url, cache_file, headers_list):
    """
    Function that requests the HTML from the server
    :param url: URL to get request from
    :param cache_file: Cache file
    :param headers_list: List of HTTP headers
    :return: Text form of HTML of the url
    """
    url_cache_temp = open(cache_file, 'a', encoding='utf-8')  # opening the cache in append mode
    url_cache_temp.write(str(url) + '\n')  # adding the url just requested to the cache
    url_cache_temp.close()
    page = requests.get(url, headers=headers_list)  # requesting the url
    return page.text  # returning the HTML in the form of text from the url


def get_html(url, cache_contents, cache_file, headers_list):
    """
    Function that determines if the url has already been requested
    :param url: URL to get request from
    :param cache_contents: Cache contents
    :param cache_file: Cache file
    :param headers_list: List of HTTP headers
    :return: HTML string of the website
    """
    html = ''  # init HTML string
    print("Getting URL of page " + url[37:-1])
    if url not in cache_contents:  # if the url is not already cached
        sleep(2)  # 2 second delay
        html = get_html_from_web(url, cache_file, headers_list)  # requesting the website's HTML with
        # get_html_from_web function
        print("URL successfully retrieved")
    else:  # if the url is cached
        sleep(1)  # 1 second delay
        print("URL already in Org List")
    return html


def write_file(html_list, html_file):
    """
    Function that writes the HTML to a file to cache the actual HTML content of a url
    :param html_list: List of HTMLs retrieved
    :param html_file: HTML cache file
    :return: None
    """
    file_w = open(html_file, 'a', encoding='utf-8')  # opening the HTML cache file in append mode
    for html in html_list:  # going through each HTML
        if html != '':  # if the HTML is not blank
            file_w.write(str(html.encode('utf-8', 'ignore')))  # write it in the file
    file_w.close()


def get_programs_list():
    """
    Function that gives the links of all programs from the search page in a list
    :return: List of all programs' links
    """
    if path.exists('SGS_programs_list.txt'):  # checking if the path exists
        programs_search_file = open('SGS_programs_list.txt', 'r', encoding='utf-8')  # opening the file in read
        soups = programs_search_file.readlines()  # splitting it at </html> to separate the HTMLs of each page
        soups.pop()  # removing the last item as it is empty
        programs_search_file.close()
        soups = [i.strip('\n') for i in soups]  # removing all '\n'

        return get_links(soups)
    else:
        print("Sorry, that file does not exist. Please run caching_sgs_programs_list.py")


def get_links(soup_list):
    """
    Function that takes in a list of HTMLs for the search pages and returns every programs' link from that page
    :param soup_list: List of pages' HTML
    :return: List of all programs'' links
    """

    for i in range(len(soup_list)):  # going through the list of HTMLs
        soup_list[i] = soup_list[i] + '</html>'  # adding </html> at the end of each item so it's readable by bs4
        soup_list[i] = BeautifulSoup(soup_list[i], 'html.parser')  # replacing each HTML with its parsed version

    link_list = []  # init list of links inside each HTML

    for i in range(len(soup_list)):  # going through the HTML list
        for link in soup_list[i].find_all('a'):  # finding all instances of a-class attributes, which are links
            link_list.append(link.get('href'))  # here the tag is 'href' and .get gets the link associated with that tag

    programs_links = []  # init list of individual organization links

    for i in link_list:  # going through the list of links
        if 'https://www.sgs.utoronto.ca/programs/' in i:  # checking if 'id' is in the url and 'index' is not
            programs_links.append(i)  # adding the link

    programs_links = programs_links[1:-1]  # removing first and last, as they are just general programs list links
    # print(programs_links)
    return programs_links


file_list = ['SGS_programs.txt', 'SGS_programs_cache.txt']  # init file list

for i in file_list:  # checking if those lists exist
    check_file_exist(i)

wipe(file_list)  # calling wipe function

url_cache_contents = read_cache(file_list[1])  # reading the programs' urls that have been cached

programs = []  # init programs list
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/91.0.4472.164 Safari/537.36',
    'Accept-Language': 'en-US, en-CA',
    'Accept-Encoding': 'gzip, deflate',
    'Accept': 'text/html',
    'Referer': 'http://www.google.com/'}

# for i in range(26): # 26 organizations per search page. Don't run this until we contact ULife admin

for i in range(0, 131):  # grabbing the 1st to 9th programs from the search list
    programs.append(get_html(get_programs_list()[i], url_cache_contents, file_list[1], headers))

write_file(programs, file_list[0])  # writing HTMLs to HTML cache file
