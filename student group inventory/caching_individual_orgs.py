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
    print("Would you like to wipe the Orgs Data? y/n")
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
    page.encoding = 'utf-8'
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
    print("Getting URL of page " + url[53:])
    if url not in cache_contents:  # if the url is not already cached
        sleep(2)  # 2 second delay
        html = get_html_from_web(url, cache_file,
                                 headers_list)  # requesting the website's HTML with get_html_from_web function
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
            file_w.write(str(html))  # write it in the file
    file_w.close()


def get_org_list():
    """
    Function that gives the links of all organizations from the search page in a list
    :return: List of all organizations' links
    """
    if path.exists('org_search_list_ulife2.txt'):  # checking if the path exists
        org_search_file = open('org_search_list_ulife2.txt', 'r', encoding='utf-8')  # opening the file in read
        soups = org_search_file.readlines()  # splitting it at </html> to separate the HTMLs of each page
        soups.pop()  # removing the last item as it is empty
        org_search_file.close()
        soups = [i.strip('\n') for i in soups]  # removing all '\n'

        return get_links(soups)
    else:
        print("Sorry, that file does not exist. Please run caching_ulife_search.py")


def get_links(soup_list):
    """
    Function that takes in a list of HTMLs for the search pages and returns every organization link from that page
    :param soup_list: List of pages' HTML
    :return: List of all organizations' links
    """

    for i in range(len(soup_list)):  # going through the list of HTMLs
        soup_list[i] = soup_list[i] + '</html>'  # adding </html> at the end of each item so it's readable by bs4
        soup_list[i] = BeautifulSoup(soup_list[i], 'html.parser')  # replacing each HTML with its parsed version

    link_list = []  # init list of links inside each HTML

    for i in range(len(soup_list)):  # going through the HTML list
        for link in soup_list[i].find_all('a'):  # finding all instances of a-class attributes, which are links
            link_list.append(link.get('href'))  # here the tag is 'href' and .get gets the link associated with that tag

    # right now link_list is a list of the subdirectories and everything after that in each url on the page
    # we will go through those and check if they match the format of an organization's url
    # organization urls look like: https://www.ulife.utoronto.ca/organizations/view/id/21090

    org_links = []  # init list of individual organization links

    for i in link_list:  # going through the list of links
        if 'id' in i and 'index' not in i:  # checking if 'id' is in the url and 'index' is not
            org_links.append('https://www.ulife.utoronto.ca/' + i)  # adding the link

    return org_links



file_list = ['org_html_list.txt', 'org_url_cache.txt']  # init file list

for i in file_list:  # checking if those lists exist
    check_file_exist(i)

wipe(file_list)  # calling wipe function

url_cache_contents = read_cache(file_list[1])  # reading the organizations' urls that have been cached

orgs = []  # init organizations list
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/91.0.4472.164 Safari/537.36',
    'Accept-Language': 'en-US, en-CA',
    'Accept-Encoding': 'gzip, deflate',
    'Accept': 'text/html',
    'Referer': 'http://www.google.com/'}


for i in range(0, 2):  # grabbing the 1st to 9th organization from the search list
    orgs.append(get_html(get_org_list()[i], url_cache_contents, file_list[1], headers))

# Note: The first two organizations on the first search page are problematic as they are blank. The following line
# removes these organizations
#orgs = orgs[2:]

write_file(orgs, file_list[0])  # writing HTMLs to HTML cache file
