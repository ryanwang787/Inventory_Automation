import requests
import sys
from os import path
from time import sleep


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
    print("Would you like to wipe the Search List Data? y/n")
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


def read_cache(cache_file):
    """
    Function that reads the cache to see which urls were already scraped
    :param cache_file: cache file
    :return: List of cached urls
    """
    url_cache_file = open(cache_file, 'r', encoding='utf-8')  # opening file
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
    print("Getting URL of page " + url)
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
            file_w.write(html)  # write it in the file
    file_w.close()


file_list = ['org_search_list_ulife.txt', 'url_cache_ulife.txt']  # init file list

for i in range(len(file_list)):  # checking if the files exist
    check_file_exist(file_list[i])

wipe(file_list)  # calling wipe function with the list of files

url_cache_contents = read_cache(file_list[1])  # calling the read cache function to get which urls have been cached

page_links = []  # init page links list
pages = []  # init page list
# list of headers
# make sure you change the user agent to your user agent
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/91.0.4472.164 Safari/537.36',
    'Accept-Language': 'en-US, en-CA',
    'Accept-Encoding': 'gzip, deflate',
    'Accept': 'text/html',
    'Referer': 'http://www.google.com/'}

'''for i in range(47): # this is for all the search pages. Don't run this until we contact ULife admin
    page_links.append('https://www.ulife.utoronto.ca/organizations/list/page/' + str(i + 1) + '/type/all')'''

for i in range(0, 3):  # 0 to 1 is for the first search page
    # adding the urls to be requested to the page links list
    page_links.append('https://www.ulife.utoronto.ca/organizations/list/page/' + str(i + 1) + '/type/all')
#page_links.append('https://www.ulife.utoronto.ca/organizations/list/page/18/type/all')

for i in range(len(page_links)):  # going through the page links list
    # adding the HTML of each of those pages from the page link list to the page list
    pages.append(get_html(page_links[i], url_cache_contents, file_list[1], headers))

write_file(pages, file_list[0])  # writing the HTMLs to the HTML cache file
