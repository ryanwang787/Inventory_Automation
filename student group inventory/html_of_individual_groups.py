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
    print("Would you like to wipe the Orgs Cached Data? y/n")
    inp = sys.stdin.readline()  # user input

    if inp.strip() == 'y':
        for file in files:  # going through each file
            file_wipe_list = open(file, 'w')  # opening it in write mode
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
    page = requests.get(url, headers=headers_list)  # requesting the url
    page.encoding = 'utf-8'
    url_cache_temp = open(cache_file, 'a', encoding='utf-8')  # opening the cache in append mode
    url_cache_temp.write(str(url) + '\n')  # adding the url just requested to the cache
    url_cache_temp.close()
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
    print("Getting URL of page " + url[29:])
    if url not in cache_contents:  # if the url is not already cached
        sleep(1)  # 2 second delay
        html = get_html_from_web(url, cache_file,
                                 headers_list)  # requesting the website's HTML with get_html_from_web function
        print("URL successfully retrieved")
    else:  # if the url is cached
        #sleep(0.1)  # 1 second delay
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


def get_url_from_txt(filepath):
    if path.exists(filepath):
        txt = open(filepath, encoding='utf-8')
        urls = txt.readlines()
        txt.close()

        urls = [i.strip('\n') for i in urls]

        return urls
    else:
        print('Path does not exist. Try again with a valid text file')


filename = 'org_url_list.txt'
org_urls = get_url_from_txt(filename)
print(len(org_urls))

cache_files = ['orgs_urls_cache.txt', 'orgs_html_cache.txt']

for i in cache_files:
    check_file_exist(i)

wipe(cache_files)

url_cache = read_cache(cache_files[0])

#orgs_html = []
header = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/91.0.4472.164 Safari/537.36',
    'Accept-Language': 'en-US, en-CA',
    'Accept-Encoding': 'gzip, deflate',
    'Accept': 'text/html',
    'Referer': 'http://www.google.com/'}

#print(org_urls[:5])

#print(len(org_urls))

'''for i in range(250):
    orgs_html = get_html(org_urls[i], url_cache, cache_files[0], header)
    write_file(orgs_html, cache_files[1])'''

'''for i in range(250, 500):
    orgs_html = get_html(org_urls[i], url_cache, cache_files[0], header)
    write_file(orgs_html, cache_files[1])'''

'''for i in range(500, 750):
    orgs_html = get_html(org_urls[i], url_cache, cache_files[0], header)
    write_file(orgs_html, cache_files[1])'''

for i in range(750, len(org_urls)):
    orgs_html = get_html(org_urls[i], url_cache, cache_files[0], header)
    write_file(orgs_html, cache_files[1])