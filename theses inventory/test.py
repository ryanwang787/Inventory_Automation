import sys
import requests

'''headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.164 Safari/537.36',
           'Accept-Language': 'en-US, en-CA',
           'Accept-Encoding': 'gzip, deflate',
           'Accept': 'text/html',
           'Referer': 'http://www.google.com/'}

url = 'https://webscraper.io/test-sites/e-commerce/allinone'

r = requests.get(url, headers=headers)
print(r.request.headers)

x = 'test var'

def adding(list):
    sum = 0
    for i in list:
        sum += i
    return sum

file1 = open('test file.txt', 'w')
for i in range(5):
    file1.write(str(i))
file1.close()

file2 = open('test file.txt', 'r')
content1 = file2.readline()
file2.close()

print(type(content1))
print(content1)

print('3' in content1)

user_in = sys.stdin.readline()
if user_in.strip().lower() == 'x':
    print('yes')

sent = ['Vannina"What Happens Beyond the Talking": A Critical Policy Analysis of Sport and Recreation Equity Policies in Higher EducationSocial Justice Education2020-11-01This thesis is a critical policy analysis of a Canadian urban university sport and recreation department’s equity policies. Through a documentary analysis']
title = '"What Happens Beyond the Talking": A Critical Policy Analysis of Sport and Recreation Equity Policies in Higher Education'
sent1 = ['']
print(title in sent1[0])

word = 'socio economic'
sent2 = 'the cost of the socio  economic crisis was massive'
sent3 = 'the cost of the socio economic crisis was massive'
print(word in sent2)
print(word in sent3)

sent2 = ' '.join(sent2.split())
#re.sub(' +', ' ', sent2)
print(type(str(type(sent2))))
if type(sent2) is str:
    print('ye')

print(False is False)

myList = [''] * 6
print(myList)

word = input("Choose a num ")
age = 2000 - int(word)
print(age)

s = 'nourish'
s1 = 'malnourished'
print(s in s1)


a = 'carbon capture'
a1 = 'carbon capturing is a techniquye'
print(a in a1)

nums = [0, 1, 2, 3, 4, 5, 6, 7]
print(nums[2:])

new_keywords = open('new keywords list.txt', 'r', encoding='utf-8')
words = new_keywords.read()
word_list = words.split(', ')
print(words)
print(word_list)

words = ['The', 'quick', 'brown', 'fox']
print(', '.join(words))'''

lis = [2, 2]

print(lis[-2])