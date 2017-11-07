import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
from pandas import ExcelWriter
from pandas import ExcelFile

df = pd.read_excel('categories.xlsx', sheetname='Sheet1')
print(df.columns)
# print(df['link'])


#Read parent link
number=0
for i in df.index:
    number = number+1
    link =df['link'][i]
    print(number)

    sublink = link[:21]
    page = requests.get(link)
    soup = BeautifulSoup(page.content, 'html.parser')
    all_content = soup.find_all(class_="msht-row-head")
    all_topic = soup.find_all(class_="msht-row-title")
    # print(all_content)

    # Find best category
    t = 0
    somestring = 'برترین‌ها'
    for s in all_topic:

        if s.string.strip() == somestring:
            # print(t)
            # print(s.string)
            k = t

        t = t + 1
    top = all_content[k]
    top_a = top.find('a')
    top_link = top_a['href'].strip()
    next_link = sublink + top_link

    page1 = requests.get(next_link)
    soup1 = BeautifulSoup(page1.content, 'html.parser')
    all_content1 = soup1.find_all(class_="msht-app")

    max_app = len(all_content1)
    applist = []
    for j in all_content1:
        apps = j.find('a')
        apps_link = apps['href'].strip()
        s = apps_link[5:-6]

        applist.append(s)

    print(df['category'][i])

    df2 = pd.DataFrame({df['category'][i]:applist})
    book = load_workbook('result.xlsx')
    writer = pd.ExcelWriter('result.xlsx', engine='openpyxl')
    writer.book=book
    df2.to_excel(writer,df['category'][i])
    writer.save()
# print(df['category'][i])






