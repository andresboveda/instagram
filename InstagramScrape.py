#!/usr/bin/python3.8
from instascrape import *  #Note that in order to install this package via pip or on your IDE, it's under the name "insta-scrape"

headers = {
    "user-agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Mobile Safari/537.36 Edg/87.0.664.57",
    "cookie": "sessionid=COPY YOUR SESSION ID HERE;"
}

alca = Profile('https://www.instagram.com/carlitosalcarazz/')
alca.scrape(headers=headers)
n1 = 'Carlos Alcaraz'
link1 = 'https://www.instagram.com/carlitosalcarazz/'
f1 = alca.followers
alca_recent_posts = alca.get_recent_posts()
alca_posts_data = [post.to_dict() for post in alca_recent_posts]
alca_likes = [d['likes'] for d in alca_posts_data]
e1 = ((sum(alca_likes[1:6]))/5)/f1

badosa = Profile('https://www.instagram.com/paulabadosa/')
badosa.scrape(headers=headers)
n2 = 'Paula Badosa'
link2 = 'https://www.instagram.com/paulabadosa/'
f2 = badosa.followers
badosa_recent_posts = badosa.get_recent_posts()
badosa_posts_data = [post.to_dict() for post in badosa_recent_posts]
badosa_likes = [d['likes'] for d in badosa_posts_data]
e2 = ((sum(badosa_likes[1:6]))/5)/f2

ansu = Profile('https://www.instagram.com/ansufati/')
ansu.scrape(headers=headers)
n3 = 'Ansu Fati'
link3 = 'https://www.instagram.com/ansufati/'
f3 = ansu.followers
ansu_recent_posts = ansu.get_recent_posts()
ansu_posts_data = [post.to_dict() for post in ansu_recent_posts]
ansu_likes = [d['likes'] for d in ansu_posts_data]
e3 = ((sum(ansu_likes[1:6]))/5)/f3

list_names = [n1,n2,n3]
list_links = [link1,link2,link3]
list_follow = [f1,f2,f3]
list_engage = [e1,e2,e3]

import xlsxwriter

workbook = xlsxwriter.Workbook(r'C:\Path\To\Location\FileName.xlsx')
worksheet = workbook.add_worksheet()

row = 1
row2 = 1
row3 = 1
row4 = 1
col = 0
bold = workbook.add_format({'bold': True})
worksheet.write('A1', 'NAME', bold)
worksheet.write('B1', 'IG LINK', bold)
worksheet.write('C1', 'FOLLOWERS', bold)
worksheet.write('D1', 'POST ENGAGEMENT AVG', bold)

for item in list_names:
    worksheet.write(row, col, item)
    row += 1
for elem in list_links:
    worksheet.write(row2, col + 1, elem)
    row2 += 1
for elem in list_follow:
    worksheet.write(row3, col + 2, elem)
    row3 += 1
for elem in list_engage:
    worksheet.write(row4, col + 3, elem)
    row4 += 1


workbook.close()
