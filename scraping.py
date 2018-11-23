import urllib.request
from bs4 import BeautifulSoup
import openpyxl as excel
import sys

pages = input('Pages :')
savename = input('savename :')

countForExcel = 1
wb = excel.Workbook()
ws = wb.active

titleList = []

for num in range(1,int(pages)):
    url = 'https://www.cotta.jp/products/list.php?category_id=00'+str(num)

    try:
        html = urllib.request.urlopen(url)
        soup = BeautifulSoup(html, 'html.parser')
        title = soup.find('title').text.replace(' | お菓子・パン材料・ラッピングの通販【cotta＊コッタ】','')

        if title in titleList:
            continue
        else:
            titleList.append(title)
            ws['A' + str(countForExcel)] = title
            countForExcel += 1

        sys.stdout.write('.')
        sys.stdout.flush()

    except urllib.error.HTTPError:
        continue

wb.save('results/'+ savename+'.xlsx')
print('\nDone!')
