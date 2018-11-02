## -*- cording: utf-8 -*-
import urllib
import urllib.request
from bs4 import BeautifulSoup
import openpyxl as excel
import sys
import concurrent.futures

# ------------線内を自由に変更してください----------
# 取得したいページ数を以下に格納してください。
pages = 1000
# 保存するときの名前を以下に格納してください。
# デフォルトは - にしてあります。
savename = '-'
# --------------------------------------------------


# 新規ワークブックオブジェクトを生成する
wb = excel.Workbook()
# アクティブシートを得る
ws = wb.active
# シート名を変更する


def scrall(url):
    tpe = concurrent.futures.ThreadPoolExecutor(max_workers=64)
    for num in range(1,pages+1):
        url = 'https://www.cotta.jp/products/list.php?category_id=00' + str(num)

    def scr_per_page(url):

        titleList = []
        countForExcel = 1

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

        except urllib.error.HTTPError:
            continue


        sys.stdout.write('.')
        sys.stdout.flush()
        return

    tpe.shutdown()




scrall(input('URLを入力してください: '))
