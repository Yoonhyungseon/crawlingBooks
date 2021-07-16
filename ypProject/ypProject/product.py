
import requests
from bs4 import BeautifulSoup
import re

def find_product2(url, cate, startrow, ws, param=['text', 'href'], tag='img'):
    ws.Columns(1).ColumnWidth = 25
    ws.Columns(2).ColumnWidth = 140
    ws.Columns(3).ColumnWidth = 140
    lists_text1 = []
    lists_href1 = []
    lists_href2 = []

    html = requests.get(url=url)
    soup = BeautifulSoup(html.content, "html.parser")
    row = 0
    ws.Cells(row+startrow + 1, 1).Value = cate
    for product in soup.select('#book_img > span > img.paddingtop5'):
        #print(product)
        if "preview_page" in product.get('onclick'):
            href1_re = []
            href2_re = []
            href1 = product.get('onclick').split(',')
            href2 = href1[0].split('(')
            print(href2[1])
            # lists_href2.append(href2[1].replace('\'',' '))
            lists_href2.append(href2[1])

    for product in soup.select('div dl dt a'):
        if "fFormAction" in product.get('href'):
            href1_re = []
            href2_re = []
            href1 = product.get('href').split(',')
            href2 = href1[0].split('(')

            for i in range(2):
                href1_re.append(re.sub(r"[^a-zA-Z0-9]", "", href1[i]))
            for i in range(2):
                href2_re.append(re.sub(r"[^a-zA-Z0-9]", "", href2[i]))

            href_str = "https://www.ypbooks.co.kr/book.yp?bookcd=" + href1_re[1]
        text = product.get_text()

        lists_text1.append(text)

        lists_href1.append(href_str)

        ws.Cells(row + startrow + 1, 2).Value = lists_text1[row]
        ws.Cells(row + startrow + 1, 3).Value = lists_href1[row]
        ws.Cells(row + startrow + 1, 4).Value = lists_href2[row]

        row = row + 1

    if (row >= 2):
        return row
    else:
        return 0
