import requests
from bs4 import BeautifulSoup
import re

def find_product2(url, cate, startrow, ws, param=['alt', 'src', 'href'], tag='img'):
    ws.Columns(1).ColumnWidth = 25
    ws.Columns(2).ColumnWidth = 140
    ws.Columns(3).ColumnWidth = 140
    lists_alt1 = []
    lists_href1 = []
    lists_href0 = []

    html = requests.get(url=url)
    soup = BeautifulSoup(html.content, "html.parser")
    row = 0
    ws.Cells(row+startrow + 1, 1).Value = cate
    for product in soup.select('div.cover a'):
        for img in product(tag):
            href1_re = []
            href2_re = []
            href = product.get(param[2])
            alt = img.get(param[0])
            if "/product/viewBookDetail.ink" in href:
                href = "http://used.kyobobook.co.kr" + href

            if "/digital/ebook/ebookDetail.ink?barcode" in href:
                href = "https://digital.kyobobook.co.kr" + href

            if ("goDetail(" in href):
                href = "http://pod.kyobobook.co.kr/podBook/podBookDetailView.ink?barcode=" + href[20:-1]

            keywords = ["goDetailProductNotAgeRn", "makeDetailUrl", "goDetailProductNotAge", "parent.fn_DetailView", "parent.fn_DetailView", "goDetailView", "openDetailProductNotAge", "goDetailProduct"]

            if any(keyword in href for keyword in keywords):

                href1 = href.split(',')
                href2 = href1[0].split('(')

                for i in range(3):
                    href1_re.append(re.sub(r"[^a-zA-Z0-9]", "", href1[i]))
                for i in range(2):
                    href2_re.append(re.sub(r"[^a-zA-Z0-9]", "", href2[i]))

                href_str = "http://www.kyobobook.co.kr/product/detailViewKor.laf?mallGb=" + href2_re[1] + "&ejkGb=" + href2_re[1] + "&linkClass=" + href1_re[1] + "&barcode=" + href1_re[2]
                href_isbn = href1_re[2]
            else:
                href_str = href

            lists_alt1.append(alt)
            lists_alt2 = list(dict.fromkeys(lists_alt1))

            lists_href0.append("\'"+href_isbn+"\'")
            lists_href1.append(href_str)
            lists_href2 = list(dict.fromkeys(lists_href1))

            if((lists_alt1 == lists_alt2) and (lists_href1 == lists_href2)):
                ws.Cells(row+startrow + 1, 2).Value = lists_alt2[row]
                ws.Cells(row+startrow + 1, 3).Value = lists_href2[row]
                ws.Cells(row+startrow + 1, 4).Value = lists_href0[row]

                row = row + 1
            else:
                del lists_alt1[-1]
                del lists_href1[-1]

    if (row >= 2):
        return row
    else:
        return 0
