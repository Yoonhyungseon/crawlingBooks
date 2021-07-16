
import requests
from bs4 import BeautifulSoup
import win32com.client
from product import find_product2
import re


"""
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
"""


html = requests.get("https://www.ypbooks.co.kr/search_field.yp?category_type=1&depth=4")

bsObject = BeautifulSoup(html.content, "html.parser")

print('\n포괄 카테고리\n')

lists_category = []
lists_href = []
for link in bsObject.find_all('a', class_="text"):
    #ws.Cells(i, 1).Value = link.text.strip()
    #ws.Cells(i, 2).Value = link.get('href')
    i = i + 1

print('\n세부 카테고리\n')
j = 1
for link in bsObject.select('table td ul li a'):
    j += 1

    lists_category.append(link.text.strip())
    if "fNext" in link.get('href'):
        href1_re = []
        href2_re = []
        href1 = link.get('href').split(',')
        href2 = href1[0].split('(')

        for i in range(2):
            href1_re.append(re.sub(r"[^a-zA-Z0-9]", "", href1[i]))
        for i in range(2):
            href2_re.append(re.sub(r"[^a-zA-Z0-9]", "", href2[i]))

        href_str = "https://www.ypbooks.co.kr/search.yp?catesearch=true&collection=books_kor&sortField=DATE&" + href2_re[1] + "=" + href1_re[1]
        #ws.Cells(j, 1).Value = link.text.strip()
        #ws.Cells(j, 2).Value = href_str
        lists_href.append(href_str)

    else:
        lists_href.append(link.get('href'))

#print(len(lists_href))
#wb.SaveAs('C:/Users/choi/Desktop/크롤링 엑셀 파일')
#excel.Quit()

excel1 = win32com.client.Dispatch("Excel.Application")
excel1.Visible = True
wb1 = excel1.Workbooks.Add()
ws1 = wb1.Worksheets("Sheet1")


row = 0
for idx in range(len(lists_href)):
    row += find_product2(lists_href[idx], lists_category[idx], row, ws=ws1)
    print(lists_href[idx], ' : ', row)

#wb1.SaveAs('C:/Users/choi/Desktop/crawing')
excel1.Quit()


