import requests
from bs4 import BeautifulSoup
import win32com.client
from product import find_product2

"""
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
"""

html = requests.get('http://www.kyobobook.co.kr/index.laf')

bsObject = BeautifulSoup(html.content, "html.parser")

print('\n포괄 카테고리\n')
i = 1
lists_category = []
lists_href = []
for link in bsObject.find_all('a', class_="text"):
    #ws.Cells(i, 1).Value = link.text.strip()
    #ws.Cells(i, 2).Value = link.get('href')
    i = i + 1

print('\n세부 카테고리\n')

for link in bsObject.select('ul.category > li > a'):
    #ws.Cells(i, 1).Value = link.text.strip()
    #ws.Cells(i, 2).Value = link.get('href')
    lists_category.append(link.text.strip())
    lists_href.append(link.get('href'))
    i = i + 1

print(lists_href)
# print(len(lists_href))
#wb.SaveAs('C:/Users/choi/Desktop/크롤링 엑셀 파일')
#excel.Quit()


excel1 = win32com.client.Dispatch("Excel.Application")
excel1.Visible = True
wb1 = excel1.Workbooks.Add()
ws1 = wb1.Worksheets("Sheet1")

row = 0
for idx in range(len(lists_href)):
    row += find_product2(lists_href[idx], lists_category[idx], row, ws=ws1)
    #print(lists_href[idx], ' : ', row)

#wb1.SaveAs('C:/Users/choi/Desktop/crawing')
excel1.Quit()
