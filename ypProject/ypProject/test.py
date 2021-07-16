import requests
import re
from bs4 import BeautifulSoup
import csv
url = "https://www.ypbooks.co.kr/book_arrange.yp?targetpage=book_week_best&pagetype=5&depth=1"
headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36"} 
res = requests.get(url, headers=headers)
res.raise_for_status()
 
soup = BeautifulSoup(res.text,'html.parser')
items=soup.select("#rightArea01 dt div")

powerbanklist = [] # 라스트 생성
# print(items)
for item in items:
    temp=[]
 
    link = item.select_one("a.btn.type02")["href"]
    link = re.findall("\d+",link)
    print(link)

    temp.append(link)
    powerbanklist.append(temp)
    print(len(powerbanklist))
 
with open('powerbanklist_select.csv',"w", encoding="utf-8-sig", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(['링크'])
    writer.writerows(powerbanklist)
f.close
