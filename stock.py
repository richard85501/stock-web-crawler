# import urllib.request as req
# url = "https://histock.tw/stock/brokerprofit.aspx?bno=1470"
# # 建立一個request物件 附加req.Request headers 的資訊
# request = req.Request(url,headers={
#     "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36"
# })
# with req.urlopen(request) as response:
#     data = response.read().decode("utf-8")
# # print(data)

# # 解析原始碼
# import bs4
# root=bs4.BeautifulSoup(data,'html.parser')
# print(root.title.string)
# /*********************
print("Exercise 1")
import requests
from bs4 import BeautifulSoup
import csv
import time
from os.path import exists

html = requests.get("https://histock.tw/stock/brokerprofit.aspx?bno=1650")
bsObj= BeautifulSoup(html.content, "lxml")
file_name ="E2-1-1-1"+"各國貨幣"+".csv" 

for single_tr in bsObj.find("", {"id":"CPHB1_bt1_g"}).find("tbody").findAll("tr"):
    cell = single_tr.findAll("td")
    rate1 = cell[1].contents[0]
print(rate1)