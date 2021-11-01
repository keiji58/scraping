# 必要なものをインポート
import requests
import chromedriver_binary
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import bs4
from bs4 import BeautifulSoup
import time
import pyautogui as pag
import openpyxl

# 検索ワードと順位を調べたい店舗名を指定
key = "高崎市 パスタ" 
myshop ="シャンゴ 問屋町本店"
# あとで、任意のExcelのセル参照させるようにしよう。


# Chromeをシークレットモードで用意
option = Options()                          # オプションを用意
option.add_argument('--incognito')          # シークレットモードの設定を付与
driver = webdriver.Chrome(options=option)   # Chromeを準備(optionでシークレットモードにしている）

# Googleマップを開く
url = 'https://www.google.co.jp/maps/'
driver.get(url)
driver.maximize_window()
time.sleep(5)


# 指定した検索ワードを入力し検索。
id = driver.find_element_by_id("searchboxinput")
id.send_keys(key)
time.sleep(1)

search_button = driver.find_element_by_xpath("//*[@id='searchbox-searchbutton']")
search_button.click()
time.sleep(4)


# スクロール。これ以外出来なかった。
# スクロールしないと8位以下が表示されない。
pag.moveTo(404,300)
pag.dragTo(404,1079,2,button="left")
time.sleep(1)


# ここからbs4でスクレイピング
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, 'lxml')
a = soup.find_all("a")

# 店舗名リストを作成
shoplist = []
for i in range(len(a)):
  if a[i].has_attr("aria-label"):
    shoplist.append(str(a[i].attrs["aria-label"]))
del shoplist[0:2] # 先頭2つは関係ないので削除

# 広告は順位から除く
if len(shoplist) <= 20: 
  for i in shoplist:
    print(str(shoplist.index(i) + 1) + "." + i)
else: # 20店舗より多い場合、先頭は広告と判定
  ad = len(shoplist) - 20
  for i in shoplist[0:ad]:
    print("広告." + i)
  for i in shoplist[ad:len(shoplist)]:
    print(str(shoplist.index(i) + 1 - ad) + "." + i)
  del shoplist[0:ad]

# 自店の順位を調べる
if myshop in shoplist:
  print(myshop + " は " + str(shoplist.index(myshop) + 1) + "位です！")
else:
  print(myshop + " は " + "圏外です。")




