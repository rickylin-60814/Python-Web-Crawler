import pandas as pd 
import csv
import requests,time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By

browser = webdriver.Firefox(executable_path='geckodriver')
browser.maximize_window()
browser.get('https://ulist.moe.gov.tw/Query/AjaxQuery/Discipline/073')#跳轉要的網頁
time.sleep(1)

url1 = 'https://ulist.moe.gov.tw/Query/AjaxQuery/Discipline/073'
df1 = pd.read_html(url1,encoding='utf-8')
name=[]
gakubu=[]
tel=[]
extension=[]
fax=[]
email=[]
site=[]
page=0

for ii in range(11):
	for i in range(0,20,1):#從0開始到20,每次+2
		test1=browser.current_window_handle # 獲取當前頁面句柄
		print(test1) # 輸出當前窗口句柄(主頁)
		bs=BeautifulSoup(browser.page_source, "lxml")#取得資料
		browser.find_elements_by_class_name("Normallink")[i].click()
		time.sleep(1.2)

		print('function3:switch to manipulate the new page window')
		handles=browser.window_handles  # 獲取瀏覽器中的所有窗口
		page=page+1
		browser.switch_to.window(handles[-(page)])
		print(handles)
		print(page)
		print(handles[page])
		print(handles[-1])
		test1=browser.current_window_handle # 獲取當前頁面句柄
		print(test1) # 輸出當前窗口句柄

		bs=BeautifulSoup(browser.page_source, "lxml")
		print("===============================")

		namae = bs.find_all('li',{'class':"active"})#取得資料
		namaedata=namae[0].text.split()#取得學校名稱
		name.append(namaedata[0])
		inmagi = bs.select('table',{'id':"one-column-emphasis"})#取得學校資訊
		inmagidata=inmagi[0].text.split()

		tel.append(inmagidata[16])
		extension.append(inmagidata[19])
		fax.append(inmagidata[22])
		email.append(inmagidata[24])
		site.append(inmagidata[26])

		print(name)
		print(inmagidata)
		print(tel)
		handles = browser.window_handles
		browser.switch_to_window(handles[0])#切换回窗口A
		if ii == 10 and i==4:
			break
	if ii ==10:
		break
	test1=browser.current_window_handle  # 獲取當前頁面句柄
	print(test1) # 輸出當前窗口句柄
	bs=BeautifulSoup(browser.page_source, "lxml")#取得資料
	browser.find_element(By.LINK_TEXT,'下一頁').click()#點選名為 "下一頁"的按鈕
	time.sleep(1.2)#呼叫時間模擬人 爬蟲風險



browser.quit() # 關閉爬蟲機器人
data = pd.DataFrame({'學校和科系名稱':name,'電話':tel,'分機':extension,
	'傳真':fax,'e-mail':email,'網址':site})
data.to_excel('大專院校建築相關資訊.xlsx',encoding='cp950')