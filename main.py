# from lazyspider.lazyheaders import LazyHeaders
# import requests
from selenium import webdriver  # 从selenium导入webdriver
import time
from selenium.webdriver.chrome.options import Options
import random

# import win32api
# import win32con
# import win32gui
# import win32process

wechat_name = "南京吃喝玩乐"
chrome_path = "C:\Program Files (x86)\Google\Chrome\Application"

file_data_name = "data_" + wechat_name + ".csv"
webdriver_path = chrome_path + "\chromedriver.exe"

chrome_options = Options()
# chrome_options.add_argument('--headless')
# chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--save-page-as-mhtml')
driver = webdriver.Chrome(
    executable_path=webdriver_path,
    options=chrome_options)
driver.get('https://weixin.sogou.com/')  # 获取页面
inputBox = driver.find_element_by_id('query')  # 获取输入框
inputBox.send_keys(wechat_name)
searchButton = driver.find_element_by_class_name('swz')  # 获取搜索按钮
searchButton.click()
toolShow = driver.find_element_by_id('tool_show')
toolShow.click()
toolSearchButton = driver.find_element_by_id('search')
toolSearchButton.click()
toolSearchInputBox = driver.find_element_by_class_name('s-sea')
toolSearchInputBox.send_keys(wechat_name)
toolSearchEnter = driver.find_element_by_id('search_enter')
toolSearchEnter.click()

with open(file_data_name, "w") as file:
    file.writelines("标题\t链接")
    file.write("\n")

while True:
    pagebarContainer = driver.find_element_by_id('pagebar_container')
    print(pagebarContainer)
    pageNext = driver.find_element_by_id('sogou_next')
    print(pageNext)
    newsList = driver.find_element_by_class_name('news-list')
    links = newsList.find_elements_by_tag_name('li')
    for link in links:
        txtBox = link.find_element_by_class_name('txt-box')
        h3 = txtBox.find_element_by_tag_name('h3')
        a = h3.find_element_by_tag_name('a')
        a.click()
        handles = driver.window_handles
        driver.switch_to.window(handles[-1])
        print(driver.title)
        print(driver.current_url)
        with open(file_data_name, "a") as file:
            file.write(driver.title)
            file.write("\t")
            file.write(driver.current_url)
            file.write("\n")

        time.sleep(random.randint(5, 10))

        driver.close()
        driver.switch_to.window(handles[0])
    if pageNext is None:
        break
    pageNext.click()
