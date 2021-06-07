from os import makedirs, path
from time import sleep
import shutil

import requests
from selenium import webdriver  # 从selenium导入webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.by import By

# import urllib.request as request

data_path = "data/rongchuang"


def need_login(driver: webdriver):
    return driver.title == "登录页"


def login(driver: webdriver):
    driver.find_element_by_xpath("/html/body/div[2]/div/div[3]/form/div[1]/input").send_keys("***")
    driver.find_element_by_xpath("/html/body/div[2]/div/div[3]/form/div[2]/input").send_keys("***")
    driver.find_element_by_xpath("/html/body/div[2]/div/div[3]/form/div[3]/input[2]").click()


def open_url(driver: webdriver,
             url: str):
    driver.get(url)  # 获取页面
    if need_login(driver):
        login(driver)


def get_sync_elements(driver: webdriver, xpath: str):
    WebDriverWait(driver, 30, 0.5).until(expected_conditions.presence_of_element_located(
        (By.XPATH, xpath)))
    return driver.find_elements_by_xpath(xpath)


def get_sync_element(driver: webdriver, xpath: str):
    WebDriverWait(driver, 30, 0.5).until(expected_conditions.presence_of_element_located(
        (By.XPATH, xpath)))
    return driver.find_element_by_xpath(xpath)


def get_async_element(driver: webdriver, xpath: str):
    element = None
    count = 0
    while element is None:
        try:
            element = driver.find_element_by_xpath(xpath)
        except:
            for _ in range(count + 1):
                driver.execute_script('window.scrollTo(0,document.body.scrollHeight)')
                sleep(0.5)
            driver.refresh()
            count += 1
    return element


def handle_jump_page(driver: webdriver, data_dir_: str):
    lis = get_sync_elements(driver, "/html/body/div[1]/div[2]/div/div[2]/div[2]/div/div[2]/ul/li[*]")
    for i, li in enumerate(lis):
        # jump_data_upload_list_img = driver.find_element_by_xpath(
        #     f"/html/body/div[1]/div[2]/div/div[2]/div[2]/div/div[2]/ul/li[{i + 1}]/div[2]/div[3]")
        # jump_data_upload_list_img = get_sync_element(driver,
        #                                              f"/html/body/div[1]/div[2]/div/div[2]/div[2]/div/div[2]/ul/li[{i + 1}]/div[2]/div[3]")
        # img_dict = {}
        # _as = jump_data_upload_list_img.find_elements_by_xpath("a[*]")
        # for _a in _as:
        #     img = _a.find_element_by_xpath("img")
        #     title = img.get_attribute("title")
        #     img_dict[title] = img
        #
        # if "下载" in img_dict:
        #     continue
        # a1 = driver.find_element_by_xpath(
        #     f"/html/body/div[1]/div[2]/div/div[2]/div[2]/div/div[2]/ul/li[{i + 1}]/div[2]/div[3]/a[1]")
        a1 = get_sync_element(driver,
                              f"/html/body/div[1]/div[2]/div/div[2]/div[2]/div/div[2]/ul/li[{i + 1}]/div[2]/div[3]/a[1]")
        a2 = get_sync_element(driver,
                              f"/html/body/div[1]/div[2]/div/div[2]/div[2]/div/div[2]/ul/li[{i + 1}]/div[2]/div[3]/a[2]")
        file_name = li.find_element_by_xpath("div[1]/a").text
        if a2.find_element_by_xpath("img").get_attribute("title") == "下载":
            print(f"handle_jump_page {(data_dir_, file_name)} can download.")
            continue

        # if "阅读" in img_dict:
        if a1.find_element_by_xpath("img").get_attribute("title") == "阅读":
            file_name_ = file_name.replace('.', '_')
            data_dir__ = path.join(data_dir_, file_name_)
            makedirs(data_dir__, exist_ok=True)

            # img_dict["阅读"].click()
            print(a1.get_attribute("href"))
            js = "window.open('" + a1.get_attribute("href") + "');"  # 新窗口打开链接
            driver.execute_script(js)
            # 切换到新标签页的window
            driver.switch_to.window(driver.window_handles[-1])

            handle_doc_page(driver, data_dir__, file_name)

            driver.close()
            # 切换到新标签页的window
            driver.switch_to.window(driver.window_handles[-1])
        pass


def handle_doc_page(driver: webdriver, data_dir_: str, file_name: str):
    (file, ext) = path.splitext(file_name)
    if ext in [".doc", ".docx"]:
        handle_word_file(driver, data_dir_)
        # print(f"handle_doc_page {data_dir_} is done.")
    elif ext in [".pdf", ".ppt", ".pptx"]:
        handle_pdf_file(driver, data_dir_)
        # print(f"handle_doc_page {data_dir_} is done.")
    elif ext in [".xls", ".xlsx"]:
        handle_excel_file(driver, data_dir_)
        # print(f"handle_doc_page {data_dir_} is done.")
    else:
        print(f"handle_doc_page {data_dir_} not support!")
        pass
    pass


def handle_word_file(driver: webdriver, data_dir_: str):
    page_tds = get_sync_elements(driver, "/html/body/div[1]/div[2]/div[2]/div[1]/table/tbody/tr[*]/td")
    for i, page_td in enumerate(page_tds):
        requests_session.headers.update({
            "Referer": driver.current_url
        })
        # resp = requests_session.get(
        #     f"{driver.current_url}&viewer=htmlviewer&filekey=toHTML-Aspose_page-{i + 1}-svg&pageNum={i + 1}"
        # )
        resp = None
        while resp is None:
            try:
                resp = requests_session.get(
                    f"{driver.current_url}&viewer=htmlviewer&filekey=toHTML-Aspose_page-{i + 1}-svg&pageNum={i + 1}",
                    stream=True
                )
            except:
                sleep(0.5)
        html = resp.text
        file_path = path.join(data_dir_, f"{i}.html")
        with open(file_path, "w", encoding='utf-8') as output_file:
            output_file.write(html)
        print(f"handle_word_file saved {file_path}")


def handle_pdf_file(driver: webdriver, data_dir_: str):
    page_tds = get_sync_elements(driver, "/html/body/div[1]/div[2]/div[2]/div[1]/table/tbody/tr[*]/td")
    for i, page_td in enumerate(page_tds):
        requests_session.headers.update({
            "Referer": driver.current_url
        })
        resp = None
        while resp is None:
            try:
                resp = requests_session.get(
                    f"{driver.current_url}&method=view&filekey=toHTML-Aspose_page-{i + 1}-img&pageNum={i + 1}",
                    stream=True
                )
            except:
                sleep(0.5)
        file_path = path.join(data_dir_, f"{i}.jpg")
        with open(file_path, "wb") as output_file:
            output_file.write(resp.content)
        print(f"handle_pdf_file saved {file_path}")


def handle_excel_file(driver: webdriver, data_dir_: str):
    li_spans = get_sync_elements(driver, "/html/body/div/div[3]/div/div[1]/div[2]/ul/li[*]/span")
    for i, li_span in enumerate(li_spans):
        requests_session.headers.update({
            "Referer": driver.current_url
        })
        # resp = requests_session.get(
        #     f"{driver.current_url}&viewer=htmlviewer&filekey=toHTML-Aspose_page-{i + 1}"
        # )
        resp = None
        while resp is None:
            try:
                resp = requests_session.get(
                    f"{driver.current_url}&viewer=htmlviewer&filekey=toHTML-Aspose_page-{i + 1}",
                    stream=True
                )
            except:
                sleep(0.5)
        html = resp.text
        file_path = path.join(data_dir_, f"{li_span.text}.html")
        with open(file_path, "w", encoding='utf-8') as output_file:
            output_file.write(html)
        print(f"handle_word_file saved {file_path}")


chrome_path = "C:\Program Files (x86)\Google\Chrome\Application"

webdriver_path = chrome_path + "\chromedriver.exe"

chrome_options = Options()
# chrome_options.add_argument('--headless')
# chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--save-page-as-mhtml')
driver = webdriver.Chrome(
    executable_path=webdriver_path,
    options=chrome_options)
# driver.maximize_window()  # 窗口最大化

open_url(driver, 'http://oa.sunac.com.cn/sys/portal/index/systemCenter2.jsp?fdId=16e2ec3b5d9863862448e2e43afa7925')
# 创建一个requests session对象
requests_session = requests.Session()
# 从driver中获取cookie列表（是一个列表，列表的每个元素都是一个字典）
cookies = driver.get_cookies()
# 把cookies设置到session中
for cookie in cookies:
    requests_session.cookies.set(cookie['name'], cookie['value'])
requests_session.headers.update({
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.77 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
    "Referer": driver.current_url,
    "Accept-Language": "zh-CN,zh;q=0.9"
})

jump_items = get_sync_elements(driver, "/html/body/div[1]/div[2]/div/div[2]/div/div[*]")
for jump_item in jump_items:
    jumps = jump_item.find_elements_by_xpath("div[*]")
    for jump in jumps:
        jump_title = jump.find_element_by_xpath("h3").text
        jump_list = jump.find_elements_by_xpath("div/div[*]/a")
        data_dir = path.join(data_path, jump_title)
        makedirs(data_dir, exist_ok=True)
        for jump_a in jump_list:
            title = jump_a.get_attribute("title")
            # href = jump_a.get_attribute("href")
            data_dir_ = path.join(data_dir, title)
            makedirs(data_dir_, exist_ok=True)
            jump_a.click()
            # 切换到新标签页的window
            driver.switch_to.window(driver.window_handles[-1])

            handle_jump_page(driver, data_dir_)

            driver.close()
            # 切换到新标签页的window
            driver.switch_to.window(driver.window_handles[-1])

driver.quit()
