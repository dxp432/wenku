import selenium
from selenium import webdriver
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import selenium.webdriver.support.ui as ui
import random
from selenium.common.exceptions import TimeoutException
import xlwt
import xlrd
import xlrd
import xlwt
from xlutils import copy
import re
import tkinter as tk  # 使用Tkinter前需要先导入
import os
import arrow
from auto import *

def is_here(value):
    with open('all_url.txt', "rb+") as f:
        # lines = f.readlines()  # 读取所有行
        for lines in f.readlines():
            lines1 = str(lines)
            if value in lines1:
                return True
    return False

def read_time():
     with open('time.txt', "rb") as f:
        line = f.readline().decode('utf-8')  # 读取所有行
        return str(line)

def txt_append(value):
    f= open("all_url.txt","a+")
    f.write(value + "\n")
    f.close()

def judge_url(my_url):
    """
    该函数判断一个字符串是否为百度文库
    :param url:
    :return:
    """
    if 'wenku.baidu' in my_url or 'wk.baidu' in my_url :
        return True
    else:
        return False


def new_dl_file():
    test_report = 'C:\\Users\\dxp\\Downloads\\'
    print(test_report)
    lists = os.listdir(test_report) #列出目录的下所有文件和文件夹保存到lists_
    is_exist = False
    for list in lists:
        if 'crdownload' in list:
            is_exist = True
            print('有文件名字叫未确认 871574.crdownload，正在下载中。。。。。。')

    while is_exist:
        is_exist = False
        lists = os.listdir(test_report)
        for list in lists:
            if 'crdownload' in list:
                is_exist = True
                print('有文件名字叫 未确认 871574.crdownload，正在下载中。。。。。。')
                time.sleep(1)
    print('文件已经下载好了。')
    lists.sort(key=lambda fn:os.path.getmtime(test_report + "\\" + fn))  # 按时间排序
    file_new = os.path.join(test_report,lists[-1]) # 获取最新的文件保存到file_new
    print(file_new)
    return file_new


def up_load_yunpan(file_new, driver, url_raw):
    computer_prtsc('myScreencap.png')
    computer_matchImgClick('myScreencap.png', 'menu.png')
    time.sleep(2)
    computer_prtsc('myScreencap.png')
    computer_matchImgClick('myScreencap.png', 'yunpan.png')
    time.sleep(2)
    computer_prtsc('myScreencap.png')
    computer_matchImgClick('myScreencap.png', 'yunpan_upload.png')
    time.sleep(2)
    computer_copy_2_paste(file_new)
    computer_enter()
    # 给点上传的时间
    time.sleep(10)
    # 获取文件名1606641542655.docx
    my_file = file_new.split('\\')[4]
    computer_prtsc('myScreencap.png')
    computer_matchImgClick('myScreencap.png', 'yunpan_search.png')
    time.sleep(1)
    computer_copy_2_paste(my_file)
    time.sleep(0.2)
    computer_enter()
    time.sleep(1)
    computer_prtsc('myScreencap.png')
    computer_matchImgClick('myScreencap.png', 'yunpan_select.png')
    time.sleep(0.5)
    computer_prtsc('myScreencap.png')
    computer_matchImgClick('myScreencap.png', 'yunpan_share.png')
    time.sleep(0.5)
    computer_prtsc('myScreencap.png')
    computer_matchImgClick('myScreencap.png', 'yunpan_share_ok.png')
    time.sleep(3)
    computer_prtsc('myScreencap.png')
    computer_matchImgClick('myScreencap.png', 'yunpan_copy_url.png')
    time.sleep(0.1)
    computer_matchImgClick('myScreencap.png', 'yunpan_close.png')
    time.sleep(0.5)
    computer_prtsc('myScreencap.png')
    computer_matchImgClick('myScreencap.png', 'yunpan_home.png')
    time.sleep(0.1)

    # 回复留言
    # driver.refresh()
    computer_prtsc('myScreencap.png')
    computer_matchImgClick('myScreencap.png', 'menu.png')
    time.sleep(2)
    computer_prtsc('myScreencap.png')
    computer_matchImgClick('myScreencap.png', 'web_wx.png')
    time.sleep(2)
    all_msg = driver.find_elements_by_xpath('//div[contains(@class, "wxMsg")]')
    print('all_msg')
    for msg in all_msg:
        if url_raw in msg.get_attribute('textContent'):
            my_index = all_msg.index(msg)
            phones = driver.find_elements_by_xpath('//a[contains(@class, "icon18_common reply_gray js_reply")]')
            for phone in phones:
                if phones.index(phone) == my_index:
                    phone.click()
                    print('点击')
                    time.sleep(2)
                    computer_ctrl_v()
                    time.sleep(1)
                    computer_enter()


def get_file(my_url, driver):
    url_raw = my_url
    # 打开网页
    driver_get_file = webdriver.Chrome(r'chromedriver.exe')#開啟
    my_url = str(my_url).replace('baidu.com','baiduvvv.com')
    driver_get_file.get(my_url)
    driver_get_file.refresh()
    driver_get_file.set_window_position(0, 0)
    driver_get_file.set_window_size(960,1080)

    # 点击下载
    dl_btn = driver_get_file.find_element_by_xpath('//*[@id="subbtn"]')
    dl_btn.click()
    time.sleep(5)
    dl_ing = True
    while dl_ing:
        computer_prtsc('myScreencap.png')
        if computer_if_matchImg('myScreencap.png','dl_ing.png'):
            print('还在下载中......')
            time.sleep(5)
        if computer_if_matchImg('myScreencap.png','submit.png'):
            print('输入验证码')
            computer_prtsc('myScreencap.png')
            computer_matchImgClick('myScreencap.png', 'menu.png')
            time.sleep(2)
            computer_prtsc('myScreencap.png')
            computer_matchImgClick('myScreencap.png', 'wechat.png')
            time.sleep(2)
            computer_prtsc('myScreencap.png')
            computer_matchImgClick('myScreencap.png', 'wechat_2.png')
            computer_type_input('2')
            computer_enter()
            time.sleep(2)
            computer_prtsc('myScreencap.png')
            computer_matchImg_doubleClick('myScreencap.png', 'wechat_code.png')
            time.sleep(1)
            computer_ctrl_c()
            time.sleep(1)
            # 验证码是:3694,请尽快使用！
            # 得到3694
            wincld.OpenClipboard()
            d = wincld.GetClipboardData(win32con.CF_TEXT)
            wincld.CloseClipboard()
            my_code = str(d).split(':')[1]
            my_code = my_code.split(',')[0]

            computer_matchImgClick('myScreencap.png', 'menu.png')
            time.sleep(2)
            computer_prtsc('myScreencap.png')
            computer_matchImgClick('myScreencap.png', 'web.png')
            time.sleep(2)
            computer_prtsc('myScreencap.png')
            computer_matchImgClick('myScreencap.png', 'web_enter_code.png')
            computer_copy_2_paste(my_code)
            computer_matchImgClick('myScreencap.png','submit.png')
            time.sleep(2)


        if computer_if_matchImg('myScreencap.png','dl_done.png'):
            print('下载已经完成')
            dl_ing = False
            my_new_file = new_dl_file()
            print(my_new_file)
            up_load_yunpan(my_new_file, driver, url_raw)

def write_txt_append(value, driver):
    if judge_url(value):
        if is_here(value):
            print('url已经存在')
        else:
            # 保留HTTP:// 去掉中文、特殊字符。
            value = 'http' + str(value).split('http')[1]
            txt_append(value)
            get_file(value, driver)


def write_excel_xls_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数

    # print(sheet2.nrows)
    # sheet的名称，行数，列数
    # print sheet2.name,sheet2.nrows,sheet2.ncols
    # for i in range(worksheet.nrows):
    # # # 获取整行和整列的值（数组）
    #     rows =worksheet.cell(i, 0).value.encode('utf-8') # 获取第i行内容
    #     rows = rows.strip()
    #     value = value.strip()
    #     if value == rows:
    #         break
    if value in worksheet.col_values(0):
        print('存在')
        pass
    else:
        if judge_url(value):
            new_workbook = copy.copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
            new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
            # for i in range(0, index):
            #     for j in range(0, len(value[i])):
            new_worksheet.write(rows_old, 0, value)  # 追加写入数据，注意是从i+rows_old行开始写入
            new_workbook.save(path)  # 保存工作簿
            print("xls格式表格【追加】写入数据成功！")
    

def get_text():
    # driver = webdriver.Chrome(r'C:\\Program Files (x86)\\Google\\Chrome\Application\\chromedriver.exe')#開啟
    driver = webdriver.Chrome(r'chromedriver.exe')#開啟

    driver.get("https://mp.weixin.qq.com/")
    driver.set_window_position(660, 0)
    driver.set_window_size(660,1080)
    # driver.maximize_window()

    # 创建一个workbook 设置编码
    workbook = xlwt.Workbook(encoding = 'utf-8')
    # 创建一个worksheet
    worksheet = workbook.add_sheet('Worksheet')
    workbook.save('Excel_test.xls')

    input("输入任意值，然后按回车，以便开始爬取URL:")

    # 消息管理
    msg_btn = driver.find_element_by_xpath('//*[@id="menuBar"]/li[6]/ul/li[1]/a')
    msg_btn.click()
    time.sleep(1)
    ddown_btn = driver.find_element_by_xpath('//*[@id="dayselect"]/a')
    ddown_btn.click()
    time.sleep(1)
    # 今天
    today_btn = driver.find_element_by_xpath('//*[@id="dayselect"]/div/ul/li[2]/a')
    today_btn.click()
    time.sleep(2)
    
    while True:
        driver.refresh()

        # 判断时间
        all_times = driver.find_elements_by_xpath('//div[@class="message_time"]')
        for time_tr in all_times:
            if arrow.get(time_tr.get_attribute('textContent'), 'HH:mm') >= arrow.get(read_time(), 'HH:mm'):
                my_index = all_times.index(time_tr)
                phones = driver.find_elements_by_xpath('//div[@class="message_content text"]//div')
                for phone in phones:
                    if phones.index(phone) == my_index:
                        write_txt_append(phone.get_attribute('textContent'), driver)
                print("正在读取")

        time.sleep(5)

get_text()
# up_load_yunpan(new_dl_file())
