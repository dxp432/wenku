## 实时获取微信公众账号粉丝留言，返回百度文库下载链接
### 思路：
1. 用selenium打开微信公众账号登录界面。
2. 等待，用微信扫一扫登录管理员后台，继续。
3. 开始实时刷新获取微信公众账号粉丝留言。
4. 对留言的处理。

### 步骤讲解：
1. 用selenium打开“小鹏同学”微信公众账号登录界面。

```python
driver = webdriver.Chrome(r'chromedriver.exe')
driver.get("https://mp.weixin.qq.com/")
```

2. 等待，用微信扫一扫登录管理员后台，继续。登录后自动进入消息管理，如下图：

```python
input("输入任意值，然后按回车，以便开始爬取URL:")

# 消息管理
msg_btn = driver.find_element_by_xpath('//*[@id="menuBar"]/li[6]/ul/li[1]/a')
msg_btn.click()
time.sleep(1)
ddown_btn = driver.find_element_by_xpath('//*[@id="dayselect"]/a')
ddown_btn.click()
time.sleep(1)
# 选择今天
today_btn = driver.find_element_by_xpath('//*[@id="dayselect"]/div/ul/li[2]/a')
today_btn.click()
time.sleep(2)
```

3. 开始实时刷新获取微信公众账号粉丝留言。
下面的代码主要是筛选出时间，并且用xpath找到当天规定时间内人所有的留言。

```python
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
```

4. 对留言的处理。

我这里是把留言都保存到txt文件

```python
def txt_append(value):
    f= open("all_url.txt","a+")
    f.write(value + "\n")
    f.close()
```
判断留言是否之前在txt文件，
没有的话就把这个留言当作新的留言，
如果留言是“百度文库的链接”，
就把这个百度文库下载下来，
上传到百度网盘得到分享链接，
再通过微信公众平台的消息管理界面的回复分享链接给粉丝。

测试：
发送百度文库链接到“小鹏同学”微信公众账号，
2分钟之后就能获取到百度文库的下载链接。