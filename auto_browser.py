from selenium import webdriver
import time
import os
import pandas
import subprocess
import win32clipboard as w
import win32con
import datetime
import sys

while True:
    normal_hour = [7, 10]
    danger_hour = [0, 11, 12, 13, 14]
    now = datetime.datetime.now()
    now_hour = now.hour
    now_minute = now.minute

    danger_flag = now_hour in danger_hour and (5 <= now_minute < 10 or 35 <= now_minute < 40)
    normal_flag = now_hour in normal_hour and 5 <= now_minute < 10

    flag = danger_flag or normal_flag

    print(now.strftime("%Y-%m-%d %H:%M:%S"), flag)

    if flag:

        driver = webdriver.Chrome()

        driver.get('https://reported.17wanxiao.com/login.html')

        time.sleep(3)

        driver.find_element_by_name('username').send_keys('username')
        driver.find_element_by_name('miracle').send_keys('password')
        driver.find_element_by_id('btnSubmit').click()

        time.sleep(3)

        # 未打卡明细
        driver.get('https://reported.17wanxiao.com/index.html#sys/unreported2.0.html')
        time.sleep(10)
        driver.switch_to.frame(driver.find_element_by_xpath('//*[@id="rrapp"]/div[1]/section[2]/iframe'))
        driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/a[2]').click()

        time.sleep(20)

        # 获取文件名和创建文件的时间
        file_path = 'C:\\Users\\zhinushannan\\Downloads'
        lists = os.listdir(file_path)  # 列出目录的下所有文件和文件夹保存到lists
        lists.sort(key=lambda fn: os.path.getmtime(file_path + "\\" + fn))  # 按时间排序
        file_name = os.path.join(file_path, lists[-1])  # 获取最新的文件保存到file_new
        file_time = time.localtime(os.path.getmtime(file_name))

        # 编辑需要发送的信息
        data = pandas.read_excel(file_name)['姓名'].values

        if len(data) == 0:
            driver.quit()
            sys.exit()

        msg = '健康打卡自动提醒\n尚未打卡的同学如下：\n' + \
              str(data) + \
              '\n已打卡的同学请忽略此条消息\n' + \
              time.strftime("%Y-%m-%d %H:%M:%S", file_time) + \
              '\nDesign By 释治怒   Powered By kwcoder.cn'

        print(msg)

        # 将信息导入粘贴板
        w.OpenClipboard()
        w.EmptyClipboard()
        w.SetClipboardData(win32con.CF_UNICODETEXT, msg)
        w.CloseClipboard()

        # 发送
        subprocess.call('cscript autosend.vbs')

        driver.quit()

    time.sleep(300)

