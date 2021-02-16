from selenium import webdriver
import time
import os
import pandas
import subprocess
import win32clipboard as w
import win32con
import datetime

while True:
    # 获取文件名
    file_path = 'C:\\Users\\zhinushannan\\Downloads'
    lists = os.listdir(file_path)  # 列出目录的下所有文件和文件夹保存到lists
    lists.sort(key=lambda fn: os.path.getmtime(file_path + "\\" + fn))  # 按时间排序
    file_name = os.path.join(file_path, lists[-1])  # 获取最新的文件保存到file_new
    file_time = time.localtime(os.path.getmtime(file_name))

    # 编辑需要发送的信息
    data = pandas.read_excel(file_name)['姓名'].values
    msg = '自动提醒打卡程序代码测试\n所有需要健康打卡的同学如下：\n' + \
          str(data) + \
          '\n请忽略此条消息\n' + \
          time.strftime("%Y-%m-%d %H:%M:%S", file_time) + \
          '\nDesign By 释治怒   Powered By kwcoder.cn'

    # 将信息导入粘贴板
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, msg)
    w.CloseClipboard()

    # 发送
    subprocess.call('cscript autosend.vbs')

    time.sleep(10)
