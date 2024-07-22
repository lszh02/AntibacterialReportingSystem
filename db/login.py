import os
import time

import win32api
from selenium.webdriver.common.by import By

from config import login_info_path


def read_login_info():
    if os.path.exists(login_info_path):
        with open(login_info_path, 'r') as f:
            lines = f.readlines()
            username = lines[0].strip()
            password = lines[1].strip()
            return username, password
    else:
        print('读取登陆文件出错！')


def login(web_driver=None, url="http://y.chinadtc.org.cn/login", account=None, pwd=None):
    if account is None:
        account, pwd = read_login_info()

    web_driver.get(url)  # 打开网址
    web_driver.find_element(By.CSS_SELECTOR, "#account").clear()  # 清除输入框数据
    web_driver.find_element(By.CSS_SELECTOR, "#account").send_keys(account)  # 输入账号
    web_driver.find_element(By.CSS_SELECTOR, "#accountPwd").clear()  # 清除输入框数据
    web_driver.find_element(By.CSS_SELECTOR, "#accountPwd").send_keys(pwd)  # 输入密码
    web_driver.find_element(By.CSS_SELECTOR, "#loginBtn").click()  # 单击登录
    print('请手动选择时间和上报模块！完成后单击右键继续……')
    while True:
        time.sleep(0.001)
        if win32api.GetKeyState(0x02) < 0:
            # up = 0 or 1, down = -127 or -128
            break
