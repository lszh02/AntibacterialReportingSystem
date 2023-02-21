import os
import time

import win32api
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

current_path = os.path.dirname(__file__)


def login(web_driver, url="http://y.chinadtc.org.cn/login", account='440306311001', pwd='NYDyjk233***'):
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


def delete_record(del_num, web_driver):
    for i in range(del_num):
        if i > 0 and i % 100 == 0:
            # 每删除100条记录需点击“下一页”
            wait.until(ec.element_to_be_clickable((By.CSS_SELECTOR, 'div.pagination a[title="下一页"]'))).click()
        # 点击删除处方
        # wait.until(
        #     ec.element_to_be_clickable((By.CSS_SELECTOR, '#outpatientTable tr:nth-child(2) a[title="删除"]'))).click()

        # 点击删除DDD
        wait.until(
            ec.element_to_be_clickable((By.CSS_SELECTOR, '#prevMedicalTable tr:nth-child(2) a[title="删除"]'))).click()

        # 点击确定删除
        wait.until(ec.alert_is_present())
        web_driver.switch_to.alert.accept()

        # 点击确定
        wait.until(ec.alert_is_present())
        web_driver.switch_to.alert.accept()
        print(f'已删除{i + 1}条记录！')
    print(f'完成！共删除{del_num}条记录！')


if __name__ == '__main__':
    web_driver = webdriver.Chrome()  # 启动浏览器
    wait_time = 60  # 等待网页相应时间
    web_driver.implicitly_wait(wait_time)  # 隐式等待
    wait = WebDriverWait(web_driver, wait_time, poll_frequency=0.2)  # 显式等待
    # 登录
    login(web_driver)

    input_text = input('需要删除多少条记录？————>')
    delete_record(int(input_text), web_driver)

    input_text = input('如需继续删除，请输入数字!  退出请按回车————>')
    while True:
        if input_text == '':
            print('bye')
            break
        else:
            delete_record(int(input_text), web_driver)
            input_text = input('如需继续删除，请输入数字!  退出请按回车————>')
