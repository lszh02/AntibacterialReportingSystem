import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

current_path = os.path.dirname(__file__)

driver = webdriver.Chrome()  # 启动浏览器
wait_time = 60  # 等待网页相应时间
driver.implicitly_wait(wait_time)  # 隐式等待
wait = WebDriverWait(driver, wait_time, poll_frequency=0.2)  # 显式等待


def longin(url="http://y.chinadtc.org.cn/login", account="440306311001", pwd="NYDyjk233***"):
    driver.get(url)  # 打开网址
    driver.find_element(By.CSS_SELECTOR, "#account").clear()  # 清除输入框数据
    driver.find_element(By.CSS_SELECTOR, "#account").send_keys(account)  # 输入账号
    driver.find_element(By.CSS_SELECTOR, "#accountPwd").clear()  # 清除输入框数据
    driver.find_element(By.CSS_SELECTOR, "#accountPwd").send_keys(pwd)  # 输入密码
    driver.find_element(By.CSS_SELECTOR, "#loginBtn").click()  # 单击登录

    driver.find_element(By.CSS_SELECTOR, 'input[value="确定"]').click()  # 单击登录
    wait.until(ec.alert_is_present())
    driver.switch_to.alert.accept()

    driver.find_element(By.CSS_SELECTOR, 'a[title="录入功能"]').click()  # 单击登录
    driver.find_element(By.CSS_SELECTOR, 'a[href="/entering/mjz/index/mjztype/1"]').click()  # 单击"门诊处方用药录入"


def delete_record(del_num, web_driver=driver):
    for i in range(del_num):
        if i > 0 and i % 100 == 0:
            # 每删除100条记录需点击“下一页”
            wait.until(ec.element_to_be_clickable((By.CSS_SELECTOR, 'div.pagination a[title="下一页"]'))).click()
        # 点击删除
        wait.until(
            ec.element_to_be_clickable((By.CSS_SELECTOR, '#outpatientTable tr:nth-child(2) a[title="删除"]'))).click()

        # 点击确定删除
        wait.until(ec.alert_is_present())
        web_driver.switch_to.alert.accept()

        # 点击确定
        wait.until(ec.alert_is_present())
        web_driver.switch_to.alert.accept()
        print(f'已删除{i + 1}条记录！')
    print(f'完成！共删除{del_num}条记录！')


if __name__ == '__main__':
    longin()
    del_num = input('需要删除多少条记录？————>')
    delete_record(int(del_num))
