import json
import time

import win32api
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait

driver = webdriver.Chrome()  # 启动浏览器
driver.implicitly_wait(20)  # 隐式等待

driver.get("http://y.chinadtc.org.cn/login")  # 打开网址
driver.find_element(By.CSS_SELECTOR, "#account").clear()  # 清除输入框数据
driver.find_element(By.CSS_SELECTOR, "#account").send_keys('440306311001')  # 输入账号
driver.find_element(By.CSS_SELECTOR, "#accountPwd").clear()  # 清除输入框数据
driver.find_element(By.CSS_SELECTOR, "#accountPwd").send_keys('NYDyjk233***')  # 输入密码
driver.find_element(By.CSS_SELECTOR, "#loginBtn").click()  # 单击登录

driver.find_element(By.CSS_SELECTOR, 'input[value="确定"]').click()  # 单击登录
time.sleep(0.5)
driver.switch_to.alert.accept()

driver.find_element(By.CSS_SELECTOR, 'a[title="录入功能"]').click()  # 单击登录
driver.find_element(By.CSS_SELECTOR, 'a[href="/entering/mjz/index/mjztype/1"]').click()  # 单击"门诊处方用药录入"

# 创建Select对象
department = Select(driver.find_element(By.ID, "department"))
# 通过 Select 对象选中科室
department.select_by_visible_text('口腔科')

# 创建Select对象
age_sel = Select(driver.find_element(By.ID, "ageSel"))
# 通过 Select 对象选中年龄
age_sel.select_by_visible_text('月')
# 输入年龄
driver.find_element(By.ID, "age").send_keys('5')

# 选择性别
driver.find_element(By.ID, 'sexM').click()  # 'sexM'为男
time.sleep(1)
driver.find_element(By.ID, 'sexW').click()  # 'sexW'为女

# 输入金额
driver.find_element(By.ID, 'outAmount').send_keys('256.12')

# 输入药品数量
driver.find_element(By.ID, 'outDrugs').send_keys('5')

# 判断是否注射剂
driver.find_element(By.ID, 'infusionY').click()
driver.find_element(By.ID, 'infusionN').click()
driver.find_element(By.ID, 'infusionY').click()
driver.find_element(By.ID, 'infusionNum').send_keys('2')

# 输入诊断
driver.find_element(By.ID, 'diagnosisName1').click()
driver.find_element(By.ID, 'searchDiagnosis').send_keys('腹痛')
driver.find_element(By.CSS_SELECTOR, '.diagnosisShade input[value="查询"]').click()  # 单击查询
diagnosis_list = driver.find_elements(By.CSS_SELECTOR, '#ceng-diag table .nameHtml a')  # 单击选择
for dia in diagnosis_list:
    if dia.text == '严重腹痛伴腹部强直':
        dia.click()
        break
else:
    print('请输入诊断！')
    while True:
        time.sleep(0.001)
        if win32api.GetKeyState(0x02) < 0:
            # up = 0 or 1, down = -127 or -128
            break
# 保存数据
driver.find_element(By.CSS_SELECTOR, 'input[value="保存门诊处方用药情况调查表"]').click()  # 单击保存
time.sleep(0.5)
driver.switch_to.alert.accept()

# 判断是否有抗菌药物
driver.find_element(By.CSS_SELECTOR, '#outpatientTable tr:nth-child(2) div:nth-child(2)').click()  # 单击保存
driver.find_element(By.CSS_SELECTOR, '#outpatientTable input[value="录入详细信息"]').click()  # 单击录入

# 录入抗菌药物
driver.find_element(By.ID, 'medicineName').click()
driver.find_element(By.ID, 'searchDrugs').send_keys('罗红霉素\n')
while True:
    time.sleep(0.001)
    if win32api.GetKeyState(0x02) < 0:
        # up = 0 or 1, down = -127 or -128
        break

# driver.find_element(By.ID, 'diagnosisName2').click()
# driver.find_element(By.ID, 'diagnosisName3').click()
# driver.find_element(By.ID, 'diagnosisName4').click()
# driver.find_element(By.ID, 'diagnosisName5').click()


input()
# # 退出
# driver.quit()
