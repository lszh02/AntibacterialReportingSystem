import os.path

# import pyautogui
import time
import win32api
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

current_path = os.path.dirname(__file__)
res_path = os.path.join(os.path.abspath(os.path.join(current_path, '../..')), 'res')

driver = webdriver.Chrome()  # 启动浏览器
driver.implicitly_wait(60)  # 隐式等待
wait = WebDriverWait(driver, 60)  # 显式等待

driver.get("http://y.chinadtc.org.cn/login")  # 打开网址
driver.find_element(By.CSS_SELECTOR, "#account").clear()  # 清除输入框数据
driver.find_element(By.CSS_SELECTOR, "#account").send_keys('440306311001')  # 输入账号
driver.find_element(By.CSS_SELECTOR, "#accountPwd").clear()  # 清除输入框数据
driver.find_element(By.CSS_SELECTOR, "#accountPwd").send_keys('NYDyjk233***')  # 输入密码
driver.find_element(By.CSS_SELECTOR, "#loginBtn").click()  # 单击登录

driver.find_element(By.CSS_SELECTOR, 'input[value="确定"]').click()  # 单击登录
wait.until(ec.alert_is_present())
driver.switch_to.alert.accept()

driver.find_element(By.CSS_SELECTOR, 'a[title="录入功能"]').click()  # 单击登录
driver.find_element(By.CSS_SELECTOR, 'a[href="/entering/mjz/index/mjztype/1"]').click()  # 单击"门诊处方用药录入"


def mouse_click(img, click_times=1, l_or_r="left"):
    # 定义鼠标事件
    # pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159
    while True:
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=click_times, interval=0.2, duration=0.2, button=l_or_r)
            break
        print("未找到匹配图片,0.2秒后重试")
        time.sleep(0.2)


class PrescriptionReport:
    def __init__(self, webdriver, one_prescription_info, dep_dict, ddd_drug_dict):
        """
        传入selenium的webdriver，一条处方信息，科室字典，抗菌药字典，执行网页自动上报。
        :param webdriver: selenium的webdriver
        :param one_prescription_info: 一条处方信息
        :param dep_dict: 科室字典
        :param ddd_drug_dict: 抗菌药字典
        """
        self.webdriver = webdriver
        self.prescription_info = one_prescription_info
        self.dep_dict = dep_dict
        self.ddd_drug_dict = ddd_drug_dict

    def do_report(self):
        # 选择科室
        self.input_department_name()

        # 输入年龄
        self.input_age()

        # 选择性别
        self.input_gender()

        # 输入处方金额
        self.input_total_money()

        # 输入药品总数
        self.input_quantity_of_drugs()

        # 判断是否注射剂
        self.injection_or_not()

        # 输入诊断
        self.input_diagnosis()

        # 保存数据
        self.webdriver.find_element(By.CSS_SELECTOR, 'input[value="保存门诊处方用药情况调查表"]').click()  # 单击保存
        wait.until(ec.alert_is_present())
        self.webdriver.switch_to.alert.accept()
        print("保存数据")

        # 判断是否有抗菌药物
        self.antibacterial_or_not()

    def input_department_name(self):
        # 通过科室字典dep_dict获取科室对应的网络上报名称
        dep_web_name = self.dep_dict.get(self.prescription_info.get("department"))

        # 通过Select对象,选中对应科室
        Select(self.webdriver.find_element(By.ID, "department")).select_by_visible_text(dep_web_name)
        print(f'选择科室：{self.prescription_info.get("department")}')
        return f'选择科室：{self.prescription_info.get("department")}'

    def input_age(self):
        # 创建Select对象
        age_sel = Select(self.webdriver.find_element(By.ID, "ageSel"))  # 可选项有‘岁、月、周、天‘
        age = self.prescription_info.get("age")
        if '岁' in age:
            age = age.split('岁')[0]
            age_sel.select_by_visible_text('岁')
            self.webdriver.find_element(By.ID, "age").send_keys(age)
            return f'输入年龄:{age}岁'
        elif '月' in age:
            age = age.split('月')[0]
            age_sel.select_by_visible_text('月')
            self.webdriver.find_element(By.ID, "age").send_keys(age)
            return f'输入年龄:{age}月'
        elif '周' in age:
            age = age.split('周')[0]
            age_sel.select_by_visible_text('周')
            self.webdriver.find_element(By.ID, "age").send_keys(age)
            return f'输入年龄:{age}周'
        elif '天' in age:
            age = age.split('天')[0]
            age_sel.select_by_visible_text('天')
            self.webdriver.find_element(By.ID, "age").send_keys(age)
            return f'输入年龄:{age}天'

    def input_gender(self):
        gender = self.prescription_info.get("gender")
        if gender == 'man':
            self.webdriver.find_element(By.ID, 'sexM').click()  # 'sexM'为男
            return f"选择性别：男"
        elif gender == 'woman':
            self.webdriver.find_element(By.ID, 'sexW').click()  # 'sexW'为女
            return f"选择性别：女"

    def input_total_money(self):
        money = self.prescription_info.get("total_money")
        self.webdriver.find_element(By.ID, 'outAmount').send_keys(round(money, 2))

        if money > 10000:
            # fixme
            # pyautogui.alert(text='药品金额大于10000元！', title='请确认：', button='YES')
            pass
        return f'输入处方金额:{round(money, 2)}'

    def input_quantity_of_drugs(self):
        drugs_count = len(self.prescription_info.get("drug_info"))
        self.webdriver.find_element(By.ID, 'outDrugs').send_keys(drugs_count)
        return f'输入药品数量:{drugs_count}'

    def injection_or_not(self):
        inj_count = 0
        for drug in self.prescription_info.get('drug_info'):
            for i in ['注射', '狂犬病疫苗', '破伤风']:
                if i in drug.get('drug_name'):
                    inj_count += 1
                    break
        if inj_count != 0:
            self.webdriver.find_element(By.ID, 'infusionY').click()
            self.webdriver.find_element(By.ID, 'infusionNum').send_keys(inj_count)
            return f'输入注射剂数量:{inj_count}'
        else:
            return '该处方中无注射剂'

    def input_diagnosis(self):
        diagnosis_list = self.prescription_info.get("diagnosis")
        # 诊断可以输入1-5个
        for i in range(min(len(diagnosis_list), 5)):
            diagnosis = diagnosis_list[i]
            if '癌' in diagnosis_list[i]:
                diagnosis.replace('癌', '肿瘤')
            if '泌尿系感染' in diagnosis_list[i]:
                diagnosis.replace('泌尿系感染', '泌尿道感染')

            self.webdriver.find_element(By.ID, 'diagnosisName1').click()
            self.webdriver.find_element(By.ID, 'searchDiagnosis').send_keys(diagnosis)
            self.webdriver.find_element(By.CSS_SELECTOR, '.diagnosisShade input[value="查询"]').click()
            # 获取网络诊断列表，与输入的诊断进行匹配
            web_diagnosis_list = self.webdriver.find_elements(By.CSS_SELECTOR, '#ceng-diag table .nameHtml a')
            for web_diagnosis in web_diagnosis_list:
                if web_diagnosis.text == diagnosis:
                    web_diagnosis.click()
                    print(f'输入诊断：{diagnosis}')
                    break
            else:
                print('请手动输入诊断！')
                while True:
                    time.sleep(0.001)
                    if win32api.GetKeyState(0x02) < 0:
                        # up = 0 or 1, down = -127 or -128
                        break
        return f'输入诊断：{diagnosis_list}'

    def antibacterial_or_not(self):
        drug_list = self.prescription_info.get('drug_info')
        for one_drug_info in drug_list:
            drug_name = one_drug_info.get('drug_name')
            if drug_name in self.ddd_drug_dict:
                self.webdriver.find_element(By.CSS_SELECTOR,
                                            '#outpatientTable tr:nth-child(2) div:nth-child(2)').click()  # 单击“有”
                self.webdriver.find_element(By.CSS_SELECTOR, '#outpatientTable input[value="录入详细信息"]').click()  # 单击录入
                self.input_antibacterial(drug_list)
                break
        else:
            print('该处方中无抗菌药物！')
            return '该处方中无抗菌药物！'

    def input_antibacterial(self, drug_list):
        antibacterial_list = []
        for one_drug_info in drug_list:
            drug_name = one_drug_info.get('drug_name')
            if drug_name in self.ddd_drug_dict:
                antibacterial_list.append(drug_name)
                # 输入抗菌药名称
                self.webdriver.find_element(By.ID, 'medicineName').click()
                self.webdriver.find_element(By.ID, 'searchDrugs').send_keys(self.ddd_drug_dict.get(drug_name))
                # fixme 尚未实现自动选择药品规格
                while True:
                    time.sleep(0.001)
                    if win32api.GetKeyState(0x02) < 0:
                        # up = 0 or 1, down = -127 or -128
                        break

                # 输入抗菌药金额
                self.webdriver.find_element(By.ID, 'amountOutpatient').send_keys(one_drug_info.get('money'))

                # 输入抗菌药数量
                self.webdriver.find_element(By.ID, 'totalMedicine').send_keys(one_drug_info.get('quantity'))
                # fixme 尚未实现该处细节操作自动化
                while True:
                    time.sleep(0.001)
                    if win32api.GetKeyState(0x02) < 0:
                        # up = 0 or 1, down = -127 or -128
                        break

                # 保存抗菌药物
                self.webdriver.find_element(By.ID, "saveOutpatientDetail('mz')").click()
                wait.until(ec.alert_is_present())
                self.webdriver.switch_to.alert.accept()
                print(f"输入抗菌药物:{drug_name}")

        self.webdriver.find_element(By.CSS_SELECTOR, 'input[value = "返回门诊处方用药情况调查表"]').click()
        return f'输入抗菌药物{len(antibacterial_list)}个：{antibacterial_list}'


class JzPrescriptionReport(PrescriptionReport):
    def input_department_name(self):
        pass


if __name__ == '__main__':
    data = None
    dep_dict = None
    ddd_drug_dict = None

    r = PrescriptionReport(driver, data, dep_dict, ddd_drug_dict)
    r.do_report()
