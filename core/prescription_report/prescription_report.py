import os.path

import time
import win32api
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

from core.ddd_report.ddd_report import DDDReport
from db import database
from db.database import read_excel

current_path = os.path.dirname(__file__)
res_path = os.path.join(os.path.abspath(os.path.join(current_path, '../..')), 'res')


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

    # web_driver.find_element(By.ID, 'report').click()  # 单击登录
    # # 选择时间
    # web_driver.find_element(By.CSS_SELECTOR, 'input[value="确定"]').click()  # 单击登录
    # wait.until(ec.alert_is_present())
    # web_driver.switch_to.alert.accept()
    #
    # web_driver.find_element(By.CSS_SELECTOR, 'a[title="录入功能"]').click()  # 单击登录
    # if mode == 1:
    #     web_driver.find_element(By.CSS_SELECTOR, 'a[href="/entering/mjz/index/mjztype/1"]').click()  # 单击"门诊处方用药录入"
    # elif mode == 2:
    #     web_driver.find_element(By.CSS_SELECTOR, 'a[href="/entering/mjz/index/mjztype/2"]').click()  # 单击"急诊处方用药录入"
    # elif mode == 3:
    #     web_driver.find_element(By.CSS_SELECTOR, 'a[href="/entering/kjyxh/"]').click()  # 单击"抗菌药物消耗情况录入"


class PrescriptionReport:
    def __init__(self, one_prescription_info, dep_dict, ddd_drug_dict, web_driver, wait):
        """
        传入一条处方信息，科室字典，抗菌药字典和webdriver对象，执行网页自动上报。
        :param one_prescription_info: 一条处方信息
        :param dep_dict: 科室字典
        :param ddd_drug_dict: 抗菌药字典
        :param web_driver: selenium的webdriver
        """
        self.prescription_info = one_prescription_info
        self.dep_dict = dep_dict
        self.ddd_drug_dict = ddd_drug_dict
        self.web_driver = web_driver
        self.wait = wait

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
        self.save_data()

        # 判断是否有抗菌药物
        self.antibacterial_or_not()

    def input_department_name(self):
        # 通过科室字典dep_dict获取科室对应的网络上报名称
        dep_web_name = self.dep_dict.get(self.prescription_info.get("department"))

        # 通过Select对象,选中对应科室
        Select(self.web_driver.find_element(By.ID, "department")).select_by_visible_text(dep_web_name)
        print(f'选择科室：{self.prescription_info.get("department")}')
        return f'选择科室：{self.prescription_info.get("department")}'

    def input_age(self):
        # 创建Select对象
        age_sel = Select(self.web_driver.find_element(By.ID, "ageSel"))  # 可选项有‘岁、月、周、天‘
        age = self.prescription_info.get("age")
        if '岁' in age:
            age = age.split('岁')[0]
            age_sel.select_by_visible_text('岁')
            self.web_driver.find_element(By.ID, "age").clear()  # 清除输入框数据
            self.web_driver.find_element(By.ID, "age").send_keys(age)
            print(f'输入年龄:{age}岁')
            return f'输入年龄:{age}岁'
        elif '月' in age:
            age = age.split('月')[0]
            age_sel.select_by_visible_text('月')
            self.web_driver.find_element(By.ID, "age").clear()  # 清除输入框数据
            self.web_driver.find_element(By.ID, "age").send_keys(age)
            print(f'输入年龄:{age}月')
            return f'输入年龄:{age}月'
        elif '周' in age:
            age = age.split('周')[0]
            age_sel.select_by_visible_text('周')
            self.web_driver.find_element(By.ID, "age").clear()  # 清除输入框数据
            self.web_driver.find_element(By.ID, "age").send_keys(age)
            print(f'输入年龄:{age}周')
            return f'输入年龄:{age}周'
        elif '天' in age:
            age = age.split('天')[0]
            age_sel.select_by_visible_text('天')
            self.web_driver.find_element(By.ID, "age").clear()  # 清除输入框数据
            self.web_driver.find_element(By.ID, "age").send_keys(age)
            print(f'输入年龄:{age}天')
            return f'输入年龄:{age}天'

    def input_gender(self):
        gender = self.prescription_info.get("gender")
        if gender == 'man':
            self.web_driver.find_element(By.ID, 'sexM').click()  # 'sexM'为男
            print(f"选择性别：男")
            return f"选择性别：男"
        elif gender == 'woman':
            self.web_driver.find_element(By.ID, 'sexW').click()  # 'sexW'为女
            print(f"选择性别：女")
            return f"选择性别：女"

    def input_total_money(self):
        money = self.prescription_info.get("total_money")
        self.web_driver.find_element(By.ID, 'outAmount').clear()  # 清除输入框数据
        self.web_driver.find_element(By.ID, 'outAmount').send_keys(round(money, 2))
        if money > 10000:
            self.wait.until(ec.alert_is_present())
            self.web_driver.switch_to.alert.accept()

        print(f'输入处方金额:{round(money, 2)}')
        return f'输入处方金额:{round(money, 2)}'

    def input_quantity_of_drugs(self):
        drugs_count = len(self.prescription_info.get("drug_info"))
        self.web_driver.find_element(By.ID, 'outDrugs').clear()  # 清除输入框数据
        self.web_driver.find_element(By.ID, 'outDrugs').send_keys(drugs_count)
        print(f'输入药品数量:{drugs_count}')
        return f'输入药品数量:{drugs_count}'

    def injection_or_not(self):
        inj_count = 0
        for drug in self.prescription_info.get('drug_info'):
            for i in ['注射', '狂犬病疫苗', '破伤风']:
                if i in drug.get('drug_name'):
                    inj_count += 1
                    break
        if inj_count != 0:
            self.web_driver.find_element(By.ID, 'infusionY').click()
            self.web_driver.find_element(By.ID, 'infusionNum').clear()  # 清除输入框数据
            self.web_driver.find_element(By.ID, 'infusionNum').send_keys(inj_count)
            print(f'输入注射剂数量:{inj_count}')
            return f'输入注射剂数量:{inj_count}'
        else:
            print('该处方中无注射剂')
            return '该处方中无注射剂'

    def input_diagnosis(self):
        diagnosis_list = self.prescription_info.get("diagnosis")
        # 诊断可以输入1-5个
        for i in range(min(len(diagnosis_list), 5)):
            diagnosis = diagnosis_list[i]
            if '癌' in diagnosis:
                diagnosis = diagnosis.replace('癌', '肿瘤')
            if '泌尿系感染' in diagnosis:
                diagnosis = diagnosis.replace('泌尿系感染', '泌尿道感染')

            self.web_driver.find_element(By.ID, 'diagnosisName' + f'{i + 1}').click()
            self.web_driver.find_element(By.ID, 'searchDiagnosis').send_keys(diagnosis)
            self.web_driver.find_element(By.CSS_SELECTOR, '.diagnosisShade input[value="查询"]').click()

            # 获取网络诊断列表，与输入的诊断进行匹配
            try:
                web_diagnosis_list = self.wait.until(
                    ec.visibility_of_all_elements_located((By.CSS_SELECTOR, "#ceng-diag table .nameHtml a")))  # 每一行

                # TODO 将无法与网络系统匹配的诊断导出，以便后续分析。
                # with open('diagnosis_cant_input.txt', 'a') as f:
                #     f.write(diagnosis + '\n')

                # 自动录入诊断可能无响应，此处获取手动修改后的输入诊断
                diagnosis_change = self.web_driver.find_element(By.ID, 'searchDiagnosis').get_attribute('value')
                for web_diagnosis in web_diagnosis_list:
                    # print(f'------{web_diagnosis.text}--------')
                    if web_diagnosis.text == diagnosis_change:
                        web_diagnosis.click()
                        print(f'输入诊断：{diagnosis}')
                        break
                else:
                    print('无匹配诊断，请手动输入！完成后单击右键继续……')
                    while True:
                        time.sleep(0.001)
                        if win32api.GetKeyState(0x02) < 0:
                            # up = 0 or 1, down = -127 or -128
                            break
            except Exception as e:
                print(f'输入诊断出错！错误信息为：{e}')
                print('请手动输入！完成后单击右键继续……')
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
                self.web_driver.find_element(By.CSS_SELECTOR,
                                             '#outpatientTable tr:nth-child(2) div:nth-child(2)').click()  # 单击“有”
                self.web_driver.find_element(By.CSS_SELECTOR,
                                             '#outpatientTable input[value="录入详细信息"]').click()  # 单击录入
                self.input_antibacterial(drug_list)
                break
        else:
            print('该处方中无抗菌药物！')
            return '该处方中无抗菌药物！'

    def input_antibacterial(self, drug_list):
        antibacterial_list = []
        for one_drug_info in drug_list:
            drug_name = one_drug_info.get('drug_name')
            drug_specification = one_drug_info.get('specifications')
            if drug_name in self.ddd_drug_dict:
                antibacterial_list.append(drug_name)
                # 输入抗菌药名称
                self.web_driver.find_element(By.ID, 'medicineName').click()
                self.web_driver.find_element(By.ID, 'searchDrugs').send_keys(self.ddd_drug_dict.get(drug_name))
                self.web_driver.find_element(By.CSS_SELECTOR, '#searchDrugs+input[value="查询"]').click()

                # 获取网络抗菌药物列表，与输入的药品进行匹配（名称、规格）
                web_drug_rows = self.wait.until(
                    ec.visibility_of_all_elements_located((By.CSS_SELECTOR, "#ceng-drug table tr")))  # 每一行
                one_row_unit = ''
                for one_row in web_drug_rows:
                    one_row_name = one_row.find_element(By.CSS_SELECTOR,
                                                        "#ceng-drug table td:nth-child(2)").text  # 药品网络名称
                    one_row_spec = one_row.find_element(By.CSS_SELECTOR,
                                                        "#ceng-drug table td:nth-child(3)").text  # 药品网络规格
                    one_row_unit = one_row.find_element(By.CSS_SELECTOR,
                                                        "#ceng-drug table td:nth-child(4)").text  # 药品网络单位
                    if self.ddd_drug_dict.get(drug_name) == one_row_name and drug_specification.split('*')[
                        0] == one_row_spec:
                        one_row.find_element(By.CSS_SELECTOR, "#ceng-drug table td:nth-child(6) a").click()
                        break
                else:
                    print('请选择相应的药品！右键继续……')
                    while True:
                        time.sleep(0.001)
                        if win32api.GetKeyState(0x02) < 0:
                            # up = 0 or 1, down = -127 or -128
                            break

                # 输入其他细节
                self.input_other_details(one_drug_info, one_row_unit)

                # TODO 判断抗菌药物限制级别

                # 保存抗菌药物
                self.web_driver.find_element(By.CSS_SELECTOR, 'input[value="保存抗菌药详细信息录入"]').click()
                self.wait.until(ec.alert_is_present())
                self.web_driver.switch_to.alert.accept()
                print(f"输入抗菌药物:{drug_name}")

        self.goto_main_ui()
        return f'输入抗菌药物{len(antibacterial_list)}个：{antibacterial_list}'

    def input_other_details(self, one_drug_info, one_row_unit):
        freq_web_list = ['即刻', '1/日', '2/日', '3/日', '4/日', 'q2h', 'q6h', 'q8h', 'q12h', '每晚', '其他']
        way_web_list = ['静脉滴注', '静脉泵入', '静脉推注', '肌肉注射', '静脉注射', '皮下注射', '球后注射',
                        '结膜下注射', '眼内注射', '直肠给药', '雾化吸入', '肠道准备',
                        '口服', '外用', '滴鼻', '滴耳', '滴眼', '鞘内注射', '腹膜透析', '皮试']
        dose_unit_web_list = ['克', '毫克', '万单位', '滴', 'ml', '片', '支', '粒', '瓶', '包', '袋']
        freq_dict = {'ONCE': '即刻',
                     'QD': '1/日',
                     'BID': '2/日',
                     'TID': '3/日',
                     'QID': '4/日',
                     'Q2H': 'q2h',
                     'Q6H': 'q6h',
                     'Q8H': 'q8h',
                     'Q12H': 'q2h',
                     'QN': '每晚',
                     '(空白)': '其他'}
        way_dict = {'口服': '口服',
                    '血液透析 ': '腹膜透析',
                    '外用': '外用',
                    '肌肉注射': '肌肉注射',
                    '吸入': '雾化吸入',
                    '涂眼睑内': '外用',
                    '滴眼': '滴眼',
                    '皮下注射(不带费用、耗材）': '皮下注射',
                    '静脉注射(麻醉科专用)': '静脉注射',
                    '塞肛': '外用',
                    '含服': '口服',
                    '局麻用': '肌肉注射',
                    '喷喉': '外用'}
        dose_unit_dict = {'丸': '粒', 'ug': '粒', '吸': '粒', 'ml': 'ml', '片': '片', 'g': '克', '粒': '粒',
                          'mg': '毫克'}

        try:
            # 输入抗菌药金额
            self.web_driver.find_element(By.ID, 'amountOutpatient').send_keys(one_drug_info.get('money'))

            # 输入抗菌药总数量和剂量
            self.web_driver.find_element(By.ID, 'totalMedicine').send_keys(one_drug_info.get('quantity'))
            self.web_driver.find_element(By.ID, 'onceMeter').send_keys(one_drug_info.get('dose'))

            # 选择数量单位和剂量单位
            Select(self.web_driver.find_element(By.ID, "totalMedicineUnit")).select_by_visible_text(one_row_unit)
            Select(self.web_driver.find_element(By.ID, "onceMeterUnit")).select_by_visible_text(
                dose_unit_dict.get(one_drug_info.get('doseUnit')))

            # 选择用法（频次）
            Select(self.web_driver.find_element(By.ID, "medicineFrequency")).select_by_visible_text(
                freq_dict.get(one_drug_info.get('frequency'), '其他'))

            # 选择给药途径
            Select(self.web_driver.find_element(By.ID, "medicineWay")).select_by_visible_text(
                way_dict.get(one_drug_info.get('usage'), '口服'))
        except Exception as e:
            print(f'输入抗菌药时出错，报错信息：{e}')
        print(f'请确认抗菌药输入无误后单击右键继续……')
        while True:
            time.sleep(0.001)
            if win32api.GetKeyState(0x02) < 0:
                # up = 0 or 1, down = -127 or -128
                break

    def save_data(self):
        self.web_driver.find_element(By.CSS_SELECTOR, 'input[value="保存门诊处方用药情况调查表"]').click()  # 单击保存
        self.wait.until(ec.alert_is_present())
        self.web_driver.switch_to.alert.accept()
        print("保存数据")
        return "保存数据"

    def goto_main_ui(self):
        self.web_driver.find_element(By.CSS_SELECTOR, 'input[value = "返回门诊处方用药情况调查表"]').click()


class JzPrescriptionReport(PrescriptionReport):
    def goto_main_ui(self):
        self.web_driver.find_element(By.CSS_SELECTOR, 'input[value = "返回急诊处方用药情况调查表"]').click()

    def save_data(self):
        self.web_driver.find_element(By.CSS_SELECTOR, 'input[value="保存急诊处方用药情况调查表"]').click()  # 单击保存
        self.wait.until(ec.alert_is_present())
        self.web_driver.switch_to.alert.accept()
        print("保存数据")
        return "保存数据"


if __name__ == '__main__':
    driver = webdriver.Chrome()  # 启动浏览器
    wait_time = 60  # 等待网页相应时间
    driver.implicitly_wait(wait_time)  # 隐式等待
    wait = WebDriverWait(driver, wait_time, poll_frequency=0.2)  # 显式等待

    login(driver)
    # 打开excel文件，从sheet4获取处方基本信息，从sheet5获取处方药品信息
    excel_path = r'D:\张思龙\药事\抗菌药物监测\2022年\2022年12月'
    file_name = r'门诊处方点评（100张）-20221216上午.xls'
    base_sheet = read_excel(rf"{excel_path}\{file_name}", 'Sheet3')
    drug_sheet = read_excel(rf"{excel_path}\{file_name}", 'Sheet4')

    # 实例化处方数据
    presc_data = database.Prescription(base_sheet, drug_sheet).get_prescription_data()
    # 获取字典
    dep_dict = database.Prescription.get_dep_dict()
    ddd_drug_dict = DDDReport.get_ddd_drug_dict()
    # 断点续录
    record_completed = int(input('已录入记录条数为？'))
    for one_presc in presc_data[record_completed:]:
        r = PrescriptionReport(one_presc, dep_dict, ddd_drug_dict, driver, wait)
        r.do_report()
    input('录入完成！确认输入Yes')
