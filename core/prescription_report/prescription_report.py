import json
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
from core.delete_record.delete import login
from db.database import read_excel

current_path = os.path.dirname(__file__)


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
        self.modifying_words = self.get_modifying_words()
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

            # 对诊断进行预处理
            if '癌' in diagnosis:
                diagnosis = diagnosis.replace('癌', '肿瘤')
            if '泌尿系感染' in diagnosis:
                diagnosis = diagnosis.replace('泌尿系感染', '泌尿道感染')
            # 去掉诊断中的前后缀（修饰词）
            for i in self.modifying_words:
                if i in diagnosis:
                    diagnosis = diagnosis.replace(i, '')
                    continue

            self.web_driver.find_element(By.ID, 'diagnosisName' + f'{i + 1}').click()
            self.web_driver.find_element(By.ID, 'searchDiagnosis').send_keys(diagnosis)
            self.web_driver.find_element(By.CSS_SELECTOR, '.diagnosisShade input[value="查询"]').click()

            # 获取网络诊断列表，与输入的诊断进行匹配
            try:
                web_diagnosis_list = self.wait.until(
                    ec.visibility_of_all_elements_located((By.CSS_SELECTOR, "#ceng-diag table .nameHtml a")))  # 每一行

                # 医生的诊断可能无法与网络系统诊断匹配，如不能匹配，可手动修改后再查询。
                # 此处获取能查询到结果的“输入内容”，判断是医生的原始诊断 还是修改后的诊断？
                input_diagnosis_text = self.web_driver.find_element(By.ID, 'searchDiagnosis').get_attribute('value')
                if input_diagnosis_text != diagnosis:
                    # 将无法与网络系统匹配的诊断导出，以便后续分析。
                    with open(os.path.join(os.path.dirname(__file__), r"..\..\db\diagnosis_cant_input.txt"), 'a',
                              encoding='utf-8') as f:
                        f.write(diagnosis + '>>>' + input_diagnosis_text + '\n')

                for web_diagnosis in web_diagnosis_list:
                    if web_diagnosis.text == input_diagnosis_text:
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

                    # 药品名称匹配
                    if self.ddd_drug_dict.get(drug_name) == one_row_name:

                        local_drug_spec = drug_specification.split('*')[0]

                        # 下面多处的切片操作可能出现非预期的结果。

                        # 如果本地药品规格为“2ml:0.5mg”格式，则只取质量0.5mg，舍弃体积2ml:
                        if 'ml:' in local_drug_spec:
                            local_drug_spec = local_drug_spec.split('ml:')[1]
                        elif 'ml：' in local_drug_spec:
                            local_drug_spec = local_drug_spec.split('ml：')[1]

                        # 如果本地药品规格为“80mg(8万单位)”格式，则只取质量80mg，舍弃体积(8万单位):
                        if local_drug_spec.endswith(')'):
                            local_drug_spec = local_drug_spec.split('(')[0]

                        # 如果本地药品规格为“2g/支”格式，则只取质量2g，舍弃体积/支:
                        if '/支' in local_drug_spec:
                            local_drug_spec = local_drug_spec.split('/支')[0]
                        elif '/袋' in local_drug_spec:
                            local_drug_spec = local_drug_spec.split('/袋')[0]
                        elif '/瓶' in local_drug_spec:
                            local_drug_spec = local_drug_spec.split('/瓶')[0]

                        print('校正后的本地规格为', local_drug_spec)

                        # 药品规格完全匹配
                        if local_drug_spec == one_row_spec:
                            one_row.find_element(By.CSS_SELECTOR, "#ceng-drug table td:nth-child(6) a").click()
                            break

                        # 将药品规格中的'mg'转换成'g'后再匹配
                        if 'mg' in local_drug_spec and 'g' in one_row_spec:
                            try:
                                # mg——>g，此处切片可能导致float()里面不是数字，导致报错。
                                if float(local_drug_spec[:-2]) / 1000 == float(one_row_spec.split('g')[0]):
                                    one_row.find_element(By.CSS_SELECTOR, "#ceng-drug table td:nth-child(6) a").click()
                                    break
                            except Exception:
                                pass

                        # 将药品规格中整型与浮点型统一：如3g和3.0g
                        elif '.0g' in one_row_spec:
                            try:
                                # 此处切片可能导致float()里面不是数字，导致报错。
                                if float(local_drug_spec[:-1]) == float(one_row_spec.split('g')[0]):
                                    one_row.find_element(By.CSS_SELECTOR, "#ceng-drug table td:nth-child(6) a").click()
                                    break
                            except Exception:
                                pass
                else:
                    print('请选择相应的药品！右键继续……')
                    while True:
                        time.sleep(0.001)
                        if win32api.GetKeyState(0x02) < 0:
                            # up = 0 or 1, down = -127 or -128
                            break

                # 输入其他细节
                self.input_other_details(one_drug_info, one_row_unit)

                # 急诊处方需要判断抗菌药物限制级别
                self.determine_the_level_of_antimicrobial_restriction(drug_name)

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

    def determine_the_level_of_antimicrobial_restriction(self, drug_name):
        pass

    def save_data(self):
        self.web_driver.find_element(By.CSS_SELECTOR, 'input[value="保存门诊处方用药情况调查表"]').click()  # 单击保存
        self.wait.until(ec.alert_is_present())
        self.web_driver.switch_to.alert.accept()
        print("保存数据")
        return "保存数据"

    def goto_main_ui(self):
        self.web_driver.find_element(By.CSS_SELECTOR, 'input[value = "返回门诊处方用药情况调查表"]').click()

    @staticmethod
    def get_modifying_words():
        try:
            # 读取诊断前后缀文件，返回前后缀（两个）列表
            with open(file=os.path.join(os.path.dirname(__file__), r"..\..\db\modifying_words"), mode='r',
                      encoding='utf-8') as f:
                modifying_words = json.load(f)
                return modifying_words
        except Exception as e:
            print('读取诊断前后缀文件时出错：', e)


class JzPrescriptionReport(PrescriptionReport):
    def goto_main_ui(self):
        self.web_driver.find_element(By.CSS_SELECTOR, 'input[value = "返回急诊处方用药情况调查表"]').click()

    def save_data(self):
        self.web_driver.find_element(By.CSS_SELECTOR, 'input[value="保存急诊处方用药情况调查表"]').click()  # 单击保存
        self.wait.until(ec.alert_is_present())
        self.web_driver.switch_to.alert.accept()
        print("保存数据")
        return "保存数据"

    def determine_the_level_of_antimicrobial_restriction(self, drug_name):
        # 制作抗菌药物的限制级和特殊级的字典
        restricted_grade_antibacterial_drugs = ['头孢克肟', '头孢地尼', '拉氧头孢', '头孢孟多', '头孢他啶', '头孢哌酮',
                                                '哌拉西林', '莫西沙星', '妥布霉素', '链霉素', '夫西地酸', '依替米星',
                                                '厄他培南', '伏立康唑', '氟康唑']
        special_grade_antibacterial_drugs = ['美罗培南', '亚胺培南', '替加环素', '万古霉素', '利奈唑胺', '替考拉宁',
                                             '两性霉素', '伏立康唑', '卡泊芬净', '泊沙康唑']

        # 通过检索药品名称是否在相应字典之中，进而录入。
        for i in restricted_grade_antibacterial_drugs:
            if i in drug_name:
                self.web_driver.find_element(By.ID, 'yxianzhi').click()
                break

        for i in special_grade_antibacterial_drugs:
            if i in drug_name:
                self.web_driver.find_element(By.ID, 'teshu').click()
                break


if __name__ == '__main__':
    driver = webdriver.Chrome()  # 启动浏览器
    wait_time = 60  # 等待网页相应时间
    driver.implicitly_wait(wait_time)  # 隐式等待
    wait = WebDriverWait(driver, wait_time, poll_frequency=0.2)  # 显式等待

    # 登录
    login_info_path = os.path.join(os.path.join(os.path.dirname(__file__), '../..'), 'login_info.txt')
    if os.path.exists(login_info_path):
        with open(login_info_path, 'r') as f:
            lines = f.readlines()
            username_input = lines[0].strip()
            password_input = lines[1].strip()
    else:
        print('读取登陆文件出错！')
    login(driver, account=username_input, pwd=password_input)

    # 打开excel文件，从sheet4获取处方基本信息，从sheet5获取处方药品信息
    excel_path = r'D:\张思龙\1.药事\抗菌药物监测\2023年\2023年8月'
    file_name = r'202308门诊下.xls'
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
