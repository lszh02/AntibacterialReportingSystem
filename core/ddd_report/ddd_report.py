# -*- coding: utf-8 -*-
import copy
import json
import os

import time
import win32api
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.wait import WebDriverWait

from db.database import read_excel
from db.login import login

current_path = os.path.dirname(__file__)
db_path = os.path.join(os.path.abspath(os.path.join(current_path, '../..')), 'db')


class DDDData:
    def __init__(self, drug_data_sheet):
        """
        :param drug_data_sheet: 药品信息表
        """
        self._drug_data_sheet = drug_data_sheet

    def get_ddd_data(self):
        """
        根据excel表格中的数据提取药名、规格、用量、金额等。
        :return: DDD药品信息的列表
        """
        ddd_data = []
        # drug_data用来保存药品的信息：药品名称、规格、用量、金额等
        ddd_info_dict = dict.fromkeys(['row_num', 'drug_name', 'specifications', 'quantity', 'price', 'money'])

        row_num = 1
        while row_num < self._drug_data_sheet.nrows:
            ddd_info_dict['row_num'] = row_num
            ddd_info_dict['drug_name'] = self._drug_data_sheet.cell(row_num, 0).value
            ddd_info_dict['specifications'] = self._drug_data_sheet.cell(row_num, 1).value
            ddd_info_dict['quantity'] = self._drug_data_sheet.cell(row_num, 2).value
            ddd_info_dict['price'] = self._drug_data_sheet.cell(row_num, 3).value
            ddd_info_dict['money'] = self._drug_data_sheet.cell(row_num, 4).value
            ddd_data.append(copy.deepcopy(ddd_info_dict))  # 一条药品信息，存入列表
            row_num += 1
        return ddd_data


class DDDReport:
    def __init__(self, ddd_data, record_completed, web_driver, wait):
        self.ddd_data = ddd_data
        self.antibacterial_drugs_dict = DDDReport.get_antibacterial_drugs_dict()
        self.record_completed = record_completed
        self.web_driver = web_driver
        self.wait = wait

    @staticmethod
    def get_antibacterial_drugs_dict():
        try:
            with open(file=rf'{db_path}\antibacterial_drugs_dict.json', mode='r', encoding='utf-8') as f:
                dep_dict = json.load(f)
                print('成功读取DDD药品字典！')
                return dep_dict
        except Exception as e:
            print('打开DDD药品字典时出错：', e)

    @staticmethod
    def update_antibacterial_drugs_dict(ddd_drug_name_dict):
        try:
            with open(file=rf'{db_path}\antibacterial_drugs_dict.json', mode='w', encoding='utf-8') as f:
                json.dump(ddd_drug_name_dict, f, ensure_ascii=False, indent=2)
                print(rf'已更新DDD药品字典，并保存于{db_path}\antibacterial_drugs_dict.json')
        except Exception as e:
            print('已更新DDD药品字典，但以json格式保存科室字典时出错：', e)

    def do_report(self):
        for one_info in self.ddd_data[self.record_completed:]:
            # 输入药品名称
            self.input_drug_name(one_info)

            # 输入药品数量
            self.input_drug_count(one_info)

            # 输入药品金额
            self.input_drug_money(one_info)

            # 保存数据
            self.save_data()

            self.record_completed += 1

    def input_drug_name(self, one_drug_info):
        drug_name = one_drug_info.get('drug_name')
        drug_specification = one_drug_info.get('specifications')
        # 输入抗菌药名称
        self.web_driver.find_element(By.ID, 'medicineName').click()
        if drug_name in self.antibacterial_drugs_dict:
            self.web_driver.find_element(By.ID, 'searchDrugs').send_keys(self.antibacterial_drugs_dict.get(drug_name))
            self.web_driver.find_element(By.CSS_SELECTOR, '#searchDrugs+input[value="查询"]').click()
        else:
            self.web_driver.find_element(By.ID, 'searchDrugs').send_keys(drug_name)
            self.web_driver.find_element(By.CSS_SELECTOR, '#searchDrugs+input[value="查询"]').click()
            print(f'该药品规格为：{drug_specification}')
            value = input(f'请输入{drug_name}在上报系统中的名字——>')
            # 增加一条，更新字典
            self.antibacterial_drugs_dict[drug_name] = value
            DDDReport.update_antibacterial_drugs_dict(self.antibacterial_drugs_dict)

        # 获取网络抗菌药物列表，与输入的药品进行匹配（名称、规格）
        self.matching_drugs(drug_name, drug_specification)
        print(f"输入药品名称：{drug_name}")
        return f"输入药品名称：{drug_name}"

    def input_drug_count(self, one_drug_info):
        self.web_driver.find_element(By.ID, 'medicineNums').send_keys(one_drug_info.get('quantity'))
        print("输入药品总数:", one_drug_info.get('quantity'))
        return f"输入药品数量：{one_drug_info.get('quantity')}"

    def input_drug_money(self, one_drug_info):
        self.web_driver.find_element(By.ID, 'totalMoney').send_keys(round(one_drug_info.get('money'), 2))
        print("输入药品金额:", round(one_drug_info.get('money'), 2))
        return f"输入药品金额：{round(one_drug_info.get('money'), 2)}"

    def save_data(self):
        self.web_driver.find_element(By.CSS_SELECTOR, 'input[value="保存季度住院病人抗菌药物使用情况调查表"]').click()  # 单击保存
        self.wait.until(ec.alert_is_present())
        self.web_driver.switch_to.alert.accept()
        print("保存数据")
        return "保存数据"

    def matching_drugs(self, drug_name, drug_specification):
        # 获取网络抗菌药物列表，与输入的药品进行匹配（名称、规格）
        web_drug_rows = self.wait.until(
            ec.visibility_of_all_elements_located((By.CSS_SELECTOR, "#ceng-drug table tr")))  # 每一行
        # web_drug_rows = self.web_driver.find_elements(By.CSS_SELECTOR, "#ceng-drug table tr")  # 每一行
        for i in range(len(web_drug_rows)):
            one_row_name = self.web_driver.find_element(By.CSS_SELECTOR,
                                                        f"#ceng-drug table tr:nth-child({i + 1}) td:nth-child(2)").text  # 药品网络名称
            one_row_spec = self.web_driver.find_element(By.CSS_SELECTOR,
                                                        f"#ceng-drug table tr:nth-child({i + 1}) td:nth-child(3)").text  # 药品网络规格
            # 药品名称匹配
            if self.antibacterial_drugs_dict.get(drug_name) == one_row_name:

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
                    self.web_driver.find_element(By.CSS_SELECTOR,
                                                 f"#ceng-drug table tr:nth-child({i + 1}) td:nth-child(6) a").click()
                    break

                # 将药品规格中的'mg'转换成'g'后再匹配
                if 'mg' in local_drug_spec and 'g' in one_row_spec:
                    try:
                        # mg——>g，此处切片可能导致float()里面不是数字，导致报错。
                        if float(local_drug_spec[:-2]) / 1000 == float(one_row_spec.split('g')[0]):
                            self.web_driver.find_element(By.CSS_SELECTOR,
                                                         f"#ceng-drug table tr:nth-child({i + 1}) td:nth-child(6) a").click()
                            break
                    except Exception:
                        pass

                # 将药品规格中整型与浮点型统一：如3g和3.0g
                elif '.0g' in one_row_spec:
                    try:
                        # 此处切片可能导致float()里面不是数字，导致报错。
                        if float(local_drug_spec[:-1]) == float(one_row_spec.split('g')[0]):
                            self.web_driver.find_element(By.CSS_SELECTOR,
                                                         f"#ceng-drug table tr:nth-child({i + 1}) td:nth-child(6) a").click()
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
        return f"输入药品名称：{drug_name}"


if __name__ == '__main__':
    # 打开文件，获取sheet页
    excel_path = r'D:\张思龙\1.药事\抗菌药物监测\2023年'
    file_name = "2023年第一季度抗菌药物使用明细(1).xls"
    worksheet = read_excel(rf"{excel_path}\{file_name}", 'Sheet2')

    # 实例化处方数据
    ddd_data = DDDData(worksheet).get_ddd_data()

    web_driver = webdriver.Chrome()  # 启动浏览器
    wait_time = 100  # 等待网页相应时间
    web_driver.implicitly_wait(wait_time)  # 隐式等待
    wait = WebDriverWait(web_driver, wait_time, poll_frequency=0.2)  # 显式等待

    # 登录
    login(web_driver)

    # 断点续录
    record_completed = int(input('已录入记录条数为？'))

    report = DDDReport(ddd_data, record_completed, web_driver, wait)
    report.do_report()
    print(f"————已遍历所有处方，共计{record_completed}条！")
