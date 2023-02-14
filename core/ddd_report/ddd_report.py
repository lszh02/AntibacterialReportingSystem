# -*- coding: utf-8 -*-
import copy
import json
import os

import pyautogui
import time
import pyperclip
import win32api
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec

from db.database import read_excel

current_path = os.path.dirname(__file__)
db_path = os.path.join(os.path.abspath(os.path.join(current_path, '../..')), 'db')
res_path = os.path.join(os.path.abspath(os.path.join(current_path, '../..')), 'res')


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
        self.ddd_drug_dict = DDDReport.get_ddd_drug_dict()
        self.record_completed = record_completed
        self.web_driver = web_driver
        self.wait = wait

    @staticmethod
    def get_ddd_drug_dict():
        try:
            with open(file=rf'{db_path}\ddd_drug_dict.json', mode='r', encoding='utf-8') as f:
                dep_dict = json.load(f)
                print('成功读取DDD药品字典！')
                return dep_dict
        except Exception as e:
            print('打开DDD药品字典时出错：', e)

    @staticmethod
    def update_ddd_drug_dict(ddd_drug_name_dict):
        try:
            with open(file=rf'{db_path}\ddd_drug_dict.json', mode='w', encoding='utf-8') as f:
                json.dump(ddd_drug_name_dict, f, ensure_ascii=False, indent=2)
                print(rf'已更新DDD药品字典，并保存于{db_path}\ddd_drug_dict.json')
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
        # 输入抗菌药名称
        self.web_driver.find_element(By.ID, 'medicineName').click()
        drug_name = one_drug_info.get('drug_name')
        if drug_name in self.ddd_drug_dict:
            self.web_driver.find_element(By.ID, 'searchDrugs').send_keys(self.ddd_drug_dict.get(drug_name))
            self.web_driver.find_element(By.CSS_SELECTOR, '#searchDrugs+input[value="查询"]').click()
        else:
            print(f'该药品规格为：{one_drug_info.get("specifications")}')
            value = input(f'请输入{drug_name}在上报系统中的名字——>')
            # 增加一条，更新字典
            self.ddd_drug_dict[drug_name] = value
            DDDReport.update_ddd_drug_dict(self.ddd_drug_dict)

        print('请选择相应的药品！右键继续……')
        while True:
            time.sleep(0.001)
            if win32api.GetKeyState(0x02) < 0:
                # up = 0 or 1, down = -127 or -128
                break
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
