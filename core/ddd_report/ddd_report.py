# -*- coding: utf-8 -*-
import copy
import json
import os

import pyautogui
import time
import pyperclip
import win32api

from db.database import read_excel

current_path = os.path.dirname(__file__)
db_path = os.path.join(os.path.abspath(os.path.join(current_path, '../..')), 'db')
res_path = os.path.join(os.path.abspath(os.path.join(current_path, '../..')), 'res')


def mouse_click(img, click_times=1, l_or_r="left"):
    # 定义鼠标事件
    # pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159
    while True:
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=click_times, interval=0.2, duration=0.2, button=l_or_r)
            break
        print("未找到匹配图片,0.5秒后重试")
        time.sleep(0.5)


def get_ddd_drug_dict():
    try:
        with open(file=rf'{db_path}\ddd_drug_dict.json', mode='r', encoding='utf-8') as f:
            dep_dict = json.load(f)
            print('成功读取DDD药品字典！')
            return dep_dict
    except Exception as e:
        print('打开DDD药品字典时出错：', e)


def update_ddd_drug_dict(ddd_drug_name_dict):
    try:
        with open(file=rf'{db_path}\ddd_drug_dict.json', mode='w', encoding='utf-8') as f:
            json.dump(ddd_drug_name_dict, f, ensure_ascii=False, indent=2)
            print(rf'已更新DDD药品字典，并保存于{db_path}\ddd_drug_dict.json')
    except Exception as e:
        print('已更新DDD药品字典，但以json格式保存科室字典时出错：', e)


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


def input_drug_money(one_drug_info):
    mouse_click(rf"{res_path}/image/ddd_image/drugs_money.png")
    money = one_drug_info.get('money')
    pyperclip.copy(round(money, 2))
    pyautogui.hotkey('ctrl', 'v')
    print("输入药品金额:", money)
    return f"输入药品金额：{money}"


def input_drug_count(one_drug_info):
    mouse_click(rf"{res_path}/image/ddd_image/drugs_count.png")
    quantity = one_drug_info.get('quantity')
    pyperclip.copy(quantity)
    pyautogui.hotkey('ctrl', 'v')
    print("输入药品总数:", quantity)
    return f"输入药品数量：{quantity}"


class DDDReport:
    def __init__(self, ddd_data, ddd_drug_dict, record_completed):
        self.ddd_data = ddd_data
        self.ddd_drug_dict = ddd_drug_dict
        self.record_completed = record_completed

    def do_report(self):
        for one_info in self.ddd_data[self.record_completed:]:

            # 输入药品名称
            self.input_drug_name(one_info)

            # 输入药品数量
            input_drug_count(one_info)

            # 输入药品金额
            input_drug_money(one_info)

            # 保存数据
            mouse_click(rf"{res_path}/image/ddd_image/save.png")
            time.sleep(0.3)
            mouse_click(rf"{res_path}/image/ddd_image/enter.png")
            print("保存数据")
            self.record_completed += 1

    def input_drug_name(self, one_drug_info):
        mouse_click(rf"{res_path}/image/ddd_image/drug_name.png")
        mouse_click(rf"{res_path}/image/ddd_image/input_box.png")

        drug_name = one_drug_info.get('drug_name')
        if drug_name in self.ddd_drug_dict:
            pyperclip.copy(self.ddd_drug_dict.get(drug_name))
            pyautogui.hotkey('ctrl', 'v')
            mouse_click(rf"{res_path}/image/ddd_image/search.png")
        else:
            pyperclip.copy(drug_name)
            pyautogui.hotkey('ctrl', 'v')
            mouse_click(rf"{res_path}/image/ddd_image/search.png")
            print(f'该药品规格为：{one_drug_info.get("specifications")}')
            value = input(f'请输入{drug_name}在上报系统中的名字——>')
            # 增加一条，更新字典
            self.ddd_drug_dict[drug_name] = value
            update_ddd_drug_dict(self.ddd_drug_dict)

        while True:
            time.sleep(0.001)
            if win32api.GetKeyState(0x02) < 0:
                # up = 0 or 1, down = -127 or -128
                break
        return f"输入药品名称：{drug_name}"


if __name__ == '__main__':
    # 打开文件，获取sheet页
    excel_path = r'C:\Users\sloan\Desktop\report0'
    file_name = "最终数据-第二季度 (1).xls"
    worksheet = read_excel(rf"{excel_path}\{file_name}", 'Sheet2')
    print(rf"已打开工作表：{excel_path}\{file_name},获取Sheet1")

    # 实例化处方数据
    ddd_data = DDDData(worksheet).get_ddd_data()
    ddd_drug_name_dict = get_ddd_drug_dict()
    # 断点续录
    record_completed = int(input('已录入记录条数为？'))

    report = DDDReport(ddd_data, ddd_drug_name_dict, record_completed)
    report.do_report()
    print(f"————已遍历所有处方，共计{record_completed}条！")
