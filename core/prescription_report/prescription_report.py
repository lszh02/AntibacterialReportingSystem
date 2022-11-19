import os.path

import pyautogui
import time
import pyperclip
import win32api

from core.ddd_report.ddd_report import get_ddd_drug_dict
from db import database

from db.database import read_excel, get_dep_dict

current_path = os.path.dirname(__file__)
res_path = os.path.join(os.path.abspath(os.path.join(current_path, '../..')), 'res')


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
    def __init__(self, one_prescription_info, dep_dict, ddd_drug_dict):
        """
        读取一条处方信息，执行上报。
        :param one_prescription_info: 一条处方信息
        :param dep_dict: 科室字典
        :param ddd_drug_dict: 抗菌药字典
        """
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
        mouse_click(rf"{res_path}/image/prescription_image/save.png")
        mouse_click(rf"{res_path}/image/prescription_image/enter.png")
        print("保存数据")

        # 判断是否有抗菌药物
        self.antibacterial_or_not()

    def input_department_name(self):
        dep_chinese_name = self.prescription_info.get("department")
        dep_pic_name = self.dep_dict.get(dep_chinese_name)
        out_of_range = ['dep_huxi', 'dep_xinxiongwai', 'dep_yan', 'dep_xueye', 'dep_xinnei',
                        'dep_shennei', 'dep_zhongyi', 'dep_erbihou', 'dep_ganranxingjibing']

        mouse_click(rf'{res_path}/image/prescription_image/dep_1.png')
        print(f"点击：所属科室")

        if dep_pic_name in out_of_range:
            # 向下拖动滚动条
            mouse_click(rf'{res_path}/image/prescription_image/dep_2.png')
            pyautogui.dragRel(0, 200, duration=0.2)
        mouse_click(rf"{res_path}/image/prescription_image/dep_image/{dep_pic_name}.png")
        return f'选择科室：{dep_chinese_name}'

    def input_age(self):
        mouse_click(rf"{res_path}/image/prescription_image/age.png")
        print("点击：年龄")
        age = self.prescription_info.get("age")
        if '岁' in age:
            age = age.split("岁")[0]
            pyperclip.copy(age)
            pyautogui.hotkey('ctrl', 'v')
            return f"输入年龄:{age}岁"
        elif '月' in age:
            age = age.split("月")[0]
            pyperclip.copy(age)
            pyautogui.hotkey('ctrl', 'v')
            mouse_click(rf"{res_path}/image/prescription_image/age_year.png")
            mouse_click(rf"{res_path}/image/prescription_image/age_month.png")
            return f"输入年龄:{age}月"
        elif '天' in age:
            age = age.split("天")[0]
            pyperclip.copy(age)
            pyautogui.hotkey('ctrl', 'v')
            mouse_click(rf"{res_path}/image/prescription_image/age_year.png")
            mouse_click(rf"{res_path}/image/prescription_image/age_day.png")
            return f"输入年龄:{age}天"

    def input_gender(self):
        gender = self.prescription_info.get("gender")
        mouse_click(rf'{res_path}/image/prescription_image/{gender}.png')
        if gender == 'man':
            return f"选择性别：男"
        elif gender == 'woman':
            return f"选择性别：女"

    def input_total_money(self):
        mouse_click(rf"{res_path}/image/prescription_image/money.png")
        print("点击：处方金额")
        money = self.prescription_info.get("total_money")
        pyperclip.copy(round(money, 2))
        pyautogui.hotkey('ctrl', 'v')
        if money > 10000:
            pyautogui.alert(text='药品金额大于10000元！', title='请确认：', button='YES')
        return f"输入处方金额:{round(money, 2)}"

    def input_quantity_of_drugs(self):
        # 输入药品总数
        mouse_click(rf"{res_path}/image/prescription_image/drugs_count.png")
        print("点击：药品品种数")
        drugs_count = len(self.prescription_info.get("drug_info"))
        print(drugs_count)
        pyperclip.copy(drugs_count)
        pyautogui.hotkey('ctrl', 'v')
        return f"输入药品数量:{drugs_count}"

    def injection_or_not(self):
        inj_count = 0
        for drug in self.prescription_info.get('drug_info'):
            if '注射' in drug.get('drug_name') or '狂犬病疫苗' in drug.get('drug_name') or '破伤风' in drug.get('drug_name'):
                inj_count += 1
        if inj_count != 0:
            mouse_click(rf"{res_path}/image/prescription_image/inj01.png")
            mouse_click(rf"{res_path}/image/prescription_image/inj02.png")
            pyperclip.copy(inj_count)
            pyautogui.hotkey('ctrl', 'v')
            return f"输入注射剂数量:{inj_count}"
        else:
            return '该处方中无注射剂'

    def input_diagnosis(self):
        mouse_click(rf"{res_path}/image/prescription_image/diagnosis1.png")
        print("点击：诊断1")
        diagnosis = self.prescription_info.get("diagnosis")

        # FIXME 多诊断时只取前三个，其他诊断不录入
        diagnosis1 = diagnosis[0]
        if '癌' in diagnosis1:
            diagnosis1 = diagnosis1.replace('癌', '肿瘤')
        if '泌尿系感染' in diagnosis1:
            diagnosis1 = diagnosis1.replace('泌尿系感染', '泌尿道感染')

        pyperclip.copy(f"{diagnosis1}")
        pyautogui.hotkey('ctrl', 'v')
        mouse_click(rf"{res_path}/image/prescription_image/search.png")
        print('请输入诊断！')
        while True:
            time.sleep(0.001)
            if win32api.GetKeyState(0x02) < 0:
                # up = 0 or 1, down = -127 or -128
                break

        if len(diagnosis) > 1:
            mouse_click(rf"{res_path}/image/prescription_image/diagnosis2.png")
            print("点击：诊断2")
            diagnosis2 = diagnosis[1]
            if '癌' in diagnosis2:
                diagnosis2 = diagnosis2.replace('癌', '肿瘤')
            if '泌尿系感染' in diagnosis2:
                diagnosis2 = diagnosis2.replace('泌尿系感染', '泌尿道感染')

            pyperclip.copy(f"{diagnosis2}")
            pyautogui.hotkey('ctrl', 'v')
            mouse_click(rf"{res_path}/image/prescription_image/search.png")
            print('请输入诊断！')
            while True:
                time.sleep(0.001)
                if win32api.GetKeyState(0x02) < 0:
                    # up = 0 or 1, down = -127 or -128
                    break
        
        if len(diagnosis) > 2:
            mouse_click(rf"{res_path}/image/prescription_image/diagnosis3.png")
            print("点击：诊断3")
            diagnosis3 = diagnosis[2]
            if '癌' in diagnosis3:
                diagnosis3 = diagnosis3.replace('癌', '肿瘤')
            if '泌尿系感染' in diagnosis3:
                diagnosis3 = diagnosis3.replace('泌尿系感染', '泌尿道感染')

            pyperclip.copy(f"{diagnosis3}")
            pyautogui.hotkey('ctrl', 'v')
            mouse_click(rf"{res_path}/image/prescription_image/search.png")
            print('请输入诊断！')
            while True:
                time.sleep(0.001)
                if win32api.GetKeyState(0x02) < 0:
                    # up = 0 or 1, down = -127 or -128
                    break
        return f'输入诊断：{diagnosis}'

    def antibacterial_or_not(self):
        drug_list = self.prescription_info.get('drug_info')
        for one_drug_info in drug_list:
            drug_name = one_drug_info.get('drug_name')
            if drug_name in self.ddd_drug_dict:
                mouse_click(rf"{res_path}/image/prescription_image/has_antibacterial.png")
                mouse_click(rf"{res_path}/image/prescription_image/input_antibacterial.png")
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
                mouse_click(rf"{res_path}/image/prescription_image/input_antibacterial_name.png")
                pyperclip.copy(self.ddd_drug_dict.get(drug_name))
                pyautogui.hotkey('ctrl', 'v')
                mouse_click(rf"{res_path}/image/prescription_image/search.png")
                while True:
                    time.sleep(0.001)
                    if win32api.GetKeyState(0x02) < 0:
                        # up = 0 or 1, down = -127 or -128
                        break

                # 输入抗菌药金额
                mouse_click(rf"{res_path}/image/prescription_image/input_antibacterial_money.png")
                pyperclip.copy(one_drug_info.get('money'))
                pyautogui.hotkey('ctrl', 'v')

                # 输入抗菌药数量
                mouse_click(rf"{res_path}/image/prescription_image/input_antibacterial_quantity.png")
                pyperclip.copy(one_drug_info.get('quantity'))
                pyautogui.hotkey('ctrl', 'v')
                while True:
                    time.sleep(0.001)
                    if win32api.GetKeyState(0x02) < 0:
                        # up = 0 or 1, down = -127 or -128
                        break

                # 保存抗菌药物
                mouse_click(rf"{res_path}/image/prescription_image/save_antibacterial.png")
                mouse_click(rf"{res_path}/image/prescription_image/enter.png")
                print(f"输入抗菌药物:{drug_name}")
        mouse_click(rf"{res_path}/image/prescription_image/return_to_main.png")
        return f'输入抗菌药物{len(antibacterial_list)}个：{antibacterial_list}'


class JzPrescriptionReport(PrescriptionReport):
    def input_department_name(self):
        dep_chinese_name = self.prescription_info.get("department")
        dep_pic_name = self.dep_dict.get(dep_chinese_name)
        out_of_range = ['dep_xiaoernei', 'dep_ICU', 'dep_fuchan', 'dep_shenjingnei', 'dep_shenjingwai', 'dep_kouqiang',
                        'dep_putongnei', 'dep_miniaowai', 'dep_gu', 'dep_puwai']

        if dep_pic_name == 'dep_jizhen':
            print(f'选择科室：{dep_chinese_name}')
        else:
            mouse_click(rf'{res_path}/image/jizhen_image/dep_image/dep_jizhen.png')
            print(f"点击：所属科室")

            if dep_pic_name in out_of_range:
                # 向上拖动滚动条
                mouse_click(rf'{res_path}/image/jizhen_image/scroll_bar.png')
                pyautogui.dragRel(0, -200, duration=0.2)
            mouse_click(rf'{res_path}/image/jizhen_image/dep_image/{dep_pic_name}.png')
        return f'选择科室：{dep_chinese_name}'


if __name__ == '__main__':
    # 打开excel文件，从sheet4获取处方基本信息，从sheet5获取处方药品信息
    excel_path = r'D:\张思龙\药事\抗菌药物监测\2022年\2022年10月'
    file_name = r'门诊处方点评20221016下午.xls'
    base_sheet = read_excel(rf"{excel_path}\{file_name}", 'Sheet4')
    drug_sheet = read_excel(rf"{excel_path}\{file_name}", 'Sheet5')

    # 实例化处方数据
    presc_data = database.Prescription(base_sheet, drug_sheet).get_prescription_data()
    # 获取科室字典
    dep_dict = get_dep_dict()
    ddd_drug_dict = get_ddd_drug_dict()
    # 断点续录
    record_completed = int(input('已录入记录条数为？'))
    for one_presc in presc_data[record_completed:]:
        report = PrescriptionReport(one_presc, dep_dict, ddd_drug_dict)
        report.do_report()
    print(f"————已遍历所有处方，共计{record_completed}条！")
