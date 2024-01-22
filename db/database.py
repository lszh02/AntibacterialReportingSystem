import copy
import json
import os
import random
import time

import ngender
import win32api
import xlrd

current_path = os.path.dirname(__file__)


def read_excel(excel_file, sheet_name):
    try:
        excel_file = xlrd.open_workbook(excel_file)
        sheet = excel_file.sheet_by_name(sheet_name)
        print(f"已打开工作表,获取{sheet_name}")
        return sheet
    except Exception as e:
        print('打开文件出错，错误为：', e)


class Prescription:
    def __init__(self, prescription_data_sheet, drug_data_sheet):
        """
        获取Execl中的处方基本信息sheet和用药信息sheet，如果科室有新增，将触发科室字典更新。
        :param prescription_data_sheet: 处方基本信息表
        :param drug_data_sheet: 药品信息表
        """
        self._prescription_data_sheet = prescription_data_sheet
        self._drug_data_sheet = drug_data_sheet
        self.department_dict = self.update_department_dict()

    def get_prescription_data(self, path=None):
        """
        根据excel表格中的数据提取处方信息，包括行号、处方ID、科室、患者姓名、性别、年龄、诊断、总金额、药品信息等。
        将提取的信息以json格式保存。
        :param path: 以json格式保存的路径
        :return: 处方信息的列表
        """
        prescription_data = []
        prescription_info_dict = dict.fromkeys(["row_num", 'prescription_id', 'department', 'patient_name',
                                                'age', 'diagnosis', "gender", 'drug_info', "total_money"])

        row_num = 1
        while row_num < self._prescription_data_sheet.nrows:
            # 获取行号、处方ID、科室、患者姓名、性别、年龄、诊断、总金额、药品信息等。
            prescription_info_dict['row_num'] = row_num
            prescription_info_dict['prescription_id'] = self._prescription_data_sheet.cell(row_num, 0).value
            prescription_info_dict['department'] = self._prescription_data_sheet.cell(row_num, 1).value
            prescription_info_dict['patient_name'] = self._prescription_data_sheet.cell(row_num, 2).value
            prescription_info_dict['age'] = self._prescription_data_sheet.cell(row_num, 3).value
            prescription_info_dict['diagnosis'] = self._get_diagnosis_info(row_num)
            prescription_info_dict['gender'] = self._guess_gender(row_num, prescription_info_dict['department'],
                                                                  prescription_info_dict['patient_name'],
                                                                  prescription_info_dict['diagnosis'])
            prescription_info_dict['drug_info'] = self._get_drug_data(prescription_info_dict['prescription_id'])
            prescription_info_dict['total_money'] = Prescription.get_total_money(prescription_info_dict['drug_info'])
            # 将一条处方信息字典添加进prescription_data列表
            prescription_data.append(copy.deepcopy(prescription_info_dict))

            row_x = 1
            while row_num + row_x < self._prescription_data_sheet.nrows:
                if self._prescription_data_sheet.cell_type(row_num + row_x, 0) == 0:
                    row_x += 1
                else:
                    break
            row_num += row_x

        if path:
            try:
                with open(file=rf'{path}\prescription_data.json', mode='w', encoding='utf-8') as f:
                    json.dump(prescription_data, f, ensure_ascii=False, indent=2)
            except Exception as e:
                print('保存json出错：', e)
            print(f'已提取所有处方信息，并以json格式保存于{path}')
            return prescription_data
        else:
            print('已返回处方信息列表')
            return prescription_data

    @staticmethod
    def save_department_dict(department_dict):
        try:
            with open(file=rf'{current_path}\department_dict.json', mode='w', encoding='utf-8') as f:
                json.dump(department_dict, f, ensure_ascii=False, indent=2)
                print(rf'已更新科室字典，并保存于{current_path}\department_dict.json')
        except Exception as e:
            print('已更新科室字典，但以json格式保存科室字典时出错：', e)

    @staticmethod
    def get_department_dict():
        try:
            with open(file=rf'{current_path}\department_dict.json', mode='r', encoding='utf-8') as f:
                dep_dict = json.load(f)
                print(rf'于{current_path}\department_dict.json成功读取科室字典！')
                return dep_dict
        except Exception as e:
            print('读取科室字典时出错：', e)

    def update_department_dict(self):
        """(手动）更新科室字典。"""
        department_dict = Prescription.get_department_dict()
        l1 = len(department_dict)
        row_num = 1
        while row_num < self._prescription_data_sheet.nrows:
            # 获取Excel表中所有科室名称
            dep_chinese_name = self._prescription_data_sheet.cell(row_num, 1).value
            if dep_chinese_name not in department_dict:
                dep_pic_name = input(f'{dep_chinese_name}  未关联对应科室字典，请输入"dep_name格式“！')
                department_dict[dep_chinese_name] = dep_pic_name  # 增加一条，更新字典
                print(f"科室字典新增一条：{dep_chinese_name}:{dep_pic_name}")
                print(f'请截图并命名为{dep_pic_name}.png，存入res/image/menzhen(或jizhen)_image/dep_image文件夹中！点击右键继续')
                while True:
                    time.sleep(0.001)
                    if win32api.GetKeyState(0x02) < 0:
                        # up = 0 or 1, down = -127 or -128
                        break
            row_x = 1
            while row_num + row_x < self._prescription_data_sheet.nrows:
                if self._prescription_data_sheet.cell_type(row_num + row_x, 0) == 0:
                    row_x += 1
                else:
                    break
            row_num += row_x

        l2 = len(department_dict)
        if l2 > l1:
            Prescription.save_department_dict(department_dict)
        else:
            print('读取科室信息无新增，科室字典无需更新！')
        return department_dict

    @staticmethod
    def get_total_money(drug_info):
        total_money = 0
        for drug in drug_info:
            total_money += drug.get('money')
        return total_money

    def _get_diagnosis_info(self, row_num):
        diagnosis_list = []
        if ',' in self._prescription_data_sheet.cell(row_num, 4).value:
            # 急诊处方的诊断信息是以','分割保存在一个excel单元格中
            diagnosis_list = self._prescription_data_sheet.cell(row_num, 4).value.split(',')
        else:
            # 门诊处方的诊断信息分别保存在不同的excel单元格中（多行）
            diagnosis_list.append(self._prescription_data_sheet.cell(row_num, 4).value)
        # 门诊处方有多个诊断时（多行），需要下探，列表追加诊断
        row_x = 1
        while row_num + row_x < self._prescription_data_sheet.nrows:
            if self._prescription_data_sheet.cell_type(row_num + row_x, 0) == 0:
                diagnosis_list.append(self._prescription_data_sheet.cell(row_num + row_x, 4).value)
                row_x += 1
            else:
                break
        return diagnosis_list

    def _get_drug_data(self, prescription_id):
        """
        根据excel表格中的处方ID提取每张处方的药品信息，包括药品名称、规格、用法、剂量、剂量单位、频率、数量、金额等。
        将提取的信息以json格式保存。
        :param prescription_id: 处方ID
        :return: 该处方的药品信息
        """
        drug_data = []
        # drug_data用来保存药品的信息：药品名称、规格、用法、剂量、剂量单位、频率、数量、金额等
        drug_info_dict = dict.fromkeys(
            ['drug_name', 'specifications', 'usage', 'dose', 'doseUnit', 'frequency', 'quantity', 'money'])

        row_num = 1
        while row_num < self._drug_data_sheet.nrows:
            if prescription_id == self._drug_data_sheet.cell(row_num, 0).value:  # 匹配ID
                drug_info_dict['drug_name'] = self._drug_data_sheet.cell(row_num, 1).value  # 取第2列：药名
                drug_info_dict['specifications'] = self._drug_data_sheet.cell(row_num, 2).value  # 取第3列：规格
                drug_info_dict['usage'] = self._drug_data_sheet.cell(row_num, 3).value  # 取第4列：用法
                drug_info_dict['dose'] = self._drug_data_sheet.cell(row_num, 4).value  # 取第5列：剂量
                drug_info_dict['doseUnit'] = self._drug_data_sheet.cell(row_num, 5).value  # 取第6列：剂量单位
                drug_info_dict['frequency'] = self._drug_data_sheet.cell(row_num, 6).value  # 取第7列：频率
                drug_info_dict['quantity'] = self._drug_data_sheet.cell(row_num, 7).value  # 取第8列：数量
                drug_info_dict['money'] = self._drug_data_sheet.cell(row_num, 8).value  # 取第9列：金额
                drug_data.append(copy.deepcopy(drug_info_dict))  # 一条药品信息，存入列表

                row_x = 1
                while row_num + row_x < self._drug_data_sheet.nrows:
                    # 一张处方多个药品时，列表追加其他药品信息
                    if self._drug_data_sheet.cell_type(row_num + row_x, 0) == 0:
                        drug_info_dict['drug_name'] = self._drug_data_sheet.cell(row_num + row_x, 1).value  # 取第2列：药名
                        drug_info_dict['specifications'] = self._drug_data_sheet.cell(row_num + row_x,
                                                                                      2).value  # 取第3列：规格
                        drug_info_dict['usage'] = self._drug_data_sheet.cell(row_num + row_x, 3).value  # 取第4列：用法
                        drug_info_dict['dose'] = self._drug_data_sheet.cell(row_num + row_x, 4).value  # 取第5列：剂量
                        drug_info_dict['doseUnit'] = self._drug_data_sheet.cell(row_num + row_x, 5).value  # 取第6列：剂量单位
                        drug_info_dict['frequency'] = self._drug_data_sheet.cell(row_num + row_x, 6).value  # 取第7列：频率
                        drug_info_dict['quantity'] = self._drug_data_sheet.cell(row_num + row_x, 7).value  # 取第8列：数量
                        drug_info_dict['money'] = self._drug_data_sheet.cell(row_num + row_x, 8).value  # 取第9列：金额
                        drug_data.append(copy.deepcopy(drug_info_dict))
                        row_x += 1
                    else:
                        break
                break
            row_num += 1
        return drug_data

    def _guess_gender(self, row_num, department, patient_name, diagnosis_list):
        gender = random.choice(["man", "woman"])
        # 根据名字猜性别
        try:
            if ngender.guess(patient_name)[0] == 'male':
                gender = "man"
            elif ngender.guess(patient_name)[0] == 'female':
                gender = "woman"
        except AssertionError:
            print('-' * 20 + f'第{row_num}行的{patient_name}可能是外国人，无法猜测性别！' + '-' * 20)

        # 根据诊断猜性别
        for diagnosis in diagnosis_list:
            for i in {'子宫', '卵巢', '阴道', '妊娠', '孕', '乳腺'}:
                if i in diagnosis:
                    gender = "woman"
                    break
            for i in {'包皮', '龟头', '睾丸', '前列腺', '阴茎'}:
                if i in diagnosis:
                    gender = "man"
                    break

        # 如果科室为妇产科，则性别为女
        if self.department_dict.get(department) == 'dep_fuchan':
            gender = "woman"
        return gender

