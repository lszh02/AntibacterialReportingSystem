# -*- encoding = utf-8 -*-
import os
import sys
import time

import pyautogui
import pyperclip
import win32api
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import pyqtSignal, QThread, QObject
from PyQt5.QtGui import QIcon, QStandardItemModel
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QInputDialog, QMessageBox, QLineEdit, QWidget, QHeaderView, \
    QAbstractItemView, QTableView

from core.ddd_report.ddd_report import DDDData, DDDReport
from core.prescription_report.prescription_report import mouse_click, PrescriptionReport, JzPrescriptionReport
from db.database import read_excel, Prescription
from res.UI.MainWindow import Ui_MainWindow

current_path = os.path.dirname(__file__)
res_path = os.path.join(current_path, 'res')


class PrescriptionUpdateDep(Prescription, QWidget):
    def __init__(self, prescription_data_sheet, drug_data_sheet):
        QWidget.__init__(self)
        Prescription.__init__(self, prescription_data_sheet, drug_data_sheet)

    def update_dep_dict(self):
        print('重写更新科室字典被调用')
        dep_dict = Prescription.get_dep_dict()
        l1 = len(dep_dict)
        row_num = 1
        while row_num < self._prescription_data_sheet.nrows:
            # 获取Excel表中所有科室名称
            dep_chinese_name = self._prescription_data_sheet.cell(row_num, 1).value
            if dep_chinese_name not in dep_dict:
                # 第三个参数表示显示类型，可选，有正常（QLineEdit.Normal）、密碼（ QLineEdit. Password）、不显示（ QLineEdit. NoEcho）三种情况
                dep_pic_name, ok = QInputDialog.getText(self, "科室字典需更新", f'{dep_chinese_name} 未关联对应字典，请输入:',
                                                        QLineEdit.Normal, "dep_name")
                dep_dict[dep_chinese_name] = dep_pic_name  # 增加一条，更新字典

                # 后两项分别为按钮(以|隔开)、默认按钮(省略则默认为第一个按钮)
                # 选择Yes代码中replay为16384，选择No则replay为65536
                reply = QMessageBox.warning(self, "字典更新需截图",
                                            f'请截图并命名为{dep_pic_name}.png，存入res/image/menzhen(或jizhen)_image/dep_image文件夹中！',
                                            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

            row_x = 1
            while row_num + row_x < self._prescription_data_sheet.nrows:
                if self._prescription_data_sheet.cell_type(row_num + row_x, 0) == 0:
                    row_x += 1
                else:
                    break
            row_num += row_x

        l2 = len(dep_dict)
        if l2 > l1:
            Prescription.save_dep_dict(dep_dict)
        else:
            print('读取科室信息无新增，科室字典无需更新！')
        return dep_dict


class DDDReportByUI(DDDReport, QObject):
    ddd_drug_sig = pyqtSignal(dict)
    ddd_progress_sig = pyqtSignal(str)
    ddd_update_sig = pyqtSignal(str)
    start_record_sig = pyqtSignal(str)
    isPause = False
    ddd_drug_name = ''
    start_record = None

    def __init__(self, ddd_data):
        DDDReport.__init__(self, ddd_data, None)
        QObject.__init__(self)

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
            # mouse_click(rf"{res_path}/image/ddd_image/search.png")
            self.ddd_update_sig.emit(f'{drug_name}')  # 发送信号：输入药品名称
            self.isPause = True
            print('暂停上报，等待更新药品字典')
            while True:
                if self.isPause:
                    time.sleep(0.1)
                    continue
                else:
                    print('更新药品字典')
                    self.ddd_drug_dict[drug_name] = self.ddd_drug_name
                    DDDReportByUI.update_ddd_drug_dict(self.ddd_drug_dict)
                    break

        while True:
            time.sleep(0.001)
            if win32api.GetKeyState(0x02) < 0:
                # up = 0 or 1, down = -127 or -128
                break
        return f"输入药品名称：{drug_name}"

    def do_report(self):
        self.start_record_sig.emit('从哪一条开始？')  # 发送信号：从哪一条开始？
        # 等待UI传参
        while self.start_record is None:
            time.sleep(0.01)

        # 遍历剩余信息
        for one_info in self.ddd_data[self.start_record:]:
            self.ddd_drug_sig.emit(one_info)  # 发送信号：一条数据信息
            self.ddd_progress_sig.emit(f"—————开始填报第{self.start_record + 1}条记录！—————")  # 发送信号：进度信息
            self.ddd_progress_sig.emit(self.input_drug_name(one_info))  # 发送信号：输入药品名称
            self.ddd_progress_sig.emit(DDDReportByUI.input_drug_count(one_info))  # 发送信号：输入药品数量
            self.ddd_progress_sig.emit(DDDReportByUI.input_drug_money(one_info))  # 发送信号：输入药品金额
            time.sleep(0.2)
            mouse_click(rf"{res_path}/image/ddd_image/save.png")
            time.sleep(0.3)
            mouse_click(rf"{res_path}/image/ddd_image/enter.png")
            self.ddd_progress_sig.emit("保存数据！")  # 发送信号：进度信息

            self.start_record += 1
            self.ddd_progress_sig.emit(f"—————已填报{self.start_record}条记录！—————")  # 发送信号：进度信息
            self.ddd_progress_sig.emit('')  # 空一行
        self.ddd_progress_sig.emit(f'填报完毕！  共计{self.start_record}条！')
        # self.finished_sig.emit()


class PrescriptionReportThread(QThread):
    # 信号是类变量，必须在类中定义，不能在实例方法中定义，否则后面发射信号和连接槽方法时都会报错。
    presc_sig = pyqtSignal(dict)
    prescription_progress_sig = pyqtSignal(str)

    def __init__(self, data_type, data, record_completed):
        super(PrescriptionReportThread, self).__init__()
        self.data_type = data_type
        self.data = data
        self.record_completed = record_completed

    def run(self):
        # 获取科室字典和抗菌药物字典
        dep_dict = Prescription.get_dep_dict()
        ddd_drug_dict = DDDReport.get_ddd_drug_dict()

        # 遍历剩余处方信息
        for one_presc in self.data[self.record_completed:]:
            self.presc_sig.emit(one_presc)  # 发送信号：一条处方信息
            if self.data_type == 1:
                # 按门诊处方上报
                report = PrescriptionReport(one_presc, dep_dict, ddd_drug_dict)
            elif self.data_type == 2:
                # 按急诊处方上报
                report = JzPrescriptionReport(one_presc, dep_dict, ddd_drug_dict)

            self.prescription_progress_sig.emit(
                '—' * 5 + f"开始填报第{self.record_completed + 1}条记录！" + '—' * 5)  # 发送信号：进度信息
            self.prescription_progress_sig.emit(report.input_department_name())  # 发送信号：选择科室
            self.prescription_progress_sig.emit(report.input_age())  # 发送信号：输入年龄
            self.prescription_progress_sig.emit(report.input_gender())  # 发送信号：选择性别
            self.prescription_progress_sig.emit(report.input_total_money())  # 发送信号：输入处方金额
            self.prescription_progress_sig.emit(report.input_quantity_of_drugs())  # 发送信号：输入药品总数
            self.prescription_progress_sig.emit(report.injection_or_not())  # 发送信号：判断是否注射剂
            self.prescription_progress_sig.emit(report.input_diagnosis())  # 发送信号：输入诊断
            mouse_click(rf"{res_path}/image/prescription_image/save.png")
            # time.sleep(0.3)
            mouse_click(rf"{res_path}/image/prescription_image/enter.png")
            self.prescription_progress_sig.emit("保存数据！")  # 发送信号：进度信息
            self.prescription_progress_sig.emit(report.antibacterial_or_not())  # 发送信号：判断是否有抗菌药物

            self.record_completed += 1
            self.prescription_progress_sig.emit('—' * 5 + f"已填报{self.record_completed}条记录！" + '—' * 5)  # 发送信号：进度信息
            self.prescription_progress_sig.emit('')  # 空一行
        self.prescription_progress_sig.emit(f'填报完毕！  共计{self.record_completed}条！')  # 空一行


class MyWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.stackedWidget.setCurrentIndex(0)
        self.file_btn.clicked.connect(self.file_choose)
        self.start_btn.clicked.connect(self.start_report)
        self.exit_btn.clicked.connect(lambda: sys.exit())
        self.data_type.currentIndexChanged.connect(self.data_type_changed)

    def set_ui_table_width(self):
        self.drug_info.setColumnWidth(0, 120)
        self.drug_info.setColumnWidth(1, 90)
        self.drug_info.setColumnWidth(2, 50)
        self.drug_info.setColumnWidth(3, 50)
        self.drug_info.setColumnWidth(4, 40)
        self.drug_info.setColumnWidth(5, 45)
        self.drug_info.setColumnWidth(6, 55)
        self.ddd_drug_info.setColumnWidth(0, 150)
        self.ddd_drug_info.setColumnWidth(1, 100)
        self.ddd_drug_info.setColumnWidth(2, 65)
        self.ddd_drug_info.setColumnWidth(3, 65)
        self.ddd_drug_info.setColumnWidth(4, 65)

    def data_type_changed(self):
        # 获取上报类型：0为未选择，1为门诊处方，2为急诊处方，3为DDD
        if self.data_type.currentIndex() == 3:
            self.ddd_sheet.setCurrentIndex(1)
            self.ddd_sheet.setEnabled(True)
            self.presc_sheet.setEnabled(False)
            self.drug_sheet.setEnabled(False)
        else:
            self.presc_sheet.setCurrentIndex(3)
            self.drug_sheet.setCurrentIndex(4)
            self.ddd_sheet.setEnabled(False)
            self.presc_sheet.setEnabled(True)
            self.drug_sheet.setEnabled(True)

    def file_choose(self):
        filename, filetype = QFileDialog.getOpenFileName(self, "打开文件", r"D:\张思龙\药事\抗菌药物监测\2022年", "全部文件(*.*)")
        if filename != "":
            self.file_path_text.setText(filename)

    def start_report(self):
        # 获取上报类型：0为未选择，1为门诊处方，2为急诊处方，3为DDD
        data_type = self.data_type.currentIndex()
        # 获取Excel文件
        excel_file = self.file_path_text.toPlainText()

        if data_type == 0:
            QMessageBox.warning(self, "无法上报", "未选择上报类型，请检查后重试！", QMessageBox.Yes | QMessageBox.No,
                                QMessageBox.Yes)
        elif data_type == 3:
            # 获取Sheet表格
            ddd_sheet = self.ddd_sheet.currentText()
            try:
                # 打开excel文件，获取DDD数据信息
                ddd_sheet = read_excel(excel_file, ddd_sheet)
                # 实例化处方数据
                ddd_data = DDDData(ddd_sheet).get_ddd_data()
            except Exception as e:
                print(e)
                # 后两项分别为按钮(以|隔开，共有7种按钮类型，见示例后)、默认按钮(省略则默认为第一个按钮)
                QMessageBox.warning(self, "无法上报", f"原因为：{e}", QMessageBox.Yes | QMessageBox.No,
                                    QMessageBox.Yes)
                return

            # 设置表格列宽
            self.set_ui_table_width()
            # 跳转DDD页面
            self.stackedWidget.setCurrentIndex(2)
            self.tabWidget.setCurrentIndex(0)
            self.start_ddd_report(ddd_data)

        else:
            # 获取Sheet表格
            presc_sheet = self.presc_sheet.currentText()
            drug_sheet = self.drug_sheet.currentText()
            try:
                # 打开excel文件，获取处方基本信息和处方药品信息
                base_sheet = read_excel(excel_file, presc_sheet)
                drug_sheet = read_excel(excel_file, drug_sheet)
                # 实例化处方数据，更新科室字典
                presc_data = PrescriptionUpdateDep(base_sheet, drug_sheet).get_prescription_data()
            except Exception as e:
                print(e)
                # 后两项分别为按钮(以|隔开，共有7种按钮类型，见示例后)、默认按钮(省略则默认为第一个按钮)
                QMessageBox.warning(self, "无法上报", f"原因为：{e}", QMessageBox.Yes | QMessageBox.No,
                                    QMessageBox.Yes)
                return

            # 设置表格列宽
            self.set_ui_table_width()
            # 跳转页面
            self.stackedWidget.setCurrentIndex(1)

            # 后面四个数字的作用依次是 初始值 最小值 最大值 步幅
            record_completed, ok = QInputDialog.getInt(self, "是否断点续录", "请输入已录入记录条数:", 0, 0, 10000, 1)

            # 开启新进程调用上报程序
            self.report_thread = PrescriptionReportThread(data_type, presc_data, record_completed)
            self.report_thread.presc_sig.connect(self.presc_display)  # 连接信号槽，在UI上显示处方信息
            self.report_thread.prescription_progress_sig.connect(
                lambda progress_sig: self.progress.append(progress_sig))  # 连接信号槽,在UI上显示进度
            self.report_thread.start()

    def start_ddd_report(self, ddd_data):
        # 创建一个 0行4列 的标准模型
        self.model = QStandardItemModel(0, 2)
        # 设置表头标签
        self.model.setHorizontalHeaderLabels(['药名', '规格'])

        for one_info in ddd_data:
            drug_name = QtGui.QStandardItem(one_info.get('drug_name'))
            specifications = QtGui.QStandardItem(one_info.get('specifications').split('*')[0])
            record = [drug_name, specifications]
            self.model.appendRow(record)

        self.tableView.setModel(self.model)

        self.tableView.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)  # 所有列自动拉伸，充满界面
        self.tableView.setSelectionMode(QAbstractItemView.SingleSelection)  # 设置只能选中整行
        self.tableView.setSelectionBehavior(QAbstractItemView.SelectRows)  # 设置只能选中一行
        self.tableView.setEditTriggers(QTableView.NoEditTriggers)  # 不可编辑

        # 开启新进程进行DDD上报
        self.thread = QThread()
        self.ddd_report = DDDReportByUI(ddd_data)
        self.ddd_report.moveToThread(self.thread)

        # 连接信号槽：btn_ok开启进程工作，取得当前选中行的index行号，在UI上显示处方信息和进度，当字典需要更新时调用UI跨进程传参。
        self.btn_ok.clicked.connect(self.thread.start)
        self.thread.started.connect(self.ddd_report.do_report)
        self.ddd_report.start_record_sig.connect(self.set_ddd_start_record)
        self.ddd_report.ddd_drug_sig.connect(self.ddd_display)
        self.ddd_report.ddd_progress_sig.connect(lambda ddd_progress_sig: self.ddd_progress.append(ddd_progress_sig))
        self.ddd_report.ddd_update_sig.connect(self.update_ddd_drug_name)

    def presc_display(self, presc_sig):  # 这里是接收信号
        # 将处方信息显示在UI上面
        self.id.setText(presc_sig.get('prescription_id'))
        self.dep.setText(presc_sig.get('department'))
        self.name.setText(presc_sig.get('patient_name'))
        self.age.setText(presc_sig.get('age'))
        self.gender.setCurrentIndex(0 if presc_sig.get('gender') == 'man' else 1)
        self.diagnosis.setText('')
        for i in presc_sig.get('diagnosis'):
            self.diagnosis.append(i)

        self.drug_info.clearContents()  # 仅删除表格中数据区内所有单元格的内容
        for index, one_drug in enumerate(presc_sig.get('drug_info')):
            self.drug_info.setItem(index, 0, QtWidgets.QTableWidgetItem(one_drug.get('drug_name')))
            self.drug_info.setItem(index, 1, QtWidgets.QTableWidgetItem(one_drug.get('specifications')))
            self.drug_info.setItem(index, 2, QtWidgets.QTableWidgetItem(one_drug.get('usage')))
            self.drug_info.setItem(index, 3,
                                   QtWidgets.QTableWidgetItem(str(one_drug.get('dose')) + one_drug.get('doseUnit')))
            self.drug_info.setItem(index, 4, QtWidgets.QTableWidgetItem(one_drug.get('frequency')))
            self.drug_info.setItem(index, 5, QtWidgets.QTableWidgetItem(str(one_drug.get('quantity'))))
            self.drug_info.setItem(index, 6, QtWidgets.QTableWidgetItem(str(one_drug.get('money'))))

    def ddd_display(self, ddd_drug_sig):
        # 跳转上报进度页面
        self.tabWidget.setCurrentIndex(1)
        # 将处方信息显示在UI上面
        self.ddd_drug_info.clearContents()  # 仅删除表格中数据区内所有单元格的内容
        self.ddd_drug_info.setItem(0, 0, QtWidgets.QTableWidgetItem(ddd_drug_sig.get('drug_name')))
        self.ddd_drug_info.setItem(0, 1, QtWidgets.QTableWidgetItem(ddd_drug_sig.get('specifications').split('*')[0]))
        self.ddd_drug_info.setItem(0, 2, QtWidgets.QTableWidgetItem(str(ddd_drug_sig.get('quantity'))))
        self.ddd_drug_info.setItem(0, 3, QtWidgets.QTableWidgetItem(str(ddd_drug_sig.get('price'))))
        self.ddd_drug_info.setItem(0, 4, QtWidgets.QTableWidgetItem(str(ddd_drug_sig.get('money'))))

    def update_ddd_drug_name(self, drug_name):
        ddd_drug_name, ok = QInputDialog.getText(self, "药品名称字典需更新", f'请输入{drug_name} 在上报系统中的名字:', QLineEdit.Normal, "")
        self.ddd_report.ddd_drug_name = ddd_drug_name
        self.ddd_report.isPause = False

    def set_ddd_start_record(self):
        self.ddd_report.start_record = self.tableView.currentIndex().row()  # 取得当前选中行的index


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    ui = MyWindow()
    ui.setWindowIcon(QIcon('res/UI/drug.png'))
    ui.show()
    sys.exit(app.exec_())
