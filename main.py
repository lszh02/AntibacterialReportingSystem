# -*- coding:utf-8 -*-
import os
import sys
import time
import traceback

from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import pyqtSignal, QThread, QObject
from PyQt5.QtGui import QIcon, QStandardItemModel
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QInputDialog, QMessageBox, QLineEdit, QWidget, QHeaderView, \
    QAbstractItemView, QTableView
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

from core.ddd_report.ddd_report import DDDData, DDDReport
from core.prescription_report.prescription_report import PrescriptionReport, JzPrescriptionReport
from db.database import read_excel, Prescription
from res.UI.MainWindow import Ui_MainWindow
from res.UI.LoginWindow import Ui_LoginWindow


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


class PrescriptionReportThread(QThread):
    # 信号是类变量，必须在类中定义，不能在实例方法中定义，否则后面发射信号和连接槽方法时都会报错。
    prescription_sig = pyqtSignal(dict)
    prescription_progress_sig = pyqtSignal(str)
    finished_sig = pyqtSignal()

    def __init__(self, data_type, data, record_completed, web_driver, wait):
        super().__init__()
        self.data_type = data_type
        self.data = data
        self.record_completed = record_completed
        self.web_driver = web_driver
        self.wait = wait

    def run(self):
        # 获取科室字典和抗菌药物字典
        dep_dict = Prescription.get_dep_dict()
        ddd_drug_dict = DDDReport.get_ddd_drug_dict()

        # 遍历剩余处方信息
        for one_prescription in self.data[self.record_completed:]:
            self.prescription_sig.emit(one_prescription)  # 发送信号：一条处方信息
            if self.data_type == 1:
                # 按门诊处方上报
                report = PrescriptionReport(one_prescription, dep_dict, ddd_drug_dict, self.web_driver, self.wait)
            elif self.data_type == 2:
                # 按急诊处方上报
                report = JzPrescriptionReport(one_prescription, dep_dict, ddd_drug_dict, self.web_driver, self.wait)

            self.prescription_progress_sig.emit(
                '—' * 5 + f"开始填报第{self.record_completed + 1}条记录！" + '—' * 5)  # 发送信号：进度信息
            self.prescription_progress_sig.emit(report.input_department_name())  # 发送信号：选择科室
            self.prescription_progress_sig.emit(report.input_age())  # 发送信号：输入年龄
            self.prescription_progress_sig.emit(report.input_gender())  # 发送信号：选择性别
            self.prescription_progress_sig.emit(report.input_total_money())  # 发送信号：输入处方金额
            self.prescription_progress_sig.emit(report.input_quantity_of_drugs())  # 发送信号：输入药品总数
            self.prescription_progress_sig.emit(report.injection_or_not())  # 发送信号：判断是否注射剂
            self.prescription_progress_sig.emit(report.input_diagnosis())  # 发送信号：输入诊断
            self.prescription_progress_sig.emit(report.save_data())  # 发送信号：保存数据
            # fixme 此处点击保存后，页面可能还没刷新就开始填报下一条记录，导致填报诊断时报错
            time.sleep(1)
            self.prescription_progress_sig.emit(report.antibacterial_or_not())  # 发送信号：判断是否有抗菌药物

            self.record_completed += 1
            self.prescription_progress_sig.emit('—' * 5 + f"已填报{self.record_completed}条记录！" + '—' * 5)  # 发送信号：进度信息
            self.prescription_progress_sig.emit('')  # 空一行
        self.prescription_progress_sig.emit(f'填报完毕！  共计{self.record_completed}条！')
        self.prescription_progress_sig.emit('完成上报任务，10秒后将返回主界面！')
        time.sleep(10)
        self.finished_sig.emit()


class DDDReportByUI(DDDReport, QObject):
    ddd_drug_sig = pyqtSignal(dict)
    ddd_progress_sig = pyqtSignal(str)
    ddd_update_sig = pyqtSignal(str)
    finished_sig = pyqtSignal()
    isPause = False
    ddd_drug_name = ''
    start_record = None

    def __init__(self, ddd_data, driver, driver_wait):
        self.driver = driver
        self.driver_wait = driver_wait
        QObject.__init__(self)
        DDDReport.__init__(self, ddd_data, None, self.driver, self.driver_wait)

    def do_report(self):
        # 遍历剩余信息
        for one_info in self.ddd_data[self.start_record:]:
            self.ddd_drug_sig.emit(one_info)  # 发送信号：一条数据信息
            self.ddd_progress_sig.emit(f"—————开始填报第{self.start_record + 1}条记录！—————")  # 发送信号：进度信息
            self.ddd_progress_sig.emit(self.input_drug_name(one_info))  # 发送信号：输入药品名称
            self.ddd_progress_sig.emit(self.input_drug_count(one_info))  # 发送信号：输入药品数量
            self.ddd_progress_sig.emit(self.input_drug_money(one_info))  # 发送信号：输入药品金额
            self.ddd_progress_sig.emit(self.save_data())  # 发送信号：输入药品金额
            self.ddd_progress_sig.emit("保存数据！")  # 发送信号：进度信息

            self.start_record += 1
            self.ddd_progress_sig.emit(f"—————已填报{self.start_record}条记录！—————")  # 发送信号：进度信息
            self.ddd_progress_sig.emit('')  # 空一行
        self.ddd_progress_sig.emit(f'填报完毕！  共计{self.start_record}条！')
        self.finished_sig.emit()

    def input_drug_name(self, one_drug_info):
        drug_name = one_drug_info.get('drug_name')
        drug_specification = one_drug_info.get('specifications')
        # 输入抗菌药名称
        self.web_driver.find_element(By.ID, 'medicineName').click()
        if drug_name in self.ddd_drug_dict:
            self.web_driver.find_element(By.ID, 'searchDrugs').send_keys(self.ddd_drug_dict.get(drug_name))
            self.web_driver.find_element(By.CSS_SELECTOR, '#searchDrugs+input[value="查询"]').click()
        else:
            self.web_driver.find_element(By.ID, 'searchDrugs').send_keys(drug_name)
            self.web_driver.find_element(By.CSS_SELECTOR, '#searchDrugs+input[value="查询"]').click()
            # 发送信号：输入药品名称（该药不在字典中）
            self.ddd_update_sig.emit(f'{drug_name}')
            self.isPause = True
            print('暂停上报，等待更新药品字典')
            while True:
                if self.isPause:
                    time.sleep(0.1)
                    continue
                else:
                    # 增加一条，更新字典
                    self.ddd_drug_dict[drug_name] = self.ddd_drug_name
                    DDDReportByUI.update_ddd_drug_dict(self.ddd_drug_dict)
                    break

        # 获取网络抗菌药物列表，与输入的药品进行匹配（名称、规格）
        self.matching_drugs(drug_name, drug_specification)
        print(f"输入药品名称：{drug_name}")
        return f"输入药品名称：{drug_name}"


class MainWindow(QMainWindow, Ui_MainWindow):
    input_error_sig = pyqtSignal(str)

    def __init__(self, driver, driver_wait):
        super().__init__()
        self.driver = driver
        self.driver_wait = driver_wait

        self.setupUi(self)
        self.stackedWidget.setCurrentIndex(0)

        self.file_btn.clicked.connect(self.file_choose)
        self.start_btn.clicked.connect(self.set_options)
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
            self.prescription_sheet.setEnabled(False)
            self.drug_sheet.setEnabled(False)
        else:
            self.prescription_sheet.setCurrentIndex(2)
            self.drug_sheet.setCurrentIndex(3)
            self.ddd_sheet.setEnabled(False)
            self.prescription_sheet.setEnabled(True)
            self.drug_sheet.setEnabled(True)

    def file_choose(self):
        filename, filetype = QFileDialog.getOpenFileName(self, "打开文件", r"D:\张思龙\1.药事\抗菌药物监测\2023年", "全部文件(*.*)")
        if filename != "":
            self.file_path_text.setText(filename)

    def set_options(self):
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

            # 跳转DDD页面，设置表格列宽
            self.stackedWidget.setCurrentIndex(2)
            self.tabWidget.setCurrentIndex(0)
            self.set_ui_table_width()
            self.set_drugs_list(ddd_data)

            # 开启新进程进行DDD上报
            self.ddd_report(ddd_data)

        else:
            # 获取Sheet表格
            prescription_sheet = self.prescription_sheet.currentText()
            drug_sheet = self.drug_sheet.currentText()
            try:
                # 打开excel文件，获取处方基本信息和处方药品信息
                base_sheet = read_excel(excel_file, prescription_sheet)
                drug_sheet = read_excel(excel_file, drug_sheet)
                # 实例化处方数据，更新科室字典
                prescription_data = PrescriptionUpdateDep(base_sheet, drug_sheet).get_prescription_data()
            except Exception as e:
                # print(e)
                traceback.print_exc()
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
            self.report_thread = PrescriptionReportThread(data_type, prescription_data, record_completed, self.driver,
                                                          self.driver_wait)
            self.report_thread.prescription_sig.connect(self.prescription_display)  # 连接信号槽，在UI上显示处方信息
            self.report_thread.finished_sig.connect(self.prescription_report_finished)  # 连接信号槽，处方上报结束，返回主界面
            self.report_thread.prescription_progress_sig.connect(
                lambda progress_sig: self.progress.append(progress_sig))  # 连接信号槽,在UI上显示进度
            self.report_thread.start()

    def set_drugs_list(self, ddd_data):
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

    def prescription_display(self, prescription_sig):  # 这里是接收信号
        # 将处方信息显示在UI上面
        self.id.setText(prescription_sig.get('prescription_id'))
        self.dep.setText(prescription_sig.get('department'))
        self.name.setText(prescription_sig.get('patient_name'))
        self.age.setText(prescription_sig.get('age'))
        self.gender.setCurrentIndex(0 if prescription_sig.get('gender') == 'man' else 1)
        self.diagnosis.setText('')
        for i in prescription_sig.get('diagnosis'):
            self.diagnosis.append(i)

        self.drug_info.clearContents()  # 仅删除表格中数据区内所有单元格的内容
        for index, one_drug in enumerate(prescription_sig.get('drug_info')):
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
        # 获取输入内容，更新字典。
        ddd_drug_name, ok = QInputDialog.getText(self, "药品名称字典需更新", f'请输入{drug_name} 在上报系统中的名字:', QLineEdit.Normal, "")
        if ok and ddd_drug_name:
            self.ddd_reporter.ddd_drug_name = ddd_drug_name
            self.ddd_reporter.isPause = False
        else:
            # TODO 如果输入为空或者点击取消按钮，则不更新字典。
            # 后两项分别为按钮(以|隔开，共有7种按钮类型，见示例后)、默认按钮(省略则默认为第一个按钮)
            # 选择Yes代码中replay为16384, 选择No则replay为65536
            reply = QMessageBox.question(self, "异常输入", "网络系统中无对应的品种吗？", QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == 16384:
                self.input_error_sig.emit()
                pass
            else:
                pass

    def ddd_report(self, ddd_data):
        self.thread = QThread()
        self.ddd_reporter = DDDReportByUI(ddd_data, self.driver, self.driver_wait)
        self.ddd_reporter.moveToThread(self.thread)

        # 连接信号槽：
        self.thread.started.connect(self.ddd_reporter.do_report)
        self.ddd_reporter.finished_sig.connect(self.ddd_report_finished)

        # btn_ok链接取得当前选中行的index行号并开始DDD上报工作
        self.btn_ok.clicked.connect(self.get_start_num_and_report)

        # 在UI上显示药品信息和上报进度
        self.ddd_reporter.ddd_drug_sig.connect(self.ddd_display)
        self.ddd_reporter.ddd_progress_sig.connect(
            lambda ddd_progress_sig: self.ddd_progress_text.append(ddd_progress_sig))

        # 当药品字典需要更新时调用UI跨进程传参
        self.ddd_reporter.ddd_update_sig.connect(self.update_ddd_drug_name)

    def get_start_num_and_report(self):
        # 取得当前选中行的index，不选默认为-1
        current_row = self.tableView.currentIndex().row()
        if current_row == -1:
            QMessageBox.warning(self, "请选择", "未选择开始条目，请点击药品条目后重试！", QMessageBox.Yes | QMessageBox.No,
                                QMessageBox.Yes)
        else:
            self.ddd_reporter.start_record = current_row  # 将当前选中行的index设置为上报线程中的起始数
            self.thread.start()

    def ddd_report_finished(self):
        print('上报任务完成，返回主界面！')
        self.stackedWidget.setCurrentIndex(0)
        self.thread.quit()
        self.ddd_reporter.deleteLater()
        self.thread.deleteLater()

    def prescription_report_finished(self):
        print('上报任务完成，返回主界面！')
        self.stackedWidget.setCurrentIndex(0)
        self.report_thread.quit()
        self.report_thread.deleteLater()


class LoginWindow(QMainWindow, Ui_LoginWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.login_button.clicked.connect(self.login)

        # Load saved login info if available
        if os.path.exists('login_info.txt'):
            with open('login_info.txt', 'r') as f:
                lines = f.readlines()
                self.username_input.setText(lines[0].strip())
                self.password_input.setText(lines[1].strip())

    def login(self):
        # Get username and password
        username = self.username_input.text()
        password = self.password_input.text()
        self.login_button.setEnabled(False)
        self.login_button.setText("登录中……")

        # Save username and password if remember me is checked
        if self.remember_checkbox.isChecked():
            with open('login_info.txt', 'w') as f:
                f.write(username + '\n')
                f.write(password)
        else:
            # Remove saved login info
            if os.path.exists('login_info.txt'):
                os.remove('login_info.txt')

        # Start login thread
        self.login_thread = LoginThread(username, password)
        self.login_thread.login_signal.connect(self.login_result)
        self.login_thread.start()

    def login_result(self, success):
        self.login_button.setEnabled(True)
        self.login_button.setText("登录")
        if success:
            # Close login window and open main window
            self.close()
            self.main_window = MainWindow(self.login_thread.driver, self.login_thread.driver_wait)
            self.main_window.show()
        else:
            reply = QMessageBox.warning(self, "登陆失败", "登陆失败，可能是网络延迟，是否重试？", QMessageBox.Yes | QMessageBox.No,
                                        QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.login()
            else:
                self.close()


class LoginThread(QThread):
    login_signal = pyqtSignal(bool)

    def __init__(self, username, password, driver=None, wait_time=60, parent=None):
        super().__init__(parent)
        self.username = username
        self.password = password

        self.driver = driver
        self.wait_time = wait_time
        self.driver_wait = WebDriverWait(self.driver, self.wait_time, poll_frequency=0.2)  # 显式等待

    def run(self):
        if self.driver is None:
            self.driver = webdriver.Chrome()
            self.driver.implicitly_wait(self.wait_time)  # 隐式等待
            self.driver_wait = WebDriverWait(self.driver, self.wait_time, poll_frequency=0.2)  # 显式等待

        try:
            # Use selenium to logining to website
            self.login(self.driver, url="http://y.chinadtc.org.cn/login", account=self.username, pwd=self.password)
        except Exception as e:
            print(e)
            self.login_signal.emit(False)

    def login(self, web_driver, url=None, account=None, pwd=None):
        web_driver.get(url)  # 打开网址
        web_driver.find_element(By.CSS_SELECTOR, "#account").clear()  # 清除输入框数据
        web_driver.find_element(By.CSS_SELECTOR, "#account").send_keys(account)  # 输入账号
        web_driver.find_element(By.CSS_SELECTOR, "#accountPwd").clear()  # 清除输入框数据
        web_driver.find_element(By.CSS_SELECTOR, "#accountPwd").send_keys(pwd)  # 输入密码
        web_driver.find_element(By.CSS_SELECTOR, "#loginBtn").click()  # 单击登录
        # 能定位到“退出”按钮即表示登录成功
        self.driver_wait.until(ec.presence_of_element_located((By.CSS_SELECTOR, 'a[title="退出"]')))
        self.login_signal.emit(True)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    ui = LoginWindow()
    ui.setWindowIcon(QIcon('res/UI/drug.png'))
    ui.show()
    sys.exit(app.exec_())
