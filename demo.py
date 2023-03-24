# -*- coding:utf-8 -*-
import os
import sys

from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import pyqtSignal, QThread, QObject
from PyQt5.QtGui import QIcon, QStandardItemModel
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QInputDialog, QMessageBox, QLineEdit, QWidget, QHeaderView, \
    QAbstractItemView, QTableView, QLabel, QPushButton, QCheckBox, QVBoxLayout, QApplication

# # Create login window
# class DisplayWindow(QWidget):
#     pass
#
#
# class LoginWindow(QWidget):
#     def __init__(self):
#         super().__init__()
#
#         # Set window title and size
#         self.setWindowTitle('Login')
#         self.setGeometry(100, 100, 300, 200)
#
#         # Create username and password labels and inputs
#         self.username_label = QLabel('Username:', self)
#         self.username_label.move(50, 50)
#         self.username_input = QLineEdit(self)
#         self.username_input.move(120, 50)
#         self.password_label = QLabel('Password:', self)
#         self.password_label.move(50, 100)
#         self.password_input = QLineEdit(self)
#         self.password_input.move(120, 100)
#         self.password_input.setEchoMode(QLineEdit.Password)
#
#         # Create login and exit buttons
#         self.login_button = QPushButton('Login', self)
#         self.login_button.move(50, 150)
#         self.login_button.clicked.connect(self.login)
#         self.exit_button = QPushButton('Exit', self)
#         self.exit_button.move(150, 150)
#         self.exit_button.clicked.connect(self.exit)
#
#         # Create checkbox to remember password
#         self.remember_checkbox = QCheckBox('Remember me', self)
#         self.remember_checkbox.move(50, 200)
#         self.remember_checkbox.stateChanged.connect(self.remember_password)
#
#         # Load saved login info if available
#         if os.path.exists('login_info.txt'):
#             with open('login_info.txt', 'r') as f:
#                 lines = f.readlines()
#                 self.username_input.setText(lines[0].strip())
#                 self.password_input.setText(lines[1].strip())
#
#     # Function to remember password
#     def remember_password(self, state):
#         if state == Qt.Checked:
#             # Save username and password to a file
#             with open('login_info.txt', 'w') as f:
#                 f.write(self.username_input.text() + '\n')
#                 f.write(self.password_input.text())
#         else:
#             # Remove saved login info
#             if os.path.exists('login_info.txt'):
#                 os.remove('login_info.txt')
#
#     # Function to check login credentials and open display window
#     def login(self):
#         # Check if username and password are correct
#         if self.username_input.text() == 'admin' and self.password_input.text() == 'password':
#             # Open display window
#             self.display_window = DisplayWindow()
#             self.display_window.show()
#             self.close()
#         else:
#             # Display error message
#             QMessageBox.warning(self, 'Error', 'Incorrect username or password. Please try again.')

# from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, QMessageBox
#
#
# class Login(QWidget):
#     def __init__(self):
#         super().__init__()
#         self.setWindowTitle('Login')
#         self.resize(300, 150)
#         self.username_label = QLabel('Username:')
#         self.username_input = QLineEdit()
#         self.password_label = QLabel('Password:')
#         self.password_input = QLineEdit()
#         self.password_input.setEchoMode(QLineEdit.Password)
#         self.login_button = QPushButton('Login')
#         self.login_button.clicked.connect(self.check_login)
#         layout = QVBoxLayout()
#         layout.addWidget(self.username_label)
#         layout.addWidget(self.username_input)
#         layout.addWidget(self.password_label)
#         layout.addWidget(self.password_input)
#         layout.addWidget(self.login_button)
#         self.setLayout(layout)
#
#     def check_login(self):
#         if self.username_input.text() == 'admin' and self.password_input.text() == '123456':
#             self.hide()
#             self.main_window = MainWindow()
#             self.main_window.show()
#         else:
#             QMessageBox.warning(self, 'Warning', 'Incorrect username or password!')
#
#
# class MainWindow(QWidget):
#     def __init__(self):
#         super().__init__()
#         self.setWindowTitle('Main Window')
#         self.resize(300, 150)
#         self.welcome_label = QLabel('Welcome!')
#         self.exit_button = QPushButton('Exit')
#         self.exit_button.clicked.connect(self.exit)
#         layout = QVBoxLayout()
#         layout.addWidget(self.welcome_label)
#         layout.addWidget(self.exit_button)
#         self.setLayout(layout)
#
#     def exit(self):
#         self.hide()
#         self.login_window = Login()
#         self.login_window.show()
#
#
# if __name__ == '__main__':
#     app = QApplication([])
#     login_window = Login()
#     login_window.show()
#     app.exec_()

# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     login = Login()
#     login.show()
#     sys.exit(app.exec_())
from selenium import webdriver
from selenium.webdriver.common.keys import Keys


class Login(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Login')
        self.resize(300, 150)

        self.username_label = QLabel('Username:')
        self.username_input = QLineEdit()
        self.password_label = QLabel('Password:')
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.login_button = QPushButton('Login')
        self.login_button.clicked.connect(self.login)

        layout = QVBoxLayout()
        layout.addWidget(self.username_label)
        layout.addWidget(self.username_input)
        layout.addWidget(self.password_label)
        layout.addWidget(self.password_input)
        layout.addWidget(self.login_button)

        self.setLayout(layout)

    def login(self):
        username = self.username_input.text()
        password = self.password_input.text()
        if username == 'admin' and password == 'password':
            self.hide()
            self.main = Main()
            self.main.show()
            driver = webdriver.Chrome()
            driver.get("https://www.example.com")
            elem = driver.find_element_by_name("username")
            elem.send_keys(username)
            elem = driver.find_element_by_name("password")
            elem.send_keys(password)
            elem.send_keys(Keys.RETURN)
            assert "Login successful" in driver.page_source
        else:
            self.label.setText('Login failed')


class Main(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Main')
        self.resize(300, 150)

        self.label = QLabel('Welcome Login')
        layout = QVBoxLayout()
        layout.addWidget(self.label)

        self.setLayout(layout)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    login = Login()
    login.show()
    sys.exit(app.exec_())
