# 用PyQt5生成一个附带“记住密码”功能的登陆界面，当输入账号和密码并点击“登陆”按钮后,在后台保存账号和密码，以便下次免输入
import os

from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, \
    QMessageBox, QCheckBox
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import sys


class LoginThread(QThread):
    login_signal = pyqtSignal(bool)

    def __init__(self, username, password, parent=None):
        super().__init__(parent)
        self.username = username
        self.password = password

    def run(self):
        # Use selenium to login to website
        # driver = webdriver.Chrome()
        # driver.get("https://www.example.com/login")
        # username_input = driver.find_element_by_id("username")
        # password_input = driver.find_element_by_id("password")
        # username_input.send_keys(self.username)
        # password_input.send_keys(self.password)
        # login_button = driver.find_element_by_id("login-button")
        # login_button.click()
        # success = driver.current_url == "https://www.example.com/dashboard"
        # driver.quit()
        login_success = True  # Placeholder for actual login success
        self.login_signal.emit(login_success)


class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Login")
        self.setFixedSize(300, 150)

        # Create widgets
        self.username_label = QLabel("Username:")
        self.username_input = QLineEdit()
        self.password_label = QLabel("Password:")
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.remember_checkbox = QCheckBox("Remember me")
        self.login_button = QPushButton("Login")

        # Create layout
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.username_label)
        self.layout.addWidget(self.username_input)
        self.layout.addWidget(self.password_label)
        self.layout.addWidget(self.password_input)
        self.layout.addWidget(self.remember_checkbox)
        self.layout.addWidget(self.login_button)

        # Set layout
        self.setLayout(self.layout)

        # Connect signals
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

        # Save username and password if remember me is checked
        if self.remember_checkbox.isChecked():
            # Save username and password to file or database
            with open('login_info.txt', 'w') as f:
                f.write(self.username_input.text() + '\n')
                f.write(self.password_input.text())
        else:
            # Remove saved login info
            if os.path.exists('login_info.txt'):
                os.remove('login_info.txt')

        # Start login thread
        self.login_thread = LoginThread(username, password)
        self.login_thread.login_signal.connect(self.login_result)
        self.login_thread.start()

    def login_result(self, success):
        if success:
            # Show success message box
            QMessageBox.information(self, "Login Success", "You have successfully logged in.")
            # Close login window and open main window
            self.close()
            self.main_window = MainWindow()
            self.main_window.show()
        else:
            # Show error message box
            QMessageBox.warning(self, "Login Error", "Incorrect username or password.")


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Main")
        self.setFixedSize(300, 150)

        # Create widgets
        self.label = QLabel("登陆成功")

        # Create layout
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.label)

        # Set layout
        self.setLayout(self.layout)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    login_window = LoginWindow()
    login_window.show()
    sys.exit(app.exec_())
