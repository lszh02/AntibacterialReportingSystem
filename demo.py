# Create login window
class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()

        # Set window title and size
        self.setWindowTitle('Login')
        self.setGeometry(100, 100, 300, 200)

        # Create username and password labels and inputs
        self.username_label = QLabel('Username:', self)
        self.username_label.move(50, 50)
        self.username_input = QLineEdit(self)
        self.username_input.move(120, 50)
        self.password_label = QLabel('Password:', self)
        self.password_label.move(50, 100)
        self.password_input = QLineEdit(self)
        self.password_input.move(120, 100)
        self.password_input.setEchoMode(QLineEdit.Password)

        # Create login and exit buttons
        self.login_button = QPushButton('Login', self)
        self.login_button.move(50, 150)
        self.login_button.clicked.connect(self.login)
        self.exit_button = QPushButton('Exit', self)
        self.exit_button.move(150, 150)
        self.exit_button.clicked.connect(self.exit)

        # Create checkbox to remember password
        self.remember_checkbox = QCheckBox('Remember me', self)
        self.remember_checkbox.move(50, 200)
        self.remember_checkbox.stateChanged.connect(self.remember_password)

        # Load saved login info if available
        if os.path.exists('login_info.txt'):
            with open('login_info.txt', 'r') as f:
                lines = f.readlines()
                self.username_input.setText(lines[0].strip())
                self.password_input.setText(lines[1].strip())

    # Function to remember password
    def remember_password(self, state):
        if state == Qt.Checked:
            # Save username and password to a file
            with open('login_info.txt', 'w') as f:
                f.write(self.username_input.text() + '\n')
                f.write(self.password_input.text())
        else:
            # Remove saved login info
            if os.path.exists('login_info.txt'):
                os.remove('login_info.txt')

    # Function to check login credentials and open display window
    def login(self):
        # Check if username and password are correct
        if self.username_input.text() == 'admin' and self.password_input.text() == 'password':
            # Open display window
            self.display_window = DisplayWindow()
            self.display_window.show()
            self.close()
        else            # Display error message
            QMessageBox.warning(self, 'Error', 'Incorrect username or password. Please try again.')