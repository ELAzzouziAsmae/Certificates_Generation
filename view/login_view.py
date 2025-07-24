import os
import sys
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton, QLabel, QMessageBox
)
from PyQt5.QtCore import Qt

import win32security
import win32con


def check_windows_credentials(username, password):
    try:
        if "\\" in username:
            domain, user = username.split("\\", 1)
        else:
            domain = os.environ.get("USERDOMAIN", "")
            user = username

        # print(f"[DEBUG] Trying login for: {domain}\\{user}")

        token = win32security.LogonUser(
            user,
            domain,
            password,
            win32con.LOGON32_LOGON_INTERACTIVE,
            win32con.LOGON32_PROVIDER_DEFAULT
        )
        token.Close()
        return True
    except Exception as e:
        print("[AUTH ERROR]", e)
        return False


class LoginPage(QWidget):
    def __init__(self, on_login_success):
        super().__init__()
        self.on_login_success = on_login_success

        layout = QVBoxLayout()
        layout.setContentsMargins(60, 40, 60, 40)
        layout.setSpacing(20)

        # Titre
        self.title_label = QLabel("Login")
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet("""
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            font-size: 32px;
            font-weight: 700;
            color: #2c3e50;
        """)
        layout.addWidget(self.title_label)

        # SSO - nom utilisateur Windows (non modifiable)
        self.sso_edit = QLineEdit()
        username = os.getlogin()  # ou getpass.getuser()
        self.sso_edit.setText(username)
        self.sso_edit.setReadOnly(True)
        layout.addWidget(self.sso_edit)


        # Mot de passe
        password_layout = QHBoxLayout()

        self.password_edit = QLineEdit()
        self.password_edit.setPlaceholderText("Password")
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.password_edit.returnPressed.connect(self.check_login)
        password_layout.addWidget(self.password_edit)

        self.eye_button = QPushButton("👁")
        self.eye_button.setCheckable(True)
        self.eye_button.setFixedWidth(40)
        self.eye_button.setToolTip("Show/hide password")
        self.eye_button.clicked.connect(self.toggle_password_visibility)
        password_layout.addWidget(self.eye_button)

        layout.addLayout(password_layout)


        # Bouton connexion
        self.login_button = QPushButton("Sign In")
        self.login_button.clicked.connect(self.check_login)
        layout.addWidget(self.login_button)

        # Label erreur
        self.error_label = QLabel()
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setStyleSheet("color: #e74c3c; font-weight: bold;")
        layout.addWidget(self.error_label)

        self.setLayout(layout)

        # Style global
        self.setStyleSheet("""
            QLineEdit {
                font-size: 16px;
                padding: 12px 15px;
                border-radius: 25px;
                border: 2px solid #bdc3c7;
                background: #ecf0f1;
                color: #34495e;
                selection-background-color: #3498db;
            }
            QLineEdit:focus {
                border: 2px solid #2980b9;
                background: #ffffff;
            }
            QPushButton {
                font-size: 18px;
                font-weight: 600;
                padding: 12px 0;
                border-radius: 25px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 #2980b9, stop:1 #3498db);
                color: white;
                border: none;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 #3498db, stop:1 #2980b9);
            }
            QPushButton:pressed {
                background-color: #1c5980;
            }
        """)

    def check_login(self):
        username = self.sso_edit.text()
        password = self.password_edit.text()
        
        domain = os.environ.get("USERDOMAIN", "")
        full_username = f"{domain}\\{username}"  # utilisé en interne

        if check_windows_credentials(full_username, password):
            self.error_label.setText("")
            self.on_login_success()
        else:
            self.error_label.setText("Incorrect username or password, please try again")
            self.password_edit.clear()


    def toggle_password_visibility(self):
        if self.eye_button.isChecked():
            self.password_edit.setEchoMode(QLineEdit.Normal)
        else:
            self.password_edit.setEchoMode(QLineEdit.Password)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Application with pages")

        self.login_page = LoginPage(self.login_success)

        layout = QVBoxLayout()
        layout.addWidget(self.login_page)
        self.setLayout(layout)
        self.resize(400, 300)

    def login_success(self):
        QMessageBox.information(self, "Success", "Login successful !")
        # Ici tu peux changer de page ou lancer ta suite


