import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QStackedWidget, QDesktopWidget,
    QLabel, QSpacerItem, QSizePolicy
)
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import Qt
from view.login_view import LoginPage
from view.cert_view import CertPage


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Certificate Management - Login")
        self.setMinimumSize(750, 990)
        self.resize(750, 990)

        self.stack = QStackedWidget()

        self.login_page = LoginPage(self.go_to_cert_page)
        self.cert_page = CertPage()

        self.stack.addWidget(self.login_page)
        self.stack.addWidget(self.cert_page)

        layout = QVBoxLayout()

        # Ajouter le logo avec chemin dynamique
        logo_path = resource_path("GEVernova_logo.jpg")
        self.logo_label = QLabel(self)
        pixmap = QPixmap(logo_path)
        if not pixmap.isNull():
            self.logo_pixmap = pixmap.scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.logo_label.setPixmap(self.logo_pixmap)
        else:
            self.logo_label.setText("Logo not found")
        self.logo_label.setAlignment(Qt.AlignCenter)

        layout.addWidget(self.logo_label)
        layout.addWidget(self.stack)

        spacer = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)
        layout.addItem(spacer)

        self.copyright_label = QLabel("© 2025 GE Vernova. All rights reserved.", self)
        self.copyright_label.setAlignment(Qt.AlignCenter)
        self.copyright_label.setStyleSheet("font-size: 10px; color: gray;")
        layout.addWidget(self.copyright_label)

        self.setLayout(layout)

        self.center()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def go_to_cert_page(self):
        self.stack.setCurrentWidget(self.cert_page)
        self.setWindowTitle("Certificate Management - Generation")

    def go_to_login_page(self):
        self.stack.setCurrentWidget(self.login_page)
        self.setWindowTitle("Certificate Management - Login")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()

    # Icone avec chemin dynamique
    icon_path = resource_path("app_icon.ico")
    window.setWindowIcon(QIcon(icon_path))

    window.show()
    sys.exit(app.exec())
