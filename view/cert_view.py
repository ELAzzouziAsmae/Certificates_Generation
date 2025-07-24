from PyQt5.QtWidgets import ( 
    QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog,
    QLineEdit, QFrame, QSpinBox, QDateEdit, QProgressBar, QGraphicsDropShadowEffect
)
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QColor

from Worker.cert_worker import CertificateWorker

class CertPage(QWidget):
    def __init__(self):
        super().__init__()

        self.template_path = None
        self.loaded_excel_path = None

        outer_layout = QVBoxLayout()
        outer_layout.setContentsMargins(40, 40, 40, 40)
        outer_layout.setSpacing(20)

        container = QFrame()
        container.setStyleSheet("""
            QFrame {
                background-color: #fdfdfd;
                border-radius: 16px;
                padding: 30px;
            }
        """)
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(15)
        shadow.setXOffset(0)
        shadow.setYOffset(4)
        shadow.setColor(QColor(0, 0, 0, 80))
        container.setGraphicsEffect(shadow)

        layout = QVBoxLayout(container)
        layout.setSpacing(15)
        layout.setContentsMargins(0, 5, 0, 0)  # ➜ Décalage de 5px vers le bas

        title_label = QLabel("🎓 Training Certificate Generator")
        title_label.setStyleSheet("font-size: 22px; font-weight: 600; color: #2c3e50;")
        layout.addWidget(title_label, alignment=Qt.AlignHCenter)

        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("Training title")
        layout.addWidget(self.title_edit)

        self.load_template_button = QPushButton("📄 Load PowerPoint Template")
        self.load_template_button.clicked.connect(self.load_template)
        layout.addWidget(self.load_template_button)

        self.load_excel_button = QPushButton("📊 Load Excel File")
        self.load_excel_button.clicked.connect(self.load_excel)
        layout.addWidget(self.load_excel_button)

        self.score_spinbox = QSpinBox()
        self.score_spinbox.setMinimum(0)
        self.score_spinbox.setMaximum(100)
        self.score_spinbox.setValue(80)
        self.score_spinbox.setPrefix("Min score: ")
        layout.addWidget(self.score_spinbox)

        calendar_style = """
            QCalendarWidget QWidget {
                background-color: #ffffff;
                color: #2c3e50;
                font-family: 'Segoe UI', sans-serif;
                font-size: 14px;
                border: 1px solid #bdc3c7;
                border-radius: 8px;
            }

            QCalendarWidget QAbstractItemView:enabled {
                background-color: #f9f9f9;
                color: #2c3e50;
                selection-background-color: #3498db;
                selection-color: white;
            }

            QCalendarWidget QToolButton {
                height: 28px;
                color: white;
                font-weight: bold;
                background-color: #2980b9;
                border: none;
                border-radius: 6px;
                margin: 4px;
            }

            QCalendarWidget QToolButton:hover {
                background-color: #3498db;
            }

            QCalendarWidget QToolButton::menu-indicator {
                image: none;
            }

            QCalendarWidget QMenu {
                background-color: #ffffff;
                color: #2c3e50;
                border: 1px solid #bdc3c7;
            }

            QCalendarWidget QSpinBox {
                background-color: #ffffff;
                color: #2c3e50;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                padding: 2px;
            }

            QCalendarWidget QAbstractItemView {
                outline: none;
            }
        """

        layout.addWidget(QLabel("📅 Training date between:"))
        layout.itemAt(layout.count() - 1).widget().setStyleSheet("font-size: 16px; font-weight: 600;")

        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDisplayFormat("dd/MM/yyyy")
        self.start_date.setDate(QDate.currentDate().addMonths(-1))
        self.start_date.calendarWidget().setStyleSheet(calendar_style)
        self.start_date.calendarWidget().setFixedSize(420, 280)
        layout.addWidget(self.start_date)

        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDisplayFormat("dd/MM/yyyy")
        self.end_date.setDate(QDate.currentDate())
        self.end_date.calendarWidget().setStyleSheet(calendar_style)
        self.end_date.calendarWidget().setFixedSize(420, 280)
        layout.addWidget(self.end_date)

        self.generate_button = QPushButton("⚙️  Generate Certificates")
        self.generate_button.clicked.connect(self.generate_certifications)
        layout.addWidget(self.generate_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("Progress: %p%")
        layout.addWidget(self.progress_bar)

        self.log_label = QLabel("")
        self.log_label.setWordWrap(True)
        self.log_label.setStyleSheet("color: #27ae60; font-weight: 500;")
        layout.addWidget(self.log_label)

        outer_layout.addWidget(container)
        self.setLayout(outer_layout)

        self.setStyleSheet("""
            QWidget {
                background-color: #ecf0f1;
                font-family: 'Segoe UI', sans-serif;
            }
            QLineEdit {
                font-size: 15px;
                padding: 10px;
                border-radius: 8px;
                border: 1.5px solid #bdc3c7;
            }
            QPushButton {
                font-size: 15px;
                padding: 12px;
                border-radius: 10px;
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
                background-color: #2471a3;
            }
        """)

    def load_template(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select PowerPoint Template", "", "PowerPoint Files (*.pptx)")
        if file_path:
            self.template_path = file_path
            self.log_label.setText(f"📄  Template loaded : {file_path}")

    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.loaded_excel_path = file_path
            self.log_label.setText(f"📊  Excel file loaded : {file_path}")

    def generate_certifications(self):
        if not self.template_path:
            self.log_label.setText("⚠️ Please load a PowerPoint template.")
            return

        if not self.loaded_excel_path:
            self.log_label.setText("⚠️ Please load an Excel file.")
            return

        formation_title = self.title_edit.text().strip()
        if not formation_title:
            self.log_label.setText("⚠️ Please enter a training title..")
            return

        output_folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if not output_folder:
            return

        score_min = self.score_spinbox.value()
        start = self.start_date.date().toPyDate()
        end = self.end_date.date().toPyDate()

        self.log_label.setText("🕒 Generating certificates...")
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Generating... %p%")

        self.worker = CertificateWorker(
            self.template_path,
            self.loaded_excel_path,
            formation_title,
            output_folder,
            score_min,
            start,
            end
        )

        self.worker.progress.connect(self.show_log_message)
        self.worker.progress_percent.connect(self.progress_bar.setValue)
        self.worker.finished.connect(self.show_log_message)
        self.worker.finished.connect(lambda _: self.progress_bar.setValue(100))
        self.worker.finished.connect(lambda msg: (
            self.progress_bar.setValue(100),
            self.progress_bar.setFormat("Completed!"),
            self.show_log_message(msg)
        ))

        self.progress_bar.setValue(0)
        self.worker.start()

    def show_log_message(self, message):
        self.log_label.setText(message)
