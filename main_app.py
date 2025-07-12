import sys
import os
from PyQt6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QTextEdit, QWidget,
                             QMessageBox, QFileDialog)
from PyQt6.QtCore import Qt, QSettings
from order_search import ExcelDataLoader, OrderProcessor


class OrderSearchApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Поиск заказов")
        self.setGeometry(100, 100, 700, 600)
        self.excel_file = None

        # Initialize settings
        self.settings = QSettings("MyCompany", "OrderSearchApp")

        # Initialize the data processor
        self.loader = ExcelDataLoader()
        self.processor = OrderProcessor(self.loader)

        self.init_ui()

        # Load saved file path if exists
        saved_path = self.settings.value("last_excel_file")
        if saved_path and os.path.exists(saved_path):
            self.set_excel_file(saved_path)

    def init_ui(self):
        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        # Title label
        title_label = QLabel("🔍 Поиск информации о заказе")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 20px;")
        layout.addWidget(title_label)

        # File selection section
        file_layout = QHBoxLayout()
        self.file_label = QLabel("Файл не выбран")
        self.file_label.setStyleSheet("color: gray;")

        select_file_btn = QPushButton("Выбрать файл Excel")
        select_file_btn.clicked.connect(self.select_excel_file)

        file_layout.addWidget(self.file_label, stretch=1)
        file_layout.addWidget(select_file_btn)
        layout.addLayout(file_layout)

        # Input section
        input_layout = QHBoxLayout()
        self.order_input = QLineEdit()
        self.order_input.setPlaceholderText("Введите номер заказа...")
        self.order_input.returnPressed.connect(self.search_order)
        input_layout.addWidget(self.order_input)

        search_btn = QPushButton("Поиск")
        search_btn.clicked.connect(self.search_order)
        search_btn.setEnabled(False)
        self.search_btn = search_btn
        input_layout.addWidget(search_btn)

        layout.addLayout(input_layout)

        # Result display
        self.result_display = QTextEdit()
        self.result_display.setReadOnly(True)
        self.result_display.setStyleSheet("font-family: monospace;")
        layout.addWidget(self.result_display)

        # Status bar
        self.statusBar().showMessage("Выберите файл Excel с данными")

    def select_excel_file(self):
        # Start from last used directory or home directory
        start_dir = self.settings.value("last_directory", os.path.expanduser("~"))

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл Excel",
            start_dir,
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )

        if file_path:
            self.set_excel_file(file_path)
            # Save directory for next time
            self.settings.setValue("last_directory", os.path.dirname(file_path))

    def set_excel_file(self, file_path):
        self.excel_file = file_path
        self.file_label.setText(file_path)
        self.file_label.setStyleSheet("color: black;")
        self.search_btn.setEnabled(True)
        self.statusBar().showMessage(f"Файл загружен: {file_path}")

        # Update processor to use the selected file
        class CustomExcelLoader(ExcelDataLoader):
            def load_data(self, _):
                return super().load_data(file_path)

        self.processor.data_loader = CustomExcelLoader()

        # Save the file path
        self.settings.setValue("last_excel_file", file_path)

    def search_order(self):
        if not self.excel_file:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите файл Excel")
            return

        order_number = self.order_input.text().strip()
        if not order_number:
            self.statusBar().showMessage("Введите номер заказа")
            return

        try:
            result = self.processor.process_order(order_number)
            self.result_display.setText(result)
            self.statusBar().showMessage(f"Найден заказ №{order_number}")
        except Exception as e:
            self.result_display.clear()
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")
            self.statusBar().showMessage("Ошибка при поиске заказа")

    def closeEvent(self, event):
        # Save window geometry
        self.settings.setValue("geometry", self.saveGeometry())
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)

    # Restore window geometry from settings
    settings = QSettings("MyCompany", "OrderSearchApp")
    window = OrderSearchApp()
    geometry = settings.value("geometry")
    if geometry:
        window.restoreGeometry(geometry)

    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()