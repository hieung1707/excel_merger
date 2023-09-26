import sys

from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QComboBox, QToolButton, QLineEdit, \
    QListWidget, QMessageBox

from src.model.message import Message
from src.view.generated.main_window import Ui_MainWindow


class MainView(QMainWindow):
    class WidgetName:
        EXCEL_LIST = "excel_list"
        COMMON_SHEETS_CB = "common_sheets_cb"
        OUTPUT_DIR_EDIT = "output_dir_edit"

        # buttons
        OUTPUT_DIR_BTN = "output_dir_btn"
        ADD_BTN = "add_btn"
        REMOVE_BTN = "rm_btn"
        MERGE_BTN = "merge_btn"
        CLEAR_BTN = "clear_btn"

    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # display widgets
        self.excel_list: QListWidget = self.findChild(QListWidget, self.WidgetName.EXCEL_LIST)
        self.combo_sheets_cb: QComboBox = self.findChild(QComboBox, self.WidgetName.COMMON_SHEETS_CB)
        self.output_dir_edit: QLineEdit = self.findChild(QLineEdit, self.WidgetName.OUTPUT_DIR_EDIT)

        # buttons
        self.add_btn: QPushButton = self.findChild(QPushButton, self.WidgetName.ADD_BTN)
        self.rm_btn: QPushButton = self.findChild(QPushButton, self.WidgetName.REMOVE_BTN)
        self.merge_btn: QPushButton = self.findChild(QPushButton, self.WidgetName.MERGE_BTN)
        self.clear_btn: QPushButton = self.findChild(QPushButton, self.WidgetName.CLEAR_BTN)
        self.output_dir_btn: QToolButton = self.findChild(QToolButton, self.WidgetName.OUTPUT_DIR_BTN)

    def add_row(self, file_name: str):
        model = self.excel_list.model()

        # Insert the item into the model
        model.insertRow(model.rowCount())
        model.setData(model.index(model.rowCount() - 1, 0), file_name)

    def get_selected_rows(self) -> list:
        """
        Get selected rows
        :return:
        """
        selected_indexes = self.excel_list.selectedIndexes()

        return list(sorted([index.row() for index in selected_indexes]))

    def remove_rows(self, rows: list):
        for row in reversed(rows):
            self.excel_list.takeItem(row)

    def set_common_sheets(self, sheet_list: list):
        self.combo_sheets_cb.clear()

        for idx, sheet_name in enumerate(sheet_list):
            self.combo_sheets_cb.addItem(sheet_name)

    def get_selected_item_in_combobox(self):
        current_index = self.combo_sheets_cb.currentIndex()

        return self.combo_sheets_cb.itemText(current_index)

    def create_notification_dialog(self, msg: Message):
        dialog = QMessageBox(self)
        dialog.setIcon(msg.icon)
        dialog.setWindowTitle(msg.title)
        dialog.setText(msg.message)
        dialog.setStandardButtons(QMessageBox.Ok)

        return dialog
