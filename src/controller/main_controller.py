import datetime
import os.path
import sys
import time

import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QApplication

from src import constants
from src.model.message import Messages, Message
from src.view.main_window import MainView


class MainController:
    def __init__(self):
        self.view = MainView()

        # class variables
        self.output_dir: str = None

        self.setup_listeners()
        self.list_files = []
        self.common_sheets = set()

    def setup_listeners(self):
        self.view.add_btn.clicked.connect(self.add)
        self.view.rm_btn.clicked.connect(self.remove)
        self.view.merge_btn.clicked.connect(self.merge)
        self.view.output_dir_btn.clicked.connect(self.select_output_dir)
        self.view.clear_btn.clicked.connect(self.clear)

    def add(self):
        filenames, _ = QFileDialog.getOpenFileNames(
            None,
            "QFileDialog.getOpenFileNames()",
            "",
            "Excel Files (*.xlsx)"
        )
        filenames = list(set(filenames).difference(self.list_files))
        if not filenames:
            return

        common_sheet_names = self._get_common_sheets(filenames)
        if not common_sheet_names:
            msg = self.view.create_notification_dialog(Messages.NO_COMMON_SHEET.value)
            msg.exec_()
            return

        common_sheets = self.common_sheets.intersection(common_sheet_names)
        if not self.common_sheets:
            common_sheets = common_sheet_names
        elif self.common_sheets and not common_sheets:
            msg = self.view.create_notification_dialog(Messages.MISMATCH_SHEETS.value)
            msg.exec_()
            return

        old_list = self.common_sheets.copy()
        self.common_sheets = common_sheets
        self.list_files = self.list_files + filenames

        for filename in filenames:
            self.view.add_row(filename)

        if old_list != self.common_sheets:
            self.view.set_common_sheets(list(self.common_sheets))

    def remove(self):
        rows = self.view.get_selected_rows()
        self.view.remove_rows(rows)

        for index in reversed(rows):
            self.list_files.pop(index)

        common_sheets = self._get_common_sheets(self.list_files)
        self.view.set_common_sheets(list(common_sheets))

    def merge(self):
        if not self.output_dir:
            msg = self.view.create_notification_dialog(Messages.NO_OUTPUT_DIR.value)
            msg.exec_()
            return

        if not self.list_files:
            msg = self.view.create_notification_dialog(Messages.NO_FILES_SELECTED.value)
            msg.exec_()
            return

        try:
            current_dt = datetime.datetime.now().strftime(constants.DT_FORMAT)
            export_filename = f"Export_{current_dt}.xlsx"
            export_path = os.path.join(self.output_dir, export_filename)

            df = self._merge_excels()
            df.to_excel(export_path, index=False)
            msg = self.view.create_notification_dialog(Messages.EXPORT_SUCCESS.value)
            msg.exec_()

        except Exception as e:
            msg = self.view.create_notification_dialog(
                Message.error(str(e)))
            msg.exec_()

    def select_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(None, "Select directory")

        if not dir_path:
            return

        self.output_dir = dir_path
        self.view.output_dir_edit.setText(self.output_dir)

    def clear(self):
        self.list_files.clear()
        self.common_sheets.clear()
        self.view.excel_list.clear()
        self.view.combo_sheets_cb.clear()

    def show(self):
        self.view.show()

    @staticmethod
    def _get_common_sheets(filenames: list) -> set:
        sheet_names = []

        for filename in filenames:
            excel_file = pd.ExcelFile(filename)
            sheet_names.append(excel_file.sheet_names)

        common_sheet_names = set.intersection(*map(set, sheet_names))

        return common_sheet_names

    def _merge_excels(self):
        sheet_name = self.view.get_selected_item_in_combobox()
        dfs = []

        for file_path in self.list_files:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            dfs.append(df)

        concat_df = pd.concat(dfs)

        return concat_df
