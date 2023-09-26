from dataclasses import dataclass
from enum import Enum

from PyQt5.QtWidgets import QMessageBox


@dataclass
class Message:
    title: str
    message: str
    icon: str

    @classmethod
    def error(cls, content: str):
        return cls(title="Unexpected error", message=content, icon=QMessageBox.Critical)

class Messages(Enum):
    NO_COMMON_SHEET = Message(
        title="No common sheet",
        message="Selected files has no common sheet names",
        icon=QMessageBox.Critical
    )
    MISMATCH_SHEETS = Message(
        title="No common sheet",
        message="No common sheets with existing selection",
        icon=QMessageBox.Critical
    )
    NO_OUTPUT_DIR = Message(
        title="Output directory not selected",
        message="Please select an output directory.",
        icon=QMessageBox.Critical
    )
    NO_FILES_SELECTED = Message(
        title="No file selected",
        message="Please select at least 01 excel file(s).",
        icon=QMessageBox.Critical
    )
    EXPORT_SUCCESS = Message(
        title="Merge success",
        message="Merge operation success.",
        icon=QMessageBox.Information
    )
