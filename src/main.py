import sys

from PyQt5.QtWidgets import QApplication

from src.controller.main_controller import MainController


def main():
    app = QApplication(sys.argv)
    ctr = MainController()
    ctr.show()
    ret_code = app.exec_()

    sys.exit(ret_code)


if __name__ == '__main__':
    main()
