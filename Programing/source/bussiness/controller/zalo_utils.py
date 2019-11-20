import os
import sys

import Programing.source.utility.fix_qt_import_error
from PyQt5 import QtGui, QtCore
from PyQt5.QtWidgets import QMessageBox, QApplication


class ZaloUtility:

    @staticmethod
    def dict_get_multikeys(dict_, *keys, default=None):
        for key in keys:
            if key in dict_:
                return dict_[key]
        return default

    @staticmethod
    def show_message(title, info, type_=QMessageBox.Information):
        """ show message """
        QApplication.instance()
        message_box = QMessageBox()
        message_box.setText(info)
        message_box.setWindowTitle(title)
        message_box.setWindowIcon(QtGui.QIcon(Path.ICON_PATH))
        message_box.setIcon(type_)
        message_box.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        message_box.activateWindow()
        message_box.exec_()

    @staticmethod
    def get_resource_path(relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = r"D:\Documents\UNIVERSITY\PROJECT\Python_RPA\bot-zalo-cskh\Programing\assets"
        return os.path.join(base_path, relative_path)

    @staticmethod
    def get_saved_paths(file: str):
        """ get file path that user used last time """
        if not os.path.isfile(file) or os.stat(file).st_size == 0:
            return ''
        with open(file, encoding='utf-8', mode='r') as f:
            path = f.read()
        return path

    @staticmethod
    def save_paths(source_file: str, new_path: str):
        """ save the path that has been used """
        with open(source_file, encoding='utf-8', mode='w') as f:
            f.write(new_path)


class Path:
    BOT_ROOT_FOLDER = os.path.join(
        ZaloUtility.dict_get_multikeys(os.environ, "LOCALAPPDATA", "USERPROFILE", "TEMP", "TMP", os.getcwd()), "AC_SOLUTION",
    )
    CHROME_DRIVER_PATH = os.path.join(BOT_ROOT_FOLDER, "ChromeDriver")
    ICON_PATH = ZaloUtility.get_resource_path("ac_solution.ico")
    IMG_PATH = ZaloUtility.get_resource_path("ac_solution.jpg")
    STATUS_PATH = ZaloUtility.get_resource_path("status_summary.xlsx")
