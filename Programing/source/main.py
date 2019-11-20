import sys

from PyQt5 import QtWidgets, QtGui

from Programing.source.bussiness.view.zalo_gui import Ui_Form, remove_qt_temporary_files
from Programing.source.bussiness.controller.zalo_utility import Path
from Programing.source.license.license_checker import LicenseChecker, LicenseCode

if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    CHECKER_ENDPOINT = 'http://quatangbaohiem.com/api/bot/authorize?mac='
    license_bot = LicenseChecker(base_url=CHECKER_ENDPOINT, bot_id='zalo_bot_cskh')
    license_status = license_bot.check()
    license_bot.show_status(license_status)
    if license_status == LicenseCode.OK:
        try:
            Form = QtWidgets.QWidget()
            Form.setWindowIcon(QtGui.QIcon(Path.ICON_PATH))
            ui = Ui_Form(Form)
            Form.show()
            return_value = app.exec_()
        finally:
            license_bot.cleanup()
            remove_qt_temporary_files()
            sys.exit(return_value)
    else:
        license_bot.cleanup()
        remove_qt_temporary_files()