from enum import Enum
import os
import time
import datetime

import requests
from PyQt5.QtWidgets import QMessageBox
from requests.exceptions import ConnectionError, Timeout, HTTPError

from Programing.source.license.license_utility import read_uuid, read_json, str_2_date, write_file, get_check_path
from Programing.source.bussiness.controller.zalo_utility import Utility


SECONDS_IN_DAY = 24 * 60 * 60


class LicenseCode(Enum):
    OK = 0
    NOT_AVAILABLE = 1
    INVALID = 2
    EXPIRED = 3
    IN_PROGRESS = 99999999


class Repair:
    REMOVE_FILE = 'remove'
    UPDATE_FILE = 'update'


class LicenseChecker:
    def __init__(self, base_url='localhost', test_mode=False, res=LicenseCode.NOT_AVAILABLE, bot_id='', uuid=None):
        # url to get license info from remote machine
        self.base_url = base_url
        self.bot_id = bot_id

        # icon path
        self.icon_path = Utility.get_resource_path("ac_solution.ico")
        # get uuid if it's not fed
        if uuid is None:
            self.uuid_address = read_uuid()
        else:
            self.uuid_address = uuid
        self.test_mode = test_mode
        self.license_info = None
        self.proxy_dict = None
        if test_mode:
            self.test_result = res

    def __get_license_info(self):
        if self.test_mode:
            if self.test_result == LicenseCode.NOT_AVAILABLE:
                json_str = '{"bot_id": "bot_1", "uuid":' + read_uuid() + '", "status": "idle", "notification": "", "expire_date": "2012-04-23" }'
            elif self.test_result == LicenseCode.EXPIRED:
                json_str = '{"bot_id": "bot_1", "uuid":"' + read_uuid() + '", "status": "idle", "notification": "Contact ??? for more information", "expire_date": "2012-04-23" }'
            elif self.test_result == LicenseCode.OK:
                json_str = '{"bot_id": "bot_1", "uuid":"' + read_uuid() + '", "status": "idle", "notification": "New version available", "expire_date": "2022-04-23" }'

            self.license_info = read_json(json_str)
        else:
            try:
                resp = self.resp = requests.get(self.base_url + self.uuid_address + '&bot_id=' + self.bot_id, proxies=self.proxy_dict)
                #print(self.base_url + self.uuid_address + '&bot_id=' + self.bot_id)
                if resp.status_code != 200:
                    self.license_info = None
                else:
                    self.license_info = resp.json()
            except (ConnectionError, HTTPError, Timeout):
                Utility.show_message("LỖI", "Vui lòng kiểm tra lại kết nối mạng của bạn", icon_path=self.icon_path, type_=QMessageBox.Critical)
                os._exit(0)
            except Exception:
                self.license_info = None

    def check(self):
        # check if another instance running, conditions are either temporary file exists
        # or its last modification time is more than 24 hours
        self.__get_license_info()
        #print(self.license_info)
        if not self.license_info or self.license_info['result'].lower() == 'failed':
            return LicenseCode.NOT_AVAILABLE
        else:
            expiry_date = str_2_date(self.license_info['expire_date'])
            if expiry_date is None or expiry_date < datetime.datetime.now():
                return LicenseCode.EXPIRED
            else:
                res = LicenseCode.OK
        if self.license_info['notification'] == Repair.REMOVE_FILE:
            self.cleanup()
            self.license_info['notification'] = "repaired successfully"

        path_2_check = get_check_path(self.bot_id)

        if os.path.isfile(path_2_check):
            time_elapsed = time.time() - os.path.getmtime(path_2_check)
            # if last modification is more than one day then allow to run checker
            if time_elapsed < SECONDS_IN_DAY:
                return LicenseCode.IN_PROGRESS
        else:
            write_file(path_2_check, '_', mode=0)
        # TODO UNCOMMNENT in PRODUCTION
        # self.cleanup()
        return res

    def cleanup(self):
        #  remove file
        try:
            os.remove(get_check_path(self.bot_id))
        except Exception:
            pass

    def show_status(self, status):
        message = ''
        if status == LicenseCode.NOT_AVAILABLE:
            message = " Bạn không được cấp quyền chạy trên máy tính này. \n " \
                      " Vui lòng liên hệ công ty AC Solution để được hỗ trợ"
        elif status == LicenseCode.INVALID:
            message = "Bạn không được cấp quyền chạy trên máy tính này.\n " \
                      "Vui lòng liên hệ công ty AC Solution để được hỗ trợ"
        elif status == LicenseCode.IN_PROGRESS:
            message = "Bạn không được phép chạy 2 chương trình cùng một lúc"
        elif status == LicenseCode.EXPIRED:
            message = "Bạn đã hết thời hạn sử dụng. \n" \
                      "Vui lòng liên hệ AC Solution để gia hạn sử dụng"
        elif self.license_info and self.license_info['notification'] and self.license_info['notification'] != '':
            message = "[AC_SOLUTION THÔNG BÁO]<br>" + self.license_info['notification']

        if status != LicenseCode.OK:
            Utility.show_message("LỖI", message, icon_path=self.icon_path, type_=QMessageBox.Critical)
            return
        else:
            if self.license_info['notification'] != '':
                Utility.show_message("AC Solution", message, icon_path=self.icon_path, type_=QMessageBox.Information)
            else:
                pass

    def __str__(self):
        _str = 'bot license on uuid: ' + self.uuid_address
        if self.test_mode:
            _str = _str + ' running in test view: '
            if self.test_result == LicenseCode.NOT_AVAILABLE:
                message = "Cannot get license information"
            elif self.test_result == LicenseCode.EXPIRED:
                message = "License expired!"
            else:
                message = "OK"
            _str = _str + message
        return _str

