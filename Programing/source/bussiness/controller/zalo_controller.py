import time
import os
import glob
from datetime import datetime

import Programing.source.utility.fix_qt_import_error
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, WebDriverException, NoSuchWindowException
from selenium.webdriver.common.keys import Keys
from PyQt5.QtWidgets import QMessageBox
import pyperclip
from selenium.webdriver.remote.webelement import WebElement

from Programing.source.bussiness.controller.zalo_utils import Path, ZaloUtility
from Programing.source.utility.chromedriver import ChromeDriver


class Zalo:

    ZALO_LOGIN = "https://chat.zalo.me/"
    ZALO_NOT_LOGIN = "https://id.zalo.me/account/login"
    ZALO_BLOCK_STRANGERS = "BẠN CHƯA THỂ GỬI TIN NHẮN ĐẾN NGƯỜI NÀY VÌ NGƯỜI NÀY CHẶN KHÔNG NHẬN TIN NHẮN TỪ NGƯỜI LẠ."
    ZALO_BLOCK_ME = "XIN LỖI! HIỆN TẠI TÔI KHÔNG MUỐN NHẬN TIN NHẮN."
    STATUS_FIELDS = ['Số ĐT', 'Tin nhắn văn bản', 'Hình ảnh gửi kèm', 'Ngày giờ gửi tin',
                     'Trạng thái gửi tin', 'Ngày giờ trạng thái (nếu có)', 'Ghi chú']
    ZALO_CLASS_MY_LAST_MESSAGE = ["card me last-msg has-status card--text",
                                  "card me last-msg has-status card--picture",
                                  "card me last-msg has-status card--sticker"]
    ZALO_CLASS_CUSTOMER_LAST_MESSAGE = ["card  pin-react  last-msg card--text",
                                        "card  pin-react  last-msg card--picture",
                                        "card  pin-react  last-msg card--sticker"]

    JS_DROP_FILES = "var c=arguments,b=c[0],k=c[1];c=c[2];" \
                    "for(var d=b.ownerDocument||document,l=0;;){" \
                    "var e=b.getBoundingClientRect()," \
                    "g=e.left+(k||e.width/2)," \
                    "h=e.top+(c||e.height/2)," \
                    "f=d.elementFromPoint(g,h);" \
                    "if(f&&b.contains(f))break;" \
                    "if(1<++l)throw b=Error('Element not interactable')," \
                    "b.code=15,b;b.scrollIntoView({behavior:'instant'," \
                    "block:'center',inline:'center'})}var a=d.createElement('INPUT');a.setAttribute('type','file');" \
                    "a.setAttribute('multiple','');a.setAttribute('style','position:fixed;z-index:2147483647;left:0;top:0;');" \
                    "a.onchange=function(b){a.parentElement.removeChild(a);" \
                    "b.stopPropagation();" \
                    "var c={constructor:DataTransfer,effectAllowed:'all',dropEffect:'none',types:['Files']," \
                    "files:a.files,setData:function(){},getData:function(){},clearData:function(){},setDragImage:function(){}};" \
                    "window.DataTransferItemList&&(c.items=Object.setPrototypeOf(Array.prototype.map.call(a.files," \
                    "function(a){return{constructor:DataTransferItem,kind:'file',type:a.type,getAsFile:" \
                    "function(){return a},getAsString:function(b){var c=new FileReader;c.onload=function(a){b(a.target.result)};" \
                    "c.readAsText(a)}}}),{constructor:DataTransferItemList,add:function(){},clear:function(){},remove:function(){}}));" \
                    "['dragenter','dragover','drop'].forEach(function(a){var b=d.createEvent('DragEvent');" \
                    "b.initMouseEvent(a,!0,!0,d.defaultView,0,0,0,g,h,!1,!1,!1,!1,0,null);Object.setPrototypeOf(b,null);" \
                    "b.dataTransfer=c;Object.setPrototypeOf(b,DragEvent.prototype);f.dispatchEvent(b)})};" \
                    "d.documentElement.appendChild(a);a.getBoundingClientRect();return a;"

    def __init__(self):
        self.contact_found = False
        self.login_ok = False
        self.driver = None
        self.step_wait = 1
        self.phones = []
        self.found_phones = []
        self.saved_path = os.path.join(Path.BOT_ROOT_FOLDER, "saved_paths.txt")

    def __del__(self):
        try:
            self.driver.close()
            self.driver.quit()
        except Exception:
            pass

    @staticmethod
    def create_driver(chome_driver_path="chromedriver.exe"):
        """ create driver for each profile """
        profile_dir = os.path.join(Path.BOT_ROOT_FOLDER, r"Profiles", "ZALO_CSKH")
        chrome_options = Options()
        chrome_options.add_argument(f"user-data-dir={profile_dir}")
        chrome_options.add_argument("disable-infobars")
        chrome_options.add_argument("launch-simple-driver")
        chrome_options.add_argument("start-maximized")
        driver = webdriver.Chrome(executable_path=chome_driver_path, options=chrome_options)
        return driver

    def login(self):
        """ login zalo """
        try:
            if self.driver.current_url.split("?")[0] == Zalo.ZALO_NOT_LOGIN:
                pass
        except (AttributeError, WebDriverException):
            self.driver = self.create_driver(ChromeDriver(Path.CHROME_DRIVER_PATH).get_chromedriver())
            self.driver.get(Zalo.ZALO_LOGIN)
            time.sleep(1)

        try:
            if self.driver.current_url != Zalo.ZALO_LOGIN:
                ZaloUtility.show_message('THÔNG BÁO', '<center> Bạn chưa đăng nhập zalo !!! </center> \n '
                                         '<center> Vui lòng đăng nhập để tiếp tục </center>', type_=QMessageBox.Warning)
                while self.driver.current_url != Zalo.ZALO_LOGIN:
                    pass
                self.login_ok = True
            else:
                self.login_ok = True
        except NoSuchWindowException:
            return

    def find_contact(self, nick):
        """
        find contact and go to message
        contact_found = True if contact is found
        """
        self.contact_found = True
        try:
            self.driver.find_element_by_id('contact-search-input')
        except NoSuchElementException:
            self.driver.get(self.ZALO_LOGIN)
            time.sleep(self.step_wait)

        contact = self.driver.find_element_by_id('contact-search-input')
        contact.clear()
        contact.send_keys(nick)
        time.sleep(self.step_wait)
        search_result = self.driver.find_elements_by_class_name("global-search-no-result")
        if search_result:
            # contact not found
            self.contact_found = False
            contact.send_keys((Keys.CONTROL, "a", Keys.DELETE))
            return

        contact.send_keys(Keys.ENTER)
        time.sleep(self.step_wait)

    def send_data(self, message, image_folder):
        """ send message and images to users whose nicks have found
            Customer can send single image or whole folder contained images"""
        if self.contact_found:
            if image_folder is not None:
                if str(image_folder).endswith((".png", ".PNG", ".jpg", ".JPG", ".jpeg", ".JPEG", ".tif", ".TIF")):
                    upload_images = image_folder
                else:
                    try:
                        # loop all images in image_folder
                        upload_images = [image for image in glob.glob(str(image_folder) + "/*")]
                    except Exception:
                        pass
                message_input = self.driver.find_element_by_id('richInput')
                self.drop_files(element=message_input, files=upload_images)
                time.sleep(1)

            # send message as caption of image
            if message is not None:
                message_input = self.driver.find_element_by_id("input_line_0")
                pyperclip.copy(message)
                message_input.send_keys(Keys.CONTROL, 'v')
                message_input = self.driver.find_element_by_id('input_line_0')
                message_input.send_keys(Keys.ENTER)
                time.sleep(0.5)

    def take_message_status(self, phone):
        """ get status of message that has sent
        [Ngày giờ gửi tin, Trạng thái gửi tin, Ngày giờ trạng thái (nếu có), Ghi chú]"""

        if phone not in self.found_phones:
            return [None, None, None, "Không tìm thấy nick zalo"]
        else:
            self.find_contact(phone)
            status = []
            # datetime sent message (Ex: "19:00 Hôm nay")
            chat_date = self.driver.find_elements_by_class_name('chat-date')
            str_ = chat_date[-1].get_attribute('textContent') if chat_date else None
            status.append(str_[:5] + f" - {datetime.now().strftime('%d/%m/%Y')}") if str_ else status.append(None)

            # last message of guest or me can be [text, image, sticker]
            my_status, guest_status = [], []
            for x in Zalo.ZALO_CLASS_MY_LAST_MESSAGE:
                if self.driver.execute_script(r" return document.getElementsByClassName(arguments[0])", x):
                    my_status = self.driver.execute_script(r'return document.getElementsByClassName(arguments[0])[0].innerText.split("\n")', x)
                    break
            for x in Zalo.ZALO_CLASS_CUSTOMER_LAST_MESSAGE:
                if self.driver.execute_script(r"return document.getElementsByClassName(arguments[0])", x):
                    guest_status = self.driver.execute_script(r'return document.getElementsByClassName(arguments[0])[0].innerText.split("\n")', x)
                    break
            # [Trạng thái gửi tin,  Ngày giờ trạng thái (nếu có), Ghi chú]
            if my_status:
                try:
                    status.extend([my_status[3], my_status[2], None])  # text
                except IndexError:
                    status.extend([my_status[2], my_status[1], None])  # image or sticker
            else:
                if guest_status[0].upper() == Zalo.ZALO_BLOCK_STRANGERS:
                    status.extend([None, None, "Chặn tin nhắn từ người lạ"])
                elif guest_status[0].upper() == Zalo.ZALO_BLOCK_ME:
                    status.extend([None, None, "Chặn tôi"])
                else:
                    status.extend(["Đã xem", None, None])

            return status

    def update_status_nick(self, nick):
        """ update status of nick which has statuses 'ĐÃ GỬI' """
        self.find_contact(nick)
        if not self.contact_found:
            return None
        # get status of last message of sender
        send_status = self.driver.find_elements_by_class_name('card-send-status')
        if send_status:
            return send_status[-1].get_attribute('textContent')
        else:
            # if receiver has reacted sender's message, so that means status of message is "Đã xem"
            if self.driver.find_elements_by_css_selector('.card.pin-react.last-msg.card--text'):
                return "Đã xem"
            if self.driver.find_elements_by_css_selector('.card.pin-react.last-msg.card--picture'):
                return "Đã xem"
            if self.driver.find_elements_by_css_selector('.card.pin-react.last-msg.card--sticker'):
                return "Đã xem"
            return None

    @staticmethod
    def drop_files(element: WebElement, files, offsetX=0, offsetY=0):
        driver = element.parent
        isLocal = not driver._is_remote or '127.0.0.1' in driver.command_executor._url
        paths = []

        # ensure files are present, and upload to the remote server if session is remote
        for file in (files if isinstance(files, list) else [files]):
            if not os.path.isfile(file):
                raise FileNotFoundError(file)
            paths.append(file if isLocal else element._upload(file))

        value = '\n'.join(paths)
        elm_input = driver.execute_script(Zalo.JS_DROP_FILES, element, offsetX, offsetY)
        elm_input._execute('sendKeysToElement', {'value': [value], 'text': value})