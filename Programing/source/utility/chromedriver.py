import os
import tempfile
import requests
import winreg
import urllib.request
import shutil
import zipfile
from packaging.version import Version


class ChromeDriver:
    _CHROME_KEY = winreg.HKEY_CURRENT_USER
    _CHROME_SUBKEY = "Software\\Google\\Chrome\\BLBeacon"
    _CHROME_VERSION_KEY_NAME = "version"
    _VERSION_URL = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{}"
    _CHROMEDRIVER_URL = "https://chromedriver.storage.googleapis.com/{}/chromedriver_win32.zip"

    def __init__(self, parent_path=None):
        if parent_path is None:
            self.parent_path = tempfile.mkstemp()
        else:
            self.parent_path = parent_path

    def get_chrome_major_version(self):
        try:
            with winreg.OpenKeyEx(self._CHROME_KEY, self._CHROME_SUBKEY, 0, winreg.KEY_READ) as key:
                version, _ = winreg.QueryValueEx(key, self._CHROME_VERSION_KEY_NAME)
                major_version = Version(version).release[0]
                return major_version
        except OSError as e:
            raise ValueError("Chrome is not installed") from e

    def get_corresponding_chromedriver_version(self, chrome_major_ver):
        version_url = self._VERSION_URL.format(chrome_major_ver)
        page = requests.get(version_url)
        page.raise_for_status()
        chromedriver_ver = page.text.strip()
        return chromedriver_ver

    def download_chromedriver(self, chromedriver_ver):
        chromedriver_url = self._CHROMEDRIVER_URL.format(chromedriver_ver)
        chromedriver_dir = os.path.join(self.parent_path, chromedriver_ver)
        if not os.path.exists(chromedriver_dir):
            os.makedirs(chromedriver_dir)

        zipfile_path = os.path.join(chromedriver_dir, "chromedriver_win32.zip")
        chromedriver_path = os.path.join(chromedriver_dir, "chromedriver.exe")
        urllib.request.urlretrieve(chromedriver_url, zipfile_path)

        with zipfile.ZipFile(zipfile_path) as zipf:
            zipf.extractall(chromedriver_dir)

    def get_chromedriver(self):
        chrome_major_ver = self.get_chrome_major_version()
        chromedriver_ver = self.get_corresponding_chromedriver_version(chrome_major_ver)
        chromedriver_dir = os.path.join(self.parent_path, chromedriver_ver)
        chromedriver_path = os.path.join(chromedriver_dir, "chromedriver.exe")

        if not os.path.isdir(chromedriver_dir):
            if os.path.isfile(chromedriver_dir):
                bkp_path = os.path.join(self.parent_path, "{}.bak".format(chromedriver_ver))
                shutil.move(chromedriver_dir, bkp_path)

            os.makedirs(chromedriver_dir)
            self.download_chromedriver(chromedriver_ver)
        elif not os.path.isfile(chromedriver_path):
            self.download_chromedriver(chromedriver_ver)
        return chromedriver_path
