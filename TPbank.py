import io, os, time

import pyautogui
# from google.cloud import vision
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from winreg import *

def re_captcha():
    """Detects text in the file."""
    client = vision.ImageAnnotatorClient()

    # Open image using the pillow package
    img = pyautogui.screenshot()
    # img.save(PATH)
    # initialiaze io to_bytes converter
    img_byte_arr = io.BytesIO()
    # define quality of saved array
    img.save(img_byte_arr, format='PNG', subsampling=0, quality=100)
    # converts image array to bytesarray
    content = img_byte_arr.getvalue()

    image = vision.Image(content=content)

    response = client.text_detection(image=image)
    texts = response.text_annotations

    for text in texts:
        txt = text.description
        if txt.isnumeric() and len(txt) == 5:
            return txt

    if response.error.message:
        raise Exception(
            "{}\nFor more info on error messages, check: "
            "https://cloud.google.com/apis/design/errors".format(response.error.message)
        )


class AutoDownloadTPbank:
    def __init__(self, user_name, pass_word):
        self.user_name = user_name
        self.pass_word = pass_word

        # self.excel_file_name = "TPbank_Account_Statement.xlsx"
        # with OpenKey(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
        #     self.dir_download = QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0] + "\\"
        # self.dir_sample_input = os.getcwd() + "\\sample input\\"

        self.driver = webdriver.Chrome()
        self.driver.maximize_window()

        self.runDownload()

    def loadCompleted(self, locator, timeout):
        """ check website load complete """
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, locator))
            )
            return True
        except TimeoutException:
            return False

    def clickElement(self, xpath_element):
        """ find element on website then click """
        try:
            if self.loadCompleted(xpath_element, 50):
                element = self.driver.find_element(By.XPATH, xpath_element)
                element.click()

        except NoSuchElementException:
            print("can not find element:", xpath_element)
        except Exception:
            print("can not click try perform ")
            time.sleep(10)
            # ex_element = WebDriverWait(self.driver, 30).until(
            #     EC.visibility_of_element_located((By.XPATH, xpath_element)))
            ex_element = self.driver.find_element(By.XPATH, xpath_element)
            ActionChains(self.driver).click(ex_element).perform()

    def click_select_date(self, id_btn):
        try:
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, id_btn)))
            down_arrow_btn = self.driver.find_element(By.ID, id_btn)
            down_arrow_btn.click()
            print("click:" + id_btn)
        except Exception:
            print("can't find %s, try run javaScript" % id_btn)

    def get_captcha(self):
        captcha = None
        captcha_image = '//div[@class="input-group-slot-inner"]/img'

        flag = False

        while not captcha:
            wait = WebDriverWait(self.driver, 20)
            wait.until(EC.visibility_of_element_located((By.XPATH, captcha_image)))
            captcha = re_captcha()

            if not captcha:
                refresh_captcha_btn = '//a[@class="ubtn ubg-secondary ubtn-sm ripple no-m"]'
                self.clickElement(refresh_captcha_btn)
            else:
                break
        return captcha

    def deleteAllFiles(self):
        """ delete all file in folder """
        for the_file in os.listdir(self.dir_download):
            file_path = os.path.join(self.dir_download, the_file)
            try:
                if os.path.isfile(file_path):
                    os.remove(file_path)
            except Exception as e:
                print(e)

    def download_wait(self, directory, timeout, nfiles=None):
        """
        Wait for downloads to finish with a specified timeout.

        Args
        ----
        directory : str
            The path to the folder where the files will be downloaded.
        timeout : int
            How many seconds to wait until timing out.
        nfiles : int, defaults to None
            If provided, also wait for the expected number of files.

        """
        seconds = 0
        dl_wait = True
        while dl_wait and seconds < timeout:
            time.sleep(1)
            dl_wait = False
            files = os.listdir(directory)
            if nfiles and len(files) != nfiles:
                dl_wait = True

            for fname in files:
                if fname.endswith('.crdownload'):
                    dl_wait = True

            seconds += 1
        return seconds

    # Login to TPbank
    def loginTPbank(self):
        try:
            self.driver.get("https://ebank.tpb.vn/retail/vX/")
            # captcha = self.get_captcha()
            user_ele = "/html/body/app-root/login-component/div/div[1]/div[2]/div[2]/input"
            self.clickElement(user_ele)
            user = self.driver.find_element(By.XPATH,user_ele)
            user.clear()
            password = self.driver.find_element(By.XPATH, "/html/body/app-root/login-component/div/div[1]/div[2]/div[3]/input")
            password.clear()
            # captcha_input = self.driver.find_element(By.CSS_SELECTOR, "input[formControlName=captcha]")
            button = self.driver.find_element(By.XPATH, "/html/body/app-root/login-component/div/div[1]/div[2]/div[5]/button")

            user.send_keys(self.user_name)
            password.send_keys(self.pass_word)
            # captcha_input.send_keys(captcha)
            button.click()

        except TimeoutException:
            print("Login TPbank timeout")
            time.sleep(5)
            return
        except:
            time.sleep(15)
            print("has been login TPbank - can't find element")

    def runDownload(self):
        """ start download TPbank Transaction """
        self.loginTPbank()

        account_detail_btn = '//a[@class="list-link-item has-link-arrow tk"]/div/div[@class="item-link-arrow ' \
                             'ubg-white-2"] '
        self.clickElement(account_detail_btn)

        select = '(//span[@class="select2-selection__rendered"])[2]'
        self.clickElement(select)

        other_option = '//ul[@class="select2-results__options"]/li[3]'
        self.clickElement(other_option)

        self.driver.execute_script("document.getElementById('TimeRange').removeAttribute('readonly')")

        print("search_btn")
        search_btn = '//a[@class="ubtn ubg-primary ubtn-md ripple"]/span'
        self.clickElement(search_btn)

        self.deleteAllFiles()

        print("excel_btn")
        excel_btn = '//a[@class="ubtn ubg-secondary ubtn-md ripple"]'
        self.clickElement(excel_btn)

        time.sleep(2)
        wait_time = self.download_wait(self.dir_download, 5)
        time.sleep(wait_time)

        self.driver.quit()

    def isLoginError(self):
        xpath_element = '//*[@id="maincontent"]/ng-component/div[1]/div/div[3]/div/div/div/app-login-form/div/div/div[4]/p'
        login_error = WebDriverWait(self.driver, 20).until(
            EC.visibility_of_element_located((By.XPATH, xpath_element))).text
        print("login_error", login_error)

        if login_error == 'Mã kiểm tra không chính xác. Quý khách vui lòng kiểm tra lại.':
            return True
        return False
