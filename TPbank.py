import io, os, time
import json
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

    def timeoutToken(self,):
        while True:
            element = '//*[@id="btn-step1"]/button[2]'
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

            #first login...
            first_login_element = '//*[@id="btn-step1"]/button[2]'
            if self.loadCompleted(first_login_element,20):
                self.clickElement(first_login_element)
            print("login success")
        except TimeoutException:
            print("Login TPbank timeout")
            time.sleep(5)
            return
        except:
            time.sleep(10)
            print("has been login TPbank - can't find element")

    def runDownload(self):
        """ start download TPbank Transaction """
        self.loginTPbank()

        
        # self.timeoutToken()

        account_detail_btn = '//*[@id="instruction-4"]/app-acc-slider/div/div[1]/div[1]/div/div[4]/div[2]/div/div'#click query button
        self.clickElement(account_detail_btn)

        select = '(//span[@class="select2-selection__rendered"])[2]'
        self.clickElement(select)

        # find_transaction_data = '//*[@id="content-transaction"]/li[1]/app-acc-trans-item/app-date/div'
        # self.clickElement(find_transaction_data)

        all_transaction = self.driver.find_element(By.CLASS_NAME, 'transaction-content')
    
        each = all_transaction.find_elements(By.TAG_NAME, 'li')
        table_data = []
        with open('extract_data.txt', mode='w', encoding='utf-8') as log_file:
            log_file.write(' ' + '\n')

        for row in each:
            ActionChains(self.driver).click(row).perform()
            time.sleep(4)
            modal = self.driver.find_element(By.CLASS_NAME, 'modal-body')
            modal_content = modal.find_element(By.CLASS_NAME, 'content')
            modal_transaction_info = modal_content.find_element(By.CLASS_NAME, 'transaction-info')
            info_tran = modal_transaction_info.find_elements(By.CLASS_NAME,'line-right')
            data = [tran_text.text for tran_text in info_tran]
            table_data.append(data)
            
            with open('extract_data.txt', mode='a', encoding='utf-8') as log_file:
                log_file.write(' '.join(data) + '\n')
            close_btn = '/html/body/app-root/main-component/div/div[2]/div/div/div[1]/div/app-account-transaction/div/div/div[2]/app-acc-trans-search/div[1]/app-modal[1]/div/div/div/div/div[1]/img'
            self.clickElement(close_btn)
            transaction_info = '/html/body/app-root/main-component/div/div[2]/div/div/div[1]/div/app-account-transaction/div/div/div[2]/app-acc-trans-search/div[1]/app-modal[1]/div/div/div/div/div[4]'
            print(data)
            print('----------------------------------')
        time.sleep(10)
        self.driver.quit()

    def isLoginError(self):
        xpath_element = '//*[@id="maincontent"]/ng-component/div[1]/div/div[3]/div/div/div/app-login-form/div/div/div[4]/p'
        login_error = WebDriverWait(self.driver, 20).until(
            EC.visibility_of_element_located((By.XPATH, xpath_element))).text
        print("login_error", login_error)

        if login_error == 'Mã kiểm tra không chính xác. Quý khách vui lòng kiểm tra lại.':
            return True
        return False

if __name__ == "__main__":
    file_name = f'.\\setting.json'
    with open(file_name) as file:
        info = json.load(file)
    AutoDownloadTPbank(info['USER_NAME'],info['PASSWORD'])