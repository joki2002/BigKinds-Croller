import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException  # NoSuchElementException 임포트 추가

import chromedriver_autoinstaller
import os
import json
import logging
import time
import calendar
from datetime import datetime

import openpyxl

form_class = uic.loadUiType("main.ui")[0]

class Mywindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
    
    # def setupUi(self):
        # filter add button event
        self.btnAdd.clicked.connect(self.btnAdd_Clicked)

        # line edit return event
        self.edFilter.returnPressed.connect(self.edFilter_ReturnPressed)
        
        # croller start button event
        self.btnStart.clicked.connect(self.btnStart_Clicked)

        # list view item click event
        self.lwFilter.itemClicked.connect(self.lwFilter_SelectItem)

        # progress bar set value 0
        self.pbActive.setValue(0)
        
        # set focused on line edit
        self.edFilter.setFocus()


    # filter add button event
    def btnAdd_Clicked(self):
        text = self.edFilter.text().strip()
        if text != '':
            self.lwFilter.addItem(text)
            self.edFilter.setText('')
    
    # line edit return event
    def edFilter_ReturnPressed(self):
        self.btnAdd_Clicked()

    # list view item change event
    def lwFilter_SelectItem(self, item):
        sel_index = self.lwFilter.selectedIndexes()
        for index in sel_index:            
            reply = QMessageBox.question(
                self, '검색어 항목 삭제', f'{item.text()}를 삭제하시겠습니까?',
                QMessageBox.Yes|QMessageBox.No|QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                self.lwFilter.model().removeRow(index.row())
                
    # croller start button event
    def btnStart_Clicked(self):
        items_list = []
        item_count = self.lwFilter.count()

        wb = openpyxl.Workbook()

        sheet = wb.active
        excel_data = []

        excel_data.append(['뉴스명','신문사','홈페이지링크'])
        
        for i in range(item_count):
            item = self.lwFilter.item(i)
            item_text = item.text()
            items_list.append(item_text.strip())

        chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]
        file_path = os.path.abspath(__file__)
        folder_path = os.path.dirname(file_path)
        driver_path = f'{folder_path}/{chrome_ver}/chromedriver.exe'

        settings = {
            "recentDestinations": [{
                    "id": "Save as PDF",
                    "origin": "local",
                    "account": "",
                }],
                "selectedDestinationId": "Save as PDF",
                "version": 2
            }
        prefs = {
            'printing.print_to_file': True,
            'printing.print_preview_sticky_settings.appState': json.dumps(settings),
            }
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])   # Turn Off ALL LOG
        
        chrome_options.add_argument('--enable-print-browser')
        chrome_options.add_argument('--kiosk-printing')
        chrome_options.add_argument('--window-size=1920x1080') 
        # chrome_options.add_argument('--window-size=3000x3000')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--kiosk-printing')
        chrome_options.add_argument('--headless=new')
        chrome_options.add_argument('--silent')
        chrome_options.add_argument('--log-level=3')
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36")

        chrome_options.add_experimental_option('prefs', prefs)

        # Selenium 로그 관리자(logger) 생성 및 설정
        selenium_logger = logging.getLogger('selenium.webdriver.remote.remote_connection')
        selenium_logger.setLevel(logging.WARNING)  # Selenium 로그 레벨을 WARNING으로 설정
        
        driver = webdriver.Chrome(options=chrome_options)

        # 빅카인즈 사이트 들어가기
        driver.get('https://www.bigkinds.or.kr/')
        driver.maximize_window()
        wait = WebDriverWait(driver, 30)  # 최대 10초 동안 대기

        # 상세검색 메뉴 들어가기
        detail_search_btn = driver.find_element(By.XPATH, '//*[@id="news-search-form"]/div/div[1]/button')
        detail_search_btn.click()

        # 상세검색 메뉴에서 상세검색 탭 들어가기
        detail_search2_btn = driver.find_element(By.XPATH, '//*[@id="news-search-form"]/div/div[1]/div[2]/div/div[3]/div[1]/a')
        detail_search2_btn.click()

        # 상세 검색 입력창에 리스트 내용 입력
        find_str = ','.join(items_list)
        input_edit = driver.find_element(By.XPATH, '//*[@id="orKeyword1"]')
        input_edit.send_keys(find_str)

        time.sleep(1)
        # 검색 버튼 클릭
        search_btn = driver.find_element(By.XPATH,'//*[@id="detailSrch1"]/div[7]/div/button[2]')
        search_btn.click()

        time.sleep(3)
        # 총 페이지 수 가져오기
        total_page = driver.find_element(By.XPATH,'//*[@id="news-results-tab"]/div[1]/div[2]/div/div/div/div/div[3]/div')
        try:
            total_page_int = int(total_page.get_attribute('data-page'))
            print(f'total_page_int = {total_page_int}')
        except ValueError:
            print('bring page count error')
            sys.exit(0)

        # 총 페이지 수 
        for i in range(int(total_page_int)):
            time.sleep(3)
            self.pbActive.setValue((i+1) / total_page_int * 100)

            for j in range(1, 11):
                data = []
                try:
                    # title
                    title = driver.find_element(By.XPATH, f'//*[@id="news-results"]/div[{j}]/div/div[2]/a/div/strong/span')
                    # 신문사명 and 링크
                    newspaper = driver.find_element(By.XPATH, f'//*[@id="news-results"]/div[{j}]/div/div[2]/div/div/a')
                except NoSuchElementException:
                    break

                # 뉴스 제목
                title = title.text
                # 신문사명
                newspaper_name = newspaper.text
                # 링크
                newspaper_link = newspaper.get_attribute('href')
                if(type(newspaper_link) != type(str())):
                    newspaper_link = ''

                data.append(title)
                data.append(newspaper_name)
                data.append(newspaper_link)

                print(data)

                excel_data.append(data)
            # next page btn click
            next_page = driver.find_element(By.XPATH, '//*[@id="news-results-tab"]/div[6]/div[2]/div/div/div/div/div[4]/a')
            next_page.click()
    
        # 엑셀 데이터 추가
        for row in excel_data:
            sheet.append(row)
        
        # 현재 시간 가져오기
        current_time = datetime.now()
        formatted_time = current_time.strftime("%Y-%m-%d")
        
        # 현재 스크립트 파일의 폴더 경로 가져오기
        current_script_folder = os.path.dirname(os.path.abspath(__file__))
        print(f'current folder = {current_script_folder}')

        # 파일 저장 폴더 생성
        folder_path = current_script_folder + '/Result File/'
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        
        folder_path = folder_path + str(datetime.today().year) + '/'
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            
        folder_path = folder_path + str(convert_to_month_name(datetime.today().month)) + '/'
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        save_dir = f"{folder_path}{formatted_time}_{items_list[0]}.xlsx"
        # 파일 이름 겹치는지 확인
        if os.path.exists(save_dir):
            counter = 1
            while True:
                f, f_ext = os.path.splitext(save_dir)
                adjusted_filename = f'{f}_{counter}{f_ext}'
                
                if not os.path.exists(adjusted_filename):
                    save_dir = adjusted_filename
                    break
                
                counter += 1
                
        # 엑셀 파일 저장
        wb.save(save_dir)

        # 메시지 다이얼로그 생성
        message_box = QMessageBox()
        message_box.setWindowTitle("알림")
        message_box.setText("크롤링이 완료되었습니다.")
        message_box.setIcon(QMessageBox.Information)

        # 확인 버튼 추가
        message_box.setStandardButtons(QMessageBox.Ok)

        # 다이얼로그 실행
        result = message_box.exec_()
            

def convert_to_month_name(month_number):
    # calendar.month_name 리스트를 사용하여 숫자를 월 이름으로 변환
    if 1 <= month_number <= 12:
        return calendar.month_name[month_number]
    else:
        return "Invalid month number"
    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = Mywindow()
    myWindow.show()
    app.exec_()