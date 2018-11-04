# 브라우저 실행
from selenium import webdriver
driver = webdriver.Chrome('C:/Python/chromedriver.exe')

# 상장회사검색
driver.get('http://marketdata.krx.co.kr/mdi#document=040601')

# 다운로드 버튼을 클릭
from selenium.webdriver.common.by import By

button = driver.find_element(By.XPATH, '//button[text()="Excel"]')
button.click()

import os
import time

# 다운로드 폴더로 이동
folder = 'C:/Users/Downloads/'
os.chdir(folder)

# 파일 다운로드까지 대기 (1초씩 최대 30회)
fname = 'data.xls'
for _ in range(30):
    if os.path.exists(fname):
        break
    time.sleep(1)

# 파일명 바꾸기
os.rename('data.xls', '상장회사목록.xls')

# 브라우저 종료
driver.close()