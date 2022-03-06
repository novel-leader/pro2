### 셀레니움 사용
# 셀레니움은 웹페이지 테스트 자동화 프레임워크
# 웹브라우저를 컨트롤 할 수 있는 웹드라이버 필요
# 설치된 크롬 브라우저와 같은 버전을 사용해야 함
# "chrome dreiver" 검색 -> 해당 버전 클릭 -> win32 다운
# 개발자도구 (F12) 활용

## 셀레니움 기본 설정
# pip install selenium
from selenium import webdriver
# 셀레니움 사용을 위한 import
options = webdriver.ChromeOptions()
# 옵션 설정을 위한 변수
options.headless = True
# 창 숨기기 옵션
options.add_argument("window-size=1920x1080")
# 창 싸이즈 옵션
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36")
# 코딩 접근이 막히면 사람 접근으로 속이는 옵션
browser = webdriver.Chrome()
# 웹드라이버로 브라우저 열기
browser = webdriver.Chrome(options=options)
# 옵션 적용
browser = webdriver.Chrome("./chromedriver.exe")
# 크롬 드라이버 위치가 다른 경우 수동 지정
browser.maximize_window()
# 창 최대화

## 웹페이지 이동
browser.get("https://www.naver.com")
import time
# 로딩시간 적용을 위한 import
time.sleep(3)
# 동적 웹페이지는 로딩시간 필요
print(browser.page_source)
# html 정보출력
browser.save_screenshot('naver.png')
# 스크린샷

## 브라우저 기본 컨트롤
browser.back()
# 뒤로 가기
browser.forward()
# 앞으로 가기
browser.refresh()
# 새로고침
browser.close()
# 브라우저 탭 닫기
browser.quit()
# 브라우저 창 닫기

## 요소 찾기
elem = browser.find_element_by_link_text('카페')   
# 카페 메뉴 찾기
elem.get_attribute('href')
# href 속성 가져오기
elem.get_attribute('class')
# class 속성 가져오기
elems = browser.find_elements_by_tag_name('a')
for elem in elems:
    elem.get_attribute('href')
# a 태그 모두 찾기

## 키보드 입력
elem = browser.find_element_by_id('query')
elem.clear()
elem.send_keys('')
from selenium.webdriver.common.keys import Keys
# enter키 입력을 위한 import
elem.send_keys(Keys.ENTER)

## 마우스 입력
elem.click()

## 스크롤 내리기
elem = browser.find_element(By.XPATH, '')
# 스크롤이 필요한 위치의 요소 설정
browser.execute_script('arguments[0].scrollBy(0, 1080)', elem)
# 스크롤 1080만큼 내리기

## 스크롤 무한반복
prev_height = browser.execute_script('return arguments[0].scrollHeight', elem)
while True:
    browser.execute_script('arguments[0].scrollBy(0, arguments[0].scrollHeight)', elem)
    time.sleep(1)
    # 동적 페이지 로딩시간 설정
    cur_height = browser.execute_script('return arguments[0].scrollHeight', elem)
    print(cur_height)
    # 현재 높이가 바뀌는지 확인
    if cur_height == prev_height:
        break
    prev_height = cur_height
print("스크롤 완료")
# 현재 스크롤 높이와 맨 밑으로 내린 스크롤의 높이를 비교하여 반복

## 버튼클릭
elem = browser.find_element_by_xpath('//*[@id="male"]')
if elem.is_selected() == False:
    elem.click()
    # 라디오 버튼이 선택되어 있지 않으면 클릭

## 옵션 태그 선택
elem = browser.find_element_by_xpath('//*[@id="cars"]/option[4]')
# cars에 해당하는 element를 찾고, 드롭다운 내부에 있는 4번째 옵션 선택
elem = browser.find_element_by_xpath('//*[@id="cars"]/option[text()="Audi"]')
# 옵션 중에서 텍스트가 Audi 인 항목을 선택
elem = browser.find_element_by_xpath('//*[@id="cars"]/option[contains(text(), "Au")]')
# 텍스트 값이 부분 일치하는 항목 선택하는 방법
elem.click()



### 뷰티풀숩으로 자료저장
# pip install beautifulsoup4
# pip install lxml
from bs4 import BeautifulSoup
# 뷰피풀숩 사용을 위한 import
soup = BeautifulSoup(browser.page_source, "lxml")
# html정보를 lxml로 뷰티풀숩 객체로 만듦
# lxml 구문분석 파써

## 필요한 요소 찾기
title = soup.title.get_text()
# 타이틀 태그에서 텍스트만 추출
a_href = soup.a["href"]
# a 태그의 href 어트리뷰트 값 추출
img = soup.find("img")
# 첫번째 img 태그 찾기
a_tag = soup.find_all("a")[0]
# 모든 a 태그 중 첫번째 a 태그 찾기
a_tag = soup.find("a", attrs={"class":"Nbtn_upload"})
# class="Nbtn_upload"인 a element 찾기
a_tag = soup.find(attrs={"class":"Nbtn_upload"})
# class="Nbtn_upload"인 어떤 element 찾기

## 텍스트 출력
texts = soup.find_all("div", attrs={"class":["ImZGtf mpg5gc", "Vpfmgd"]})
# 텍스트가 포함된 상위 카테고리의 요소들을 찾아 변수로 지정
for text in texts:
    title = text.find("span", attrs={"class":"OXiLu"}).get_text().strip()
    # 카테고리안에서 텍스트가 있는 요소를 찾아 변수로 지정
    # 불필요한 공백은 .strip()으로 제거

    ## 카테고리 안에 해당 텍스트가 없을 때 건너뛰기
    title = text.find("span", attrs={"class":"OXiLu"})
    # 카테고리안에서 텍스트가 있을만한 요소를 변수로 지정
    if title:
        title = title.get_text()
        # 타이틀이 있다면 텍스트 추출
    else:
        continue
        # 없으면 다음 for구문 진행

    ## 광고제품 제외
    ad_badge = text.find("span", attrs={"class":"ad-badge-text"})
    if ad_badge:
        continue
        # 광고관련 클래스가 있으면 다음 for구문 진행
    title = title.get_text()

    # 리뷰 100개 이상, 평점 4.5 이상 되는 것만 조회
    rate = text.find("em", attrs={"class":"rating"})
    rate_cnt = text.find("span", attrs={"class":"rating-total-count"}) 
    if float(rate) >= 4.5 and int(rate_cnt) >= 100:
        print(title, rate, rate_cnt)

    print(title)

## 테이블 자료 출력
data_rows = soup.find("table", attrs={"class":"tbl"}).find("tbody").find_all("tr")
for index, row in enumerate(data_rows):
    columns = row.find_all("td")
    print("=========== 매물 {} ===========".format(index+1))
    print("거래 :", columns[0].get_text().strip())
    print("면적 :", columns[1].get_text().strip())
    print("가격 :", columns[2].get_text().strip(), "(만원)")

## 텍스트 엑셀로 저장
import csv
filename = "시가총액1-200.csv"
f = open(filename, "w", encoding="utf-8-sig", newline="")
# newline=""은 파일 저장할때 불필요한 줄삽입 제거
writer = csv.writer(f)
title = "N	종목명	현재가	전일비	등락률	액면가	시가총액	상장주식수	외국인비율	거래량	PER	ROE".split("\t")
# 탭(\t)으로 나눠서 리스트 생성 ["N", "종목명", "현재가", ...]
writer.writerow(title)
# 엑셀에 리스트타입인 타이틀을 줄 단위로 저장

## 뷰티풀숩으로 이미지 저장
import urllib.request
# 이미지 저장을 위한 import
images = soup.find_all("img", attrs={"class":"img_emoticon"})
# 이미지 요소들을 찾아 변수로 지정
for idx, image in enumerate(images):
# enumerate 사용하여 idx 만듦
    image_url = image["src"]
    # 이미지 url을 찾아 변수로 저장
    urllib.request.urlretrieve(image_url, "이미지" + str(idx) + ".png")
print("이미지 다운로드 완료")




### 셀레니움 심화과정
## 요소가 iframe 안에 있을 경우
browser.switch_to.frame('IframeID')
# iframe ID 입력하여 frame 전환
browser.switch_to.default_content()
# 상위로 빠져 나옴

## 로딩시간으로 인한 no such element 해결방법
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
try:
    elem = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='content']/div[2]")))
    # 브라우저에서 해당 xpath가 나타날 때까지 10초 기다림  
    elem.click()
finally:
    browser.quit()
    # 나타나지 않으면 브라우저 종료

## 실행할 때마다 지저분한 멘트가 출력될 때
elem = browser.find_element_by_class_name('input_box')
# 위처럼 실행하면 지저분한 멘트가 나오며 아래처럼 하면 안나옴
from selenium.webdriver.common.by import By
elem = browser.find_element(By.CLASS_NAME, 'input_box').find_element(By.TAG_NAME, 'input')

## 링크클릭으로 다운로드시 저장경로 변경 설정
from selenium.webdriver.chrome.options import Options
chrome_options = Options()
chrome_options.add_experimental_option('prefs', {'download.default_directory':r'C:\python\VScode'})
browser = webdriver.Chrome(options=chrome_options)

## 브라우저간 전환(handle)/팝업창
curr_handle = browser.current_window_handle
# 실행된 현재 핸들 저장
handles = browser.window_handles
# 팝업으로 인한 모든 핸들 정보
for handle in handles:
    print(handle) # 각 핸들 정보 출력
    browser.switch_to.window(handle) # 각 핸들로 이동해서
    print(browser.title) # 브라우저의 제목 출력
    print()
# 새로 이동된 브라우저에서 뭔가 자동화 작업을 수행하고 끝나면
browser.close()
# 그 브라우저를 종료
browser.switch_to.window(curr_handle)
# 이전 핸들로 돌아오기



### 셀레니움 스크롤 심화과정
## 특정 엘리먼트로 스크롤 이동
from selenium.webdriver.common.action_chains import ActionChains
actions = ActionChains(browser)
actions.move_to_element(elem).perform()

## 좌표 정보 이용하여 이동
xy = elem.location_once_scrolled_into_view
# 엘리먼트가 있는 좌표 정보를 변수로 지정
print("type : ", type(xy), "value : ", xy)
# dict 타입이고 프린트만으로 이동하는 효과를 나타냄

## 일반적 스크롤 다운
while True:
    browser.execute_script("window.scrollTo(0, document.body.scrollHeight)")
    # 스크롤을 가장 아래로 내림
    time.sleep(3)
    # 페이지 로딩 대기
    curr_height = browser.execute_script("return document.body.scrollHeight")
    # 현재 문서 높이를 가져와서 저장
    if curr_height == prev_height:
        break
    prev_height = curr_height

## 네이버 부동산 스크롤 다운
browser.get('https://new.land.naver.com/houses?ms=37.5087476,126.9339552,16&a=VL:DDDGG:JWJT:SGJT:HOJT&e=RETAIL')
elem = browser.find_element_by_xpath('//*[@id="listContents1"]/div')
prev_height = browser.execute_script('return arguments[0].scrollHeight', elem)
while 1 :
    browser.execute_script('arguments[0].scrollTop = arguments[0].scrollHeight', elem)
    time.sleep(1)
    cur_height = browser.execute_script('return arguments[0].scrollHeight', elem)
    if cur_height == prev_height:
        break
    prev_height = cur_height