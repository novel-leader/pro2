# %%
from selenium import webdriver
options = webdriver.ChromeOptions()
# options.headless = True
options.add_argument("window-size=1920x1080")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36")
browser = webdriver.Chrome('C:/github/project2/chromedriver.exe', options=options)

from selenium.webdriver.common.keys import Keys         # enter 입력키 가져오기
from selenium.webdriver.common.by import By # browser.find_element(By.ID, 'query') 이형식으로 바꾸기
import time

import requests
from bs4 import BeautifulSoup


# 엑셀 리스트 불러오기
from openpyxl import load_workbook # 엑셀파일 불러오기
wb = load_workbook('foodlist.xlsx') # sample.xlsx 파일에서 wb 을 불러옴
ws = wb.active # 활성화된 Sheet

si = []
for x in range(2, ws.max_row + 1):
    si.append(str(ws.cell(row=x, column=2).value))  # column=1는 A열
print(si, len(si))

gu = []
for x in range(2, ws.max_row + 1):
    gu.append(str(ws.cell(row=x, column=3).value))  # column=2는 B열
print(gu, len(gu))

for i in range(0, ws.max_row-1):
    print(si[i], gu[i])

i = 0

# 페이지 이동
browser.get("https://map.naver.com/v5/search/")
time.sleep(3)

# 강서구 맛집 검색
elem = browser.find_element(By.CLASS_NAME, 'input_box').find_element(By.TAG_NAME, 'input')
## send_keys입력은 input 태그에서만 가능
elem.clear()
elem.send_keys(si[i], ' ', gu[i],' 맛집')
elem.send_keys(Keys.ENTER)

time.sleep(1)
# iframe 전환
browser.switch_to.frame('searchIframe') # frame ID 입력하여 frame 전환

# 하나 엘리먼트 지정 -> 문서 높이를 가져와서 스크롤 무한 반복
elem = browser.find_element(By.XPATH, '//*[@id="_pcmap_list_scroll_container"]')
# 하나씩 스크롤 다운
# 스크롤 무한반복
prev_height = browser.execute_script('return arguments[0].scrollHeight', elem)
while True:
    browser.execute_script('arguments[0].scrollBy(0, arguments[0].scrollHeight)', elem)
    time.sleep(1)
    cur_height = browser.execute_script('return arguments[0].scrollHeight', elem)
    if cur_height == prev_height:
        break
    prev_height = cur_height
print("스크롤 완료")


# 맛집 li 가져오기---------------------------------
# soup = BeautifulSoup(browser.page_source, "lxml")
# soup가 페이지를 잘 가져왔는지 확인 : 페이지가 안가져와져ㅠㅠ
# 동적 페이지에서는 로딩이 다 안되서 body가 저장안됨
# with open("food.html", "w", encoding="utf8") as f:
#     #f.write(soup.text) # html 문서 대충 보기
#     f.write(soup.prettify()) # html 문서를 예쁘게 출력


soup = BeautifulSoup(browser.page_source, "lxml")
# foods = soup.find_all("div", attrs={"class":["ImZGtf mpg5gc", "Vpfmgd"]})
foodname = soup.find_all("li", attrs={"class":"_1EKsQ _12tNp"})
print(len(foodname))

for food in foodname:
    # 음식점 이름 가져오기
    title = food.find("span", attrs={"class":"OXiLu"}).get_text()
    print(title)

    # 평점
    rate = food.find("em")
    if rate:
        rate = rate.get_text()
        print(rate)
    else:
        rate = 0
        print(rate)

    # 리뷰
    review = food.find_all("span", string="리뷰")
    if review:
        review = review.get_text()
        print(review)
    else:
        review = 0
        print(review)

# %%

#     # 링크
#     link = food.find("a", attrs={"class":"JC71ub"})["href"]
#     # 올바른 링크 : https://play.google.com + link

#     print(f"제목 : {title}")
#     print(f"할인 전 금액 : {original_price}")
#     print(f"할인 후 금액 : {price}")
#     print("링크 : ", "https://play.google.com" + link)
#     print("-" * 100)

# browser.switch_to.default_content() # 상위로 빠져 나옴

# browser.quit()