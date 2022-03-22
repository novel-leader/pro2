from selenium import webdriver
options = webdriver.ChromeOptions()
options.headless = True
options.add_argument("window-size=1920x1080")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36")
browser = webdriver.Chrome("C:/github/project2/chromedriver.exe", options=options)

### 이번건은 이모티콘샵에서 검색해서 나오는 모든 것 다운로드입니다.

# from selenium.webdriver.common.keys import Keys         # enter 입력키 가져오기
# from selenium.webdriver.common.by import By # browser.find_element(By.ID, 'query') 이형식으로 바꾸기

from bs4 import BeautifulSoup
import time
import urllib.request

browser.get(f"https://e.kakao.com/search?q=%EC%98%B4%ED%8C%A1%EC%9D%B4")
time.sleep(3)

# 첫번째 결과 클릭
counts = browser.find_elements_by_class_name('double_img')
print(len(counts))

for i in range(len(counts)):
    elem = browser.find_element_by_xpath(f'//*[@id="kakaoContent"]/div[2]/ol/li[{i+1}]')
    elem.click()

    time.sleep(5) # 자료 로딩 시간 필요
    soup = BeautifulSoup(browser.page_source, "lxml")
    name = soup.find("span", attrs={"class":"tit_inner"}).get_text()
    print(name)

    images = soup.find_all("img", attrs={"class":"img_emoticon"})
    for idx, image in enumerate(images):   # enumerate 사용하여 idx 만듦
        image_url = image["src"]
        urllib.request.urlretrieve(image_url, "C:/python/down/" + name +"_" + str(idx+1) + ".png")
    print(len(images), "개 다운로드 완료")

    browser.back()
    time.sleep(3) # 뒤로가기 지연


browser.quit()