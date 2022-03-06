from selenium import webdriver
options = webdriver.ChromeOptions()
options.headless = True
options.add_argument("window-size=1920x1080")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36")
browser = webdriver.Chrome(options=options)

# from selenium.webdriver.common.keys import Keys         # enter 입력키 가져오기
# from selenium.webdriver.common.by import By # browser.find_element(By.ID, 'query') 이형식으로 바꾸기

from bs4 import BeautifulSoup
import time
import urllib.request

name = "leave-it-to-me-house-keeper-ebichu"
browser.get(f"https://e.kakao.com/t/{name}")
time.sleep(5)

soup = BeautifulSoup(browser.page_source, "lxml")
images = soup.find_all("img", attrs={"class":"img_emoticon"})

for idx, image in enumerate(images):      # enumerate 사용하여 idx 만듦
    image_url = image["src"]
    urllib.request.urlretrieve(image_url, name +"_" + str(idx) + ".png")

print(len(images), "개 다운로드 완료")