### 데스크탑 자동화
# pip install pyautogui
import pyautogui


## 기본
# 화면 크기 확인
size = pyautogui.size()
print(size)

# 마우스 위치 정보 Tool
# pyautogui.mouseInfo()

# 매크로 실행 정지 : 모니터 네 모퉁이에 커서 닿기, Ctrl + Alt + Del
# 매크로 실행 정지 무시
pyautogui.FAILSAFE = False

# 모든 동작에 1초씩 정지
pyautogui.PAUSE = 1

# 대기
pyautogui.sleep(1)

# 스크린샷 저장
img = pyautogui.screenshot()
img.save("screenshot.png")



## 마우스 제어
pyautogui.moveTo(500, 500, duration=1)
# 1초동안 절대값 좌표로 이동
pyautogui.move(300, 0, duration=1)
# 1초동안 상대값 좌표로 이동
p = pyautogui.position()
print(p[0], p[1])
print(p.x, p.y)
# 현재 좌표값 확인
pyautogui.click(200, 200, duration=1)
# 1초동안 절대값 좌표로 이동 후 클릭
pyautogui.mouseDown()
# 마우스 누르기
pyautogui.mouseUp()
# 마우스 떼기
pyautogui.scroll(-1000)
# ()음수이면 아래로 1000만큼 스크롤



## 키보드 제어
pyautogui.write("12345", interval=0.2) # interval 입력 속도
pyautogui.write(["t","e","s","t","left","enter"], interval=0.2)
pyautogui.hotkey("shift", "3")
pyautogui.press("enter")

# 한글입력
# pip install pyperclip
import pyperclip
def ko_write(text):
    pyperclip.copy(text)
    pyautogui.hotkey("ctrl", "v")

ko_write("함수로 한글 쓰기")



## 이미지 찾기
# 준비작업 : 이미지 저장
# 스크린 샷 : 윈도우+Shift+S
# 그림판에 붙여넣기 -> 잘라내기 -> 저장

# 이미지로 이동하기
img1 = pyautogui.locateOnScreen("img1.png")
# img1.png의 이미지파일과 매칭되는 이미지 위치 좌표를 변수에 저장
pyautogui.moveTo(img1)
# 해당 이미지 좌표로 이동

# 체크박스 여러개 체크하기
for i in pyautogui.locateAllOnScreen("img_checkbox.png"):
    print(i)
    pyautogui.click(i, duration=1)
# 같은 이미지 여러개 찾기

# 이미지가 가변적일 경우 주변의 불변적인 이미지를 기준으로 클릭
img2 = pyautogui.locateOnScreen("img2.png")
pyautogui.click(img2.left + 200, img2.top + 0)
# 이미지 왼쪽위 모서리를 기준으로 상대적 좌표값 입력



### 이미지 찾기 속도 개선
# 1. GrayScale : 30% 속도 개선
img1 = pyautogui.locateOnScreen("img1.png", grayscale=True)

# 2. 범위지정
img1 = pyautogui.locateOnScreen("img1.png", region=(x, y, width, height))

# 3. 정확도 조정
# pip install opencv-python : confidence(신뢰성) 함수 설치
img1 = pyautogui.locateOnScreen("img1.png", confidence=0.7)
# confidence 기본값 0.999 = 99.9%



### 자동화 대상이 바로 보여지지 않는 경우 기다리기
# 1. 대상이 뜰때까지 기다려야 하는 경우
img1 = None
while img1 is None:
    img1 = pyautogui.locateOnScreen("img1.png")
    print("발견할 때까지 while 반복중")
pyautogui.moveTo(img1)
# img1값이 None이면 계속 돌고, 찾으면 img1 위치값 반환 후 while 구문 나가기

# 2. 일정 시간동안만 기다리기 (TimeOut)
import time, sys 
# 시간정보와 프로그램 종료를 위한 import
timeout = 5
start = time.time()
img1 = None
while img1 is None:
    img1 = pyautogui.locateOnScreen("img1.png")
    end = time.time()
    if end - start > timeout:
        print("시간 종료")
        sys.exit() # 프로그램 종료
    print(timeout - (end - start) , "초까지 못찾으면 종료예정")
pyautogui.moveTo(img1)
# 시작시간(start)과 종료시간(end)을 변수로 설정
# 대기시간(timeout) 설정값(5초)이 지나면 프로그램 종료



### 이미지 찾아 클릭 함수로 만들기
import time, sys

def find_target(img, timeout=10):
    start = time.time()
    target = None
    while target is None:
        target = pyautogui.locateOnScreen(img, confidence=0.8)
        end = time.time()
        if end - start > timeout:
            break
    return target

def my_click(img, timeout=10):
    target = find_target(img, timeout)
    if target:
        pyautogui.click(target, duration=0.5)
    else:
        print(f"Timeout {timeout}s. 이미지 못찾음. 프로그램 종료.")
        sys.exit()

my_click("img1.png", 5)



### 윈도우 창 제어
w = pyautogui.getActiveWindow()
# 현재 활성화된 창의 정보
print(w.title, w.size)
# 창의 제목, 크기(width, height)
print(w.left, w.right, w.top, w.bottom)
# 창의 좌표 정보

for w in pyautogui.getAllWindows():
    print(w)
for w in pyautogui.getWindowsWithTitle("제목 없음"):
    print(w)
# 모든 윈도우 확인 ->  "제목 없음" 타이틀 창 리스트 가져오기
w = pyautogui.getWindowsWithTitle("제목 없음")[0]
print(w)
# "제목없음" 타이틀 중 1번째 가져오기

if w.isActive == False:
    w.activate()
# 현재 활성화가 아니면 활성화 (맨 앞으로 가져오기)
if w.isMaximized == False:
    w.maximize()
# 현재 최대화가 아니면 최대화
if w.isMinimized == False:
    w.minimize()
# 현재 최소화가 아니면 최소화
w.restore()
# 화면 원복
w.close()
# 프로그램 종료
