### 엑셀
# pip install openpyxl
from openpyxl import Workbook


## 워크북 생성 제어
wb = Workbook()
# 새 워크북 생성
wb.save("sample.xlsx")
wb.close()


## sheet 제어
ws1 = wb.active
# 현재 활성화된 sheet 가져옴
ws1.title = "sheet1"
# sheet1의 이름을 변경
ws2 = wb.create_sheet("Sheet2", 1)
# 1번째 index에 "Sheet2" 생성
ws2.sheet_properties.tabColor = "ff66ff"
# RGB 형태로 값을 넣어주면 탭 색상 변경

ws1 = wb["Sheet1"]
# Dict형태로 Sheet 명으로 접근
print(wb.sheetnames)
# 모든 Sheet 이름 확인

# Sheet 복사
copy1 = wb.copy_worksheet(ws1)
copy1.title = "Copied Sheet1"


## 셀 제어
ws1["A1"] = "Test"
# A1에 "test" 입력

### 업무자동화 엑셀 3번 셀부터 이어서 하기