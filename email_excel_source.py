### email과 excel 사용
import smtplib
from account import *
# account.py 파일에 EMAIL_ADDRESS, EMAIL_PASSWORD 변수 설정

with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
    smtp.ehlo() # 연결이 잘 수립되는지 확인
    smtp.starttls() # 모든 내용이 암호화되어 전송
    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD) # 로그인

    subject = "test mail" # 메일 제목
    body = "mail body" # 메일 본문
    msg = f"Subject: {subject}\n{body}"

    # 발신자, 수신자, 정해진 형식의 메시지
    smtp.sendmail(EMAIL_ADDRESS, "hoocarol4@gmail.com", msg)


## EmailMessage 이용
from email.message import EmailMessage
msg = EmailMessage()
msg["Subject"] = "테스트 메일입니다" # 제목
msg["From"] = EMAIL_ADDRESS # 보내는 사람
msg["To"] = "hoocarol4@gmail.com" # 받는 사람

# 여러 명에게 메일을 보낼 때
msg["To"] = "nadocoding@gmail.com, nadocoding@gmail.com"

to_list = ["nadocoding@gmail.com", "nadocoding@gmail.com"]
msg["To"] = ", ".join(to_list)

# 참조
msg["Cc"] = "nadocoding@gmail.com"
# 비밀참조
msg["Bcc"] = "nadocoding@gmail.com"

# 본문
msg.set_content("테스트 본문입니다")

with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
    smtp.ehlo()
    smtp.starttls()
    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    smtp.send_message(msg)


## 파일첨부
msg.set_content("다운로드 하세요")
# MIME Type
with open("btn_brush.png", "rb") as f:
    msg.add_attachment(f.read(), maintype="image", subtype="png", filename=f.name)

with open("테스트.pdf", "rb") as f:
    msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=f.name)

with open("엑셀.xlsx", "rb") as f:
    msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=f.name)

with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
    smtp.ehlo()
    smtp.starttls()
    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    smtp.send_message(msg)


## 받은메일함 관리
from imap_tools import MailBox
from account import *

mailbox = MailBox("imap.gmail.com", 993)
mailbox.login(EMAIL_ADDRESS, EMAIL_PASSWORD, initial_folder="INBOX")

# limit : 최대 메일 갯수
# reverse : True 일 경우 최근 메일부터, False 일 경우 과거 메일부터
for msg in mailbox.fetch(limit=1, reverse=True):
    print("제목", msg.subject)
    print("발신자", msg.from_)
    print("수신자", msg.to)
    print("참조자", msg.cc)
    print("비밀참조자", msg.bcc)
    print("날짜", msg.date)
    print("본문", msg.text)
    print("HTML 메시지", msg.html)
    print("=" * 100)

    # 첨부 파일
    for att in msg.attachments:
        print("첨부파일 이름", att.filename)
        print("타입", att.content_type)
        print("크기", att.size)

        # 파일 다운로드
        with open("download_" + att.filename, "wb") as f:
            f.write(att.payload)
            print("첨부 파일 {} 다운로드 완료".format(att.filename))

mailbox.logout()

## 받은메일함 검색
with MailBox("imap.gmail.com", 993).login(EMAIL_ADDRESS, EMAIL_PASSWORD, initial_folder="INBOX") as mailbox:
    # 전체 메일 다 가져오기
    for msg in mailbox.fetch():
        print("[{}] {}".format(msg.from_, msg.subject))

    # 읽지 않은 메일 가져오기
    for msg in mailbox.fetch('(UNSEEN)'):
        print("[{}] {}".format(msg.from_, msg.subject))

    # 특정인이 보낸 메일 가져오기
    for msg in mailbox.fetch('(FROM nadocoding@gmail.com)', limit=3, reverse=True):
        print("[{}] {}".format(msg.from_, msg.subject))

    # 작은 따옴표로 먼저 감싸고, 실제 TEXT 부분은 큰 따옴표로 감싸주세요
    # 어떤 글자를 포함하는 메일 (제목, 본문)
    for msg in mailbox.fetch('(TEXT "test mail")'):
        print("[{}] {}".format(msg.from_, msg.subject))

    # 어떤 글자를 포함하는 메일 (제목만)
    for msg in mailbox.fetch('(SUBJECT "test mail")'):
       print("[{}] {}".format(msg.from_, msg.subject))
    
    # 어떤 글자(한글)을 포함하는 메일 필터링 (제목만)
    for msg in mailbox.fetch(limit=5, reverse=True):
        if "테스트" in msg.subject:
            print("[{}] {}".format(msg.from_, msg.subject))

    # 특정 날짜 이후의 메일
    for msg in mailbox.fetch('(SENTSINCE 07-Nov-2020)', reverse=True, limit=5):
        print("[{}] {}".format(msg.from_, msg.subject))
    
    # 특정 날짜에 온 메일
    for msg in mailbox.fetch('(ON 07-Nov-2020)', reverse=True, limit=5): 
        print("[{}] {}".format(msg.from_, msg.subject))

    # 2가지 이상의 조건을 모두 만족하는 메일 (그리고 조건)
    for msg in mailbox.fetch('(ON 07-Nov-2020 SUBJECT "test mail")', reverse=True, limit=5): 
        print("[{}] {}".format(msg.from_, msg.subject))

    # 2가지 이상의 조건 중 하나라도 만족하는 메일 (또는 조건)
    for msg in mailbox.fetch('(OR ON 07-Nov-2020 SUBJECT "test mail")', reverse=True, limit=5): 
        print("[{}] {}".format(msg.from_, msg.subject))



### project
import smtplib
from random import *
from account import *
from email.message import EmailMessage

## 메일보내기
# 제목 : 파이썬 특강 신청합니다.
# 본문 : 닉네임/전화번호 뒤 4자리 (예 : 나도코딩/1234)

nicknames = ["유재석", "박명수", "정형돈", "노홍철", "조세호"]
with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
    smtp.ehlo()
    smtp.starttls()
    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    for nickname in nicknames:
        msg = EmailMessage()
        msg["Subject"] = "파이썬 특강 신청합니다."
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = "hoocarol4@gmail.com"
        content = "/".join([nickname, str(randint(1000, 9999))])
        # content = nickname + "/" + str(randint(1000, 9999))
        msg.set_content(content)
        smtp.send_message(msg)
        print(nickname + "님이 나도코딩 계정으로 메일 발송 완료")

## 받은메일 조회
applicant_list = [] # 지원자 리스트
with MailBox("imap.gmail.com", 993).login(EMAIL_ADDRESS, EMAIL_PASSWORD, initial_folder="INBOX") as mailbox:
    index = 1 # 순번
    for msg in mailbox.fetch('(SENTSINCE 07-Nov-2020)'): # 2020년 11월 7일 이후로 온 메일 조회
        if "파이썬 특강" in msg.subject:
            nickname, phone = msg.text.strip().split("/")
            print("순번 : {} 닉네임 : {} 전화번호 : {}".format(index, nickname, phone))
            applicant_list.append((msg, index, nickname, phone))
            index += 1
for applicant in applicant_list:
    print(applicant)

## 메일 일괄 답장
import smtplib
from account import *
from imap_tools import MailBox
from email.message import EmailMessage

max_val = 3 # 최대 선정자 수
applicant_list = [] # 지원자 리스트

print("[1. 지원자 메일 조회]")
with MailBox("imap.gmail.com", 993).login(EMAIL_ADDRESS, EMAIL_PASSWORD, initial_folder="INBOX") as mailbox:
    index = 1 # 순번
    for msg in mailbox.fetch('(SENTSINCE 07-Nov-2020)'): # 2020년 11월 7일 이후로 온 메일 조회
        if "파이썬 특강" in msg.subject:
            nickname, phone = msg.text.strip().split("/")
            # print("순번 : {} 닉네임 : {} 전화번호 : {}".format(index, nickname, phone))
            applicant_list.append((msg, index, nickname, phone))
            index += 1

print("[2. 선정 / 탈락 메일 발송]")
with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
    smtp.ehlo()
    smtp.starttls()
    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    for applicant in applicant_list:
        to_addr = applicant[0].from_ # 수신 메일 주소
        # index = applicant[1]
        # nickname = applicant[2]
        # phone = applicant[3]
        index, nickname, phone = applicant[1:]

        title = None
        content = None

        if index <= max_val:
            title = "파이썬 특강 안내 [선정]"
            content = "{}님 축하드립니다. 특강 대상자로 선정되셨습니다. (선정순번 {}번)".format(nickname, index)
        else:
            title = "파이썬 특강 안내 [탈락]"
            content = "{}님 아쉽게도 탈락입니다. 취소 인원이 발생하는 경우 연락드리겠습니다. (대기순번 {}번)".format(nickname, index - max_val)

        msg = EmailMessage()
        msg["Subject"] = title
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = to_addr
        msg.set_content(content)
        smtp.send_message(msg)
        print(nickname, "님에게 메일 발송 완료")
        
## 선정대상자 엑셀 저장
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.append(["순번", "닉네임", "전화번호"])

for applicant in applicant_list[:max_val]:
    ws.append(applicant[1:])
    # index = applicant[1]
    # nickname = applicant[2]
    # phone = applicant[3]
    # ws.append([index, nickname, phone])
wb.save("result.xlsx")
print("모든 작업이 완료되었습니다.")
