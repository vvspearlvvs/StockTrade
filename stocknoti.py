import win32com.client

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 현재가 객체 구하기
objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
objStockMst.SetInputValue(0, 'A005930')   #종목 코드 - 삼성전자
objStockMst.BlockRequest()

# 현재가 통신 및 통신 에러 처리
rqStatus = objStockMst.GetDibStatus()
rqRet = objStockMst.GetDibMsg1()
print("통신상태", rqStatus, rqRet)
if rqStatus != 0:
    exit()

# 현재가 정보 조회
'''
code = objStockMst.GetHeaderValue(0)  #종목코드
name= objStockMst.GetHeaderValue(1)  # 종목명
time= objStockMst.GetHeaderValue(4)  # 시간
cprice= objStockMst.GetHeaderValue(11) # 종가
diff= objStockMst.GetHeaderValue(12)  # 대비
open= objStockMst.GetHeaderValue(13)  # 시가
high= objStockMst.GetHeaderValue(14)  # 고가
low= objStockMst.GetHeaderValue(15)   # 저가
bid = objStockMst.GetHeaderValue(17)   #매수호가
vol= objStockMst.GetHeaderValue(18)   #거래량
vol_value= objStockMst.GetHeaderValue(19)  #거래대금
# 예상 체결관련 정보

exPrice = objStockMst.GetHeaderValue(55) #예상체결가
exDiff = objStockMst.GetHeaderValue(56) #예상체결가 전일대비
exVol = objStockMst.GetHeaderValue(57) #예상체결수량
'''

offer = objStockMst.GetHeaderValue(16)  #매도호가
exFlag = objStockMst.GetHeaderValue(58) #예상체결가 구분 플래그

print("매도호가", offer)

if (exFlag == ord('0')):
    print("장 구분값: 동시호가와 장중 이외의 시간")
elif (exFlag == ord('1')) :
    print("장 구분값: 동시호가 시간")
elif (exFlag == ord('2')):
    print("장 구분값: 장중 또는 장종료")

from slacker import Slacker

slack = Slacker('xoxb-Slack토큰명')

# Send a message to #general channel
slack.chat.post_message('#stock_noti', '삼성전자 현재가(매도호가):'+str(offer))

'''
# Get users list
response = slack.users.list()
users = response.body['members']

# Upload a file
slack.files.upload('hello.txt')

# If you need to proxy the requests
proxy_endpoint = 'http://myproxy:3128'
slack = Slacker('<your-slack-api-token-goes-here>',
                http_proxy=proxy_endpoint,
                https_proxy=proxy_endpoint)

# Advanced: Use `request.Session` for connection pooling (reuse)
from requests.sessions import Session
with Session() as session:
    slack = Slacker(token, session=session)
    slack.chat.post_message('#general', 'All these requests')
    slack.chat.post_message('#general', 'go through')
    slack.chat.post_message('#general', 'a single https connection')
    '''
