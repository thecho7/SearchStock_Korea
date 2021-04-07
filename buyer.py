# Buyer 기능
# 1. 투신, 사모펀드의 시가총액 대비 순매수금액이 n%(n=1) 이상 발생한 종목 검출
# 2. 거래량이 직전 3개월 평균 대비 k배 이상(k=2) 발생한 종목 검출
# 3. 금융투자, 투신, 연기금, 사모펀드가 모두 순매수를 기록한 종목 검출
# 4. 종가 기준 상승을 기록한 종목
# 5. (1 & 2 & 3 & 4) 종목을 찾는다
# 6. 4번 기준 내림차순 정렬

import win32com.client
from connectServer import connectServer
import time

class findBuyer():
    def __init__(self, n, k):
        connectServer()
        self.obj = win32com.client.Dispatch("CpSysDib.CpSvr7254")
        self.obj.SetInputValue(1, 6)
        self.obj.SetInputValue(5, 0)
        self.n = n
        self.k = k
        self.buyer_type = {
            '전체': 0,
            '개인': 1,
            '외국인': 2,
            '기관계': 3,
            '금융투자': 4,
            '보험': 5,
            '투신': 6,
            '은행': 7,
            '기타금융': 8,
            '연기금': 9,
            '국가지자체': 10,
            '기타외인': 11,
            '사모펀드': 12,
            '기타법인': 13
        }
    
    def condFilter(self, result, momentum, cap):
        # 1. (투신 + 사모펀드)의 시가총액 대비 순매수금액이 n%(n=1) 이상 발생한 종목 검출
        cond_1 = ((result['투신'] + result['사모펀드']) * 1e6) >= cap * self.n
        # print((result['투신'] + result['사모펀드']) * 1e6, cap * self.n, cond_1)

        # 2. 거래량이 직전 3개월 평균 대비 k배 이상(k=2) 발생한 종목 검출
        cond_2 = momentum[0] > (sum(momentum[1:]) / len(momentum[1:])) * self.k
        # print( momentum[0], (sum(momentum[1:]) / len(momentum[1:])) * self.k, cond_2)
        # 3. 투신, 사모펀드가 모두 순매수를 기록한 종목 검출
        # cond_3 = (result['금융투자'] > 0) & (result['연기금'] > 0) & (result['사모펀드'] > 0)
        cond_3 = (result['투신'] > 0) & (result['사모펀드'] > 0)
        # print(result['금융투자'], result['연기금'], result['사모펀드'], cond_3)
        # 4. 종가 기준 상승을 기록한 종목 --> main에서 처리
        # 5. (1 & 2 & 3 & 4) 종목을 찾는다
        cond = cond_1 & cond_2 & cond_3
        # print(cond)
        return cond

    def getData(self, code, buyers, momentum, cap):
        self.obj.SetInputValue(0, code)
        self.obj.BlockRequest()
        result = dict()
        # 매매주체별 정보 얻기
        # 투신, 사모펀드, 금융투자, 연기금
        for b in buyers:
            result[b] = self.obj.GetDataValue(self.buyer_type[b], 0)

        return self.condFilter(result, momentum, cap) # 0, 1