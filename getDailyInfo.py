import win32com.client
from connectServer import connectServer
import time

class getItemInfoSingle():
    def __init__(self):
        # 연결 여부 체크
        connectServer()
        # 일자별 object 구하기
        self.objStockWeek = win32com.client.Dispatch("DsCbo1.StockWeek")
        self.info_field = {
            0: '날짜', 
            1: '시가', 
            2: '고가', 
            3: '저가', 
            4: '종가', 
            5: '전일대비', 
            6: '누적거래량', 
            7: '외인보유', 
            8: '외인보유전일비',
            9: '외인비중', 
            10: '등락율', 
            11: '대비부호', 
            12: '기관순매수수량', 
            13: '시간외단일가시가', 
            14: '시간외단일가고가', 
            15: '시간외단일가저가', 
            16: '시간외단일가종가', 
            17: '시간외단일가대비부호', 
            18: '시간외단일가전일대비', 
            19: '시간외단일가거래량', 
            20: '거래대금(만원)', 
            21: '외국인 순매수 수량',
            22: '시가총액'} # DsCbo1.StockWeek field에는 없음

    def checkValidity(self):   
        # 통신 결과 확인
        rqStatus = self.objStockWeek.GetDibStatus()
        rqRet = self.objStockWeek.GetDibMsg1()
        # print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
        return True

    def requestData(self, code, num_week):
        # self.objStockWeek.SetInputValue(0, 'A005930')   #종목 코드 - 삼성전자
        self.objStockWeek.SetInputValue(0, code)
        return self.getData(num_week)
        
    def getData(self, num_week):
        res = list()
        # 연속 데이터 요청
        # 예제는 5번만 연속 통신 하도록 함. 한번에 36일치를 받아옴
        week_cnt = 0
        while week_cnt < num_week:
            self.objStockWeek.BlockRequest()
            assert self.checkValidity()

            # 일자별 정보 데이터 처리
            count = self.objStockWeek.GetHeaderValue(1)  # 데이터 개수
            for i in range(count):
                # date = str(self.objStockWeek.GetDataValue(0, i))  # 일자
                res.append([self.objStockWeek.GetDataValue(j, i) for j in range(22)])

            week_cnt+=1
            if not self.objStockWeek.Continue:
                break
            # time.sleep(0.2)
        return res