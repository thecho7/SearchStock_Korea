import win32com.client
from connectServer import connectServer
from prettytable import PrettyTable

def getChart(item):
    # 연결 여부 체크
    connectServer()

    table = PrettyTable()
    # 차트 객체 구하기
    objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
    
    objStockChart.SetInputValue(0, item)   #종목 코드 - 삼성전자
    objStockChart.SetInputValue(1, ord('2')) # 개수로 조회
    objStockChart.SetInputValue(4, 100) # 최근 100일 치
    objStockChart.SetInputValue(5, [0,2,3,4,5, 8]) #날짜,시가,고가,저가,종가,거래량
    objStockChart.SetInputValue(6, ord('D')) # '차트 주가 - 일간 차트 요청
    objStockChart.SetInputValue(9, ord('1')) # 수정주가 사용
    objStockChart.BlockRequest()
    
    len = objStockChart.GetHeaderValue(3)
    
    table.field_names = ["날짜", "시가", "고가", "저가", "종가", "거래량"]
    
    for i in range(len):
        day = objStockChart.GetDataValue(0, i)
        begin = objStockChart.GetDataValue(1, i)
        high = objStockChart.GetDataValue(2, i)
        low = objStockChart.GetDataValue(3, i)
        close = objStockChart.GetDataValue(4, i)
        vol = objStockChart.GetDataValue(5, i)
        table.add_row([day, begin, high, low, close, vol])

    print(table)