import win32com.client
from connectServer import connectServer

def getStockNameAll():
    stock_info = dict()
    # 연결 여부 체크
    connectServer()
    
    # 종목코드 리스트 구하기
    objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    codeList = objCpCodeMgr.GetStockListByMarket(1) #거래소
    codeList2 = objCpCodeMgr.GetStockListByMarket(2) #코스닥
    
    # f = open('ETF는 누구인가2.txt', 'w', encoding='utf8')
    # objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
    for i, code in enumerate(codeList):   
        # objStockMst.SetInputValue(0, code)
        # objStockMst.BlockRequest()
        # cprice= objStockMst.GetHeaderValue(11) # 종가

        secondCode = objCpCodeMgr.GetStockSectionKind(code)
        name = objCpCodeMgr.CodeToName(code)
        # stdPrice = objCpCodeMgr.GetStockStdPrice(code)
        if secondCode == 1:
            stock_info[code] = name
    
    # print("코스닥 종목코드", len(codeList2))
    for i, code in enumerate(codeList2):
        secondCode = objCpCodeMgr.GetStockSectionKind(code)
        name = objCpCodeMgr.CodeToName(code)
        # stdPrice = objCpCodeMgr.GetStockStdPrice(code)
        if secondCode == 1:
            stock_info[code] = name        
    
    # print("거래소 + 코스닥 종목코드 ",len(codeList) + len(codeList2))
    return stock_info