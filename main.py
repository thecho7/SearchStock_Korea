import win32com.client
from connectServer import connectServer
from getStockNameAll import getStockNameAll
from getDailyInfo import getItemInfoSingle
from telegramBot import telegramMachine
from getMarketCap import CMarketTotal
from customIndices import netBuyCapRatioInst, continuousBuy, statistics
from stockMstInfo import stockMstGet
from buyer import findBuyer
from utils import sendTeleg, makeTable
from tqdm import tqdm
import datetime
import getEvents
import time

import os
os.environ["PATH"] = '.' + ";"+os.environ["PATH"] 

flags = dict()
search_cap_upper = 500000000000000
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr') # g_objCodeMgr.CodeToName(code)

# # 오늘 날짜 (주말의 경우는?)
today = datetime.datetime.now()
today_date = today.strftime('%Y%m%d')
stock_info = getStockNameAll()
infoSingle = getItemInfoSingle()

time.sleep(5)
objMarketTotal = CMarketTotal()
objMarketTotal.GetAllMarketTotal()

teleg = telegramMachine(telgm_token = 'YOUR_TELEGRAM_TOKEN', chat_id = 'YOUR_CHAT_ID')
sendObj = sendTeleg(teleg, top_rank=5)

# 오늘 날짜 고지
teleg.sendMsg('[{}] 정보\n'.format(today_date))

# 각 종목 시총 계산
print("각 종목 시총 구하는 중...")
info_all = dict()
for i, (code, name) in enumerate(tqdm(stock_info.items())):
    amount, cap = objMarketTotal.dataInfo[code]
    # 시총 3,000억 이하의 종목들만 탐색
    if cap <= search_cap_upper:
        info_all[code] = dict()
        info_all[code]['current'] = [amount, cap]

# 모든 종목에 대한 모든 정보 읽기
print("각 종목 기록 구하는 중...")
num_week = 2
for i, code in enumerate(tqdm(info_all.keys())):
    info_all[code]['time_series'] = infoSingle.requestData(code, num_week)
    time.sleep(0.25 * num_week)

# 조건 1
print("조건 1 구하는 중...")
fb = findBuyer(0.01, 2)
yearminmax = stockMstGet()
res_cond1 = dict()
period_1 = 60 + 1
for i, code in enumerate(tqdm(info_all.keys())):
    # 상장 후 3개월 이상 된 주식이며
    if len(info_all[code]['time_series']) >= period_1:
        # 종가 기준 전날 대비 상승한 종목에 대해서
        today_price, yesterday_price = info_all[code]['time_series'][0][4], info_all[code]['time_series'][1][4]
        price_change = today_price / yesterday_price
        
        momentum_60 = [info_all[code]['time_series'][i][6] for i in range(period_1)]
        if sum(momentum_60[1:]):
            momentum_change = momentum_60[0] / (sum(momentum_60[1:]) / (period_1 - 2))

            if price_change >= 1 and fb.getData(code, ['투신', '사모펀드', '금융투자', '연기금'], momentum_60, info_all[code]['current'][1]):
                res_cond1[code] = []
                #### res_cond1: [최근 60일 대비 당일거래량 비율(%), 어제오늘 종가 비율(%), 52주최고가여부, 52주최저가여부, 시총대비외인순매수, 시총대비기관순매수]
                res_cond1[code].append(int(100 * momentum_change))
                res_cond1[code].append('{:.2f}%'.format(float(100 * (price_change - 1))))
                for val in yearminmax.get(code, today_price).values():
                    res_cond1[code].append(val)

            time.sleep(0.4)

# 시총대비 순매수 순위 (ETF/ETN/선물/옵션 등 제외)
print("시총대비 기관/외인 순매수 구하는 중...")
buy_cap_ratio_i = dict()
buy_cap_ratio_f = dict()        
for i, (code, name) in enumerate(tqdm(stock_info.items())):
    if g_objCodeMgr.GetStockSectionKind(code) == 1 and (code in info_all.keys()):
        last_inst_buy = info_all[code]['time_series'][0][12] * info_all[code]['time_series'][0][4]
        last_foreign_buy = info_all[code]['time_series'][0][21] * info_all[code]['time_series'][0][4]
        cap = info_all[code]['current'][1]
        buy_cap_ratio_i[code] = [last_inst_buy, cap]
        buy_cap_ratio_f[code] = [last_foreign_buy, cap]
        time.sleep(0.3) # Creon API: 15초에 60종목 (4/sec)

buy_cap_ratio_i_res = netBuyCapRatioInst(buy_cap_ratio_i)
buy_cap_ratio_f_res = netBuyCapRatioInst(buy_cap_ratio_f)

for code in res_cond1.keys():
    print(g_objCodeMgr.CodeToName(code))
    res_cond1[code].append('{}%'.format(100 * round(buy_cap_ratio_i_res[code], 2)))
    res_cond1[code].append('{}%'.format(100 * round(buy_cap_ratio_f_res[code], 2)))

# 거래량 순으로 정렬
res_cond1 = dict(sorted(res_cond1.items(), key=(lambda x: x[1][0]), reverse=True))

# 거래량 변환(int --> str)
for key in res_cond1.keys():
    res_cond1[key][0] = '{}%'.format(res_cond1[key][0])

if len(res_cond1) != 0:
    sendObj.sendTable(res_cond1)

##################################################################
sendObj.send(buy_cap_ratio_i_res, 0)
sendObj.send(buy_cap_ratio_f_res, 1)

cont_inst = continuousBuy(info_all, g_objCodeMgr, idx=12)
cont_foreign = continuousBuy(info_all, g_objCodeMgr, idx=21)
sendObj.send(cont_inst, 2)
sendObj.send(cont_foreign, 3)

##################################################################
# 기관, 외인 연속매수일 별 통계
# stats = statistics()
# days = [1, 5, 20]
# stats.contBuyHist(info_all, days, 12)
# stats.contBuyHist(info_all, days, 21)
# stats.saveExcel('test2.xlsx')

## 뉴스
# objMarketWatch = getEvents.CpRpMarketWatch()
# objMarketWatch.Request('*')
# news = '<특징주 소식>\n'
# for new in objMarketWatch.result:
#     news = news + '[{}] {}\n{}\n'.format(new['시간'], new['종목명'], new['특이사항'])

# teleg.sendMsg(news)