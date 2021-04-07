import matplotlib
from openpyxl import Workbook

# input: {code1: {buy1, cap1}, code2: {buy2, cap2}, ...}
def netBuyCapRatioInst(info):
    res = dict()
    for code, val in info.items():
        buy, cap = val[0], val[1]
        if cap > 0:
            res[code] = round(buy / cap, 3)
        else:
            res[code] = 0.

    return dict(sorted(res.items(), key=lambda item: item[1], reverse=True))

def continuousBuy(info_all, g_objCodeMgr, idx):
    res = dict()
    for code, val in info_all.items():
        if g_objCodeMgr.GetStockSectionKind(code) == 1:
            for day, info in enumerate(val['time_series']):
                if info[idx] <= 0:
                    break
            res[code] = day

    return dict(sorted(res.items(), key=lambda item: item[1], reverse=True))

# 거래량의 이동평균 계산
# 5, 20, 60, 120 (1주일, 1개월, 3개월, 6개월)
# k: 각 기간을 기준으로, 직전일 이동평균 거래량 대비 k배 거래량 상승 종목
def movingAverageMomentum(code, info, k):
    period = [5, 20, 60, 120]
    mvavg = dict()
    for p in period:
        if len(info) < p:
            break

class statistics():
    def __init__(self):
        self.book = Workbook()
        # 서로 다른 통계는 다른 sheet에 저장
        self.indices = {12: '기관 연속순매수 기준 이익 횟수',
                        21: '외국인 연속순매수 기준 이익 횟수'}
        
    # 기관 및 외국인 연속 순매수/순매도일 기준 n일 뒤 각 종목 수익률 통계(Histogram)
    # idx: 12(기관), 21(외국인)
    # data: data[code]['time_series']
    def contBuyHist(self, data, days, idx):
        ws = self.book.create_sheet(self.indices[idx])
        col_title = ['{}일 후 상승 횟수'.format(day) for day in days]
        col_title.insert(0, '연속매수일')
        # ws.append(('연속매수일', '{}일 후 상승 횟수'.format(day)))
        ws.append(col_title)
        res = dict()
        for code, _ in data.items():
            seq = data[code]['time_series']
            d_cnt = 0
            # n일 후의 주가를 구할 수 없는 경우는 건너 뛴다
            for i in range(len(seq) - 1, days[0], -1):
                # 순매수 or 보합
                if seq[i][idx] > 0:
                    d_cnt = max(d_cnt, 0) + 1
                # 순매도
                elif seq[i][idx] < 0:
                    d_cnt = min(d_cnt, 0) - 1
                else:
                    d_cnt = 0    
                
                if d_cnt not in res.keys():
                    res[d_cnt] = [[0, 0]] * len(days)
                
                # day일 후는 i - day
                for j, day in enumerate(days):
                    if (i - day) >= 0:
                        print(day, seq[i - day][4], seq[i][4])
                        print(res[d_cnt])
                        res[d_cnt][j][1] += 1
                        if seq[i - day][4] >= seq[i][4]:
                            res[d_cnt][j][0] += 1
                        else:
                            res[d_cnt][j][0] -= 1
                        print(res[d_cnt])
        res = dict(sorted(res.items(), key=lambda item: item[0], reverse=False))

        for k, v in res.items():
            result = list()
            for a in v:
                result.append(a[0])
                result.append(a[1])
                result.append(round(a[0] / a[1], 3))
            # result = [round(a[0] / a[1], 3) for a in v]
            result.insert(0, int(k))
            ws.append(result)
        # matplotlib이 32bit에서 동작 안함(해결 필요)
        # matplotlib.pyplot.bar(*zip(*res.items()))
        # matplotlib.pyplot.show()

    def saveExcel(self, path):
        self.book.save(filename=path)