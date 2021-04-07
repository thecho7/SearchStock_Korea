import os
os.environ["PATH"] = '.' + ";"+os.environ["PATH"]
import win32com
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import datetime 

def makeTable(data, title, subtitle, filename):
    plt.rcParams['font.family'] = 'Malgun Gothic'
    N = len(data) - 1
    column_headers = data[0]
    row_headers = [row[0] for row in data[1:]]
    
    cell_text = [row[1:] for row in data[1:]]

    fig, ax1 = plt.subplots(figsize=(10, 2 + N / 2.5))

    rcolors = np.full(len(row_headers), 'linen')
    ccolors = np.full(len(column_headers), 'lavender')

    table = ax1.table(cellText=cell_text,
                    cellLoc='center',
                    rowLabels=row_headers,
                    rowColours=rcolors,
                    rowLoc='center',
                    colColours=ccolors,
                    colLabels=column_headers,
                    loc='center')
    table.scale(1, 2)
    table.auto_set_font_size(False)
    table.set_fontsize(12)

    for i in range(len(column_headers)):
        cell = table[0, i]
        cell.set_fontsize(8)

    ax1.axis('off')
    ax1.set_title(f'{title}\n({subtitle})', weight='bold', size=12, color='k')

    plt.savefig(filename, dpi=200, bbox_inches='tight')

class sendTeleg():
    def __init__(self, teleg_, top_rank):
        today = datetime.datetime.now()
        self.today_date = today.strftime('%Y%m%d')
        self.g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
        self.top_rank = top_rank
        self.head_msg = ['시총대비 기관 순매수 TOP (ETF/ETN/선물/옵션 등 제외)',
                        '시총대비 외국인 순매수 TOP (ETF/ETN/선물/옵션 등 제외)',
                        '기관 연속순매수일 TOP (ETF/ETN/선물/옵션 등 제외)',
                        '외국인 연속순매수일 TOP (ETF/ETN/선물/옵션 등 제외)']
        self.teleg = teleg_

    # i == 0 기관 1 외국인
    def send(self, data, idx):
        # 시총대비 순매수 (기관, 외국인)
        res_msg = '\n'
        for rank, i in enumerate(data.keys()):
            if rank >= self.top_rank:
                break
            res_msg = res_msg + '{} {}\n'.format(self.g_objCodeMgr.CodeToName(i), data[i])

        self.teleg.sendMsg(self.head_msg[idx] + res_msg)

    def sendTable(self, data):
        new_data = []
        print(data)
        for idx, (code, vals) in enumerate(data.items()):
            new_data.append([])
            new_data[idx].append(self.g_objCodeMgr.CodeToName(code))
            new_data[idx].extend(vals)

        new_data.insert(0, ['최근 60일 대비\n당일거래량 비율(%)', '어제오늘 종가 비율(%)', '52주최고가', '52주최저가', '시총대비\n외인순매수(%)', '시총대비\n기관순매수(%)'])
        title = 'Passing lane 1'
        subtitle = '{}'.format(self.today_date)
        filename = '{}_table1.png'.format(self.today_date)
        makeTable(new_data, title, subtitle, filename)
        self.teleg.sendImage(filename)
