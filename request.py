import pythoncom
import win32com.client
import win32event
 
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
 
StopEvent = win32event.CreateEvent(None, 0, 0, None)

 
class CpEvent:
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관
 
    def OnReceived(self):
        # 실시간 처리 - 현재가 주문 체결
        if self.name == 'stockmst':
            print('recieved')
            win32event.SetEvent(StopEvent)
            return
 
 
class CpCurReply:
    def __init__(self, objEvent):
        self.name = "stockmst"
        self.obj = objEvent
 
    def Subscribe(self):
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, None)
 
 
def MessagePump(timeout):
    waitables = [StopEvent]
    while 1:
        rc = win32event.MsgWaitForMultipleObjects(
            waitables,
            0,  # Wait for all = false, so it waits for anyone
            timeout, #(or win32event.INFINITE)
            win32event.QS_ALLEVENTS)  # Accepts all input
 
        if rc == win32event.WAIT_OBJECT_0:
            # Our first event listed, the StopEvent, was triggered, so we must exit
            print('stop event')
            break
 
        elif rc == win32event.WAIT_OBJECT_0 + len(waitables):
            # A windows message is waiting - take care of it. (Don't ask me
            # why a WAIT_OBJECT_MSG isn't defined < WAIT_OBJECT_0...!).
            # This message-serving MUST be done for COM, DDE, and other
            # Windowsy things to work properly!
            print('pump')
            if pythoncom.PumpWaitingMessages():
                break  # we received a wm_quit message
        elif rc == win32event.WAIT_TIMEOUT:
            print('timeout')
            return
            pass
        else:
            print('exception')
            raise RuntimeError("unexpected win32wait return value")
 
 
code = 'A005930'
 
##############################################################
#1. BlockRequest
print('#####################################')
# objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
objStockMst = win32com.client.Dispatch('CpSysDib.CpMarketWatch')
objStockMst.SetInputValue(0, code)
# objStockMst.BlockRequest()
# print('BlockRequest 로 수신 받은 데이터')
# item = {}
# item['종목명']= g_objCodeMgr.CodeToName(code)
# item['현재가'] = objStockMst.GetHeaderValue(11)  # 종가
# item['대비'] =  objStockMst.GetHeaderValue(12)  # 전일대비
# print(item)
 
# print('')
##############################################################
# 2. Request ==> 메시지 펌프 ==>  OnReceived 이벤트 수신
print('#####################################')
objReply = CpCurReply(objStockMst)
objReply.Subscribe()
 
code = 'A005930'
objStockMst.SetInputValue(0, code)
objStockMst.Request()
MessagePump(10000)
item = {}
item['종목명']= g_objCodeMgr.CodeToName(code)
item['현재가'] = objStockMst.GetHeaderValue(11)  # 종가
item['대비'] =  objStockMst.GetHeaderValue(12)  # 전일대비
print(item)