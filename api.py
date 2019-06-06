# Capital API in python insert postgresDB-taiwan_finance
import psycopg2
import pythoncom, time, os
import comtypes.client as cc  
cc.GetModule(r'D:\api\SKCOM\x64\SKCOM.dll')
import comtypes.gen.SKCOMLib as sk 
# 建立COM物件
skC = cc.CreateObject(sk.SKCenterLib, interface = sk.ISKCenterLib)
skQ = cc.CreateObject(sk.SKQuoteLib, interface = sk.ISKQuoterLib)

# Some configuration
ID = ''
PW = ''

# 顯示 event 事件, t秒內，每隔一秒，檢查有沒有事件發生
def pumpwait(t=1):
    for i in range(t):
        time.sleep(1)
        pythoncom.PumpWaitingMessages()

# 建立事件類別
class skQ_events:
    def __init__(self):
        self.KlineData=[]
    def OnConnection(self, nKind, Code):
        if Code == 0:
            if nKind == 3001 :
                print("連線中, nKind = ", nKind)
            elif Code == 0 & (nKind == 3003):
                print("連線成功, nKind = ", nKind)
    def OnNotifyKlineData(self, bstrStockNo, bstrData):
        #
        self.KlineData.append(bstrData.split(','))

# Event sink, 事件實體
EventQ = skQ_events()

# make connnection to event sink
ConnectionQ = cc.GetEvents(skQ, EventQ)

# 登入,連線報價主機
print("Login,", skC.SKCenterlib_GetReturnCodeMessage(skC.SKCenterLib_Login(ID,PW)))
print("EnterMonitor,", skC.SKCenterLib_GetReturnCodeMassege(skQ.SKQuoteLib_EnterMonitor()))
pumpwait(8)

#讀取加權日k歷史報價
# bstrStockNo 放股票代碼
# sKineType,0 = 1分鐘線, 3 = 日線288天, 4 = 完整日線, 5 = 週線, 6 = 月線
# sQutType, 0 = 舊版輸出格式, 1 = 新版輸出格式。新版格式一分日期與分時資料是個一個欄位
# 將data 站存至 EventQ.KineData
EventQ.KlineData=[]

# 請求歷史報價，參數可能有出入，自行參考官方API文件

nCode = skQ.SKQuoteLib_RequestKline('TSEA', sKlineType = 4, sOutType = 1)

# data 輸出
EventQ.KlineData[0:2]

# 取1分k歷史報價


