import win32com.client

class stockMstGet():
    def __init__(self):
        self.objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        self.item_type = {
            '52주최고가': 47,
            '52주최저가': 49
        }

    def get(self, code, price):
        result = dict()
        self.objStockMst.SetInputValue(0, code)
        self.objStockMst.BlockRequest()
        for key, val in self.item_type.items():
            target = self.objStockMst.GetHeaderValue(val)
            if key == '52주최고가':
                result[key] = 'O' if price >= target else 'X'

            if key == '52주최저가':
                result[key] = 'O' if price <= target else 'X'
      
        return result