import tkinter as tk
from tkinter import messagebox
import time
import pythoncom
import threading
import comtypes.client

# === 1. SKCOM DLL 路徑 ===
DLL_PATH = r"CapitalAPI_2.13.55_PythonExample\PythonExampleV4\SKCOM.dll"  # 改成你的 DLL 路徑
comtypes.client.GetModule(DLL_PATH)
import comtypes.gen.SKCOMLib as sk

# === 2. 建立 COM 物件 ===
skC = comtypes.client.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
skQ = comtypes.client.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib)
skO = comtypes.client.CreateObject(sk.SKOrderLib, interface=sk.ISKOrderLib)

# === 3. 報價事件 ===
class QuoteEvent:
    def OnConnection(self, nKind, nCode):
        if nKind == 3001:
            print("報價主機：連線中")
        elif nKind == 3002:
            print("報價主機：已斷線")
        elif nKind == 3003:
            print("報價主機：連線成功 (Ready)")
        else:
            print("OnConnection:", nKind, nCode)

    def OnNotifyQuote(self, sMarketNo, sStockIdx):
        stock = sk.SKSTOCK()
        skQ.SKQuoteLib_GetStockByIndex(sMarketNo, sStockIdx, stock)
        price = stock.nClose / (10 ** stock.sDecimal)
        code = stock.bstrStockNo
        # 更新 GUI
        app.update_price(code, price)

quote_handler = comtypes.client.GetEvents(skQ, QuoteEvent())

# === 4. 下單功能 ===
def send_order(full_account, stock_code, price, qty, buy_sell=0, mock=True):
    """
    buy_sell: 0 = 買, 1 = 賣
    mock: True = 模擬下單, False = 正式下單
    """
    order = sk.SKSTOCKORDER()
    order.bstrFullAccount = full_account
    order.bstrStockNo = stock_code
    order.sBuySell = buy_sell
    order.sTradeType = 0  # 現股
    order.sDayTrade = 0
    order.sPrice = int(price*100)  # 乘 100
    order.nQty = qty
    order.bstrMessage = ""

    if mock:
        ret = skO.SKOrderLib_SendStockOrderM(order)
    else:
        ret = skO.SKOrderLib_SendStockOrder(order)

    msg = skC.SKCenterLib_GetReturnCodeMessage(ret)
    return ret, msg

# === 5. GUI ===
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("群益 API Python GUI")

        # 登入區
        tk.Label(root, text="帳號:").grid(row=0, column=0)
        self.entry_account = tk.Entry(root)
        self.entry_account.grid(row=0, column=1)

        tk.Label(root, text="密碼:").grid(row=1, column=0)
        self.entry_password = tk.Entry(root, show="*")
        self.entry_password.grid(row=1, column=1)

        self.btn_login = tk.Button(root, text="登入", command=self.login)
        self.btn_login.grid(row=2, column=0, columnspan=2)

        # 報價顯示
        tk.Label(root, text="股票代碼:").grid(row=3, column=0)
        self.entry_stock = tk.Entry(root)
        self.entry_stock.grid(row=3, column=1)
        self.lbl_price = tk.Label(root, text="成交價: -")
        self.lbl_price.grid(row=4, column=0, columnspan=2)

        self.btn_subscribe = tk.Button(root, text="訂閱報價", command=self.subscribe)
        self.btn_subscribe.grid(row=5, column=0, columnspan=2)

        # 下單區
        tk.Label(root, text="下單張數:").grid(row=6, column=0)
        self.entry_qty = tk.Entry(root)
        self.entry_qty.grid(row=6, column=1)

        tk.Label(root, text="價格:").grid(row=7, column=0)
        self.entry_price = tk.Entry(root)
        self.entry_price.grid(row=7, column=1)

        self.mock_var = tk.IntVar(value=1)
        tk.Checkbutton(root, text="模擬下單", variable=self.mock_var).grid(row=8, column=0, columnspan=2)

        self.btn_buy = tk.Button(root, text="買進", command=lambda: self.order(0))
        self.btn_buy.grid(row=9, column=0)
        self.btn_sell = tk.Button(root, text="賣出", command=lambda: self.order(1))
        self.btn_sell.grid(row=9, column=1)

    # 更新報價
    def update_price(self, code, price):
        if self.entry_stock.get() == code:
            self.lbl_price.config(text=f"成交價: {price}")

    # 登入事件
    def login(self):
        account = self.entry_account.get()
        password = self.entry_password.get()
        ret = skC.SKCenterLib_Login(account, password)
        msg = skC.SKCenterLib_GetReturnCodeMessage(ret)
        messagebox.showinfo("登入結果", f"{ret}: {msg}")

    # 訂閱報價事件
    def subscribe(self):
        stock = self.entry_stock.get()
        if stock:
            skQ.SKQuoteLib_EnterMonitor()
            skQ.SKQuoteLib_RequestStocks(0, stock)
            messagebox.showinfo("訂閱", f"已訂閱 {stock}")

    # 下單事件
    def order(self, buy_sell):
        stock = self.entry_stock.get()
        price = float(self.entry_price.get())
        qty = int(self.entry_qty.get())
        account = self.entry_account.get()
        mock = bool(self.mock_var.get())

        ret, msg = send_order(account, stock, price, qty, buy_sell, mock)
        messagebox.showinfo("下單結果", f"{ret}: {msg}")

# === 6. 執行 GUI + Python COM pump ===
def com_loop():
    while True:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.1)

root = tk.Tk()
app = App(root)

# 用 thread 處理 COM pump，避免 GUI 卡住
threading.Thread(target=com_loop, daemon=True).start()

root.mainloop()
