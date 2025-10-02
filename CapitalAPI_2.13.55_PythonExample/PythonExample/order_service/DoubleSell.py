# DoubleSell
import os
import Global

skC = Global.skC
skO = Global.skO
skR = Global.skR
skQ = Global.skQ
skOSQ = Global.skOSQ
skOOQ = Global.skOOQ

# 添加COM元件支援
import comtypes.client
import comtypes.gen.SKCOMLib as sk

# tkinter
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox

# 事件
import Config
import MessageControl

# 全局變數：用於報價回調與自動賣出功能的通訊
_auto_sell_monitor = None

# 報價回調事件處理
class AutoSellQuoteEvent:
    def __init__(self):
        self.stock_prices = {}  # 儲存商品代碼對應的最新價格

    def OnNotifyQuoteLONG(self, sMarketNo, nStockidx):
        """報價更新回調"""
        try:
            # 獲取報價資料
            pSKStock = sk.SKSTOCKLONG()
            nCode = skQ.SKQuoteLib_GetStockByIndexLONG(sMarketNo, nStockidx, pSKStock)

            if nCode == 0:
                stock_no = pSKStock.bstrStockNo
                # 取得成交價
                current_price = pSKStock.nClose / 100.0  # 價格需除以100

                # 儲存最新價格
                self.stock_prices[stock_no] = current_price

                # 如果有啟動自動賣出監控，檢查價格
                global _auto_sell_monitor
                if _auto_sell_monitor:
                    _auto_sell_monitor.check_price(stock_no, current_price)
        except Exception as e:
            pass

# 建立報價事件實例
quote_event = AutoSellQuoteEvent()

# DoubleSell
class Order(Frame):
    def __init__(self, master=None, information=None):
        Frame.__init__(self)
        self.__master = master
        self.__oMsg = MessageControl.MessageControl()
        # UI variable
        self.__dOrder = dict(
            listInformation = information,
            boxAccount = ''
        )

        self.__CreateWidget()

    # 設定帳號
    def SetAccount(self, account):
        self.__dOrder['boxAccount'] = account

    # 建立元件
    def __CreateWidget(self):
        # 選擇權委託
        group = LabelFrame(self.__master, text="選擇權委託", style="Pink.TLabelframe")
        group.grid(column = 0, row = 0, padx = 5, pady = 5, columnspan = 3, sticky = 'w')

        frame = Frame(group, style="Pink.TFrame")
        frame.grid(column = 0, row = 0, padx = 5, pady = 5, sticky = 'w')

        # 商品代碼
        lbStockNo = Label(frame, style="Pink.TLabel", text = "商品代碼")
        lbStockNo.grid(column = 0, row = 0, pady = 3)
            # 輸入框
        txtStockNo = Entry(frame, width = 15)
        txtStockNo.grid(column = 0, row = 1, padx = 10, pady = 3, sticky = 'w')

        #買賣別
        lbBuySell = Label(frame, style="Pink.TLabel", text = "買賣別")
        lbBuySell.grid(column = 1, row = 0)
            # 輸入框
        boxBuySell = Combobox(frame, width = 5, state='readonly')
        boxBuySell['values'] = Config.BUYSELLSET
        boxBuySell.grid(column = 1, row = 1, padx = 10, sticky = 'w')

        # 委託條件
        lbPeriod = Label(frame, style="Pink.TLabel", text = "委託條件")
        lbPeriod.grid(column = 2, row = 0)
            # 輸入框
        boxPeriod = Combobox(frame, width = 10, state='readonly')
        boxPeriod['values'] = Config.PERIODSET['future']
        boxPeriod.grid(column = 2, row = 1, padx = 10, sticky = 'w')

        # 委託價
        lbPrice = Label(frame, style="Pink.TLabel", text = "委託價")
        lbPrice.grid(column = 3, row = 0)
            # 輸入框
        txtPrice = Entry(frame, width = 15)
        txtPrice.grid(column = 3, row = 1, padx = 10, sticky = 'w')

        # 委託量
        lbQty = Label(frame, style="Pink.TLabel", text = "委託量")
        lbQty.grid(column = 4, row = 0)
            # 輸入框
        txtQty = Entry(frame, width = 10)
        txtQty.grid(column = 4, row = 1, padx = 10, sticky = 'w')

        # 倉別
        lbNewClose = Label(frame, style="Pink.TLabel", text = "倉別")
        lbNewClose.grid(column = 5, row = 0)
            # 輸入框
        boxNewClose = Combobox(frame, width = 5, state='readonly')
        boxNewClose['values'] = Config.NEWCLOSESET['future']
        boxNewClose.grid(column = 5, row = 1, padx = 10, sticky = 'w')

        # 盤別
        lbReserved = Label(frame, style="Pink.TLabel", text = "盤別")
        lbReserved.grid(column = 6, row = 0)
            # 輸入框
        boxReserved = Combobox(frame, width = 10, state='readonly')
        boxReserved['values'] = Config.RESERVEDSET
        boxReserved.grid(column = 6, row = 1, padx = 10, sticky = 'w')

        # btnSendOrder
        btnSendOrder = Button(frame, style = "Pink.TButton", text = "同步委託")
        btnSendOrder["command"] = self.__btnSendOrder_Click
        btnSendOrder.grid(column = 7, row =  1, padx = 10)

        # btnSendOrderAsync - 非同步委託
        btnSendOrderAsync = Button(frame, style = "Pink.TButton", text = "非同步委託")
        btnSendOrderAsync["command"] = self.__btnSendOrderAsync_Click
        btnSendOrderAsync.grid(column = 8, row =  1, padx = 10)

        # ========== 第二組輸入框 ==========
        # 商品代碼2
        lbStockNo2 = Label(frame, style="Pink.TLabel", text = "商品代碼2")
        lbStockNo2.grid(column = 0, row = 3, pady = 3)
        txtStockNo2 = Entry(frame, width = 15)
        txtStockNo2.grid(column = 0, row = 4, padx = 10, pady = 3, sticky = 'w')

        #買賣別2
        lbBuySell2 = Label(frame, style="Pink.TLabel", text = "買賣別2")
        lbBuySell2.grid(column = 1, row = 3)
        boxBuySell2 = Combobox(frame, width = 5, state='readonly')
        boxBuySell2['values'] = Config.BUYSELLSET
        boxBuySell2.grid(column = 1, row = 4, padx = 10, sticky = 'w')

        # 委託條件2
        lbPeriod2 = Label(frame, style="Pink.TLabel", text = "委託條件2")
        lbPeriod2.grid(column = 2, row = 3)
        boxPeriod2 = Combobox(frame, width = 10, state='readonly')
        boxPeriod2['values'] = Config.PERIODSET['future']
        boxPeriod2.grid(column = 2, row = 4, padx = 10, sticky = 'w')

        # 委託價2
        lbPrice2 = Label(frame, style="Pink.TLabel", text = "委託價2")
        lbPrice2.grid(column = 3, row = 3)
        txtPrice2 = Entry(frame, width = 15)
        txtPrice2.grid(column = 3, row = 4, padx = 10, sticky = 'w')

        # 委託量2
        lbQty2 = Label(frame, style="Pink.TLabel", text = "委託量2")
        lbQty2.grid(column = 4, row = 3)
        txtQty2 = Entry(frame, width = 10)
        txtQty2.grid(column = 4, row = 4, padx = 10, sticky = 'w')

        # 倉別2
        lbNewClose2 = Label(frame, style="Pink.TLabel", text = "倉別2")
        lbNewClose2.grid(column = 5, row = 3)
        boxNewClose2 = Combobox(frame, width = 5, state='readonly')
        boxNewClose2['values'] = Config.NEWCLOSESET['future']
        boxNewClose2.grid(column = 5, row = 4, padx = 10, sticky = 'w')

        # 盤別2
        lbReserved2 = Label(frame, style="Pink.TLabel", text = "盤別2")
        lbReserved2.grid(column = 6, row = 3)
        boxReserved2 = Combobox(frame, width = 10, state='readonly')
        boxReserved2['values'] = Config.RESERVEDSET
        boxReserved2.grid(column = 6, row = 4, padx = 10, sticky = 'w')

        # 第二組按鈕
        btnSendOrder2 = Button(frame, style = "Pink.TButton", text = "同步委託2")
        btnSendOrder2["command"] = self.__btnSendOrder2_Click
        btnSendOrder2.grid(column = 7, row = 4, padx = 10)

        btnSendOrderAsync2 = Button(frame, style = "Pink.TButton", text = "非同步委託2")
        btnSendOrderAsync2["command"] = self.__btnSendOrderAsync2_Click
        btnSendOrderAsync2.grid(column = 8, row = 4, padx = 10)

        # 儲存輸入框物件
        self.__dOrder['txtStockNo'] = txtStockNo
        self.__dOrder['boxPeriod'] = boxPeriod
        self.__dOrder['boxBuySell'] = boxBuySell
        self.__dOrder['txtPrice'] = txtPrice
        self.__dOrder['txtQty'] = txtQty
        self.__dOrder['boxNewClose'] = boxNewClose
        self.__dOrder['boxReserved'] = boxReserved

        # 儲存第二組輸入框物件
        self.__dOrder['txtStockNo2'] = txtStockNo2
        self.__dOrder['boxPeriod2'] = boxPeriod2
        self.__dOrder['boxBuySell2'] = boxBuySell2
        self.__dOrder['txtPrice2'] = txtPrice2
        self.__dOrder['txtQty2'] = txtQty2
        self.__dOrder['boxNewClose2'] = boxNewClose2
        self.__dOrder['boxReserved2'] = boxReserved2


    # 4.下單送出
    # sPeriod, sBuySell, sTradeType, sNewClose, sReserved
    def __btnSendOrder_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇期貨帳號！')
        else:
            self.__SendOrder_Click(False)

    def __btnSendOrderAsync_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇期貨帳號！')
        else:
            self.__SendOrder_Click(True)

    def __btnSendOrder2_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇期貨帳號！')
        else:
            self.__SendOrder2_Click(False)

    def __btnSendOrderAsync2_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇期貨帳號！')
        else:
            self.__SendOrder2_Click(True)

    def __SendOrder_Click(self, bAsyncOrder):
        try:
            if self.__dOrder['boxBuySell'].get() == "買進":
                sBuySell = 0
            elif self.__dOrder['boxBuySell'].get() == "賣出":
                sBuySell = 1

            if self.__dOrder['boxPeriod'].get() == "ROD":
                sTradeType = 0
            elif self.__dOrder['boxPeriod'].get() == "IOC":
                sTradeType = 1
            elif self.__dOrder['boxPeriod'].get() == "FOK":
                sTradeType = 2

            if self.__dOrder['boxNewClose'].get() == "新倉":
                sNewClose = 0
            elif self.__dOrder['boxNewClose'].get() == "平倉":
                sNewClose = 1
            elif self.__dOrder['boxNewClose'].get() == "自動":
                sNewClose = 2

            if self.__dOrder['boxReserved'].get() == "盤中":
                sReserved = 0
            elif self.__dOrder['boxReserved'].get() == "T盤預約":
                sReserved = 1

            # 建立下單用的參數(FUTUREORDER)物件(下單時要填商品代號,買賣別,委託價,數量等等的一個物件)
            oOrder = sk.FUTUREORDER()
            # 填入完整帳號
            oOrder.bstrFullAccount =  self.__dOrder['boxAccount']
            # 填入期權代號
            oOrder.bstrStockNo = self.__dOrder['txtStockNo'].get()
            # 買賣別
            oOrder.sBuySell = sBuySell
            # ROD、IOC、FOK
            oOrder.sTradeType = sTradeType
            # 委託價
            oOrder.bstrPrice = self.__dOrder['txtPrice'].get()
            # 委託數量
            oOrder.nQty = int(self.__dOrder['txtQty'].get())
            # 新倉、平倉、自動
            oOrder.sNewClose = sNewClose
            # 盤中、T盤預約
            oOrder.sReserved = sReserved

            message, m_nCode = skO.SendOptionOrder(Global.Global_IID, bAsyncOrder, oOrder)
            self.__oMsg.SendReturnMessage("Order", m_nCode, "SendOptionOrder", self.__dOrder['listInformation'])
            if bAsyncOrder == False and m_nCode == 0:
                strMsg = "選擇權委託: " + str(message)
                self.__oMsg.WriteMessage( strMsg, self.__dOrder['listInformation'])
        except Exception as e:
            messagebox.showerror("error！", e)

    def __SendOrder2_Click(self, bAsyncOrder):
        try:
            if self.__dOrder['boxBuySell2'].get() == "買進":
                sBuySell = 0
            elif self.__dOrder['boxBuySell2'].get() == "賣出":
                sBuySell = 1

            if self.__dOrder['boxPeriod2'].get() == "ROD":
                sTradeType = 0
            elif self.__dOrder['boxPeriod2'].get() == "IOC":
                sTradeType = 1
            elif self.__dOrder['boxPeriod2'].get() == "FOK":
                sTradeType = 2

            if self.__dOrder['boxNewClose2'].get() == "新倉":
                sNewClose = 0
            elif self.__dOrder['boxNewClose2'].get() == "平倉":
                sNewClose = 1
            elif self.__dOrder['boxNewClose2'].get() == "自動":
                sNewClose = 2

            if self.__dOrder['boxReserved2'].get() == "盤中":
                sReserved = 0
            elif self.__dOrder['boxReserved2'].get() == "T盤預約":
                sReserved = 1

            # 建立下單用的參數(FUTUREORDER)物件
            oOrder = sk.FUTUREORDER()
            # 填入完整帳號
            oOrder.bstrFullAccount = self.__dOrder['boxAccount']
            # 填入期權代號
            oOrder.bstrStockNo = self.__dOrder['txtStockNo2'].get()
            # 買賣別
            oOrder.sBuySell = sBuySell
            # ROD、IOC、FOK
            oOrder.sTradeType = sTradeType
            # 委託價
            oOrder.bstrPrice = self.__dOrder['txtPrice2'].get()
            # 委託數量
            oOrder.nQty = int(self.__dOrder['txtQty2'].get())
            # 新倉、平倉、自動
            oOrder.sNewClose = sNewClose
            # 盤中、T盤預約
            oOrder.sReserved = sReserved

            message, m_nCode = skO.SendOptionOrder(Global.Global_IID, bAsyncOrder, oOrder)
            self.__oMsg.SendReturnMessage("Order", m_nCode, "SendOptionOrder (第二組)", self.__dOrder['listInformation'])
            if bAsyncOrder == False and m_nCode == 0:
                strMsg = "選擇權委託 (第二組): " + str(message)
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
        except Exception as e:
            messagebox.showerror("error！", f"第二組送單失敗: {e}")

# 期貨委託
class Future(Frame):
    def __init__(self, master=None, information=None):
        Frame.__init__(self)
        self.__master = master
        self.__oMsg = MessageControl.MessageControl()
        # UI variable
        self.__dOrder = dict(
            listInformation = information,
            boxAccount = ''
        )

        # 自動賣出監控變數
        self.__auto_sell_active = False
        self.__bought_stockNo = None
        self.__auto_sell_price = None
        self.__bought_qty = 0

        # 循環交易變數
        self.__loop_trading = False
        self.__buy_price = None
        self.__trade_qty = 0
        self.__current_state = None  # 'bought' 或 'sold'

        self.__CreateWidget()

    # 設定帳號
    def SetAccount(self, account):
        self.__dOrder['boxAccount'] = account

    # 建立元件
    def __CreateWidget(self):
        group = LabelFrame(self.__master, text="期貨委託", style="Pink.TLabelframe")
        group.grid(column = 0, row = 1, padx = 5, pady = 5, columnspan = 2, sticky = 'w')

        frame = Frame(group, style="Pink.TFrame")
        frame.grid(column = 0, row = 0, padx = 5, pady = 5, sticky = 'ew')

        # 商品代碼
        Label(frame, style="Pink.TLabel", text = "商品代碼").grid(column = 0, row = 0, pady = 3)
        txtStockNo = Entry(frame, width = 5)
        txtStockNo.grid(column = 0, row = 1, padx = 5, pady = 3)

        # 買賣別
        Label(frame, style="Pink.TLabel", text = "買賣別").grid(column = 1, row = 0)
        boxBuySell = Combobox(frame, width = 5, state='readonly')
        boxBuySell['values'] = Config.BUYSELLSET
        boxBuySell.grid(column = 1, row = 1, padx = 5)

        # 委託條件
        Label(frame, style="Pink.TLabel", text = "委託條件").grid(column = 2, row = 0)
        boxPeriod = Combobox(frame, width = 10, state='readonly')
        boxPeriod['values'] = Config.PERIODSET['future']
        boxPeriod.grid(column = 2, row = 1, padx = 5)

        # 當沖與否
        Label(frame, style="Pink.TLabel", text = "當沖與否").grid(column = 3, row = 0)
        boxFlag = Combobox(frame, width = 10, state='readonly')
        boxFlag['values'] = Config.FLAGSET['future']
        boxFlag.grid(column = 3, row = 1, padx = 5)

        # 委託價
        Label(frame, style="Pink.TLabel", text = "委託價").grid(column = 4, row = 0)
        txtPrice = Entry(frame, width = 12)
        txtPrice.grid(column = 4, row = 1, padx = 5)

        # 委託量
        Label(frame, style="Pink.TLabel", text = "委託量").grid(column = 5, row = 0)
        txtQty = Entry(frame, width = 10)
        txtQty.grid(column = 5, row = 1, padx = 5)

        # 倉別
        Label(frame, style="Pink.TLabel", text = "倉別").grid(column = 6, row = 0)
        boxNewClose = Combobox(frame, width = 5, state='readonly')
        boxNewClose['values'] = Config.NEWCLOSESET['future']
        boxNewClose.grid(column = 6, row = 1, padx = 5)

        # 盤別
        Label(frame, style="Pink.TLabel", text = "盤別").grid(column = 7, row = 0)
        boxReserved = Combobox(frame, width = 10, state='readonly')
        boxReserved['values'] = Config.RESERVEDSET
        boxReserved.grid(column = 7, row = 1, padx = 5)

        # 自動賣出價格
        Label(frame, style="Pink.TLabel", text = "自動賣出價格").grid(column = 8, row = 0, pady = 3)
        txtAutoSellPrice = Entry(frame, width = 12)
        txtAutoSellPrice.grid(column = 8, row = 1, padx = 5)

        # 循環交易選項
        self.__loop_var = IntVar(value=0)
        chkLoopTrading = Checkbutton(frame, style="Pink.TCheckbutton", text='循環交易', variable=self.__loop_var, onvalue=1, offvalue=0)
        chkLoopTrading.grid(column = 8, row = 2, padx = 5)

        # 停止監控按鈕
        btnStopMonitor = Button(frame, style = "Pink.TButton", text = "停止交易")
        btnStopMonitor["command"] = self.__btnStopMonitor_Click
        btnStopMonitor.grid(column = 8, row = 3, padx = 5, pady = 3)

        # btnSendOrder
        btnSendOrder = Button(frame, style = "Pink.TButton", text = "同步委託")
        btnSendOrder["command"] = self.__btnSendOrder_Click
        btnSendOrder.grid(column = 9, row = 0, padx = 5)

        # btnAsyncSendOrder
        btnAsyncSendOrder = Button(frame, style = "Pink.TButton", text = "非同步委託")
        btnAsyncSendOrder["command"] = self.__btnAsyncSendOrder_Click
        btnAsyncSendOrder.grid(column = 9, row = 1, padx = 5)

        # btnSendOrderCLR
        btnSendOrderCLR = Button(frame, style = "Pink.TButton", text = "同步委託(含倉別/盤別)")
        btnSendOrderCLR["command"] = self.__btnSendOrderCLR_Click
        btnSendOrderCLR.grid(column = 10, row = 0, padx = 5)

        # btnAsyncSendOrderCLR
        btnAsyncSendOrderCLR = Button(frame, style = "Pink.TButton", text = "非同步委託(含倉別/盤別)")
        btnAsyncSendOrderCLR["command"] = self.__btnAsyncSendOrderCLR_Click
        btnAsyncSendOrderCLR.grid(column = 10, row = 1, padx = 5)

        self.__dOrder['txtStockNo'] = txtStockNo
        self.__dOrder['boxPeriod'] = boxPeriod
        self.__dOrder['boxFlag'] = boxFlag
        self.__dOrder['boxBuySell'] = boxBuySell
        self.__dOrder['txtPrice'] = txtPrice
        self.__dOrder['txtQty'] = txtQty
        self.__dOrder['boxNewClose'] = boxNewClose
        self.__dOrder['boxReserved'] = boxReserved
        self.__dOrder['txtAutoSellPrice'] = txtAutoSellPrice

    # 停止監控按鈕
    def __btnStopMonitor_Click(self):
        if self.__auto_sell_active or self.__loop_trading:
            self.__auto_sell_active = False
            self.__loop_trading = False
            self.__stop_monitoring()
            messagebox.showinfo("提示", "已停止所有自動交易")
        else:
            messagebox.showinfo("提示", "目前沒有啟動交易")

    # 4.下單送出
    # sBuySell, sTradeType, sDayTrade, sNewClose, sReserved
    def __btnSendOrder_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇期貨帳號！')
        else:
            self.__SendOrder_Click(False)

    def __btnAsyncSendOrder_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇期貨帳號！')
        else:
            self.__SendOrder_Click(True)

    # 送出期貨委託
    def __SendOrder_Click(self, bAsyncOrder):
        try:
            if self.__dOrder['boxBuySell'].get() == "買進":
                sBuySell = 0
            elif self.__dOrder['boxBuySell'].get() == "賣出":
                sBuySell = 1

            if self.__dOrder['boxPeriod'].get() == "ROD":
                sTradeType = 0
            elif self.__dOrder['boxPeriod'].get() == "IOC":
                sTradeType = 1
            elif self.__dOrder['boxPeriod'].get() == "FOK":
                sTradeType = 2

            if self.__dOrder['boxFlag'].get() == "非當沖":
                sDayTrade = 0
            elif self.__dOrder['boxFlag'].get() == "當沖":
                sDayTrade = 1

            # 建立下單用的參數(FUTUREORDER)物件(下單時要填商品代號,買賣別,委託價,數量等等的一個物件)
            oOrder = sk.FUTUREORDER()
            # 填入完整帳號
            oOrder.bstrFullAccount =  self.__dOrder['boxAccount']
            # 填入期權代號
            oOrder.bstrStockNo = self.__dOrder['txtStockNo'].get()
            # 買賣別
            oOrder.sBuySell = sBuySell
            # ROD、IOC、FOK
            oOrder.sTradeType = sTradeType
            # 非當沖、當沖
            oOrder.sDayTrade = sDayTrade
            # 委託價
            oOrder.bstrPrice = self.__dOrder['txtPrice'].get()
            # 委託數量
            oOrder.nQty = int(self.__dOrder['txtQty'].get())
            # 新倉、平倉、自動
            message, m_nCode = skO.SendFutureOrder(Global.Global_IID, bAsyncOrder, oOrder)
            self.__oMsg.SendReturnMessage("Order", m_nCode, "SendFutureOrder", self.__dOrder['listInformation'])
            if bAsyncOrder == False and m_nCode == 0:
                strMsg = "期貨委託: " + str(message)
                self.__oMsg.WriteMessage( strMsg, self.__dOrder['listInformation'])

                # 如果是買進且設定了自動賣出價格，啟動監控
                if sBuySell == 0 and self.__dOrder['txtAutoSellPrice'].get().strip():
                    self.__bought_stockNo = self.__dOrder['txtStockNo'].get()
                    self.__auto_sell_price = float(self.__dOrder['txtAutoSellPrice'].get())
                    self.__bought_qty = int(self.__dOrder['txtQty'].get())
                    self.__buy_price = float(self.__dOrder['txtPrice'].get())
                    self.__trade_qty = self.__bought_qty
                    self.__auto_sell_active = True
                    self.__current_state = 'bought'

                    # 檢查是否啟用循環交易
                    if self.__loop_var.get() == 1:
                        self.__loop_trading = True
                        strMsg = f"已啟動循環交易模式"
                        self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])

                    self.__start_price_monitoring()
                    strMsg = f"已啟動自動賣出監控: {self.__bought_stockNo}, 買進價: {self.__buy_price}, 目標價格: {self.__auto_sell_price}"
                    self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
        except Exception as e:
            messagebox.showerror("error！", e)

    def __btnSendOrderCLR_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇期貨帳號！')
        else:
            self.__SendOrderCLR_Click(False)

    def __btnAsyncSendOrderCLR_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇期貨帳號！')
        else:
            self.__SendOrderCLR_Click(True)

    def __SendOrderCLR_Click(self, bAsyncOrder):
        try:
            if self.__dOrder['boxBuySell'].get() == "買進":
                sBuySell = 0
            elif self.__dOrder['boxBuySell'].get() == "賣出":
                sBuySell = 1

            if self.__dOrder['boxPeriod'].get() == "ROD":
                sTradeType = 0
            elif self.__dOrder['boxPeriod'].get() == "IOC":
                sTradeType = 1
            elif self.__dOrder['boxPeriod'].get() == "FOK":
                sTradeType = 2

            if self.__dOrder['boxFlag'].get() == "非當沖":
                sDayTrade = 0
            elif self.__dOrder['boxFlag'].get() == "當沖":
                sDayTrade = 1

            if self.__dOrder['boxNewClose'].get() == "新倉":
                sNewClose = 0
            elif self.__dOrder['boxNewClose'].get() == "平倉":
                sNewClose = 1
            elif self.__dOrder['boxNewClose'].get() == "自動":
                sNewClose = 2

            if self.__dOrder['boxReserved'].get() == "盤中":
                sReserved = 0
            elif self.__dOrder['boxReserved'].get() == "T盤預約":
                sReserved = 1

            # 建立下單用的參數(FUTUREORDER)物件(下單時要填商品代號,買賣別,委託價,數量等等的一個物件)
            oOrder = sk.FUTUREORDER()
            # 填入完整帳號
            oOrder.bstrFullAccount =  self.__dOrder['boxAccount']
            # 填入期權代號
            oOrder.bstrStockNo = self.__dOrder['txtStockNo'].get()
            # 買賣別
            oOrder.sBuySell = sBuySell
            # ROD、IOC、FOK
            oOrder.sTradeType = sTradeType
            # 非當沖、當沖
            oOrder.sDayTrade = sDayTrade
            # 委託價
            oOrder.bstrPrice = self.__dOrder['txtPrice'].get()
            # 委託數量
            oOrder.nQty = int(self.__dOrder['txtQty'].get())
            # 新倉、平倉、自動
            oOrder.sNewClose = sNewClose
            # 盤中、T盤預約
            oOrder.sReserved = sReserved
            # 送出期貨委託(含倉別/盤別)
            message, m_nCode = skO.SendFutureOrderCLR(Global.Global_IID, bAsyncOrder, oOrder)
            self.__oMsg.SendReturnMessage("Order", m_nCode, "SendFutureOrderCLR", self.__dOrder['listInformation'])
            if bAsyncOrder == False and m_nCode == 0:
                strMsg = "期貨委託帶倉別/盤別: " + str(message)
                self.__oMsg.WriteMessage( strMsg, self.__dOrder['listInformation'])

        except Exception as e:
            messagebox.showerror("error！", e)

    # 啟動價格監控
    def __start_price_monitoring(self):
        global _auto_sell_monitor
        _auto_sell_monitor = self

        try:
            # 註冊報價回調（如果尚未註冊）
            if not hasattr(skQ, '_quote_event_registered'):
                comtypes.client.GetEvents(skQ, quote_event)
                skQ._quote_event_registered = True

            # 請求訂閱報價
            nCode = skQ.SKQuoteLib_RequestStocks(
                Global.Global_IID,
                2,  # 期貨
                self.__bought_stockNo
            )

            if nCode == 0:
                strMsg = f"已訂閱 {self.__bought_stockNo} 報價，開始監控"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
            else:
                strMsg = f"訂閱報價失敗: {skC.SKCenterLib_GetReturnCodeMessage(nCode)}"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
                self.__auto_sell_active = False
        except Exception as e:
            strMsg = f"啟動價格監控錯誤: {e}"
            self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
            self.__auto_sell_active = False

    # 檢查價格並觸發自動交易
    def check_price(self, stock_no, current_price):
        """由報價回調呼叫，檢查是否達到交易條件"""
        if not self.__auto_sell_active and not self.__loop_trading:
            return

        if stock_no != self.__bought_stockNo:
            return

        # 如果目前是持倉狀態（已買進），檢查是否達到賣出價格
        if self.__current_state == 'bought':
            if current_price >= self.__auto_sell_price:
                strMsg = f"價格達到賣出目標! 當前價: {current_price}, 目標價: {self.__auto_sell_price}"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
                # 執行自動賣出
                self.__execute_auto_sell()

        # 如果是循環模式且已賣出，檢查是否達到買進價格
        elif self.__current_state == 'sold' and self.__loop_trading:
            if current_price <= self.__buy_price:
                strMsg = f"價格回落至買進價! 當前價: {current_price}, 買進價: {self.__buy_price}"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
                # 執行自動買進
                self.__execute_auto_buy()

    # 執行自動賣出
    def __execute_auto_sell(self):
        try:
            # 暫停賣出監控（避免重複觸發）
            temp_active = self.__auto_sell_active
            self.__auto_sell_active = False

            # 建立賣出委託
            oOrder = sk.FUTUREORDER()
            oOrder.bstrFullAccount = self.__dOrder['boxAccount']
            oOrder.bstrStockNo = self.__bought_stockNo
            oOrder.sBuySell = 1  # 賣出
            oOrder.sTradeType = 0  # ROD
            oOrder.sDayTrade = 0  # 非當沖
            oOrder.bstrPrice = str(self.__auto_sell_price)
            oOrder.nQty = self.__bought_qty

            message, m_nCode = skO.SendFutureOrder(Global.Global_IID, False, oOrder)

            if m_nCode == 0:
                strMsg = f"自動賣出成功: {self.__bought_stockNo}, 價格: {self.__auto_sell_price}, 數量: {self.__bought_qty}"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])

                # 如果是循環模式，切換到已賣出狀態，繼續監控買進時機
                if self.__loop_trading:
                    self.__current_state = 'sold'
                    strMsg = f"循環交易: 等待價格回落至 {self.__buy_price} 再買進"
                    self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
                else:
                    # 非循環模式，停止監控
                    self.__stop_monitoring()
            else:
                strMsg = f"自動賣出失敗: {message}"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
                # 恢復監控狀態
                self.__auto_sell_active = temp_active
        except Exception as e:
            messagebox.showerror("error！", f"自動賣出錯誤: {e}")
            self.__auto_sell_active = False

    # 執行自動買進（循環交易用）
    def __execute_auto_buy(self):
        try:
            # 建立買進委託
            oOrder = sk.FUTUREORDER()
            oOrder.bstrFullAccount = self.__dOrder['boxAccount']
            oOrder.bstrStockNo = self.__bought_stockNo
            oOrder.sBuySell = 0  # 買進
            oOrder.sTradeType = 0  # ROD
            oOrder.sDayTrade = 0  # 非當沖
            oOrder.bstrPrice = str(self.__buy_price)
            oOrder.nQty = self.__trade_qty

            message, m_nCode = skO.SendFutureOrder(Global.Global_IID, False, oOrder)

            if m_nCode == 0:
                strMsg = f"自動買進成功: {self.__bought_stockNo}, 價格: {self.__buy_price}, 數量: {self.__trade_qty}"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])

                # 切換回已買進狀態，等待賣出時機
                self.__current_state = 'bought'
                self.__auto_sell_active = True
                self.__bought_qty = self.__trade_qty
                strMsg = f"循環交易: 等待價格上漲至 {self.__auto_sell_price} 再賣出"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
            else:
                strMsg = f"自動買進失敗: {message}"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
        except Exception as e:
            messagebox.showerror("error！", f"自動買進錯誤: {e}")

    # 停止監控
    def __stop_monitoring(self):
        """停止價格監控並取消訂閱"""
        try:
            if self.__bought_stockNo:
                # 取消訂閱報價
                nCode = skQ.SKQuoteLib_RequestStocks(
                    Global.Global_IID,
                    -2,  # 負數表示取消訂閱期貨
                    self.__bought_stockNo
                )
                strMsg = f"已取消 {self.__bought_stockNo} 報價訂閱"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])

            # 重置所有狀態
            self.__loop_trading = False
            self.__current_state = None

            global _auto_sell_monitor
            if _auto_sell_monitor == self:
                _auto_sell_monitor = None
        except Exception as e:
            pass

# DoubleSell下單總介面
class DoubleSell(Frame):

    def __init__(self, information=None):
        Frame.__init__(self)
        self.__obj = dict(
            order = Order(master = self, information = information),
            future = Future(master = self, information = information)
        )

    def SetAccount(self, account):
        for _ in 'order', 'future':
            self.__obj[_].SetAccount(account)   

