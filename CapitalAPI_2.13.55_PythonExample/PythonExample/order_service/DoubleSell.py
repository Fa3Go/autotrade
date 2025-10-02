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

    def SetAccount(self, account):
        self.__dOrder['boxAccount'] = account

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

        self.__CreateWidget()

    def SetAccount(self, account):
        self.__dOrder['boxAccount'] = account

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

        # btnSendOrder
        btnSendOrder = Button(frame, style = "Pink.TButton", text = "同步委託")
        btnSendOrder["command"] = self.__btnSendOrder_Click
        btnSendOrder.grid(column = 8, row = 0, padx = 5)

        # btnAsyncSendOrder
        btnAsyncSendOrder = Button(frame, style = "Pink.TButton", text = "非同步委託")
        btnAsyncSendOrder["command"] = self.__btnAsyncSendOrder_Click
        btnAsyncSendOrder.grid(column = 8, row = 1, padx = 5)

        # btnSendOrderCLR
        btnSendOrderCLR = Button(frame, style = "Pink.TButton", text = "同步委託(含倉別/盤別)")
        btnSendOrderCLR["command"] = self.__btnSendOrderCLR_Click
        btnSendOrderCLR.grid(column = 9, row = 0, padx = 5)

        # btnAsyncSendOrderCLR
        btnAsyncSendOrderCLR = Button(frame, style = "Pink.TButton", text = "非同步委託(含倉別/盤別)")
        btnAsyncSendOrderCLR["command"] = self.__btnAsyncSendOrderCLR_Click
        btnAsyncSendOrderCLR.grid(column = 9, row = 1, padx = 5)

        self.__dOrder['txtStockNo'] = txtStockNo
        self.__dOrder['boxPeriod'] = boxPeriod
        self.__dOrder['boxFlag'] = boxFlag
        self.__dOrder['boxBuySell'] = boxBuySell
        self.__dOrder['txtPrice'] = txtPrice
        self.__dOrder['txtQty'] = txtQty
        self.__dOrder['boxNewClose'] = boxNewClose
        self.__dOrder['boxReserved'] = boxReserved

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

            message, m_nCode = skO.SendFutureOrder(Global.Global_IID, bAsyncOrder, oOrder)
            self.__oMsg.SendReturnMessage("Order", m_nCode, "SendFutureOrder", self.__dOrder['listInformation'])
            if bAsyncOrder == False and m_nCode == 0:
                strMsg = "期貨委託: " + str(message)
                self.__oMsg.WriteMessage( strMsg, self.__dOrder['listInformation'])
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

            message, m_nCode = skO.SendFutureOrderCLR(Global.Global_IID, bAsyncOrder, oOrder)
            self.__oMsg.SendReturnMessage("Order", m_nCode, "SendFutureOrderCLR", self.__dOrder['listInformation'])
            if bAsyncOrder == False and m_nCode == 0:
                strMsg = "期貨委託帶倉別/盤別: " + str(message)
                self.__oMsg.WriteMessage( strMsg, self.__dOrder['listInformation'])

        except Exception as e:
            messagebox.showerror("error！", e)

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

