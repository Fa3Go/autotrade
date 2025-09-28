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
class DuplexOrder(Frame):
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
        group = LabelFrame(self.__master, text="複式單委託", style="Pink.TLabelframe")
        group.grid(column = 0, row = 1, padx = 5, pady = 5, columnspan = 2,sticky = "w")

        frame = Frame(group, style="Pink.TFrame")
        frame.grid(column = 0, row = 0, padx = 5, pady = 5, sticky = 'w')
        # 商品代碼1
        Merchandise1 = Label(frame, style="Pink.TLabel", text = "商品代碼1")
        Merchandise1.grid(column = 0, row = 0, pady = 3)
            # 輸入框
        self.txtMerchandise1 = Entry(frame, width = 15)
        self.txtMerchandise1.grid(column = 0, row = 1, padx = 5)
        # 商品代碼2
        Merchandise2 = Label(frame, style="Pink.TLabel", text = "商品代碼2")
        Merchandise2.grid(column = 0, row = 2, pady = 3)
            # 輸入框
        self.txtMerchandise2 = Entry(frame, width = 15)
        self.txtMerchandise2.grid(column = 0, row = 3, padx = 5)

        # 買賣別1
        BuyOrSell1 = Label(frame, style="Pink.TLabel", text = "買賣別1")
        BuyOrSell1.grid(column = 1, row = 0)
            # 輸入框
        self.boxBuyOrSell1 = Combobox(frame, width = 10, state='readonly')
        self.boxBuyOrSell1['values'] = Config.BUYSELLSET
        self.boxBuyOrSell1.grid(column = 1, row = 1, padx = 5)
        # 買賣別2
        BuyOrSell2 = Label(frame, style="Pink.TLabel", text = "買賣別2")
        BuyOrSell2.grid(column = 1, row = 2)
            # 輸入框
        self.boxBuyOrSell2 = Combobox(frame, width = 10, state='readonly')
        self.boxBuyOrSell2['values'] = Config.BUYSELLSET
        self.boxBuyOrSell2.grid(column = 1, row = 3, padx = 5)

        # 委託條件
        TradeType = Label(frame, style="Pink.TLabel", text = "委託條件")
        TradeType.grid(column = 2, row = 1)
            # 輸入框
        self.TradeType = Combobox(frame, width = 10, state='readonly')
        self.TradeType['values'] = Config.PERIODSET["moving_stop_loss"]
        self.TradeType.grid(column = 2, row = 2, padx = 5)

        # 倉別
        NewOrClose = Label(frame, style="Pink.TLabel", text = "倉別")
        NewOrClose.grid(column = 3, row = 1)
            # 輸入框
        self.boxNewOrClose = Combobox(frame, width = 10, state='readonly')
        self.boxNewOrClose['values'] = Config.NEWCLOSESET["option_future"]
        self.boxNewOrClose.grid(column = 3, row = 2, padx = 5)

        # 委託價
        lbPrice = Label(frame, style="Pink.TLabel", text = "委託價")
        lbPrice.grid(column = 4, row = 1)
            # 輸入框
        self.txtPrice = Entry(frame, width = 10)
        self.txtPrice.grid(column = 4, row = 2, padx = 5)

        # 委託量
        lbQty = Label(frame, style="Pink.TLabel", text = "委託量")
        lbQty.grid(column = 5, row = 1)
            # 輸入框
        self.txtQty = Entry(frame, width = 10)
        self.txtQty.grid(column = 5, row = 2, padx = 5)

        # btnSendOrder
        btnSendOrder = Button(frame, style = "Pink.TButton", text = "同步委託")
        btnSendOrder["command"] = self.__btnSendOrder_Click
        btnSendOrder.grid(column = 6, row =  1, padx = 5)
        # btnSendOrderAsync
        btnSendOrderAsync = Button(frame, style = "Pink.TButton", text = "非同步委託")
        btnSendOrderAsync["command"] = self.__btnSendOrderAsync_Click
        btnSendOrderAsync.grid(column = 6, row =  2, padx = 5)

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

    def __SendOrder_Click(self, bAsyncOrder):
        try:
            if self.boxBuyOrSell1.get() == "買進":
                sBuySell = 0
            elif self.boxBuyOrSell1.get() == "賣出":
                sBuySell = 1

            if self.boxBuyOrSell2.get() == "買進":
                sBuySell2 = 0
            elif self.boxBuyOrSell2.get() == "賣出":
                sBuySell2 = 1

            if self.TradeType.get() == "IOC":
                sTradeType = 1
            elif self.TradeType.get() == "FOK":
                sTradeType = 2

            if self.boxNewOrClose.get() == "新倉":
                sNewClose = 0
            elif self.boxNewOrClose.get() == "平倉":
                sNewClose = 1

            # 建立下單用的參數(FUTUREORDER)物件(下單時要填商品代號,買賣別,委託價,數量等等的一個物件)
            oOrder = sk.FUTUREORDER()
            # 填入完整帳號
            oOrder.bstrFullAccount =  self.__dOrder['boxAccount']
            # 填入期權代號1
            oOrder.bstrStockNo = self.txtMerchandise1.get()
            # 填入期權代號2
            oOrder.bstrStockNo2 = self.txtMerchandise2.get()
            # 買賣別1
            oOrder.sBuySell = sBuySell
            # 買賣別2
            oOrder.sBuySell2 = sBuySell2
            # IOC、FOK
            oOrder.sTradeType = sTradeType
            # 委託價
            oOrder.bstrPrice = self.txtPrice.get()
            # 委託數量
            oOrder.nQty = int(self.txtQty.get())
            # 新倉、平倉
            oOrder.sNewClose = sNewClose

            message, m_nCode = skO.SendDuplexOrder(Global.Global_IID, bAsyncOrder, oOrder)
            self.__oMsg.SendReturnMessage("Order", m_nCode, "SendDuplexOrder", self.__dOrder['listInformation'])
            if bAsyncOrder == False and m_nCode == 0:
                strMsg = "複式單委託: " + str(message)
                self.__oMsg.WriteMessage( strMsg, self.__dOrder['listInformation'])

        except Exception as e:
            messagebox.showerror("error！", e)

# DoubleSell下單總介面
class DoubleSell(Frame):

    def __init__(self, information=None):
        Frame.__init__(self)
        self.__obj = dict(
            duplexorder = DuplexOrder(master = self, information = information),
        )

    def SetAccount(self, account):
        for _ in 'duplexorder':
            self.__obj[_].SetAccount(account)   

