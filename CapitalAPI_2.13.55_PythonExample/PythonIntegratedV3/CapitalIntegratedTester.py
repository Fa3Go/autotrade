# 群益證券 API 整合測試工具 v2.13.55
# Capital Securities Integrated Tester
# 整合 Login, Order, Quote, Reply 所有功能到單一介面

# API COM元件初始化
import comtypes.client
comtypes.client.GetModule(r'SKCOM.dll')
import comtypes.gen.SKCOMLib as sk
import ctypes
import os
import time
import math

# GUI元件
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
from tkinter import filedialog

# 群益API元件導入Python code內用的物件宣告
m_pSKCenter = comtypes.client.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
m_pSKReply = comtypes.client.CreateObject(sk.SKReplyLib, interface=sk.ISKReplyLib)
m_pSKOrder = comtypes.client.CreateObject(sk.SKOrderLib, interface=sk.ISKOrderLib)
m_pSKQuote = comtypes.client.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib)
m_pSKOSQuote = comtypes.client.CreateObject(sk.SKOSQuoteLib, interface=sk.ISKOSQuoteLib)
m_pSKOOQuote = comtypes.client.CreateObject(sk.SKOOQuoteLib, interface=sk.ISKOOQuoteLib)

# 全域變數
dictUserID = {}
dictUserID["更新帳號"] = ["無"]
bAsyncOrder = False
isDuplexOrder = False

# 設定選項
class Config:
    # 連線環境
    comboBoxSKCenterLib_SetAuthority = (
        "正式環境", "正式環境SGX", "測試環境", "測試環境SGX"
    )
    
    # 訂單報告
    comboBoxGetOrderReport = (
        "1:全部", "2:有效", "3:可消", "4:已消", "5:已成", "6:失敗", "7:合併同價格", "8:合併同商品", "9:預約"
    )
    
    # 成交回報
    comboBoxGetFulfillReport = (
        "1:全部", "2:合併同書號", "3:合併同價格", "4:合併同商品", "5:T+1成交回報"
    )
    
    # 市場類型
    comboBoxMarketType = (
        "0：TS(證券)", "1：TF(期貨)", "2：TO(選擇權)", "3：OS(複委託)", "4：OF(海外期貨)", "5：OO(海外選擇權)"
    )
    
    # 股票交易相關
    comboBoxPrime = ("上市上櫃", "興櫃")
    comboBoxPeriod = ("盤中", "盤後", "零股")
    comboBoxFlag = ("現股", "融資", "融券", "無券")
    comboBoxTradeType = ("ROD", "IOC", "FOK")
    comboBoxSpecialTradeType = ("市價", "限價")
    comboBoxBuySell = ("買進", "賣出")
    
    # 期貨交易相關
    comboBoxFuturesDayTrade = ("是", "否")
    comboBoxFuturesNewClose = ("新倉", "平倉", "自動")
    comboBoxFuturesBuySell = ("買進", "賣出")
    
    # 報價相關
    comboBoxSKQuoteLib_RequestStocks = ("上市", "上櫃", "興櫃")
    comboBoxSKQuoteLib_RequestTicks = ("股票", "期貨", "選擇權")
    comboBoxSKQuoteLib_RequestKLine = ("分K", "日K", "週K", "月K")

######################################################################################################################################
# 事件處理器

# ReplyLib事件
class SKReplyLibEvent:
    def OnReplyMessage(self, bstrUserID, bstrMessages):
        nConfirmCode = -1
        msg = f"【註冊公告OnReplyMessage】{bstrUserID}_{bstrMessages}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')
        return nConfirmCode

SKReplyEvent = SKReplyLibEvent()
SKReplyLibEventHandler = comtypes.client.GetEvents(m_pSKReply, SKReplyEvent)

# CenterLib事件
class SKCenterLibEvent:
    def OnTimer(self, nTime):
        msg = f"【OnTimer】{nTime}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')
        
    def OnShowAgreement(self, bstrData):
        msg = f"【OnShowAgreement】{bstrData}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')
        
    def OnNotifySGXAPIOrderStatus(self, nStatus, bstrOFAccount):
        msg = f"【OnNotifySGXAPIOrderStatus】{nStatus}_{bstrOFAccount}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')

SKCenterEvent = SKCenterLibEvent()
SKCenterEventHandler = comtypes.client.GetEvents(m_pSKCenter, SKCenterEvent)

# OrderLib事件
class SKOrderLibEvent:
    def OnAccount(self, bstrLogInID, bstrAccountData):
        msg = f"【OnAccount】{bstrLogInID}_{bstrAccountData}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')
        
        values = bstrAccountData.split(',')
        Account = values[1] + values[3]  # broker ID (IB)4碼 + 帳號7碼
        
        if bstrLogInID in dictUserID:
            if Account not in dictUserID[bstrLogInID]:
                dictUserID[bstrLogInID].append(Account)
        else:
            dictUserID[bstrLogInID] = [Account]
    
    def OnProxyStatus(self, bstrUserId, nCode):
        msg = f"【OnProxyStatus】{bstrUserId}_{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')
    
    def OnOpenInterest(self, bstrData):
        msg = f"【OnOpenInterest】{bstrData}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')
    
    def OnFutureRights(self, bstrData):
        msg = f"【OnFutureRights】{bstrData}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')
    
    def OnAsyncOrder(self, nThreadID, nCode, bstrMessage):
        msg = f"【OnAsyncOrder】{nThreadID}{nCode}{bstrMessage}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')
    
    def OnProxyOrder(self, nStampID, nCode, bstrMessage):
        msg = f"【OnProxyOrder】{nStampID}{nCode}{bstrMessage}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')

SKOrderEvent = SKOrderLibEvent()
SKOrderLibEventHandler = comtypes.client.GetEvents(m_pSKOrder, SKOrderEvent)

# QuoteLib事件
class SKQuoteLibEvent:
    def OnConnection(self, nKind, nCode):
        msg = f"【OnConnection】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nKind)}_{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')
    
    def OnNotifyServerTime(self, sHour, sMinute, sSecond, nTotal):
        msg = f"【OnNotifyServerTime】{sHour}:{sMinute}:{sSecond}總秒數:{nTotal}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')
    
    def OnNotifyQuote(self, sMarketNo, bstrStockidx):
        msg = f"【OnNotifyQuote】市場:{sMarketNo} 股票:{bstrStockidx}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')
        
    def OnNotifyMarketTot(self, sMarketNo, sPtr, nTime, nTotv, nTots, nTotc):
        msg = f"【OnNotifyMarketTot】市場:{sMarketNo} 指數:{sPtr} 成交量:{nTotv/100:.2f}億 筆數:{nTots} 家數:{nTotc}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')

SKQuoteEvent = SKQuoteLibEvent()
SKQuoteLibEventHandler = comtypes.client.GetEvents(m_pSKQuote, SKQuoteEvent)

######################################################################################################################################
# GUI

class CapitalIntegratedTester(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("群益證券API整合測試工具 v2.13.55")
        self.geometry("1400x900")
        self.configure(bg='lightgray')
        self.setup_widgets()
        
    def setup_widgets(self):
        # 建立Notebook頁籤
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 登入頁面
        self.login_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.login_frame, text='🔐 登入管理')
        self.setup_login_tab()
        
        # 股票下單頁面
        self.stock_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.stock_frame, text='📈 股票交易')
        self.setup_stock_tab()
        
        # 期貨下單頁面
        self.futures_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.futures_frame, text='📊 期貨交易')
        self.setup_futures_tab()
        
        # 選擇權頁面
        self.options_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.options_frame, text='🎯 選擇權交易')
        self.setup_options_tab()
        
        # 報價頁面
        self.quote_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.quote_frame, text='💹 市場報價')
        self.setup_quote_tab()
        
        # 訊息顯示區域
        self.setup_message_area()
    
    def setup_login_tab(self):
        # 主要設定框架
        main_frame = tk.LabelFrame(self.login_frame, text="連線設定", font=('Arial', 10, 'bold'), bg='lightblue')
        main_frame.pack(fill='x', padx=5, pady=5)
        
        # 連線環境選擇
        tk.Label(main_frame, text="連線環境:", bg='lightblue').grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.combo_authority = ttk.Combobox(main_frame, values=Config.comboBoxSKCenterLib_SetAuthority, state='readonly', width=15)
        self.combo_authority.grid(row=0, column=1, padx=5, pady=5)
        self.combo_authority.bind("<<ComboboxSelected>>", self.on_authority_changed)
        
        # 帳號密碼
        tk.Label(main_frame, text="UserID:", bg='lightblue').grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.entry_userid = tk.Entry(main_frame, width=20)
        self.entry_userid.grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(main_frame, text="Password:", bg='lightblue').grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.entry_password = tk.Entry(main_frame, width=20, show='*')
        self.entry_password.grid(row=2, column=1, padx=5, pady=5)
        
        # AP身分認證
        self.var_is_ap = tk.IntVar()
        self.check_is_ap = tk.Checkbutton(main_frame, text="AP/APH身分", variable=self.var_is_ap, 
                                         command=self.on_ap_checked, bg='lightblue')
        self.check_is_ap.grid(row=3, column=0, sticky='w', padx=5, pady=5)
        
        self.label_cert_id = tk.Label(main_frame, text="CustCertID:", bg='lightblue')
        self.entry_cert_id = tk.Entry(main_frame, width=20)
        
        # 登入按鈕
        self.btn_login = tk.Button(main_frame, text="🔑 登入", command=self.login, 
                                  bg='lightgreen', font=('Arial', 10, 'bold'), width=12)
        self.btn_login.grid(row=4, column=1, padx=5, pady=10)
        
        # 雙因子驗證按鈕
        self.btn_generate_cert = tk.Button(main_frame, text="🔐 雙因子驗證KEY", command=self.generate_cert, width=15)
        
        # 版本資訊
        version_info = m_pSKCenter.SKCenterLib_GetSKAPIVersionAndBit("")
        tk.Label(main_frame, text=f"API版本: {version_info}", bg='lightblue', 
                font=('Arial', 8)).grid(row=0, column=2, sticky='w', padx=20, pady=5)
        
        # 登入後管理框架
        login_mgmt_frame = tk.LabelFrame(self.login_frame, text="登入後管理", font=('Arial', 10, 'bold'), bg='lightyellow')
        login_mgmt_frame.pack(fill='x', padx=5, pady=5)
        
        tk.Label(login_mgmt_frame, text="使用者ID:", bg='lightyellow').grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.combo_userid = ttk.Combobox(login_mgmt_frame, values=list(dictUserID.keys()), state='readonly', width=15)
        self.combo_userid.grid(row=0, column=1, padx=5, pady=5)
        self.combo_userid.bind("<<ComboboxSelected>>", self.on_userid_selected)
        
        tk.Label(login_mgmt_frame, text="交易帳號:", bg='lightyellow').grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.combo_account = ttk.Combobox(login_mgmt_frame, state='readonly', width=15)
        self.combo_account.grid(row=0, column=3, padx=5, pady=5)
        
        # 連線管理按鈕
        btn_frame = tk.Frame(login_mgmt_frame, bg='lightyellow')
        btn_frame.grid(row=1, column=0, columnspan=4, pady=10)
        
        tk.Button(btn_frame, text="🔗 Proxy連線", command=self.init_proxy, bg='orange', width=12).pack(side='left', padx=5)
        tk.Button(btn_frame, text="❌ Proxy斷線", command=self.disconnect_proxy, bg='red', fg='white', width=12).pack(side='left', padx=5)
        tk.Button(btn_frame, text="🔄 Proxy重連", command=self.reconnect_proxy, bg='blue', fg='white', width=12).pack(side='left', padx=5)
        tk.Button(btn_frame, text="📡 連線報價", command=self.connect_quote, bg='green', fg='white', width=12).pack(side='left', padx=5)
        
        # 工具按鈕框架
        tool_frame = tk.LabelFrame(self.login_frame, text="系統工具", font=('Arial', 10, 'bold'), bg='lightcyan')
        tool_frame.pack(fill='x', padx=5, pady=5)
        
        tools_btn_frame = tk.Frame(tool_frame, bg='lightcyan')
        tools_btn_frame.pack(pady=5)
        
        tk.Button(tools_btn_frame, text="📁 變更LOG路徑", command=self.set_log_path, width=15).pack(side='left', padx=5)
        tk.Button(tools_btn_frame, text="📜 同意書狀態", command=self.request_agreement, width=15).pack(side='left', padx=5)
        tk.Button(tools_btn_frame, text="📝 最後LOG", command=self.get_last_log, width=15).pack(side='left', padx=5)
        tk.Button(tools_btn_frame, text="🔧 下單初始化", command=self.init_order, width=15).pack(side='left', padx=5)
    
    def setup_stock_tab(self):
        # 股票下單區域
        order_frame = tk.LabelFrame(self.stock_frame, text="股票下單", font=('Arial', 10, 'bold'), bg='lightgreen')
        order_frame.pack(fill='x', padx=5, pady=5)
        
        # 第一行
        tk.Label(order_frame, text="股票代號:", bg='lightgreen').grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.entry_stock_no = tk.Entry(order_frame, width=10, font=('Arial', 10))
        self.entry_stock_no.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(order_frame, text="交易別:", bg='lightgreen').grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.combo_stock_buysell = ttk.Combobox(order_frame, values=Config.comboBoxBuySell, state='readonly', width=8)
        self.combo_stock_buysell.grid(row=0, column=3, padx=5, pady=5)
        
        tk.Label(order_frame, text="數量:", bg='lightgreen').grid(row=0, column=4, sticky='w', padx=5, pady=5)
        self.entry_stock_qty = tk.Entry(order_frame, width=10, font=('Arial', 10))
        self.entry_stock_qty.grid(row=0, column=5, padx=5, pady=5)
        
        # 第二行
        tk.Label(order_frame, text="價格:", bg='lightgreen').grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.entry_stock_price = tk.Entry(order_frame, width=10, font=('Arial', 10))
        self.entry_stock_price.grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(order_frame, text="市場別:", bg='lightgreen').grid(row=1, column=2, sticky='w', padx=5, pady=5)
        self.combo_stock_prime = ttk.Combobox(order_frame, values=Config.comboBoxPrime, state='readonly', width=8)
        self.combo_stock_prime.grid(row=1, column=3, padx=5, pady=5)
        self.combo_stock_prime.set("上市上櫃")
        
        tk.Label(order_frame, text="交易時段:", bg='lightgreen').grid(row=1, column=4, sticky='w', padx=5, pady=5)
        self.combo_stock_period = ttk.Combobox(order_frame, values=Config.comboBoxPeriod, state='readonly', width=8)
        self.combo_stock_period.grid(row=1, column=5, padx=5, pady=5)
        self.combo_stock_period.set("盤中")
        
        # 第三行
        tk.Label(order_frame, text="現券別:", bg='lightgreen').grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.combo_stock_flag = ttk.Combobox(order_frame, values=Config.comboBoxFlag, state='readonly', width=8)
        self.combo_stock_flag.grid(row=2, column=1, padx=5, pady=5)
        self.combo_stock_flag.set("現股")
        
        tk.Label(order_frame, text="委託條件:", bg='lightgreen').grid(row=2, column=2, sticky='w', padx=5, pady=5)
        self.combo_stock_tradetype = ttk.Combobox(order_frame, values=Config.comboBoxTradeType, state='readonly', width=8)
        self.combo_stock_tradetype.grid(row=2, column=3, padx=5, pady=5)
        self.combo_stock_tradetype.set("ROD")
        
        tk.Label(order_frame, text="市限價:", bg='lightgreen').grid(row=2, column=4, sticky='w', padx=5, pady=5)
        self.combo_stock_special = ttk.Combobox(order_frame, values=Config.comboBoxSpecialTradeType, state='readonly', width=8)
        self.combo_stock_special.grid(row=2, column=5, padx=5, pady=5)
        self.combo_stock_special.set("限價")
        
        # 下單按鈕
        btn_stock_frame = tk.Frame(order_frame, bg='lightgreen')
        btn_stock_frame.grid(row=3, column=0, columnspan=6, pady=15)
        
        tk.Button(btn_stock_frame, text="📊 送出股票委託", command=self.send_stock_order, 
                 bg='orange', fg='white', font=('Arial', 10, 'bold'), width=15).pack(side='left', padx=10)
        tk.Button(btn_stock_frame, text="⚡ 非同步委託", command=self.send_stock_order_async, 
                 bg='purple', fg='white', width=12).pack(side='left', padx=10)
        
        # 查詢區域
        query_frame = tk.LabelFrame(self.stock_frame, text="查詢功能", font=('Arial', 10, 'bold'), bg='lightblue')
        query_frame.pack(fill='x', padx=5, pady=5)
        
        tk.Label(query_frame, text="委託回報:", bg='lightblue').grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.combo_order_report = ttk.Combobox(query_frame, values=Config.comboBoxGetOrderReport, state='readonly', width=12)
        self.combo_order_report.grid(row=0, column=1, padx=5, pady=5)
        self.combo_order_report.set("1:全部")
        tk.Button(query_frame, text="🔍 查詢", command=self.get_order_report, bg='lightcyan').grid(row=0, column=2, padx=5, pady=5)
        
        tk.Label(query_frame, text="成交回報:", bg='lightblue').grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.combo_fulfill_report = ttk.Combobox(query_frame, values=Config.comboBoxGetFulfillReport, state='readonly', width=12)
        self.combo_fulfill_report.grid(row=1, column=1, padx=5, pady=5)
        self.combo_fulfill_report.set("1:全部")
        tk.Button(query_frame, text="🔍 查詢", command=self.get_fulfill_report, bg='lightcyan').grid(row=1, column=2, padx=5, pady=5)
    
    def setup_futures_tab(self):
        # 期貨下單區域
        futures_frame = tk.LabelFrame(self.futures_frame, text="期貨下單", font=('Arial', 10, 'bold'), bg='lightcoral')
        futures_frame.pack(fill='x', padx=5, pady=5)
        
        # 第一行
        tk.Label(futures_frame, text="期貨代號:", bg='lightcoral').grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.entry_futures_no = tk.Entry(futures_frame, width=12, font=('Arial', 10))
        self.entry_futures_no.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(futures_frame, text="買賣別:", bg='lightcoral').grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.combo_futures_buysell = ttk.Combobox(futures_frame, values=Config.comboBoxFuturesBuySell, state='readonly', width=8)
        self.combo_futures_buysell.grid(row=0, column=3, padx=5, pady=5)
        
        tk.Label(futures_frame, text="口數:", bg='lightcoral').grid(row=0, column=4, sticky='w', padx=5, pady=5)
        self.entry_futures_qty = tk.Entry(futures_frame, width=10, font=('Arial', 10))
        self.entry_futures_qty.grid(row=0, column=5, padx=5, pady=5)
        
        # 第二行
        tk.Label(futures_frame, text="價格:", bg='lightcoral').grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.entry_futures_price = tk.Entry(futures_frame, width=12, font=('Arial', 10))
        self.entry_futures_price.grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(futures_frame, text="當沖:", bg='lightcoral').grid(row=1, column=2, sticky='w', padx=5, pady=5)
        self.combo_futures_daytrade = ttk.Combobox(futures_frame, values=Config.comboBoxFuturesDayTrade, state='readonly', width=8)
        self.combo_futures_daytrade.grid(row=1, column=3, padx=5, pady=5)
        self.combo_futures_daytrade.set("否")
        
        tk.Label(futures_frame, text="新平倉:", bg='lightcoral').grid(row=1, column=4, sticky='w', padx=5, pady=5)
        self.combo_futures_newclose = ttk.Combobox(futures_frame, values=Config.comboBoxFuturesNewClose, state='readonly', width=8)
        self.combo_futures_newclose.grid(row=1, column=5, padx=5, pady=5)
        self.combo_futures_newclose.set("新倉")
        
        # 下單按鈕
        btn_futures_frame = tk.Frame(futures_frame, bg='lightcoral')
        btn_futures_frame.grid(row=2, column=0, columnspan=6, pady=15)
        
        tk.Button(btn_futures_frame, text="🎯 送出期貨委託", command=self.send_futures_order, 
                 bg='purple', fg='white', font=('Arial', 10, 'bold'), width=15).pack(side='left', padx=10)
        tk.Button(btn_futures_frame, text="⚡ 非同步委託", command=self.send_futures_order_async, 
                 bg='darkred', fg='white', width=12).pack(side='left', padx=10)
        
        # 期貨查詢功能
        futures_query_frame = tk.LabelFrame(self.futures_frame, text="期貨查詢", font=('Arial', 10, 'bold'), bg='lightyellow')
        futures_query_frame.pack(fill='x', padx=5, pady=5)
        
        query_btn_frame = tk.Frame(futures_query_frame, bg='lightyellow')
        query_btn_frame.pack(pady=5)
        
        tk.Button(query_btn_frame, text="📋 期貨未平倉", command=self.get_open_interest, 
                 bg='lightblue', width=12).pack(side='left', padx=5, pady=5)
        tk.Button(query_btn_frame, text="💰 期貨權益數", command=self.get_futures_rights, 
                 bg='lightgreen', width=12).pack(side='left', padx=5, pady=5)
    
    def setup_options_tab(self):
        # 選擇權交易區域
        options_frame = tk.LabelFrame(self.options_frame, text="選擇權交易", font=('Arial', 10, 'bold'), bg='lightsteelblue')
        options_frame.pack(fill='x', padx=5, pady=5)
        
        tk.Label(options_frame, text="選擇權代號:", bg='lightsteelblue').grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.entry_option_no = tk.Entry(options_frame, width=15, font=('Arial', 10))
        self.entry_option_no.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(options_frame, text="買賣別:", bg='lightsteelblue').grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.combo_option_buysell = ttk.Combobox(options_frame, values=Config.comboBoxFuturesBuySell, state='readonly', width=8)
        self.combo_option_buysell.grid(row=0, column=3, padx=5, pady=5)
        
        tk.Label(options_frame, text="口數:", bg='lightsteelblue').grid(row=0, column=4, sticky='w', padx=5, pady=5)
        self.entry_option_qty = tk.Entry(options_frame, width=10, font=('Arial', 10))
        self.entry_option_qty.grid(row=0, column=5, padx=5, pady=5)
        
        tk.Label(options_frame, text="價格:", bg='lightsteelblue').grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.entry_option_price = tk.Entry(options_frame, width=15, font=('Arial', 10))
        self.entry_option_price.grid(row=1, column=1, padx=5, pady=5)
        
        # 下單按鈕
        btn_option_frame = tk.Frame(options_frame, bg='lightsteelblue')
        btn_option_frame.grid(row=2, column=0, columnspan=6, pady=15)
        
        tk.Button(btn_option_frame, text="🎯 送出選擇權委託", command=self.send_option_order, 
                 bg='darkblue', fg='white', font=('Arial', 10, 'bold'), width=18).pack(side='left', padx=10)
    
    def setup_quote_tab(self):
        # 報價連線區域
        quote_conn_frame = tk.LabelFrame(self.quote_frame, text="報價連線", font=('Arial', 10, 'bold'), bg='lightpink')
        quote_conn_frame.pack(fill='x', padx=5, pady=5)
        
        conn_btn_frame = tk.Frame(quote_conn_frame, bg='lightpink')
        conn_btn_frame.pack(pady=5)
        
        tk.Button(conn_btn_frame, text="🚀 初始化報價", command=self.init_quote, 
                 bg='orange', width=12).pack(side='left', padx=5, pady=5)
        tk.Button(conn_btn_frame, text="📡 連線報價", command=self.connect_quote, 
                 bg='green', fg='white', width=12).pack(side='left', padx=5, pady=5)
        tk.Button(conn_btn_frame, text="❌ 離線報價", command=self.disconnect_quote, 
                 bg='red', fg='white', width=12).pack(side='left', padx=5, pady=5)
        
        # 股票報價區域
        stock_quote_frame = tk.LabelFrame(self.quote_frame, text="股票報價訂閱", font=('Arial', 10, 'bold'), bg='lightcyan')
        stock_quote_frame.pack(fill='x', padx=5, pady=5)
        
        tk.Label(stock_quote_frame, text="股票代號:", bg='lightcyan').grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.entry_quote_stock = tk.Entry(stock_quote_frame, width=15, font=('Arial', 10))
        self.entry_quote_stock.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(stock_quote_frame, text="市場別:", bg='lightcyan').grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.combo_quote_market = ttk.Combobox(stock_quote_frame, values=Config.comboBoxSKQuoteLib_RequestStocks, 
                                              state='readonly', width=10)
        self.combo_quote_market.grid(row=0, column=3, padx=5, pady=5)
        self.combo_quote_market.set("上市")
        
        btn_quote_frame = tk.Frame(stock_quote_frame, bg='lightcyan')
        btn_quote_frame.grid(row=1, column=0, columnspan=4, pady=10)
        
        tk.Button(btn_quote_frame, text="📊 訂閱報價", command=self.request_stock_quote, 
                 bg='blue', fg='white', width=12).pack(side='left', padx=5)
        tk.Button(btn_quote_frame, text="❌ 取消訂閱", command=self.cancel_stock_quote, 
                 bg='red', fg='white', width=12).pack(side='left', padx=5)
        
        # K線查詢區域
        kline_frame = tk.LabelFrame(self.quote_frame, text="K線查詢", font=('Arial', 10, 'bold'), bg='lavender')
        kline_frame.pack(fill='x', padx=5, pady=5)
        
        tk.Label(kline_frame, text="商品代號:", bg='lavender').grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.entry_kline_stock = tk.Entry(kline_frame, width=15, font=('Arial', 10))
        self.entry_kline_stock.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(kline_frame, text="K線類型:", bg='lavender').grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.combo_kline_type = ttk.Combobox(kline_frame, values=Config.comboBoxSKQuoteLib_RequestKLine, 
                                            state='readonly', width=10)
        self.combo_kline_type.grid(row=0, column=3, padx=5, pady=5)
        self.combo_kline_type.set("日K")
        
        tk.Button(kline_frame, text="📈 查詢K線", command=self.request_kline, 
                 bg='purple', fg='white', width=12).grid(row=1, column=1, pady=10)
    
    def setup_message_area(self):
        # 訊息顯示區域
        message_frame = tk.LabelFrame(self, text="🔔 系統訊息", font=('Arial', 10, 'bold'), bg='lightyellow')
        message_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 方法訊息框
        tk.Label(message_frame, text="📋 方法回傳訊息:", bg='lightyellow', font=('Arial', 9, 'bold')).pack(anchor='w', padx=5)
        self.richTextBoxMethodMessage = tk.Listbox(message_frame, height=4, font=('Consolas', 8))
        self.richTextBoxMethodMessage.pack(fill='x', padx=5, pady=2)
        
        global richTextBoxMethodMessage
        richTextBoxMethodMessage = self.richTextBoxMethodMessage
        
        # 事件訊息框
        tk.Label(message_frame, text="📢 事件回報訊息:", bg='lightyellow', font=('Arial', 9, 'bold')).pack(anchor='w', padx=5)
        self.richTextBoxMessage = tk.Listbox(message_frame, height=8, font=('Consolas', 8))
        self.richTextBoxMessage.pack(fill='both', expand=True, padx=5, pady=2)
        
        global richTextBoxMessage
        richTextBoxMessage = self.richTextBoxMessage
    
    # 事件處理方法
    def on_authority_changed(self, event):
        authority_map = {"正式環境": 0, "正式環境SGX": 1, "測試環境": 2, "測試環境SGX": 3}
        nAuthorityFlag = authority_map.get(self.combo_authority.get(), 0)
        nCode = m_pSKCenter.SKCenterLib_SetAuthority(nAuthorityFlag)
        msg = f"【SKCenterLib_SetAuthority】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
    
    def on_ap_checked(self):
        if self.var_is_ap.get() == 1:
            self.label_cert_id.grid(row=3, column=1, sticky='w', padx=5, pady=5)
            self.entry_cert_id.grid(row=3, column=2, padx=5, pady=5)
            self.btn_generate_cert.grid(row=4, column=0, padx=5, pady=10)
        else:
            self.label_cert_id.grid_remove()
            self.entry_cert_id.grid_remove()
            self.btn_generate_cert.grid_remove()
    
    def on_userid_selected(self, event):
        m_pSKOrder.SKOrderLib_Initialize()
        m_pSKOrder.GetUserAccount()
        self.combo_userid['values'] = list(dictUserID.keys())
        if self.combo_userid.get() in dictUserID:
            self.combo_account['values'] = dictUserID[self.combo_userid.get()]
    
    # 登入相關方法
    def login(self):
        if not self.entry_userid.get() or not self.entry_password.get():
            messagebox.showwarning("警告", "請輸入完整的帳號密碼!")
            return
            
        nCode = m_pSKCenter.SKCenterLib_Login(self.entry_userid.get(), self.entry_password.get())
        msg = f"【SKCenterLib_Login】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
        
        if nCode == 0:
            messagebox.showinfo("登入成功", "登入成功! 請等待帳號資訊回傳...")
    
    def generate_cert(self):
        nCode = m_pSKCenter.SKCenterLib_GenerateKeyCert(self.entry_userid.get(), self.entry_cert_id.get())
        msg = f"【SKCenterLib_GenerateKeyCert】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
    
    def init_proxy(self):
        if not self.combo_userid.get():
            messagebox.showwarning("警告", "請先選擇使用者ID!")
            return
            
        nCode = m_pSKOrder.SKOrderLib_InitialProxyByID(self.combo_userid.get())
        msg = f"【SKOrderLib_InitialProxyByID】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
    
    def disconnect_proxy(self):
        nCode = m_pSKOrder.ProxyDisconnectByID(self.combo_userid.get())
        msg = f"【ProxyDisconnectByID】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
    
    def reconnect_proxy(self):
        nCode = m_pSKOrder.ProxyReconnectByID(self.combo_userid.get())
        msg = f"【ProxyReconnectByID】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
    
    def init_order(self):
        nCode = m_pSKOrder.SKOrderLib_Initialize()
        msg = f"【SKOrderLib_Initialize】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
    
    def set_log_path(self):
        folder_selected = filedialog.askdirectory(title="選擇LOG資料夾")
        if folder_selected:
            nCode = m_pSKCenter.SKCenterLib_SetLogPath(folder_selected)
            msg = f"【SKCenterLib_SetLogPath】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
            richTextBoxMethodMessage.insert('end', msg + "\n")
            richTextBoxMethodMessage.see('end')
    
    def request_agreement(self):
        if not self.combo_userid.get():
            messagebox.showwarning("警告", "請先選擇使用者ID!")
            return
            
        nCode = m_pSKCenter.SKCenterLib_RequestAgreement(self.combo_userid.get())
        msg = f"【SKCenterLib_RequestAgreement】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
    
    def get_last_log(self):
        log_info = m_pSKCenter.SKCenterLib_GetLastLogInfo()
        msg = f"【SKCenterLib_GetLastLogInfo】{log_info}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
    
    # 股票下單相關方法
    def send_stock_order(self):
        try:
            if not all([self.combo_userid.get(), self.combo_account.get(), 
                       self.entry_stock_no.get(), self.entry_stock_qty.get(), 
                       self.entry_stock_price.get()]):
                messagebox.showwarning("警告", "請填寫完整的委託資訊!")
                return
            
            # 取得介面參數
            stock_no = self.entry_stock_no.get()
            qty = int(self.entry_stock_qty.get())
            price = self.entry_stock_price.get()
            
            # 轉換參數
            prime_map = {"上市上櫃": 0, "興櫃": 1}
            period_map = {"盤中": 0, "盤後": 1, "零股": 2}
            flag_map = {"現股": 0, "融資": 1, "融券": 2, "無券": 3}
            buysell_map = {"買進": 0, "賣出": 1}
            tradetype_map = {"ROD": 0, "IOC": 1, "FOK": 2}
            special_map = {"市價": 0, "限價": 1}
            
            nPrime = prime_map.get(self.combo_stock_prime.get(), 0)
            nPeriod = period_map.get(self.combo_stock_period.get(), 0)
            nFlag = flag_map.get(self.combo_stock_flag.get(), 0)
            nBuySell = buysell_map.get(self.combo_stock_buysell.get(), 0)
            nTradeType = tradetype_map.get(self.combo_stock_tradetype.get(), 0)
            nSpecialTradeType = special_map.get(self.combo_stock_special.get(), 1)
            
            # 送出委託
            nCode = m_pSKOrder.SendStockOrder(
                self.combo_userid.get(),
                self.combo_account.get(),
                stock_no,
                nBuySell,
                qty,
                price,
                nPrime,
                nPeriod,
                nFlag,
                nTradeType,
                nSpecialTradeType
            )
            
            msg = f"【SendStockOrder】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
            richTextBoxMethodMessage.insert('end', msg + "\n")
            richTextBoxMethodMessage.see('end')
            
            if nCode == 0:
                messagebox.showinfo("下單成功", f"股票委託已送出: {stock_no}")
                
        except ValueError:
            messagebox.showerror("錯誤", "數量請輸入有效數字!")
        except Exception as e:
            messagebox.showerror("錯誤", f"下單發生錯誤: {e}")
    
    def send_stock_order_async(self):
        # 非同步股票下單實作
        global bAsyncOrder
        bAsyncOrder = True
        self.send_stock_order()
        bAsyncOrder = False
    
    def send_futures_order(self):
        try:
            if not all([self.combo_userid.get(), self.combo_account.get(),
                       self.entry_futures_no.get(), self.entry_futures_qty.get(),
                       self.entry_futures_price.get()]):
                messagebox.showwarning("警告", "請填寫完整的期貨委託資訊!")
                return
            
            # 期貨下單實作
            futures_no = self.entry_futures_no.get()
            qty = int(self.entry_futures_qty.get())
            price = self.entry_futures_price.get()
            
            buysell_map = {"買進": 0, "賣出": 1}
            daytrade_map = {"是": 1, "否": 0}
            newclose_map = {"新倉": 0, "平倉": 1, "自動": 2}
            
            nBuySell = buysell_map.get(self.combo_futures_buysell.get(), 0)
            nDayTrade = daytrade_map.get(self.combo_futures_daytrade.get(), 0)
            nNewClose = newclose_map.get(self.combo_futures_newclose.get(), 0)
            
            nCode = m_pSKOrder.SendFutureOrder(
                self.combo_userid.get(),
                self.combo_account.get(),
                futures_no,
                nBuySell,
                qty,
                price,
                nNewClose,
                nDayTrade,
                0  # nTradeType ROD
            )
            
            msg = f"【SendFutureOrder】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
            richTextBoxMethodMessage.insert('end', msg + "\n")
            richTextBoxMethodMessage.see('end')
            
            if nCode == 0:
                messagebox.showinfo("下單成功", f"期貨委託已送出: {futures_no}")
                
        except ValueError:
            messagebox.showerror("錯誤", "口數請輸入有效數字!")
        except Exception as e:
            messagebox.showerror("錯誤", f"期貨下單發生錯誤: {e}")
    
    def send_futures_order_async(self):
        # 非同步期貨下單
        global bAsyncOrder
        bAsyncOrder = True
        self.send_futures_order()
        bAsyncOrder = False
    
    def send_option_order(self):
        try:
            if not all([self.combo_userid.get(), self.combo_account.get(),
                       self.entry_option_no.get(), self.entry_option_qty.get(),
                       self.entry_option_price.get()]):
                messagebox.showwarning("警告", "請填寫完整的選擇權委託資訊!")
                return
            
            option_no = self.entry_option_no.get()
            qty = int(self.entry_option_qty.get())
            price = self.entry_option_price.get()
            
            buysell_map = {"買進": 0, "賣出": 1}
            nBuySell = buysell_map.get(self.combo_option_buysell.get(), 0)
            
            nCode = m_pSKOrder.SendOptionOrder(
                self.combo_userid.get(),
                self.combo_account.get(),
                option_no,
                nBuySell,
                qty,
                price,
                0,  # nNewClose
                0,  # nDayTrade
                0   # nTradeType
            )
            
            msg = f"【SendOptionOrder】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
            richTextBoxMethodMessage.insert('end', msg + "\n")
            richTextBoxMethodMessage.see('end')
            
            if nCode == 0:
                messagebox.showinfo("下單成功", f"選擇權委託已送出: {option_no}")
                
        except ValueError:
            messagebox.showerror("錯誤", "口數請輸入有效數字!")
        except Exception as e:
            messagebox.showerror("錯誤", f"選擇權下單發生錯誤: {e}")
    
    # 查詢相關方法
    def get_order_report(self):
        if not self.combo_userid.get() or not self.combo_account.get():
            messagebox.showwarning("警告", "請先選擇使用者ID和交易帳號!")
            return
            
        format_map = {
            "1:全部": 1, "2:有效": 2, "3:可消": 3, "4:已消": 4, 
            "5:已成": 5, "6:失敗": 6, "7:合併同價格": 7, "8:合併同商品": 8, "9:預約": 9
        }
        nFormat = format_map.get(self.combo_order_report.get(), 1)
        
        bstrResult = m_pSKOrder.GetOrderReport(self.combo_userid.get(), self.combo_account.get(), nFormat)
        msg = f"【GetOrderReport】{bstrResult}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')
    
    def get_fulfill_report(self):
        if not self.combo_userid.get() or not self.combo_account.get():
            messagebox.showwarning("警告", "請先選擇使用者ID和交易帳號!")
            return
            
        format_map = {
            "1:全部": 1, "2:合併同書號": 2, "3:合併同價格": 3, 
            "4:合併同商品": 4, "5:T+1成交回報": 5
        }
        nFormat = format_map.get(self.combo_fulfill_report.get(), 1)
        
        bstrResult = m_pSKOrder.GetFulfillReport(self.combo_userid.get(), self.combo_account.get(), nFormat)
        msg = f"【GetFulfillReport】{bstrResult}"
        richTextBoxMessage.insert('end', msg + "\n")
        richTextBoxMessage.see('end')
    
    def get_open_interest(self):
        if not self.combo_userid.get() or not self.combo_account.get():
            messagebox.showwarning("警告", "請先選擇使用者ID和交易帳號!")
            return
            
        nCode = m_pSKOrder.GetOpenInterestGW(self.combo_userid.get(), self.combo_account.get())
        msg = f"【GetOpenInterestGW】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
    
    def get_futures_rights(self):
        if not self.combo_userid.get() or not self.combo_account.get():
            messagebox.showwarning("警告", "請先選擇使用者ID和交易帳號!")
            return
            
        nCode = m_pSKOrder.GetFutureRights(self.combo_userid.get(), self.combo_account.get())
        msg = f"【GetFutureRights】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
    
    # 報價相關方法
    def init_quote(self):
        nCode = m_pSKQuote.SKQuoteLib_Initialize()
        msg = f"【SKQuoteLib_Initialize】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
    
    def connect_quote(self):
        nCode = m_pSKQuote.SKQuoteLib_EnterMonitor()
        msg = f"【SKQuoteLib_EnterMonitor】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
        
        if nCode == 0:
            messagebox.showinfo("連線成功", "報價伺服器連線成功!")
    
    def disconnect_quote(self):
        nCode = m_pSKQuote.SKQuoteLib_LeaveMonitor()
        msg = f"【SKQuoteLib_LeaveMonitor】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
    
    def request_stock_quote(self):
        if not self.entry_quote_stock.get():
            messagebox.showwarning("警告", "請輸入股票代號!")
            return
            
        stock_no = self.entry_quote_stock.get()
        market_map = {"上市": 0, "上櫃": 1, "興櫃": 2}
        nMarket = market_map.get(self.combo_quote_market.get(), 0)
        
        nCode = m_pSKQuote.SKQuoteLib_RequestStocks([nMarket], [stock_no])
        msg = f"【SKQuoteLib_RequestStocks】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
        
        if nCode == 0:
            messagebox.showinfo("訂閱成功", f"已訂閱 {stock_no} 報價")
    
    def cancel_stock_quote(self):
        stock_no = self.entry_quote_stock.get()
        if not stock_no:
            messagebox.showwarning("警告", "請輸入要取消的股票代號!")
            return
            
        nCode = m_pSKQuote.SKQuoteLib_RequestStocks([999], [stock_no])  # 999表示取消
        msg = f"【取消訂閱】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')
    
    def request_kline(self):
        if not self.entry_kline_stock.get():
            messagebox.showwarning("警告", "請輸入商品代號!")
            return
            
        stock_no = self.entry_kline_stock.get()
        kline_type_map = {"分K": 0, "日K": 1, "週K": 2, "月K": 3}
        nKLineType = kline_type_map.get(self.combo_kline_type.get(), 1)
        
        nCode = m_pSKQuote.SKQuoteLib_RequestKLine(stock_no, nKLineType, 30)  # 查詢30筆
        msg = f"【SKQuoteLib_RequestKLine】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        richTextBoxMethodMessage.insert('end', msg + "\n")
        richTextBoxMethodMessage.see('end')

######################################################################################################################################
# 主程式

if __name__ == '__main__':
    try:
        print("="*60)
        print("🚀 群益證券API整合測試工具 v2.13.55")
        print("="*60)
        
        app = CapitalIntegratedTester()
        
        # 顯示版本資訊
        version_info = m_pSKCenter.SKCenterLib_GetSKAPIVersionAndBit("")
        print(f"📌 API版本資訊: {version_info}")
        print("✅ 程式初始化完成，正在啟動GUI介面...")
        
        app.mainloop()
        
    except Exception as e:
        error_msg = f"程式執行錯誤: {e}"
        print(f"❌ {error_msg}")
        messagebox.showerror("錯誤", error_msg)