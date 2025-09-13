# 群益證券 API 整合測試工具
# Capital Securities Integrated Tester
# 整合 Login, Order, Quote, Reply 所有功能

import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox, filedialog
import comtypes.client
import Config

# 初始化 COM 元件
try:
    comtypes.client.GetModule(r'SKCOM.dll')
    import comtypes.gen.SKCOMLib as sk
    
    # 建立 COM 物件
    m_pSKCenter = comtypes.client.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
    m_pSKReply = comtypes.client.CreateObject(sk.SKReplyLib, interface=sk.ISKReplyLib)
    m_pSKOrder = comtypes.client.CreateObject(sk.SKOrderLib, interface=sk.ISKOrderLib)
    m_pSKQuote = comtypes.client.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib)
    m_pSKOSQuote = comtypes.client.CreateObject(sk.SKOSQuoteLib, interface=sk.ISKOSQuoteLib)
    m_pSKOOQuote = comtypes.client.CreateObject(sk.SKOOQuoteLib, interface=sk.ISKOOQuoteLib)
    
    print("=> 群益證券 COM 元件初始化成功")
    
except Exception as e:
    print(f"=> COM 元件初始化失敗: {e}")
    messagebox.showerror("錯誤", f"COM 元件初始化失敗:\n{e}")
    exit(1)

# 全域變數
dictUserID = {}
dictUserID["更新帳號"] = ["無"]
bAsyncOrder = False
sServer = 0  # 報價伺服器 0:主要 1:備援

class CapitalIntegratedTester:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("群益證券 API 整合測試工具 v1.0")
        self.root.geometry("1200x800")
        
        # 設定視窗圖示和樣式
        self.setup_style()
        
        # 建立選單
        self.create_menu()
        
        # 建立主要介面
        self.create_main_interface()
        
        # 初始化事件處理
        self.setup_event_handlers()
        
    def setup_style(self):
        """設定介面樣式"""
        self.root.configure(bg='#f0f0f0')
        
        # 設定 ttk 樣式
        style = ttk.Style()
        style.theme_use('clam')
        
    def create_menu(self):
        """建立選單列"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 檔案選單
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="檔案", menu=file_menu)
        file_menu.add_command(label="設定LOG路徑", command=self.set_log_path)
        file_menu.add_separator()
        file_menu.add_command(label="結束", command=self.root.quit)
        
        # 工具選單
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="工具", menu=tools_menu)
        tools_menu.add_command(label="API版本資訊", command=self.show_api_version)
        tools_menu.add_command(label="最後LOG資訊", command=self.show_last_log)
        
        # 說明選單
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="說明", menu=help_menu)
        help_menu.add_command(label="關於", command=self.show_about)
        
    def create_main_interface(self):
        """建立主要介面"""
        # 建立上方工具列
        self.create_toolbar()
        
        # 建立主要 Notebook
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 建立各個分頁
        self.create_login_tab()
        self.create_order_tabs()
        self.create_quote_tabs()
        self.create_reply_tab()
        
        # 建立下方狀態列和訊息區
        self.create_status_area()
        
    def create_toolbar(self):
        """建立工具列"""
        toolbar_frame = tk.Frame(self.root, bg='#e0e0e0', height=40)
        toolbar_frame.pack(fill='x', side='top')
        
        # API 版本顯示
        version_label = tk.Label(toolbar_frame, text=f"API版本: {self.get_api_version()}", 
                                bg='#e0e0e0', font=('Arial', 9))
        version_label.pack(side='left', padx=10, pady=5)
        
        # 連線狀態指示燈
        self.connection_status = tk.Label(toolbar_frame, text="●", fg='red', 
                                        bg='#e0e0e0', font=('Arial', 16))
        self.connection_status.pack(side='right', padx=10, pady=5)
        
        status_text = tk.Label(toolbar_frame, text="連線狀態:", bg='#e0e0e0')
        status_text.pack(side='right', pady=5)
        
    def create_login_tab(self):
        """建立登入分頁"""
        login_frame = ttk.Frame(self.notebook)
        self.notebook.add(login_frame, text='登入')
        
        # 登入設定區域
        login_settings = ttk.LabelFrame(login_frame, text="登入設定", padding=10)
        login_settings.pack(fill='x', padx=10, pady=5)
        
        # 連線環境
        ttk.Label(login_settings, text="連線環境:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.combo_authority = ttk.Combobox(login_settings, values=Config.comboBoxSKCenterLib_SetAuthority, 
                                          state='readonly', width=15)
        self.combo_authority.grid(row=0, column=1, padx=5, pady=5)
        self.combo_authority.bind("<<ComboboxSelected>>", self.on_authority_change)
        
        # 使用者帳號
        ttk.Label(login_settings, text="使用者帳號:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.entry_userid = tk.Entry(login_settings, width=20)
        self.entry_userid.grid(row=1, column=1, padx=5, pady=5)
        
        # 密碼
        ttk.Label(login_settings, text="密碼:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.entry_password = tk.Entry(login_settings, show='*', width=20)
        self.entry_password.grid(row=2, column=1, padx=5, pady=5)
        
        # AP/APH 身分
        self.var_ap = tk.IntVar()
        self.check_ap = tk.Checkbutton(login_settings, text="AP/APH身分", variable=self.var_ap,
                                     command=self.on_ap_check)
        self.check_ap.grid(row=3, column=0, columnspan=2, sticky='w', padx=5, pady=5)
        
        # 憑證ID (初始隱藏)
        self.label_certid = ttk.Label(login_settings, text="憑證ID:")
        self.entry_certid = tk.Entry(login_settings, width=20)
        
        # 登入按鈕
        btn_frame = tk.Frame(login_settings)
        btn_frame.grid(row=5, column=0, columnspan=3, pady=10)
        
        self.btn_login = tk.Button(btn_frame, text="登入", command=self.login, 
                                 bg='#4CAF50', fg='white', font=('Arial', 10, 'bold'))
        self.btn_login.pack(side='left', padx=5)
        
        self.btn_generate_cert = tk.Button(btn_frame, text="雙因子驗證KEY", command=self.generate_cert)
        # 初始隱藏
        
        # 帳號管理區域
        account_frame = ttk.LabelFrame(login_frame, text="帳號管理", padding=10)
        account_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(account_frame, text="選擇使用者:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.combo_userid = ttk.Combobox(account_frame, values=list(dictUserID.keys()), 
                                       state='readonly', width=15)
        self.combo_userid.grid(row=0, column=1, padx=5, pady=5)
        self.combo_userid.bind("<<ComboboxSelected>>", self.on_userid_change)
        
        ttk.Label(account_frame, text="交易帳號:").grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.combo_account = ttk.Combobox(account_frame, state='readonly', width=15)
        self.combo_account.grid(row=0, column=3, padx=5, pady=5)
        
        # 代理伺服器控制
        proxy_frame = tk.Frame(account_frame)
        proxy_frame.grid(row=1, column=0, columnspan=4, pady=10)
        
        tk.Button(proxy_frame, text="初始化Proxy", command=self.init_proxy).pack(side='left', padx=5)
        tk.Button(proxy_frame, text="斷線Proxy", command=self.disconnect_proxy).pack(side='left', padx=5)
        tk.Button(proxy_frame, text="重連Proxy", command=self.reconnect_proxy).pack(side='left', padx=5)
        tk.Button(proxy_frame, text="建立SGX專線", command=self.add_sgx_socket).pack(side='left', padx=5)
        
    def create_order_tabs(self):
        """建立下單相關分頁"""
        # 下單主分頁
        order_frame = ttk.Frame(self.notebook)
        self.notebook.add(order_frame, text='下單')
        
        # 建立下單子分頁
        order_notebook = ttk.Notebook(order_frame)
        order_notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 台股下單
        self.create_ts_order_tab(order_notebook)
        # 期貨下單
        self.create_tf_order_tab(order_notebook)
        # 海外股票
        self.create_os_order_tab(order_notebook)
        # 海外期貨
        self.create_of_order_tab(order_notebook)
        # 查詢功能
        self.create_query_tab(order_notebook)
        
    def create_ts_order_tab(self, parent):
        """建立台股下單分頁"""
        ts_frame = ttk.Frame(parent)
        parent.add(ts_frame, text='台股下單')
        
        # 下單設定區域
        order_settings = ttk.LabelFrame(ts_frame, text="台股下單設定", padding=10)
        order_settings.pack(fill='x', padx=10, pady=5)
        
        row = 0
        # 股票代號
        ttk.Label(order_settings, text="股票代號:").grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.entry_ts_stock = tk.Entry(order_settings, width=10)
        self.entry_ts_stock.grid(row=row, column=1, padx=5, pady=3)
        
        # 買賣別
        ttk.Label(order_settings, text="買賣別:").grid(row=row, column=2, sticky='w', padx=5, pady=3)
        self.combo_ts_buysell = ttk.Combobox(order_settings, values=Config.comboBoxTSBuySell, 
                                           state='readonly', width=10)
        self.combo_ts_buysell.grid(row=row, column=3, padx=5, pady=3)
        
        row += 1
        # 委託價格
        ttk.Label(order_settings, text="委託價格:").grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.entry_ts_price = tk.Entry(order_settings, width=10)
        self.entry_ts_price.grid(row=row, column=1, padx=5, pady=3)
        
        # 委託數量
        ttk.Label(order_settings, text="委託數量:").grid(row=row, column=2, sticky='w', padx=5, pady=3)
        self.entry_ts_qty = tk.Entry(order_settings, width=10)
        self.entry_ts_qty.grid(row=row, column=3, padx=5, pady=3)
        
        row += 1
        # 價格旗標
        ttk.Label(order_settings, text="價格旗標:").grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.combo_ts_priceflag = ttk.Combobox(order_settings, values=Config.comboBoxTSPriceFlag, 
                                             state='readonly', width=10)
        self.combo_ts_priceflag.grid(row=row, column=1, padx=5, pady=3)
        
        # 交易別
        ttk.Label(order_settings, text="交易別:").grid(row=row, column=2, sticky='w', padx=5, pady=3)
        self.combo_ts_tradetype = ttk.Combobox(order_settings, values=Config.comboBoxTSTradeType, 
                                             state='readonly', width=10)
        self.combo_ts_tradetype.grid(row=row, column=3, padx=5, pady=3)
        
        row += 1
        # 倉別
        ttk.Label(order_settings, text="倉別:").grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.combo_ts_period = ttk.Combobox(order_settings, values=Config.comboBoxTSPeriod, 
                                          state='readonly', width=10)
        self.combo_ts_period.grid(row=row, column=1, padx=5, pady=3)
        
        # 下單按鈕
        btn_frame = tk.Frame(order_settings)
        btn_frame.grid(row=row, column=2, columnspan=2, pady=10)
        
        tk.Button(btn_frame, text="送出委託", command=self.send_ts_order, 
                 bg='#2196F3', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
        tk.Button(btn_frame, text="非同步委託", command=self.send_ts_async_order).pack(side='left', padx=5)
        
    def create_tf_order_tab(self, parent):
        """建立期貨下單分頁"""
        tf_frame = ttk.Frame(parent)
        parent.add(tf_frame, text='期貨下單')
        
        # 下單設定區域
        order_settings = ttk.LabelFrame(tf_frame, text="期貨下單設定", padding=10)
        order_settings.pack(fill='x', padx=10, pady=5)
        
        row = 0
        # 商品代號
        ttk.Label(order_settings, text="商品代號:").grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.entry_tf_stock = tk.Entry(order_settings, width=12)
        self.entry_tf_stock.grid(row=row, column=1, padx=5, pady=3)
        
        # 買賣別
        ttk.Label(order_settings, text="買賣別:").grid(row=row, column=2, sticky='w', padx=5, pady=3)
        self.combo_tf_buysell = ttk.Combobox(order_settings, values=Config.comboBoxTFBuySell, 
                                           state='readonly', width=10)
        self.combo_tf_buysell.grid(row=row, column=3, padx=5, pady=3)
        
        row += 1
        # 委託價格
        ttk.Label(order_settings, text="委託價格:").grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.entry_tf_price = tk.Entry(order_settings, width=12)
        self.entry_tf_price.grid(row=row, column=1, padx=5, pady=3)
        
        # 委託數量
        ttk.Label(order_settings, text="委託數量:").grid(row=row, column=2, sticky='w', padx=5, pady=3)
        self.entry_tf_qty = tk.Entry(order_settings, width=10)
        self.entry_tf_qty.grid(row=row, column=3, padx=5, pady=3)
        
        row += 1
        # 新平倉
        ttk.Label(order_settings, text="新平倉:").grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.combo_tf_newclose = ttk.Combobox(order_settings, values=Config.comboBoxTFNewClose, 
                                            state='readonly', width=10)
        self.combo_tf_newclose.grid(row=row, column=1, padx=5, pady=3)
        
        # 委託條件
        ttk.Label(order_settings, text="委託條件:").grid(row=row, column=2, sticky='w', padx=5, pady=3)
        self.combo_tf_daytrade = ttk.Combobox(order_settings, values=Config.comboBoxTFDayTrade, 
                                            state='readonly', width=10)
        self.combo_tf_daytrade.grid(row=row, column=3, padx=5, pady=3)
        
        row += 1
        # 下單按鈕
        btn_frame = tk.Frame(order_settings)
        btn_frame.grid(row=row, column=0, columnspan=4, pady=10)
        
        tk.Button(btn_frame, text="送出期貨委託", command=self.send_tf_order, 
                 bg='#FF9800', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
        tk.Button(btn_frame, text="複式單委託", command=self.send_tf_duplex_order).pack(side='left', padx=5)
        
    def create_os_order_tab(self, parent):
        """建立海外股票下單分頁"""
        os_frame = ttk.Frame(parent)
        parent.add(os_frame, text='海外股票')
        
        # 實作海外股票下單介面
        order_settings = ttk.LabelFrame(os_frame, text="海外股票下單設定", padding=10)
        order_settings.pack(fill='x', padx=10, pady=5)
        
        # 基本設定
        ttk.Label(order_settings, text="股票代號:").grid(row=0, column=0, sticky='w', padx=5, pady=3)
        self.entry_os_stock = tk.Entry(order_settings, width=12)
        self.entry_os_stock.grid(row=0, column=1, padx=5, pady=3)
        
        ttk.Label(order_settings, text="交易市場:").grid(row=0, column=2, sticky='w', padx=5, pady=3)
        self.combo_os_market = ttk.Combobox(order_settings, values=Config.comboBoxOverseaMarket, 
                                          state='readonly', width=10)
        self.combo_os_market.grid(row=0, column=3, padx=5, pady=3)
        
        # 下單按鈕
        btn_frame = tk.Frame(order_settings)
        btn_frame.grid(row=2, column=0, columnspan=4, pady=10)
        
        tk.Button(btn_frame, text="送出海外股票委託", command=self.send_os_order, 
                 bg='#9C27B0', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
        
    def create_of_order_tab(self, parent):
        """建立海外期貨下單分頁"""
        of_frame = ttk.Frame(parent)
        parent.add(of_frame, text='海外期貨')
        
        # 實作海外期貨下單介面
        order_settings = ttk.LabelFrame(of_frame, text="海外期貨下單設定", padding=10)
        order_settings.pack(fill='x', padx=10, pady=5)
        
        # 基本設定
        ttk.Label(order_settings, text="商品代號:").grid(row=0, column=0, sticky='w', padx=5, pady=3)
        self.entry_of_stock = tk.Entry(order_settings, width=12)
        self.entry_of_stock.grid(row=0, column=1, padx=5, pady=3)
        
        # 下單按鈕
        btn_frame = tk.Frame(order_settings)
        btn_frame.grid(row=2, column=0, columnspan=4, pady=10)
        
        tk.Button(btn_frame, text="送出海外期貨委託", command=self.send_of_order, 
                 bg='#795548', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
        
    def create_query_tab(self, parent):
        """建立查詢功能分頁"""
        query_frame = ttk.Frame(parent)
        parent.add(query_frame, text='查詢')
        
        # 查詢設定區域
        query_settings = ttk.LabelFrame(query_frame, text="查詢設定", padding=10)
        query_settings.pack(fill='x', padx=10, pady=5)
        
        # 委託回報查詢
        ttk.Label(query_settings, text="委託回報:").grid(row=0, column=0, sticky='w', padx=5, pady=3)
        self.combo_order_report = ttk.Combobox(query_settings, values=Config.comboBoxGetOrderReport, 
                                             state='readonly', width=15)
        self.combo_order_report.grid(row=0, column=1, padx=5, pady=3)
        tk.Button(query_settings, text="查詢委託", command=self.query_order_report).grid(row=0, column=2, padx=5, pady=3)
        
        # 成交回報查詢
        ttk.Label(query_settings, text="成交回報:").grid(row=1, column=0, sticky='w', padx=5, pady=3)
        self.combo_fulfill_report = ttk.Combobox(query_settings, values=Config.comboBoxGetFulfillReport, 
                                               state='readonly', width=15)
        self.combo_fulfill_report.grid(row=1, column=1, padx=5, pady=3)
        tk.Button(query_settings, text="查詢成交", command=self.query_fulfill_report).grid(row=1, column=2, padx=5, pady=3)
        
        # 其他查詢
        btn_frame = tk.Frame(query_settings)
        btn_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        tk.Button(btn_frame, text="即時庫存", command=self.query_balance).pack(side='left', padx=5)
        tk.Button(btn_frame, text="集保庫存", command=self.query_balance_query).pack(side='left', padx=5)
        tk.Button(btn_frame, text="資券配額", command=self.query_margin_limit).pack(side='left', padx=5)
        tk.Button(btn_frame, text="損益查詢", command=self.query_profit_loss).pack(side='left', padx=5)
        
    def create_quote_tabs(self):
        """建立報價相關分頁"""
        # 報價主分頁
        quote_frame = ttk.Frame(self.notebook)
        self.notebook.add(quote_frame, text='報價')
        
        # 建立報價子分頁
        quote_notebook = ttk.Notebook(quote_frame)
        quote_notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 台股報價
        self.create_tw_quote_tab(quote_notebook)
        # 海外股票報價
        self.create_os_quote_tab(quote_notebook)
        # 海外選擇權報價
        self.create_oo_quote_tab(quote_notebook)
        # 市場統計
        self.create_market_stats_tab(quote_notebook)
        
    def create_tw_quote_tab(self, parent):
        """建立台股報價分頁"""
        tw_frame = ttk.Frame(parent)
        parent.add(tw_frame, text='台股報價')
        
        # 報價訂閱區域
        subscribe_frame = ttk.LabelFrame(tw_frame, text="報價訂閱", padding=10)
        subscribe_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(subscribe_frame, text="股票代號:").grid(row=0, column=0, sticky='w', padx=5, pady=3)
        self.entry_quote_stock = tk.Entry(subscribe_frame, width=12)
        self.entry_quote_stock.grid(row=0, column=1, padx=5, pady=3)
        
        btn_frame = tk.Frame(subscribe_frame)
        btn_frame.grid(row=0, column=2, columnspan=2, padx=10)
        
        tk.Button(btn_frame, text="連線報價", command=self.connect_quote, bg='#4CAF50', fg='white').pack(side='left', padx=5)
        tk.Button(btn_frame, text="訂閱報價", command=self.subscribe_quote).pack(side='left', padx=5)
        tk.Button(btn_frame, text="取消訂閱", command=self.unsubscribe_quote).pack(side='left', padx=5)
        tk.Button(btn_frame, text="訂閱逐筆", command=self.subscribe_tick).pack(side='left', padx=5)
        
        # 報價顯示區域
        quote_display = ttk.LabelFrame(tw_frame, text="即時報價顯示", padding=10)
        quote_display.pack(fill='both', expand=True, padx=10, pady=5)
        
        # 建立報價顯示表格
        columns = ('代號', '名稱', '成交價', '漲跌', '漲跌幅', '成交量', '買價', '賣價', '時間')
        self.quote_tree = ttk.Treeview(quote_display, columns=columns, show='headings', height=10)
        
        for col in columns:
            self.quote_tree.heading(col, text=col)
            self.quote_tree.column(col, width=80)
            
        scrollbar_quote = ttk.Scrollbar(quote_display, orient='vertical', command=self.quote_tree.yview)
        self.quote_tree.configure(yscrollcommand=scrollbar_quote.set)
        
        self.quote_tree.pack(side='left', fill='both', expand=True)
        scrollbar_quote.pack(side='right', fill='y')
        
    def create_os_quote_tab(self, parent):
        """建立海外股票報價分頁"""
        os_frame = ttk.Frame(parent)
        parent.add(os_frame, text='海外股票報價')
        
        # 實作海外股票報價功能
        subscribe_frame = ttk.LabelFrame(os_frame, text="海外股票報價訂閱", padding=10)
        subscribe_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(subscribe_frame, text="股票代號:").grid(row=0, column=0, sticky='w', padx=5, pady=3)
        self.entry_os_quote_stock = tk.Entry(subscribe_frame, width=12)
        self.entry_os_quote_stock.grid(row=0, column=1, padx=5, pady=3)
        
        tk.Button(subscribe_frame, text="訂閱海外報價", command=self.subscribe_os_quote).grid(row=0, column=2, padx=5, pady=3)
        
    def create_oo_quote_tab(self, parent):
        """建立海外選擇權報價分頁"""
        oo_frame = ttk.Frame(parent)
        parent.add(oo_frame, text='海外選擇權報價')
        
        # 實作海外選擇權報價功能
        subscribe_frame = ttk.LabelFrame(oo_frame, text="海外選擇權報價訂閱", padding=10)
        subscribe_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(subscribe_frame, text="選擇權代號:").grid(row=0, column=0, sticky='w', padx=5, pady=3)
        self.entry_oo_quote_stock = tk.Entry(subscribe_frame, width=12)
        self.entry_oo_quote_stock.grid(row=0, column=1, padx=5, pady=3)
        
        tk.Button(subscribe_frame, text="訂閱選擇權報價", command=self.subscribe_oo_quote).grid(row=0, column=2, padx=5, pady=3)
        
    def create_market_stats_tab(self, parent):
        """建立市場統計分頁"""
        stats_frame = ttk.Frame(parent)
        parent.add(stats_frame, text='市場統計')
        
        # 上市統計
        tse_frame = ttk.LabelFrame(stats_frame, text="上市統計", padding=10)
        tse_frame.pack(fill='x', padx=10, pady=5)
        
        # 建立統計標籤
        stats_labels = [
            ("總成交金額(億):", "labelnTotv"),
            ("總成交筆數:", "labelnTots"), 
            ("總成交量:", "labelnTotc"),
            ("上漲家數:", "labelnUp"),
            ("下跌家數:", "labelnDown"),
            ("平盤家數:", "labelnNoChange")
        ]
        
        for i, (label_text, var_name) in enumerate(stats_labels):
            row = i // 3
            col = (i % 3) * 2
            ttk.Label(tse_frame, text=label_text).grid(row=row, column=col, sticky='w', padx=5, pady=3)
            label = ttk.Label(tse_frame, text="0", font=('Arial', 10, 'bold'))
            label.grid(row=row, column=col+1, sticky='w', padx=5, pady=3)
            setattr(self, var_name, label)
            
    def create_reply_tab(self):
        """建立回報分頁"""
        reply_frame = ttk.Frame(self.notebook)
        self.notebook.add(reply_frame, text='回報')
        
        # 回報控制區域
        control_frame = ttk.LabelFrame(reply_frame, text="回報控制", padding=10)
        control_frame.pack(fill='x', padx=10, pady=5)
        
        btn_frame = tk.Frame(control_frame)
        btn_frame.pack()
        
        tk.Button(btn_frame, text="連線回報", command=self.connect_reply).pack(side='left', padx=5)
        tk.Button(btn_frame, text="斷線回報", command=self.disconnect_reply).pack(side='left', padx=5)
        tk.Button(btn_frame, text="清除前日回報", command=self.clear_reply).pack(side='left', padx=5)
        
        # 回報顯示區域
        display_frame = ttk.LabelFrame(reply_frame, text="回報顯示", padding=10)
        display_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # 建立回報分類 Notebook
        reply_notebook = ttk.Notebook(display_frame)
        reply_notebook.pack(fill='both', expand=True)
        
        # 各市場回報分頁
        markets = [("證券", "TS"), ("期貨", "TF"), ("選擇權", "TO"), 
                  ("海外股票", "OS"), ("海外期貨", "OF"), ("全部訊息", "ALL")]
        
        self.reply_text_widgets = {}
        for market_name, market_code in markets:
            market_frame = ttk.Frame(reply_notebook)
            reply_notebook.add(market_frame, text=market_name)
            
            text_widget = tk.Text(market_frame, height=15, wrap='none')
            scrollbar_v = ttk.Scrollbar(market_frame, orient='vertical', command=text_widget.yview)
            scrollbar_h = ttk.Scrollbar(market_frame, orient='horizontal', command=text_widget.xview)
            text_widget.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)
            
            text_widget.pack(side='left', fill='both', expand=True)
            scrollbar_v.pack(side='right', fill='y')
            scrollbar_h.pack(side='bottom', fill='x')
            
            self.reply_text_widgets[market_code] = text_widget
            
    def create_status_area(self):
        """建立狀態區域和訊息顯示"""
        # 訊息顯示區域
        message_frame = ttk.LabelFrame(self.root, text="系統訊息", padding=5)
        message_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 建立分頁顯示不同類型訊息
        message_notebook = ttk.Notebook(message_frame)
        message_notebook.pack(fill='both', expand=True)
        
        # 方法訊息
        method_frame = ttk.Frame(message_notebook)
        message_notebook.add(method_frame, text='方法訊息')
        
        self.method_text = tk.Text(method_frame, height=8, wrap='none')
        scrollbar_method = ttk.Scrollbar(method_frame, orient='vertical', command=self.method_text.yview)
        self.method_text.configure(yscrollcommand=scrollbar_method.set)
        
        self.method_text.pack(side='left', fill='both', expand=True)
        scrollbar_method.pack(side='right', fill='y')
        
        # 事件訊息
        event_frame = ttk.Frame(message_notebook)
        message_notebook.add(event_frame, text='事件訊息')
        
        self.event_text = tk.Text(event_frame, height=8, wrap='none')
        scrollbar_event = ttk.Scrollbar(event_frame, orient='vertical', command=self.event_text.yview)
        self.event_text.configure(yscrollcommand=scrollbar_event.set)
        
        self.event_text.pack(side='left', fill='both', expand=True)
        scrollbar_event.pack(side='right', fill='y')
        
        # 設定全域變數供事件處理使用
        global richTextBoxMethodMessage, richTextBoxMessage
        richTextBoxMethodMessage = self.method_text
        richTextBoxMessage = self.event_text
        
    def setup_event_handlers(self):
        """設定事件處理器"""
        self.setup_reply_events()
        self.setup_center_events()
        self.setup_order_events()
        self.setup_quote_events()
        
    def setup_reply_events(self):
        """設定 Reply 事件處理"""
        class SKReplyLibEvent:
            def OnReplyMessage(self, bstrUserID, bstrMessages):
                msg = f"【OnReplyMessage】{bstrUserID}_{bstrMessages}"
                self.parent.add_event_message(msg)
                return -1
                
            def OnNewData(self, bstrUserID, bstrData):
                msg = f"【OnNewData】{bstrUserID}_{bstrData}"
                self.parent.add_event_message(msg)
                self.parent.process_reply_data(bstrData)
                
            def OnComplete(self, bstrUserID):
                msg = f"【OnComplete】{bstrUserID}_回報連線&資料正常"
                self.parent.add_event_message(msg)
                self.parent.connection_status.config(fg='green')
                
        self.reply_event = SKReplyLibEvent()
        self.reply_event.parent = self
        self.reply_event_handler = comtypes.client.GetEvents(m_pSKReply, self.reply_event)
        
    def setup_center_events(self):
        """設定 Center 事件處理"""
        class SKCenterLibEvent:
            def OnTimer(self, nTime):
                # 可以選擇是否顯示 Timer 事件
                pass
                
            def OnShowAgreement(self, bstrData):
                msg = f"【OnShowAgreement】{bstrData}"
                self.parent.add_event_message(msg)
                
        self.center_event = SKCenterLibEvent()
        self.center_event.parent = self
        self.center_event_handler = comtypes.client.GetEvents(m_pSKCenter, self.center_event)
        
    def setup_order_events(self):
        """設定 Order 事件處理"""
        class SKOrderLibEvent:
            def OnAccount(self, bstrLogInID, bstrAccountData):
                msg = f"【OnAccount】{bstrLogInID}_{bstrAccountData}"
                self.parent.add_event_message(msg)
                self.parent.process_account_data(bstrLogInID, bstrAccountData)
                
            def OnAsyncOrder(self, nThreadID, nCode, bstrMessage):
                msg = f"【OnAsyncOrder】ThreadID:{nThreadID} Code:{nCode} {bstrMessage}"
                self.parent.add_event_message(msg)
                
            def OnProxyOrder(self, nStampID, nCode, bstrMessage):
                msg = f"【OnProxyOrder】StampID:{nStampID} Code:{nCode} {bstrMessage}"
                self.parent.add_event_message(msg)
                
        self.order_event = SKOrderLibEvent()
        self.order_event.parent = self
        self.order_event_handler = comtypes.client.GetEvents(m_pSKOrder, self.order_event)
        
    def setup_quote_events(self):
        """設定 Quote 事件處理"""
        class SKQuoteLibEvent:
            def OnConnection(self, nKind, nCode):
                kind_msg = "登入" if nKind == 3003 else "登出" if nKind == 3021 else f"Kind:{nKind}"
                msg = f"【Quote連線】{kind_msg} Code:{nCode}"
                self.parent.add_event_message(msg)
                
                # 更新連線狀態
                if nKind == 3003 and nCode == 0:  # 登入成功
                    self.parent.connection_status.config(fg='green')
                elif nKind == 3021:  # 登出
                    self.parent.connection_status.config(fg='red')
                    
            def OnNotifyQuote(self, sStockidx, sPtr):
                # 處理即時報價更新
                self.parent.process_quote_data(sStockidx, sPtr)
                
            def OnNotifyHistoryTicks(self, sStockidx, nPtr, lDate, lTimehms, nBidAskFlag, nClose, nQty):
                # 處理歷史逐筆資料
                msg = f"【歷史逐筆】代號:{sStockidx} 時間:{lTimehms} 價格:{nClose} 量:{nQty}"
                self.parent.add_event_message(msg)
                
            def OnNotifyTicks(self, sStockidx, nPtr, lDate, lTimehms, nBidAskFlag, nClose, nQty):
                # 處理即時逐筆資料
                msg = f"【即時逐筆】代號:{sStockidx} 時間:{lTimehms} 價格:{nClose} 量:{nQty}"
                self.parent.add_event_message(msg)
                
            def OnNotifyBest5(self, sStockidx, nPtr, nBestBid1, nBestBidQty1, nBestBid2, nBestBidQty2, nBestBid3, nBestBidQty3, nBestBid4, nBestBidQty4, nBestBid5, nBestBidQty5, nBestAsk1, nBestAskQty1, nBestAsk2, nBestAskQty2, nBestAsk3, nBestAskQty3, nBestAsk4, nBestAskQty4, nBestAsk5, nBestAskQty5):
                # 處理最佳五檔資料
                msg = f"【最佳五檔】代號:{sStockidx} 買1:{nBestBid1}({nBestBidQty1}) 賣1:{nBestAsk1}({nBestAskQty1})"
                self.parent.add_event_message(msg)
                
            def OnNotifyMarketTot(self, sMarketNo, sPtr, nTime, nTotv, nTots, nTotc):
                # 處理市場統計資料
                if hasattr(self.parent, 'labelnTotv'):
                    if sMarketNo == 0:  # 上市
                        self.parent.labelnTotv.config(text=str(nTotv / 100.00))
                        self.parent.labelnTots.config(text=str(nTots))
                        self.parent.labelnTotc.config(text=str(nTotc))
                        
            def OnNotifyMarketHighLowNoWarrant(self, sMarketNo, sPtr, nTime, nUp, nDown, nHigh, nLow, nNoChange, nUpNoW, nDownNoW, nHighNoW, nLowNoW, nNoChangeNoW):
                # 處理漲跌家數統計
                if hasattr(self.parent, 'labelnUp'):
                    if sMarketNo == 0:  # 上市
                        self.parent.labelnUp.config(text=str(nUp))
                        self.parent.labelnDown.config(text=str(nDown))
                        self.parent.labelnNoChange.config(text=str(nNoChange))
                
        self.quote_event = SKQuoteLibEvent()
        self.quote_event.parent = self
        self.quote_event_handler = comtypes.client.GetEvents(m_pSKQuote, self.quote_event)
        
    # ========== 輔助方法 ==========
    def add_method_message(self, message):
        """添加方法訊息"""
        self.method_text.insert('end', message + "\n")
        self.method_text.see('end')
        
    def add_event_message(self, message):
        """添加事件訊息"""
        self.event_text.insert('end', message + "\n")
        self.event_text.see('end')
        
    def get_api_version(self):
        """取得 API 版本"""
        try:
            return m_pSKCenter.SKCenterLib_GetSKAPIVersionAndBit("xxxxxxxxxx")
        except:
            return "2.13.55_x64"
            
    def process_account_data(self, login_id, account_data):
        """處理帳號資料"""
        global dictUserID
        
        values = account_data.split(',')
        if len(values) >= 4:
            account = values[1] + values[3]  # broker ID + 帳號
            
            if login_id in dictUserID:
                if account not in dictUserID[login_id]:
                    dictUserID[login_id].append(account)
            else:
                dictUserID[login_id] = [account]
                
        # 更新 ComboBox
        self.combo_userid['values'] = list(dictUserID.keys())
        
    def process_reply_data(self, data):
        """處理回報資料"""
        try:
            values = data.split(',')
            if len(values) > 1:
                market_type = values[1]
                
                # 根據市場類型分流顯示
                if market_type in self.reply_text_widgets:
                    self.reply_text_widgets[market_type].insert('end', data + "\n")
                    self.reply_text_widgets[market_type].see('end')
                    
                # 也顯示在全部訊息中
                if "ALL" in self.reply_text_widgets:
                    self.reply_text_widgets["ALL"].insert('end', data + "\n")
                    self.reply_text_widgets["ALL"].see('end')
        except:
            pass
            
    def process_quote_data(self, stock_idx, ptr):
        """處理報價資料"""
        try:
            # 取得股票基本資料
            stock_info = m_pSKQuote.SKQuoteLib_GetStockByIndex(stock_idx)
            
            if stock_info:
                # 解析股票資料
                values = stock_info.split(',')
                if len(values) >= 10:
                    stock_no = values[0]      # 股票代號
                    stock_name = values[1]    # 股票名稱
                    close_price = values[2]   # 成交價
                    change = values[3]        # 漲跌
                    change_pct = values[4]    # 漲跌幅
                    volume = values[5]        # 成交量
                    bid_price = values[6]     # 買價
                    ask_price = values[7]     # 賣價
                    time_str = values[8]      # 時間
                    
                    # 更新報價表格
                    for item in self.quote_tree.get_children():
                        item_values = self.quote_tree.item(item)['values']
                        if item_values and item_values[0] == stock_no:
                            # 更新該股票的報價資料
                            new_values = (stock_no, stock_name, close_price, change, 
                                        change_pct, volume, bid_price, ask_price, time_str)
                            self.quote_tree.item(item, values=new_values)
                            break
                    
                    # 顯示報價更新訊息
                    msg = f"【報價更新】{stock_no} {stock_name} 價格:{close_price} 漲跌:{change}"
                    self.add_event_message(msg)
                    
        except Exception as e:
            msg = f"【報價處理錯誤】{str(e)}"
            self.add_event_message(msg)
        
    # ========== 事件處理方法 ==========
    def on_authority_change(self, event):
        """連線環境變更"""
        authority_text = self.combo_authority.get()
        authority_map = {
            "正式環境": 0,
            "正式環境SGX": 1,
            "測試環境": 2,
            "測試環境SGX": 3
        }
        
        if authority_text in authority_map:
            nCode = m_pSKCenter.SKCenterLib_SetAuthority(authority_map[authority_text])
            msg = f"【SKCenterLib_SetAuthority】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
            self.add_method_message(msg)
            
    def on_ap_check(self):
        """AP身分檢查框變更"""
        if self.var_ap.get() == 1:
            self.label_certid.grid(row=4, column=0, sticky='w', padx=5, pady=5)
            self.entry_certid.grid(row=4, column=1, padx=5, pady=5)
            self.btn_generate_cert.pack(side='left', padx=5)
        else:
            self.label_certid.grid_remove()
            self.entry_certid.grid_remove()
            self.btn_generate_cert.pack_forget()
            
    def on_userid_change(self, event):
        """使用者變更"""
        selected_user = self.combo_userid.get()
        if selected_user in dictUserID:
            self.combo_account['values'] = dictUserID[selected_user]
            
            # 初始化下單物件
            nCode = m_pSKOrder.SKOrderLib_Initialize()
            msg = f"【SKOrderLib_Initialize】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
            self.add_method_message(msg)
            
            # 取得交易帳號
            nCode = m_pSKOrder.GetUserAccount()
            msg = f"【GetUserAccount】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
            self.add_method_message(msg)
            
    # ========== 登入相關方法 ==========
    def login(self):
        """登入"""
        user_id = self.entry_userid.get()
        password = self.entry_password.get()
        
        if not user_id or not password:
            messagebox.showwarning("警告", "請輸入使用者帳號和密碼")
            return
            
        nCode = m_pSKCenter.SKCenterLib_Login(user_id, password)
        msg = f"【SKCenterLib_Login】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        self.add_method_message(msg)
        
        if nCode == 0:
            self.connection_status.config(fg='orange')
            messagebox.showinfo("成功", "登入成功！")
        else:
            messagebox.showerror("錯誤", f"登入失敗: {msg}")
            
    def generate_cert(self):
        """產生雙因子驗證憑證"""
        user_id = self.entry_userid.get()
        cert_id = self.entry_certid.get()
        
        if not user_id or not cert_id:
            messagebox.showwarning("警告", "請輸入使用者帳號和憑證ID")
            return
            
        nCode = m_pSKCenter.SKCenterLib_GenerateKeyCert(user_id, cert_id)
        msg = f"【SKCenterLib_GenerateKeyCert】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        self.add_method_message(msg)
        
    def init_proxy(self):
        """初始化代理伺服器"""
        user_id = self.combo_userid.get()
        if not user_id:
            messagebox.showwarning("警告", "請先選擇使用者")
            return
            
        nCode = m_pSKOrder.SKOrderLib_InitialProxyByID(user_id)
        msg = f"【SKOrderLib_InitialProxyByID】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        self.add_method_message(msg)
        
    def disconnect_proxy(self):
        """斷線代理伺服器"""
        user_id = self.combo_userid.get()
        if not user_id:
            messagebox.showwarning("警告", "請先選擇使用者")
            return
            
        nCode = m_pSKOrder.ProxyDisconnectByID(user_id)
        msg = f"【ProxyDisconnectByID】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        self.add_method_message(msg)
        
    def reconnect_proxy(self):
        """重連代理伺服器"""
        user_id = self.combo_userid.get()
        if not user_id:
            messagebox.showwarning("警告", "請先選擇使用者")
            return
            
        nCode = m_pSKOrder.ProxyReconnectByID(user_id)
        msg = f"【ProxyReconnectByID】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        self.add_method_message(msg)
        
    def add_sgx_socket(self):
        """建立SGX專線"""
        user_id = self.combo_userid.get()
        account = self.combo_account.get()
        
        if not user_id or not account:
            messagebox.showwarning("警告", "請先選擇使用者和帳號")
            return
            
        nCode = m_pSKOrder.AddSGXAPIOrderSocket(user_id, account)
        msg = f"【AddSGXAPIOrderSocket】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        self.add_method_message(msg)
        
    # ========== 下單相關方法 ==========
    def send_ts_order(self):
        """送出台股委託"""
        # 檢查必要欄位
        user_id = self.combo_userid.get()
        account = self.combo_account.get()
        stock = self.entry_ts_stock.get()
        price = self.entry_ts_price.get()
        qty = self.entry_ts_qty.get()
        
        if not all([user_id, account, stock, price, qty]):
            messagebox.showwarning("警告", "請填寫完整的委託資料")
            return
            
        try:
            # 取得下單參數
            buysell_text = self.combo_ts_buysell.get()
            priceflag_text = self.combo_ts_priceflag.get()  
            tradetype_text = self.combo_ts_tradetype.get()
            period_text = self.combo_ts_period.get()
            
            # 轉換參數
            buysell = 0 if "買進" in buysell_text else 1
            
            # 價格旗標轉換
            priceflag_map = {"0:限價": 0, "1:市價": 1, "2:限價(LMT)": 2}
            priceflag = priceflag_map.get(priceflag_text, 0)
            
            # 交易別轉換
            tradetype_map = {"0:現股": 0, "3:融資": 3, "4:融券": 4, "8:無券賣出": 8}
            tradetype = tradetype_map.get(tradetype_text, 0)
            
            # 倉別轉換  
            period_map = {"0:ROD": 0, "3:IOC": 3, "4:FOK": 4}
            period = period_map.get(period_text, 0)
            
            # 建立委託物件
            order_obj = comtypes.client.CreateObject(sk.SKSTOCKORDER)
            order_obj.bstrFullAccount = account
            order_obj.bstrStockNo = stock
            order_obj.nOrderType = buysell  # 0:買進 1:賣出
            order_obj.bstrPrice = price
            order_obj.nQty = int(qty)
            order_obj.nPriceFlag = priceflag  # 價格旗標
            order_obj.nTradeType = tradetype  # 交易別
            order_obj.nPeriod = period  # 倉別
            order_obj.bstrOrderNo = ""  # 委託書號(新單留空)
            
            # 送出委託
            nCode = m_pSKOrder.SendStockOrder(user_id, bAsyncOrder, order_obj)
            
            if nCode == 0:
                msg = f"【台股下單成功】代號:{stock} 價格:{price} 數量:{qty} 買賣:{buysell_text}"
                self.add_method_message(msg)
                messagebox.showinfo("成功", "台股委託已送出")
                
                # 清空欄位
                self.entry_ts_stock.delete(0, 'end')
                self.entry_ts_price.delete(0, 'end') 
                self.entry_ts_qty.delete(0, 'end')
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【台股下單失敗】{error_msg}"
                self.add_method_message(msg)
                messagebox.showerror("錯誤", f"下單失敗: {error_msg}")
                
        except Exception as e:
            msg = f"【台股下單錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"下單發生錯誤: {str(e)}")
        
    def send_ts_async_order(self):
        """送出台股非同步委託"""
        global bAsyncOrder
        
        # 暫時設為非同步模式
        original_async = bAsyncOrder
        bAsyncOrder = True
        
        try:
            # 執行下單邏輯
            self.send_ts_order()
        finally:
            # 恢復原來的同步設定
            bAsyncOrder = original_async
        
    def send_tf_order(self):
        """送出期貨委託"""
        # 檢查必要欄位
        user_id = self.combo_userid.get()
        account = self.combo_account.get()
        stock = self.entry_tf_stock.get()
        price = self.entry_tf_price.get()
        qty = self.entry_tf_qty.get()
        
        if not all([user_id, account, stock, price, qty]):
            messagebox.showwarning("警告", "請填寫完整的期貨委託資料")
            return
            
        try:
            # 取得下單參數
            buysell_text = self.combo_tf_buysell.get()
            newclose_text = self.combo_tf_newclose.get()
            daytrade_text = self.combo_tf_daytrade.get()
            
            # 轉換參數
            buysell = 0 if "買進" in buysell_text else 1
            
            # 新平倉轉換
            newclose_map = {"0:新倉": 0, "1:平倉": 1, "2:自動": 2}
            newclose = newclose_map.get(newclose_text, 0)
            
            # 委託條件轉換
            daytrade_map = {"0:ROD": 0, "3:IOC": 3, "4:FOK": 4}
            daytrade = daytrade_map.get(daytrade_text, 0)
            
            # 建立期貨委託物件
            order_obj = comtypes.client.CreateObject(sk.SKFUTUREORDER)
            order_obj.bstrFullAccount = account
            order_obj.bstrStockNo = stock
            order_obj.nOrderType = buysell  # 0:買進 1:賣出
            order_obj.bstrPrice = price
            order_obj.nQty = int(qty)
            order_obj.nNewClose = newclose  # 新平倉
            order_obj.nDayTrade = daytrade  # 委託條件
            order_obj.nReserved = 0  # 預約
            order_obj.bstrOrderNo = ""  # 委託書號(新單留空)
            
            # 送出期貨委託
            nCode = m_pSKOrder.SendFutureOrder(user_id, bAsyncOrder, order_obj)
            
            if nCode == 0:
                msg = f"【期貨下單成功】代號:{stock} 價格:{price} 數量:{qty} 買賣:{buysell_text} 新平倉:{newclose_text}"
                self.add_method_message(msg)
                messagebox.showinfo("成功", "期貨委託已送出")
                
                # 清空欄位
                self.entry_tf_stock.delete(0, 'end')
                self.entry_tf_price.delete(0, 'end')
                self.entry_tf_qty.delete(0, 'end')
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【期貨下單失敗】{error_msg}"
                self.add_method_message(msg)
                messagebox.showerror("錯誤", f"期貨下單失敗: {error_msg}")
                
        except Exception as e:
            msg = f"【期貨下單錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"期貨下單發生錯誤: {str(e)}")
        
    def send_tf_duplex_order(self):
        """送出複式單委託"""
        # 檢查必要欄位
        user_id = self.combo_userid.get()
        account = self.combo_account.get()
        stock = self.entry_tf_stock.get()
        price = self.entry_tf_price.get()
        qty = self.entry_tf_qty.get()
        
        if not all([user_id, account, stock, price, qty]):
            messagebox.showwarning("警告", "請填寫完整的複式單委託資料")
            return
            
        try:
            # 建立複式單物件
            duplex_obj = comtypes.client.CreateObject(sk.SKFUTUREORDER)
            duplex_obj.bstrFullAccount = account
            duplex_obj.bstrStockNo = stock
            duplex_obj.bstrPrice = price
            duplex_obj.nQty = int(qty)
            
            # 送出複式單
            nCode = m_pSKOrder.SendDuplexOrder(user_id, duplex_obj, duplex_obj)
            
            if nCode == 0:
                msg = f"【複式單下單成功】代號:{stock} 價格:{price} 數量:{qty}"
                self.add_method_message(msg)
                messagebox.showinfo("成功", "複式單委託已送出")
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【複式單下單失敗】{error_msg}"
                self.add_method_message(msg)
                messagebox.showerror("錯誤", f"複式單下單失敗: {error_msg}")
                
        except Exception as e:
            msg = f"【複式單下單錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"複式單下單發生錯誤: {str(e)}")
        
    def send_os_order(self):
        """送出海外股票委託"""
        user_id = self.combo_userid.get()
        account = self.combo_account.get()
        stock = self.entry_os_stock.get()
        
        if not all([user_id, account, stock]):
            messagebox.showwarning("警告", "請填寫完整的海外股票委託資料")
            return
            
        try:
            # 建立海外股票委託物件
            order_obj = comtypes.client.CreateObject(sk.SKOVERSEAORDER)
            order_obj.bstrFullAccount = account
            order_obj.bstrStockNo = stock
            order_obj.bstrOrderType = "B"  # B:買進 S:賣出 (預設買進)
            order_obj.bstrPrice = "100.00"  # 預設價格
            order_obj.nQty = 100  # 預設數量
            order_obj.nTradeType = 0  # 交易別
            order_obj.nPriceFlag = 0  # 價格旗標
            order_obj.bstrOrderNo = ""  # 委託書號(新單留空)
            
            # 送出海外股票委託
            nCode = m_pSKOrder.SendOverSeaOrder(user_id, bAsyncOrder, order_obj)
            
            if nCode == 0:
                msg = f"【海外股票下單成功】代號:{stock} 帳號:{account}"
                self.add_method_message(msg)
                messagebox.showinfo("成功", "海外股票委託已送出")
                
                # 清空欄位
                self.entry_os_stock.delete(0, 'end')
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【海外股票下單失敗】{error_msg}"
                self.add_method_message(msg)
                messagebox.showerror("錯誤", f"海外股票下單失敗: {error_msg}")
                
        except Exception as e:
            msg = f"【海外股票下單錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"海外股票下單發生錯誤: {str(e)}")
        
    def send_of_order(self):
        """送出海外期貨委託"""
        user_id = self.combo_userid.get()
        account = self.combo_account.get()
        stock = self.entry_of_stock.get()
        
        if not all([user_id, account, stock]):
            messagebox.showwarning("警告", "請填寫完整的海外期貨委託資料")
            return
            
        try:
            # 建立海外期貨委託物件
            order_obj = comtypes.client.CreateObject(sk.SKOVERSEAFUTUREORDER)
            order_obj.bstrFullAccount = account
            order_obj.bstrStockNo = stock
            order_obj.bstrOrderType = "B"  # B:買進 S:賣出 (預設買進)
            order_obj.bstrPrice = "100.00"  # 預設價格
            order_obj.nQty = 1  # 預設數量
            order_obj.bstrNewClose = "0"  # 新平倉 (0:新倉)
            order_obj.bstrPeriod = "R"  # 委託條件 (R:ROD)
            order_obj.bstrOrderNo = ""  # 委託書號(新單留空)
            
            # 送出海外期貨委託
            nCode = m_pSKOrder.SendOverSeaFutureOrder(user_id, bAsyncOrder, order_obj)
            
            if nCode == 0:
                msg = f"【海外期貨下單成功】代號:{stock} 帳號:{account}"
                self.add_method_message(msg)
                messagebox.showinfo("成功", "海外期貨委託已送出")
                
                # 清空欄位
                self.entry_of_stock.delete(0, 'end')
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【海外期貨下單失敗】{error_msg}"
                self.add_method_message(msg)
                messagebox.showerror("錯誤", f"海外期貨下單失敗: {error_msg}")
                
        except Exception as e:
            msg = f"【海外期貨下單錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"海外期貨下單發生錯誤: {str(e)}")
        
    # ========== 查詢相關方法 ==========
    def query_order_report(self):
        """查詢委託回報"""
        user_id = self.combo_userid.get()
        account = self.combo_account.get()
        
        if not user_id or not account:
            messagebox.showwarning("警告", "請先選擇使用者和帳號")
            return
            
        # 取得查詢格式
        report_type = self.combo_order_report.get()
        format_map = {
            "1:全部": 1, "2:有效": 2, "3:可消": 3, "4:已消": 4, 
            "5:已成": 5, "6:失敗": 6, "7:合併同價格": 7, "8:合併同商品": 8, "9:預約": 9
        }
        
        format_num = format_map.get(report_type, 1)
        
        try:
            result = m_pSKOrder.GetOrderReport(user_id, account, format_num)
            msg = f"【GetOrderReport】{result}"
            self.add_event_message(msg)
        except Exception as e:
            self.add_method_message(f"【GetOrderReport錯誤】{e}")
            
    def query_fulfill_report(self):
        """查詢成交回報"""
        user_id = self.combo_userid.get()
        account = self.combo_account.get()
        
        if not user_id or not account:
            messagebox.showwarning("警告", "請先選擇使用者和帳號")
            return
            
        # 取得查詢格式
        report_type = self.combo_fulfill_report.get()
        format_map = {
            "1:全部": 1, "2:合併同書號": 2, "3:合併同價格": 3, 
            "4:合併同商品": 4, "5:T+1成交回報": 5
        }
        
        format_num = format_map.get(report_type, 1)
        
        try:
            result = m_pSKOrder.GetFulfillReport(user_id, account, format_num)
            msg = f"【GetFulfillReport】{result}"
            self.add_event_message(msg)
        except Exception as e:
            self.add_method_message(f"【GetFulfillReport錯誤】{e}")
            
    def query_balance(self):
        """查詢即時庫存"""
        user_id = self.combo_userid.get()
        account = self.combo_account.get()
        
        if not user_id or not account:
            messagebox.showwarning("警告", "請先選擇使用者和帳號")
            return
            
        try:
            nCode = m_pSKOrder.GetRealBalanceReport(user_id, account)
            
            if nCode == 0:
                msg = f"【即時庫存查詢】已送出查詢請求，請等候回報"
                self.add_method_message(msg)
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【即時庫存查詢失敗】{error_msg}"
                self.add_method_message(msg)
                messagebox.showerror("錯誤", f"即時庫存查詢失敗: {error_msg}")
                
        except Exception as e:
            msg = f"【即時庫存查詢錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"即時庫存查詢發生錯誤: {str(e)}")
        
    def query_balance_query(self):
        """查詢集保庫存"""
        user_id = self.combo_userid.get()
        account = self.combo_account.get()
        
        if not user_id or not account:
            messagebox.showwarning("警告", "請先選擇使用者和帳號")
            return
            
        try:
            nCode = m_pSKOrder.GetBalanceQuery(user_id, account)
            
            if nCode == 0:
                msg = f"【集保庫存查詢】已送出查詢請求，請等候回報"
                self.add_method_message(msg)
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【集保庫存查詢失敗】{error_msg}"
                self.add_method_message(msg)
                messagebox.showerror("錯誤", f"集保庫存查詢失敗: {error_msg}")
                
        except Exception as e:
            msg = f"【集保庫存查詢錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"集保庫存查詢發生錯誤: {str(e)}")
        
    def query_margin_limit(self):
        """查詢資券配額"""
        user_id = self.combo_userid.get()
        account = self.combo_account.get()
        
        if not user_id or not account:
            messagebox.showwarning("警告", "請先選擇使用者和帳號")
            return
            
        try:
            nCode = m_pSKOrder.GetMarginPurchaseAmountLimit(user_id, account)
            
            if nCode == 0:
                msg = f"【資券配額查詢】已送出查詢請求，請等候回報"
                self.add_method_message(msg)
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【資券配額查詢失敗】{error_msg}"
                self.add_method_message(msg)
                messagebox.showerror("錯誤", f"資券配額查詢失敗: {error_msg}")
                
        except Exception as e:
            msg = f"【資券配額查詢錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"資券配額查詢發生錯誤: {str(e)}")
        
    def query_profit_loss(self):
        """查詢損益"""
        user_id = self.combo_userid.get()
        account = self.combo_account.get()
        
        if not user_id or not account:
            messagebox.showwarning("警告", "請先選擇使用者和帳號")
            return
            
        try:
            nCode = m_pSKOrder.GetProfitLossGWReport(user_id, account)
            
            if nCode == 0:
                msg = f"【損益查詢】已送出查詢請求，請等候回報"
                self.add_method_message(msg)
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【損益查詢失敗】{error_msg}"
                self.add_method_message(msg)
                messagebox.showerror("錯誤", f"損益查詢失敗: {error_msg}")
                
        except Exception as e:
            msg = f"【損益查詢錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"損益查詢發生錯誤: {str(e)}")
        
    # ========== 報價相關方法 ==========
    def connect_quote(self):
        """建立報價連線"""
        user_id = self.combo_userid.get()
        if not user_id:
            messagebox.showwarning("警告", "請先登入並選擇使用者")
            return
            
        try:
            # 1. 先進入監控模式
            nCode = m_pSKQuote.SKQuoteLib_EnterMonitor()
            if nCode == 0:
                msg = f"【進入監控模式成功】"
                self.add_method_message(msg)
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【進入監控模式失敗】{error_msg}"
                self.add_method_message(msg)
                
            # 2. 連線報價伺服器
            nCode = m_pSKQuote.SKQuoteLib_ConnectByID(user_id)
            if nCode == 0:
                msg = f"【報價伺服器連線成功】使用者: {user_id}"
                self.add_method_message(msg)
                messagebox.showinfo("成功", "報價伺服器連線成功！現在可以訂閱報價")
                self.connection_status.config(fg='green')
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【報價伺服器連線失敗】{error_msg}"
                self.add_method_message(msg)
                messagebox.showerror("錯誤", f"報價伺服器連線失敗: {error_msg}")
                
        except Exception as e:
            msg = f"【報價連線錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"報價連線發生錯誤: {str(e)}")
    
    def subscribe_quote(self):
        """訂閱報價"""
        stock = self.entry_quote_stock.get()
        if not stock:
            messagebox.showwarning("警告", "請輸入股票代號")
            return
            
        try:
            # 檢查是否已登入
            user_id = self.combo_userid.get()
            if not user_id:
                messagebox.showwarning("警告", "請先登入並選擇使用者")
                return
                
            # 訂閱股票報價 (psPageNo: 頁碼，從0開始)
            psPageNo, nCode = m_pSKQuote.SKQuoteLib_RequestStocks(0, stock)
            
            if nCode == 0:
                msg = f"【報價訂閱成功】股票代號: {stock}"
                self.add_method_message(msg)
                
                # 加入到報價顯示表格
                self.quote_tree.insert('', 'end', values=(stock, '', '', '', '', '', '', '', ''))
                
                # 清空輸入欄
                self.entry_quote_stock.delete(0, 'end')
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【報價訂閱失敗】{error_msg}"
                self.add_method_message(msg)
                messagebox.showerror("錯誤", f"報價訂閱失敗: {error_msg}")
                
        except Exception as e:
            msg = f"【報價訂閱錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"報價訂閱發生錯誤: {str(e)}")
        
    def unsubscribe_quote(self):
        """取消訂閱報價"""
        # 取得選中的項目
        selected_items = self.quote_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "請先選擇要取消的股票")
            return
            
        try:
            for item in selected_items:
                values = self.quote_tree.item(item)['values']
                if values:
                    stock = values[0]  # 股票代號
                    
                    # 取消訂閱 (使用相同的方法，但會取消訂閱)
                    psPageNo, nCode = m_pSKQuote.SKQuoteLib_RequestStocks(0, stock)
                    
                    if nCode == 0:
                        msg = f"【取消訂閱成功】股票代號: {stock}"
                        self.add_method_message(msg)
                        
                        # 從表格中移除
                        self.quote_tree.delete(item)
                    else:
                        error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                        msg = f"【取消訂閱失敗】{stock}: {error_msg}"
                        self.add_method_message(msg)
                        
        except Exception as e:
            msg = f"【取消訂閱錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"取消訂閱發生錯誤: {str(e)}")
        
    def subscribe_tick(self):
        """訂閱逐筆"""
        stock = self.entry_quote_stock.get()
        if not stock:
            messagebox.showwarning("警告", "請輸入股票代號")
            return
            
        try:
            # 檢查是否已登入
            user_id = self.combo_userid.get()
            if not user_id:
                messagebox.showwarning("警告", "請先登入並選擇使用者")
                return
                
            # 訂閱逐筆資料 (psPageNo: 頁碼，從0開始)
            psPageNo, nCode = m_pSKQuote.SKQuoteLib_RequestTicks(0, stock)
            
            if nCode == 0:
                msg = f"【逐筆訂閱成功】股票代號: {stock}"
                self.add_method_message(msg)
                messagebox.showinfo("成功", f"已訂閱 {stock} 逐筆資料")
                
                # 清空輸入欄
                self.entry_quote_stock.delete(0, 'end')
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【逐筆訂閱失敗】{error_msg}"
                self.add_method_message(msg)
                messagebox.showerror("錯誤", f"逐筆訂閱失敗: {error_msg}")
                
        except Exception as e:
            msg = f"【逐筆訂閱錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"逐筆訂閱發生錯誤: {str(e)}")
        
    def subscribe_os_quote(self):
        """訂閱海外股票報價"""
        msg = "【海外股票報價訂閱】功能開發中"
        self.add_method_message(msg)
        
    def subscribe_oo_quote(self):
        """訂閱海外選擇權報價"""
        msg = "【海外選擇權報價訂閱】功能開發中"
        self.add_method_message(msg)
        
    # ========== 回報相關方法 ==========
    def connect_reply(self):
        """連線回報"""
        user_id = self.combo_userid.get()
        if not user_id:
            messagebox.showwarning("警告", "請先選擇使用者")
            return
            
        try:
            nCode = m_pSKReply.SKReplyLib_ConnectByID(user_id)
            
            if nCode == 0:
                msg = f"【回報連線成功】使用者: {user_id}"
                self.add_method_message(msg)
                messagebox.showinfo("成功", "回報伺服器連線成功")
            else:
                error_msg = m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
                msg = f"【回報連線失敗】{error_msg}"
                self.add_method_message(msg)
                messagebox.showerror("錯誤", f"回報連線失敗: {error_msg}")
                
        except Exception as e:
            msg = f"【回報連線錯誤】{str(e)}"
            self.add_method_message(msg)
            messagebox.showerror("錯誤", f"回報連線發生錯誤: {str(e)}")
        
    def disconnect_reply(self):
        """斷線回報"""
        user_id = self.combo_userid.get()
        if not user_id:
            messagebox.showwarning("警告", "請先選擇使用者")
            return
            
        nCode = m_pSKReply.SKReplyLib_CloseByID(user_id)
        msg = f"【SKReplyLib_CloseByID】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        self.add_method_message(msg)
        
    def clear_reply(self):
        """清除前日回報"""
        for text_widget in self.reply_text_widgets.values():
            text_widget.delete(1.0, 'end')
        self.add_method_message("【清除回報】已清除所有回報訊息")
        
    # ========== 工具選單方法 ==========
    def set_log_path(self):
        """設定LOG路徑"""
        folder_path = filedialog.askdirectory(title="選擇LOG存放資料夾")
        if folder_path:
            nCode = m_pSKCenter.SKCenterLib_SetLogPath(folder_path)
            msg = f"【SKCenterLib_SetLogPath】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
            self.add_method_message(msg)
            messagebox.showinfo("成功", f"LOG路徑已設定為: {folder_path}")
            
    def show_api_version(self):
        """顯示API版本資訊"""
        version = self.get_api_version()
        messagebox.showinfo("API版本資訊", f"群益證券API版本: {version}")
        
    def show_last_log(self):
        """顯示最後LOG資訊"""
        try:
            log_info = m_pSKCenter.SKCenterLib_GetLastLogInfo()
            msg = f"【SKCenterLib_GetLastLogInfo】{log_info}"
            self.add_method_message(msg)
        except Exception as e:
            self.add_method_message(f"【取得LOG資訊錯誤】{e}")
            
    def show_about(self):
        """顯示關於對話框"""
        about_text = """群益證券 API 整合測試工具 v1.0
        
整合功能:
• 登入驗證管理
• 多市場下單交易 (台股/期貨/海外)
• 即時報價訂閱
• 交易回報監控
• 查詢功能

基於: 群益證券 API v2.13.55"""
        
        messagebox.showinfo("關於", about_text)
        
    def run(self):
        """啟動應用程式"""
        self.root.mainloop()

# 程式進入點
if __name__ == "__main__":
    try:
        app = CapitalIntegratedTester()
        print("=> 群益證券 API 整合測試工具已啟動")
        app.run()
    except Exception as e:
        print(f"=> 程式啟動失敗: {e}")
        messagebox.showerror("錯誤", f"程式啟動失敗:\n{e}")