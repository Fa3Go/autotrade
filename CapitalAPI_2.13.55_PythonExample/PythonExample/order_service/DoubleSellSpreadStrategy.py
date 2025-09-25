# 雙賣價差單策略模組
import os
import Global
import threading
import time
from datetime import datetime

skC = Global.skC
skO = Global.skO
skR = Global.skR
skQ = Global.skQ
skOSQ = Global.skOSQ
skOOQ = Global.skOOQ

# 第二種讓群益API元件可導入Python code內用的物件宣告
import comtypes.client
import comtypes.gen.SKCOMLib as sk

# 畫視窗用物件
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox

# 載入其他物件
import Config
import MessageControl

# 嘗試載入Quote模組進行價格抓取
try:
    # 先載入Quote專用的Config模組
    import importlib.util
    import sys
    import os
    quote_service_path = os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))), "Quote_Service")
    quote_config_path = os.path.join(quote_service_path, "Config.py")
    spec = importlib.util.spec_from_file_location("QuoteConfig", quote_config_path)
    QuoteConfig = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(QuoteConfig)

    # 暫時保存當前的Config模組
    original_config = None
    if 'Config' in sys.modules:
        original_config = sys.modules['Config']

    # 將Quote的Config注入到sys.modules中
    sys.modules['Config'] = QuoteConfig

    # 載入Quote模組
    sys.path.append(quote_service_path)
    import Quote as QuoteModule

    # 恢復原始的Config模組
    if original_config:
        sys.modules['Config'] = original_config
    elif 'Config' in sys.modules:
        del sys.modules['Config']

    QUOTE_AVAILABLE = True
    print("Quote模組載入成功 - 可抓取實際報價")

except Exception as e:
    QuoteModule = None
    QUOTE_AVAILABLE = False
    print(f"Quote模組載入失敗: {e} - 將使用模擬價格")

# 雙賣價差單策略類別
class DoubleSellSpreadStrategy(Frame):
    def __init__(self, information=None):
        Frame.__init__(self)
        self.__oMsg = MessageControl.MessageControl()
        self.__information = information

        # UI變數
        self.__dOrder = dict(
            listInformation=information,
            boxAccount='',
            # Call選擇權相關
            call_stock_no='',
            call_price='',
            call_qty='',
            # Put選擇權相關
            put_stock_no='',
            put_price='',
            put_qty='',
            # 策略參數
            spread_price='',  # 價差委託價
            trade_type='ROD',  # 委託條件
            new_close='新倉',   # 倉別
            # MIT委託參數
            mit1_enabled=False,  # MIT委託1啟用
            mit1_stock_no='',    # MIT委託1商品代碼
            mit1_trigger_price='',  # MIT委託1觸發價
            mit1_deal_price='',     # MIT委託1委託價
            mit1_qty='',           # MIT委託1數量
            mit1_buy_sell='買進',   # MIT委託1買賣別
            mit2_enabled=False,    # MIT委託2啟用
            mit2_stock_no='',      # MIT委託2商品代碼
            mit2_trigger_price='', # MIT委託2觸發價
            mit2_deal_price='',    # MIT委託2委託價
            mit2_qty='',          # MIT委託2數量
            mit2_buy_sell='賣出'   # MIT委託2買賣別
        )

        self.__CreateWidget()

        # 報價相關
        self.current_call_price = 0.0
        self.current_put_price = 0.0
        self.spread_value = 0.0


        # MIT自動交易相關
        self.mit_auto_trading = {1: False, 2: False}  # MIT1和MIT2的自動交易狀態
        self.mit_position_state = {1: 0, 2: 0}        # MIT1和MIT2的部位狀態
        self.mit_current_price = {1: 0.0, 2: 0.0}    # MIT1和MIT2的當前價格
        self.mit_trade_locks = {1: threading.Lock(), 2: threading.Lock()}  # MIT交易鎖

        # 啟動報價監控
        if QUOTE_AVAILABLE:
            self.__start_quote_monitoring()

    def SetAccount(self, account):
        """設定交易帳號"""
        self.__dOrder['boxAccount'] = account

    def __CreateWidget(self):
        """建立GUI介面"""
        # 主要框架
        main_group = LabelFrame(self, text="雙賣價差策略委託", style="Pink.TLabelframe")
        main_group.grid(column=0, row=0, padx=5, pady=5, sticky='ew')

        # Call選擇權區域
        call_frame = LabelFrame(main_group, text="看漲", style="Pink.TLabelframe")
        call_frame.grid(column=0, row=0, padx=5, pady=5, sticky='w')

        self.__create_option_widgets(call_frame, 'call')

        # Put選擇權區域
        put_frame = LabelFrame(main_group, text="看跌", style="Pink.TLabelframe")
        put_frame.grid(column=1, row=0, padx=5, pady=5, sticky='w')

        self.__create_option_widgets(put_frame, 'put')

        # 策略參數區域
        strategy_frame = LabelFrame(main_group, text="策略參數", style="Pink.TLabelframe")
        strategy_frame.grid(column=0, row=1, columnspan=2, padx=5, pady=5, sticky='ew')

        self.__create_strategy_widgets(strategy_frame)

        # 下單按鈕區域
        button_frame = Frame(main_group, style="Pink.TFrame")
        button_frame.grid(column=0, row=2, columnspan=2, padx=5, pady=5)

        self.__create_order_buttons(button_frame)

        # MIT委託區域
        mit_frame = LabelFrame(main_group, text="期貨MIT委託", style="Pink.TLabelframe")
        mit_frame.grid(column=0, row=3, columnspan=2, padx=5, pady=5, sticky='ew')

        self.__create_mit_widgets(mit_frame)

        # 監控資訊區域
        monitor_frame = LabelFrame(main_group, text="即時監控", style="Pink.TLabelframe")
        monitor_frame.grid(column=0, row=4, columnspan=2, padx=5, pady=5, sticky='ew')

        self.__create_monitor_widgets(monitor_frame)

    def __create_option_widgets(self, parent, option_type):
        """建立選擇權輸入控件"""
        frame = Frame(parent, style="Pink.TFrame")
        frame.grid(column=0, row=0, padx=5, pady=5, sticky='w')

        # 商品代碼
        lbStockNo = Label(frame, style="Pink.TLabel", text="商品代碼")
        lbStockNo.grid(column=0, row=0, pady=3)

        txtStockNo = Entry(frame, width=15)
        txtStockNo.grid(column=0, row=1, padx=5, pady=3, sticky='w')
        setattr(self, f'txt{option_type.title()}StockNo', txtStockNo)

        # 委託價
        lbPrice = Label(frame, style="Pink.TLabel", text="委託價")
        lbPrice.grid(column=1, row=0, pady=3)

        txtPrice = Entry(frame, width=10)
        txtPrice.grid(column=1, row=1, padx=5, pady=3, sticky='w')
        setattr(self, f'txt{option_type.title()}Price', txtPrice)

        # 委託量
        lbQty = Label(frame, style="Pink.TLabel", text="委託量")
        lbQty.grid(column=2, row=0, pady=3)

        txtQty = Entry(frame, width=8)
        txtQty.grid(column=2, row=1, padx=5, pady=3, sticky='w')
        setattr(self, f'txt{option_type.title()}Qty', txtQty)

        # 即時價格顯示
        lbCurrentPrice = Label(frame, style="Pink.TLabel", text="即時價格")
        lbCurrentPrice.grid(column=3, row=0, pady=3)

        lbPriceValue = Label(frame, style="Pink.TLabel", text="0.00", foreground="blue")
        lbPriceValue.grid(column=3, row=1, padx=5, pady=3)
        setattr(self, f'lb{option_type.title()}CurrentPrice', lbPriceValue)

        # 盤別（借用OptionOrder.py的概念）
        lbReserved = Label(frame, style="Pink.TLabel", text="盤別")
        lbReserved.grid(column=0, row=2, pady=3)

        boxReserved = Combobox(frame, width=8, state='readonly')
        boxReserved['values'] = ['盤中', 'T盤預約']
        boxReserved.set('盤中')
        boxReserved.grid(column=0, row=3, padx=5, pady=3, sticky='w')
        setattr(self, f'box{option_type.title()}Reserved', boxReserved)

    def __create_strategy_widgets(self, parent):
        """建立策略參數控件"""
        frame = Frame(parent, style="Pink.TFrame")
        frame.grid(column=0, row=0, padx=5, pady=5, sticky='w')

        # 價差委託價
        lbSpreadPrice = Label(frame, style="Pink.TLabel", text="價差委託價")
        lbSpreadPrice.grid(column=0, row=0, pady=3)

        self.txtSpreadPrice = Entry(frame, width=10)
        self.txtSpreadPrice.grid(column=0, row=1, padx=5, pady=3, sticky='w')

        # 委託條件
        lbTradeType = Label(frame, style="Pink.TLabel", text="委託條件")
        lbTradeType.grid(column=1, row=0, pady=3)

        self.boxTradeType = Combobox(frame, width=8, state='readonly')
        self.boxTradeType['values'] = ['ROD', 'IOC', 'FOK']
        self.boxTradeType.set('ROD')
        self.boxTradeType.grid(column=1, row=1, padx=5, pady=3, sticky='w')

        # 倉別
        lbNewClose = Label(frame, style="Pink.TLabel", text="倉別")
        lbNewClose.grid(column=2, row=0, pady=3)

        self.boxNewClose = Combobox(frame, width=8, state='readonly')
        self.boxNewClose['values'] = ['新倉', '平倉', '自動']
        self.boxNewClose.set('新倉')
        self.boxNewClose.grid(column=2, row=1, padx=5, pady=3, sticky='w')

        # 自動計算價差按鈕
        btnCalcSpread = Button(frame, style="Pink.TButton", text="自動計算價差")
        btnCalcSpread["command"] = self.__calculate_spread
        btnCalcSpread.grid(column=3, row=1, padx=5, pady=3)

    def __create_mit_widgets(self, parent):
        """建立MIT委託控件"""
        main_frame = Frame(parent, style="Pink.TFrame")
        main_frame.grid(column=0, row=0, padx=5, pady=5, sticky='ew')

        # MIT委託1
        mit1_frame = LabelFrame(main_frame, text="MIT委託1 (停損/停利)", style="Pink.TLabelframe")
        mit1_frame.grid(column=0, row=0, padx=5, pady=5, sticky='ew')

        self.__create_single_mit_widget(mit1_frame, 1)

        # MIT委託2
        mit2_frame = LabelFrame(main_frame, text="MIT委託2 (停損/停利)", style="Pink.TLabelframe")
        mit2_frame.grid(column=1, row=0, padx=5, pady=5, sticky='ew')

        self.__create_single_mit_widget(mit2_frame, 2)

    def __create_single_mit_widget(self, parent, mit_num):
        """建立單個MIT委託控件"""
        frame = Frame(parent, style="Pink.TFrame")
        frame.grid(column=0, row=0, padx=5, pady=5, sticky='w')

        # 啟用MIT委託
        chkEnabled = Checkbutton(frame, text="啟用MIT委託", style="Pink.TCheckbutton")
        chkEnabled.grid(column=0, row=0, columnspan=3, padx=5, pady=3, sticky='w')
        setattr(self, f'chkMit{mit_num}Enabled', chkEnabled)

        # 商品代碼
        lbStockNo = Label(frame, style="Pink.TLabel", text="商品代碼")
        lbStockNo.grid(column=0, row=1, pady=3)

        txtStockNo = Entry(frame, width=12)
        txtStockNo.grid(column=0, row=2, padx=5, pady=3, sticky='w')
        setattr(self, f'txtMit{mit_num}StockNo', txtStockNo)

        # 買賣別 (借用OptionOrder.py的完整選項)
        lbBuySell = Label(frame, style="Pink.TLabel", text="買賣別")
        lbBuySell.grid(column=1, row=1, pady=3)

        boxBuySell = Combobox(frame, width=8, state='readonly')
        try:
            boxBuySell['values'] = Config.BUYSELLSET
        except:
            boxBuySell['values'] = ['買進', '賣出']
        boxBuySell.set('買進' if mit_num == 1 else '賣出')
        boxBuySell.grid(column=1, row=2, padx=5, pady=3, sticky='w')
        setattr(self, f'boxMit{mit_num}BuySell', boxBuySell)

        # 委託量
        lbQty = Label(frame, style="Pink.TLabel", text="委託量")
        lbQty.grid(column=2, row=1, pady=3)

        txtQty = Entry(frame, width=8)
        txtQty.grid(column=2, row=2, padx=5, pady=3, sticky='w')
        setattr(self, f'txtMit{mit_num}Qty', txtQty)

        # 觸發價
        lbTriggerPrice = Label(frame, style="Pink.TLabel", text="觸發價")
        lbTriggerPrice.grid(column=0, row=3, pady=3)

        txtTriggerPrice = Entry(frame, width=10)
        txtTriggerPrice.grid(column=0, row=4, padx=5, pady=3, sticky='w')
        setattr(self, f'txtMit{mit_num}TriggerPrice', txtTriggerPrice)

        # 委託價
        lbDealPrice = Label(frame, style="Pink.TLabel", text="委託價")
        lbDealPrice.grid(column=1, row=3, pady=3)

        txtDealPrice = Entry(frame, width=10)
        txtDealPrice.grid(column=1, row=4, padx=5, pady=3, sticky='w')
        setattr(self, f'txtMit{mit_num}DealPrice', txtDealPrice)

        # 委託條件
        lbTradeType = Label(frame, style="Pink.TLabel", text="委託條件")
        lbTradeType.grid(column=2, row=3, pady=3)

        boxTradeType = Combobox(frame, width=8, state='readonly')
        boxTradeType['values'] = ['IOC', 'FOK']
        boxTradeType.set('IOC')
        boxTradeType.grid(column=2, row=4, padx=5, pady=3, sticky='w')
        setattr(self, f'boxMit{mit_num}TradeType', boxTradeType)

        # MIT委託按鈕
        btnSendMIT = Button(frame, style="Pink.TButton", text=f"送出MIT{mit_num}")
        if mit_num == 1:
            btnSendMIT["command"] = lambda: self.__SendMITOrder_Click(1)
        else:
            btnSendMIT["command"] = lambda: self.__SendMITOrder_Click(2)
        btnSendMIT.grid(column=0, row=5, padx=5, pady=5)

        # 取消MIT委託按鈕
        btnCancelMIT = Button(frame, style="Pink.TButton", text=f"取消MIT{mit_num}")
        if mit_num == 1:
            btnCancelMIT["command"] = lambda: self.__CancelMITOrder_Click(1)
        else:
            btnCancelMIT["command"] = lambda: self.__CancelMITOrder_Click(2)
        btnCancelMIT.grid(column=1, row=5, padx=5, pady=5)

        # 自動交易區域
        auto_frame = LabelFrame(frame, text=f"MIT{mit_num}自動交易", style="Pink.TLabelframe")
        auto_frame.grid(column=0, row=6, columnspan=3, padx=5, pady=5, sticky='ew')

        # 啟用自動交易
        chkAutoTrade = Checkbutton(auto_frame, text="啟用自動交易", style="Pink.TCheckbutton")
        chkAutoTrade.grid(column=0, row=0, columnspan=2, padx=5, pady=3, sticky='w')
        setattr(self, f'chkMit{mit_num}AutoTrade', chkAutoTrade)

        # 買入價格
        lbBuyPrice = Label(auto_frame, style="Pink.TLabel", text="買入價格")
        lbBuyPrice.grid(column=0, row=1, pady=3)

        txtBuyPrice = Entry(auto_frame, width=10)
        txtBuyPrice.grid(column=0, row=2, padx=5, pady=3, sticky='w')
        setattr(self, f'txtMit{mit_num}BuyPrice', txtBuyPrice)

        # 賣出價格
        lbSellPrice = Label(auto_frame, style="Pink.TLabel", text="賣出價格")
        lbSellPrice.grid(column=1, row=1, pady=3)

        txtSellPrice = Entry(auto_frame, width=10)
        txtSellPrice.grid(column=1, row=2, padx=5, pady=3, sticky='w')
        setattr(self, f'txtMit{mit_num}SellPrice', txtSellPrice)

        # 當前價格顯示
        lbCurrentPrice = Label(auto_frame, style="Pink.TLabel", text="當前價格")
        lbCurrentPrice.grid(column=2, row=1, pady=3)

        lbPriceValue = Label(auto_frame, style="Pink.TLabel", text="0.00", foreground="blue", font=("Arial", 10, "bold"))
        lbPriceValue.grid(column=2, row=2, padx=5, pady=3)
        setattr(self, f'lbMit{mit_num}CurrentPrice', lbPriceValue)

        # 部位狀態
        lbPosition = Label(auto_frame, style="Pink.TLabel", text="部位狀態")
        lbPosition.grid(column=0, row=3, pady=3)

        lbPositionValue = Label(auto_frame, style="Pink.TLabel", text="空單", foreground="green")
        lbPositionValue.grid(column=0, row=4, padx=5, pady=3)
        setattr(self, f'lbMit{mit_num}Position', lbPositionValue)

        # 自動交易控制按鈕
        btnStartAuto = Button(auto_frame, style="Pink.TButton", text="啟動自動")
        if mit_num == 1:
            btnStartAuto["command"] = lambda: self.__StartMITAutoTrading(1)
        else:
            btnStartAuto["command"] = lambda: self.__StartMITAutoTrading(2)
        btnStartAuto.grid(column=1, row=4, padx=5, pady=3)

        btnStopAuto = Button(auto_frame, style="Pink.TButton", text="停止自動")
        if mit_num == 1:
            btnStopAuto["command"] = lambda: self.__StopMITAutoTrading(1)
        else:
            btnStopAuto["command"] = lambda: self.__StopMITAutoTrading(2)
        btnStopAuto.grid(column=2, row=4, padx=5, pady=3)


    def __create_order_buttons(self, parent):
        """建立下單按鈕"""
        # 同步下單
        btnSendOrder = Button(parent, style="Pink.TButton", text="同步委託下單")
        btnSendOrder["command"] = self.__btnSendOrder_Click
        btnSendOrder.grid(column=0, row=0, padx=5, pady=3)

        # 非同步下單
        btnSendOrderAsync = Button(parent, style="Pink.TButton", text="非同步委託下單")
        btnSendOrderAsync["command"] = self.__btnSendOrderAsync_Click
        btnSendOrderAsync.grid(column=1, row=0, padx=5, pady=3)

        # 取消委託
        btnCancelOrder = Button(parent, style="Pink.TButton", text="取消委託")
        btnCancelOrder["command"] = self.__btnCancelOrder_Click
        btnCancelOrder.grid(column=2, row=0, padx=5, pady=3)

    def __create_monitor_widgets(self, parent):
        """建立監控資訊控件"""
        frame = Frame(parent, style="Pink.TFrame")
        frame.grid(column=0, row=0, padx=5, pady=5, sticky='ew')

        # 當前價差
        lbCurrentSpread = Label(frame, style="Pink.TLabel", text="當前價差:")
        lbCurrentSpread.grid(column=0, row=0, padx=5, pady=3)

        self.lbSpreadValue = Label(frame, style="Pink.TLabel", text="0.00", foreground="red", font=("Arial", 12, "bold"))
        self.lbSpreadValue.grid(column=1, row=0, padx=5, pady=3)

        # 策略狀態
        lbStatus = Label(frame, style="Pink.TLabel", text="策略狀態:")
        lbStatus.grid(column=2, row=0, padx=5, pady=3)

        self.lbStrategyStatus = Label(frame, style="Pink.TLabel", text="待機中", foreground="green")
        self.lbStrategyStatus.grid(column=3, row=0, padx=5, pady=3)

        # 損益試算
        lbPnL = Label(frame, style="Pink.TLabel", text="預估損益:")
        lbPnL.grid(column=4, row=0, padx=5, pady=3)

        self.lbPnLValue = Label(frame, style="Pink.TLabel", text="0.00", foreground="black")
        self.lbPnLValue.grid(column=5, row=0, padx=5, pady=3)

    def __calculate_spread(self):
        """自動計算價差委託價"""
        try:
            if QUOTE_AVAILABLE:
                # 使用即時報價計算
                spread = self.current_call_price - self.current_put_price
                self.txtSpreadPrice.delete(0, END)
                self.txtSpreadPrice.insert(0, f"{spread:.2f}")
                self.spread_value = spread
                self.__update_spread_display()
            else:
                # 使用輸入的委託價計算
                call_price = float(self.txtCallPrice.get() or "0")
                put_price = float(self.txtPutPrice.get() or "0")
                spread = call_price - put_price
                self.txtSpreadPrice.delete(0, END)
                self.txtSpreadPrice.insert(0, f"{spread:.2f}")
        except ValueError:
            messagebox.showerror("錯誤", "請確認價格輸入格式正確!")

    def __start_quote_monitoring(self):
        """啟動報價監控"""
        def monitor_quotes():
            while True:
                try:
                    # 安全地獲取GUI元件的值
                    try:
                        call_stock_no = self.txtCallStockNo.get() if hasattr(self, 'txtCallStockNo') else ""
                        put_stock_no = self.txtPutStockNo.get() if hasattr(self, 'txtPutStockNo') else ""
                    except:
                        call_stock_no = ""
                        put_stock_no = ""

                    if call_stock_no and put_stock_no and QUOTE_AVAILABLE:
                        # 使用Quote模組獲取報價
                        call_quote = QuoteModule.get_quote(call_stock_no)
                        put_quote = QuoteModule.get_quote(put_stock_no)

                        if call_quote and put_quote:
                            self.current_call_price = float(call_quote.get('price', 0))
                            self.current_put_price = float(put_quote.get('price', 0))

                            # 使用after方法安全地更新GUI
                            def update_gui():
                                try:
                                    if hasattr(self, 'lbCallCurrentPrice'):
                                        self.lbCallCurrentPrice.config(text=f"{self.current_call_price:.2f}")
                                    if hasattr(self, 'lbPutCurrentPrice'):
                                        self.lbPutCurrentPrice.config(text=f"{self.current_put_price:.2f}")

                                    # 更新價差
                                    self.spread_value = self.current_call_price - self.current_put_price
                                    self.__update_spread_display()
                                except Exception as e:
                                    print(f"GUI更新錯誤: {e}")

                            # 在主線程中執行GUI更新
                            self.after_idle(update_gui)

                    time.sleep(1)  # 每秒更新一次
                except Exception as e:
                    print(f"報價監控錯誤: {e}")
                    time.sleep(5)

        # 在背景執行緒中執行監控
        monitor_thread = threading.Thread(target=monitor_quotes, daemon=True)
        monitor_thread.start()

    def __update_spread_display(self):
        """更新價差顯示"""
        self.lbSpreadValue.config(text=f"{self.spread_value:.2f}")

        # 根據價差變化調整顏色
        if self.spread_value > 0:
            self.lbSpreadValue.config(foreground="red")  # 正價差紅色
        elif self.spread_value < 0:
            self.lbSpreadValue.config(foreground="green")  # 負價差綠色
        else:
            self.lbSpreadValue.config(foreground="black")  # 零價差黑色

    def __start_auto_trading(self):
        """啟動自動交易"""
        try:
            if not self.__validate_auto_trading_inputs():
                return

            if self.__dOrder['boxAccount'] == '':
                messagebox.showerror("錯誤", '請選擇期貨帳號!')
                return

            self.auto_trading_active = True
            self.chkAutoTrading.state(['selected'])
            messagebox.showinfo("自動交易", "自動交易已啟動")
            self.lbStrategyStatus.config(text="自動交易中", foreground="blue")

            # 啟動自動交易監控
            self.__start_auto_trading_monitor()

        except Exception as e:
            messagebox.showerror("錯誤", f"啟動自動交易失敗: {str(e)}")

    def __stop_auto_trading(self):
        """停止自動交易"""
        self.auto_trading_active = False
        self.chkAutoTrading.state(['!selected'])
        messagebox.showinfo("自動交易", "自動交易已停止")
        self.lbStrategyStatus.config(text="待機中", foreground="green")

    def __reset_position(self):
        """重設部位"""
        self.position_state = 0
        self.lbPositionState.config(text="空單", foreground="green")
        messagebox.showinfo("重設部位", "部位已重設為空單")

    def __validate_auto_trading_inputs(self):
        """驗證自動交易輸入"""
        try:
            if not self.txtAutoStockNo.get():
                messagebox.showerror("錯誤", "請輸入自動交易商品代碼!")
                return False

            if not self.txtAutoBuyPrice.get():
                messagebox.showerror("錯誤", "請輸入買入價格!")
                return False

            if not self.txtAutoSellPrice.get():
                messagebox.showerror("錯誤", "請輸入賣出價格!")
                return False

            if not self.txtAutoQty.get():
                messagebox.showerror("錯誤", "請輸入交易數量!")
                return False

            # 驗證數值格式
            buy_price = float(self.txtAutoBuyPrice.get())
            sell_price = float(self.txtAutoSellPrice.get())
            qty = int(self.txtAutoQty.get())

            if buy_price >= sell_price:
                messagebox.showerror("錯誤", "買入價格必須小於賣出價格!")
                return False

            if qty <= 0:
                messagebox.showerror("錯誤", "交易數量必須大於0!")
                return False

            return True

        except ValueError:
            messagebox.showerror("錯誤", "價格或數量格式錯誤!")
            return False

    def __start_auto_trading_monitor(self):
        """啟動自動交易監控"""
        def auto_trading_monitor():
            while self.auto_trading_active:
                try:
                    stock_no = self.txtAutoStockNo.get()

                    if stock_no and QUOTE_AVAILABLE:
                        # 獲取當前價格
                        quote = QuoteModule.get_quote(stock_no)
                        if quote:
                            current_price = float(quote.get('price', 0))
                            self.current_auto_price = current_price

                            # 更新UI顯示
                            self.lbCurrentAutoPrice.config(text=f"{current_price:.2f}")

                            # 執行自動交易邏輯
                            self.__execute_auto_trading_logic(current_price)

                    time.sleep(1)  # 每秒檢查一次

                except Exception as e:
                    print(f"自動交易監控錯誤: {e}")
                    time.sleep(5)

        # 在背景執行緒中執行監控
        auto_monitor_thread = threading.Thread(target=auto_trading_monitor, daemon=True)
        auto_monitor_thread.start()

    def __execute_auto_trading_logic(self, current_price):
        """執行自動交易邏輯"""
        try:
            with self.auto_trade_lock:  # 使用鎖防止重複下單
                buy_price = float(self.txtAutoBuyPrice.get())
                sell_price = float(self.txtAutoSellPrice.get())

                # 當前為空單且價格等於買入價格時，執行買入
                if self.position_state == 0 and abs(current_price - buy_price) < 0.01:
                    if self.__execute_auto_buy():
                        self.position_state = 1
                        self.lbPositionState.config(text="多單", foreground="red")
                        print(f"自動買入執行 - 價格: {current_price}")

                # 當前為多單且價格等於賣出價格時，執行賣出
                elif self.position_state == 1 and abs(current_price - sell_price) < 0.01:
                    if self.__execute_auto_sell():
                        self.position_state = 0
                        self.lbPositionState.config(text="空單", foreground="green")
                        print(f"自動賣出執行 - 價格: {current_price}")

        except Exception as e:
            print(f"自動交易邏輯錯誤: {e}")

    def __execute_auto_buy(self):
        """執行自動買入"""
        try:
            # 建立委託單物件
            oOrder = sk.FUTUREORDER()

            # 填入帳號資訊
            oOrder.bstrFullAccount = self.__dOrder['boxAccount']

            # 填入商品代號
            oOrder.bstrStockNo = self.txtAutoStockNo.get()

            # 買進
            oOrder.sBuySell = 0

            # 委託條件 (IOC)
            oOrder.sTradeType = 1

            # 新倉
            oOrder.sNewClose = 0

            # 非當沖
            oOrder.sDayTrade = 0

            # 委託價 (市價單)
            oOrder.bstrPrice = "M"

            # 委託數量
            oOrder.nQty = int(self.txtAutoQty.get())

            # 發送委託單
            message, m_nCode = skO.SendFutureOrder(Global.Global_IID, True, oOrder)

            if m_nCode == 0:
                strMsg = f"自動買入委託成功: {message}"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
                return True
            else:
                print(f"自動買入委託失敗，錯誤代碼: {m_nCode}")
                return False

        except Exception as e:
            print(f"自動買入錯誤: {e}")
            return False

    def __execute_auto_sell(self):
        """執行自動賣出"""
        try:
            # 建立委託單物件
            oOrder = sk.FUTUREORDER()

            # 填入帳號資訊
            oOrder.bstrFullAccount = self.__dOrder['boxAccount']

            # 填入商品代號
            oOrder.bstrStockNo = self.txtAutoStockNo.get()

            # 賣出
            oOrder.sBuySell = 1

            # 委託條件 (IOC)
            oOrder.sTradeType = 1

            # 平倉
            oOrder.sNewClose = 1

            # 非當沖
            oOrder.sDayTrade = 0

            # 委託價 (市價單)
            oOrder.bstrPrice = "M"

            # 委託數量
            oOrder.nQty = int(self.txtAutoQty.get())

            # 發送委託單
            message, m_nCode = skO.SendFutureOrder(Global.Global_IID, True, oOrder)

            if m_nCode == 0:
                strMsg = f"自動賣出委託成功: {message}"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
                return True
            else:
                print(f"自動賣出委託失敗，錯誤代碼: {m_nCode}")
                return False

        except Exception as e:
            print(f"自動賣出錯誤: {e}")
            return False

    def __btnSendOrder_Click(self):
        """同步下單按鈕點擊事件"""
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("錯誤", '請選擇期貨帳號!')
        else:
            self.__SendDuplexOrder(False)

    def __btnSendOrderAsync_Click(self):
        """非同步下單按鈕點擊事件"""
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("錯誤", '請選擇期貨帳號!')
        else:
            self.__SendDuplexOrder(True)

    def __btnCancelOrder_Click(self):
        """取消委託按鈕點擊事件"""
        messagebox.showinfo("取消委託", "取消委託功能待實作")

    def __SendDuplexOrder(self, bAsyncOrder):
        """執行雙賣價差複式委託下單"""
        try:
            # 驗證輸入
            if not self.__validate_inputs():
                return

            # 更新策略狀態
            self.lbStrategyStatus.config(text="下單中...", foreground="orange")

            # 準備委託條件參數 (使用新的映射函數)
            sTradeType = self.__get_trade_type()
            sNewClose = self.__get_new_close()

            # 建立複式委託單物件
            oOrder = sk.FUTUREORDER()

            # 填入帳號資訊
            oOrder.bstrFullAccount = self.__dOrder['boxAccount']

            # Call選擇權 (第一腳 - 賣出)
            oOrder.bstrStockNo = self.txtCallStockNo.get().strip()
            oOrder.sBuySell = 1  # 賣出 (雙賣策略)

            # Put選擇權 (第二腳 - 賣出)
            oOrder.bstrStockNo2 = self.txtPutStockNo.get().strip()
            oOrder.sBuySell2 = 1  # 賣出 (雙賣策略)

            # 委託條件
            oOrder.sTradeType = sTradeType

            # 委託價格 (使用價差委託價)
            oOrder.bstrPrice = self.txtSpreadPrice.get().strip()

            # 委託數量 (使用Call的數量，假設Call和Put數量相同)
            oOrder.nQty = int(self.txtCallQty.get().strip())

            # 倉別
            oOrder.sNewClose = sNewClose

            # 當沖標記 (預設非當沖)
            oOrder.sDayTrade = 0

            # 設定盤別 (如果有的話)
            if hasattr(self, 'boxCallReserved'):
                call_reserved = self.__get_reserved_code(self.boxCallReserved.get())
                oOrder.sReserved = call_reserved

            # 發送複式委託單
            message, m_nCode = skO.SendDuplexOrder(Global.Global_IID, bAsyncOrder, oOrder)

            # 處理回傳結果 (借用OptionOrder.py的錯誤處理邏輯)
            self.__oMsg.SendReturnMessage("DoubleSellSpread", m_nCode, "SendDuplexOrder", self.__dOrder['listInformation'])

            if bAsyncOrder == False and m_nCode == 0:
                strMsg = f"雙賣價差策略委託成功: {message}"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
                self.lbStrategyStatus.config(text="委託成功", foreground="green")
                messagebox.showinfo("下單成功", strMsg)
                print(f"[SUCCESS] {strMsg}")
            elif m_nCode != 0:
                error_msg = self.__get_error_message(m_nCode)
                self.lbStrategyStatus.config(text="委託失敗", foreground="red")
                full_error_msg = f"錯誤代碼: {m_nCode}\n錯誤訊息: {error_msg}"
                messagebox.showerror("下單失敗", full_error_msg)
                print(f"[ERROR] 下單失敗 - {full_error_msg}")
                self.__oMsg.WriteMessage(f"雙賣價差策略委託失敗: {full_error_msg}", self.__dOrder['listInformation'])
            else:
                self.lbStrategyStatus.config(text="委託送出", foreground="blue")
                print("[INFO] 非同步委託已送出，等待回覆中...")

        except Exception as e:
            self.lbStrategyStatus.config(text="下單錯誤", foreground="red")
            messagebox.showerror("錯誤", f"下單發生錯誤: {str(e)}")

    def __validate_inputs(self):
        """驗證輸入資料 (借用OptionOrder.py的驗證邏輯)"""
        try:
            # 檢查帳號
            if not self.__dOrder['boxAccount']:
                messagebox.showerror("錯誤", "請選擇期貨帳號!")
                return False

            # 檢查商品代碼
            call_stock_no = self.txtCallStockNo.get().strip()
            put_stock_no = self.txtPutStockNo.get().strip()

            if not call_stock_no:
                messagebox.showerror("錯誤", "請輸入Call選擇權商品代碼!")
                return False

            if not put_stock_no:
                messagebox.showerror("錯誤", "請輸入Put選擇權商品代碼!")
                return False

            # 驗證商品代碼格式 (選擇權代碼通常包含字母和數字)
            if len(call_stock_no) < 3 or len(put_stock_no) < 3:
                messagebox.showerror("錯誤", "商品代碼格式不正確!")
                return False

            # 檢查委託價
            call_price_str = self.txtCallPrice.get().strip()
            put_price_str = self.txtPutPrice.get().strip()
            spread_price_str = self.txtSpreadPrice.get().strip()

            if not call_price_str:
                messagebox.showerror("錯誤", "請輸入Call選擇權委託價!")
                return False

            if not put_price_str:
                messagebox.showerror("錯誤", "請輸入Put選擇權委託價!")
                return False

            if not spread_price_str:
                messagebox.showerror("錯誤", "請輸入價差委託價!")
                return False

            # 檢查委託量
            call_qty_str = self.txtCallQty.get().strip()
            put_qty_str = self.txtPutQty.get().strip()

            if not call_qty_str:
                messagebox.showerror("錯誤", "請輸入Call選擇權委託量!")
                return False

            if not put_qty_str:
                messagebox.showerror("錯誤", "請輸入Put選擇權委託量!")
                return False

            # 驗證數值格式並檢查合理性
            call_price = float(call_price_str)
            put_price = float(put_price_str)
            spread_price = float(spread_price_str)
            call_qty = int(call_qty_str)
            put_qty = int(put_qty_str)

            # 檢查價格合理性
            if call_price <= 0 or put_price <= 0:
                messagebox.showerror("錯誤", "委託價必須大於0!")
                return False

            # 檢查數量合理性
            if call_qty <= 0 or put_qty <= 0:
                messagebox.showerror("錯誤", "委託量必須大於0!")
                return False

            if call_qty > 9999 or put_qty > 9999:
                messagebox.showerror("錯誤", "委託量不能超過9999!")
                return False

            # 檢查委託條件選擇
            if not self.boxTradeType.get():
                messagebox.showerror("錯誤", "請選擇委託條件!")
                return False

            # 檢查倉別選擇
            if not self.boxNewClose.get():
                messagebox.showerror("錯誤", "請選擇倉別!")
                return False

            # 檢查Call和Put的盤別選擇
            if hasattr(self, 'boxCallReserved') and not self.boxCallReserved.get():
                messagebox.showerror("錯誤", "請選擇Call選擇權盤別!")
                return False

            if hasattr(self, 'boxPutReserved') and not self.boxPutReserved.get():
                messagebox.showerror("錯誤", "請選擇Put選擇權盤別!")
                return False

            # 檢查委託量是否相同 (雙賣策略要求)
            if call_qty != put_qty:
                result = messagebox.askyesno("警告", "雙賣策略建議Call和Put委託量相同!\n是否繼續下單?")
                if not result:
                    return False

            return True

        except ValueError as ve:
            messagebox.showerror("錯誤", f"數值格式錯誤: {str(ve)}")
            return False
        except Exception as e:
            messagebox.showerror("錯誤", f"輸入驗證失敗: {str(e)}")
            return False

    def __get_trade_type(self):
        """取得委託條件代碼"""
        trade_type_map = {
            "ROD": 0,
            "IOC": 1,
            "FOK": 2
        }
        return trade_type_map.get(self.boxTradeType.get(), 0)

    def __get_new_close(self):
        """取得倉別代碼"""
        new_close_map = {
            "新倉": 0,
            "平倉": 1,
            "自動": 2
        }
        return new_close_map.get(self.boxNewClose.get(), 0)

    def __get_error_message(self, error_code):
        """取得錯誤訊息 (借用OptionOrder.py的錯誤處理概念)"""
        error_messages = {
            0: "成功",
            -1: "一般性錯誤",
            -2: "帳號錯誤",
            -3: "商品代碼錯誤",
            -4: "買賣別錯誤",
            -5: "委託條件錯誤",
            -6: "委託價格錯誤",
            -7: "委託數量錯誤",
            -8: "倉別錯誤",
            -9: "盤別錯誤",
            -10: "帳號未登入",
            -11: "權限不足",
            -12: "餘額不足",
            -13: "超過可委託量",
            -14: "商品停止交易",
            -15: "非交易時段",
            -99: "未知錯誤"
        }
        return error_messages.get(error_code, f"錯誤代碼 {error_code}")

    def __get_buy_sell_code(self, buy_sell_text):
        """取得買賣別代碼 (借用OptionOrder.py的邏輯)"""
        if buy_sell_text == "買進":
            return 0
        elif buy_sell_text == "賣出":
            return 1
        else:
            return 0  # 預設買進

    def __get_reserved_code(self, reserved_text):
        """取得盤別代碼 (借用OptionOrder.py的邏輯)"""
        if reserved_text == "盤中":
            return 0
        elif reserved_text == "T盤預約":
            return 1
        else:
            return 0  # 預設盤中

    # MIT委託相關方法
    def __SendMITOrder_Click(self, mit_num):
        """MIT委託下單按鈕點擊事件"""
        if not self.__validate_mit_inputs(mit_num):
            return
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("錯誤", '請選擇期貨帳號!')
        else:
            self.__SendMITOrder(mit_num, False)

    def __CancelMITOrder_Click(self, mit_num):
        """取消MIT委託按鈕點擊事件"""
        messagebox.showinfo("取消MIT委託", f"取消MIT委託{mit_num}功能待實作")

    def __validate_mit_inputs(self, mit_num):
        """驗證MIT委託輸入資料"""
        try:
            # 檢查是否啟用
            chk_enabled = getattr(self, f'chkMit{mit_num}Enabled')
            if not chk_enabled.instate(['selected']):
                messagebox.showwarning("提醒", f"MIT委託{mit_num}未啟用!")
                return False

            # 檢查商品代碼
            stock_no = getattr(self, f'txtMit{mit_num}StockNo').get()
            if not stock_no:
                messagebox.showerror("錯誤", f"請輸入MIT委託{mit_num}商品代碼!")
                return False

            # 檢查觸發價
            trigger_price = getattr(self, f'txtMit{mit_num}TriggerPrice').get()
            if not trigger_price:
                messagebox.showerror("錯誤", f"請輸入MIT委託{mit_num}觸發價!")
                return False

            # 檢查委託價
            deal_price = getattr(self, f'txtMit{mit_num}DealPrice').get()
            if not deal_price:
                messagebox.showerror("錯誤", f"請輸入MIT委託{mit_num}委託價!")
                return False

            # 檢查委託量
            qty = getattr(self, f'txtMit{mit_num}Qty').get()
            if not qty:
                messagebox.showerror("錯誤", f"請輸入MIT委託{mit_num}委託量!")
                return False

            # 驗證數值格式
            float(trigger_price)
            float(deal_price)
            int(qty)

            return True

        except ValueError:
            messagebox.showerror("錯誤", f"MIT委託{mit_num}數值輸入格式錯誤!")
            return False

    def __SendMITOrder(self, mit_num, bAsyncOrder):
        """執行MIT委託下單"""
        try:
            # 更新策略狀態
            self.lbStrategyStatus.config(text=f"MIT{mit_num}下單中...", foreground="orange")

            # 取得買賣別 (使用新的映射函數)
            buy_sell_text = getattr(self, f'boxMit{mit_num}BuySell').get()
            sBuySell = self.__get_buy_sell_code(buy_sell_text)

            # 取得委託條件 (使用IOC=1, FOK=2的映射)
            trade_type_text = getattr(self, f'boxMit{mit_num}TradeType').get()
            sTradeType = 1 if trade_type_text == "IOC" else 2  # IOC=1, FOK=2

            # 建立MIT委託單物件
            oOrder = sk.FUTUREORDER()

            # 填入帳號資訊
            oOrder.bstrFullAccount = self.__dOrder['boxAccount']

            # 填入商品代號
            oOrder.bstrStockNo = getattr(self, f'txtMit{mit_num}StockNo').get()

            # 買賣別
            oOrder.sBuySell = sBuySell

            # 委託條件
            oOrder.sTradeType = sTradeType

            # 新倉(MIT委託通常是新倉)
            oOrder.sNewClose = 0

            # 非當沖
            oOrder.sDayTrade = 0

            # 委託價
            oOrder.bstrPrice = getattr(self, f'txtMit{mit_num}DealPrice').get()

            # 委託數量
            oOrder.nQty = int(getattr(self, f'txtMit{mit_num}Qty').get())

            # 觸發價
            oOrder.bstrTrigger = getattr(self, f'txtMit{mit_num}TriggerPrice').get()

            # 發送MIT委託單
            message, m_nCode = skO.SendFutureMITOrder(Global.Global_IID, bAsyncOrder, oOrder)

            # 處理回傳結果
            self.__oMsg.SendReturnMessage(f"MIT{mit_num}Order", m_nCode, "SendFutureMITOrder", self.__dOrder['listInformation'])

            if bAsyncOrder == False and m_nCode == 0:
                strMsg = f"MIT{mit_num}委託成功: {message}"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
                self.lbStrategyStatus.config(text=f"MIT{mit_num}委託成功", foreground="green")
                messagebox.showinfo("MIT下單成功", strMsg)
            elif m_nCode != 0:
                self.lbStrategyStatus.config(text=f"MIT{mit_num}委託失敗", foreground="red")
                messagebox.showerror("MIT下單失敗", f"錯誤代碼: {m_nCode}")
            else:
                self.lbStrategyStatus.config(text=f"MIT{mit_num}委託送出", foreground="blue")

        except Exception as e:
            self.lbStrategyStatus.config(text=f"MIT{mit_num}下單錯誤", foreground="red")
            messagebox.showerror("錯誤", f"MIT{mit_num}下單發生錯誤: {str(e)}")


    # MIT自動交易功能
    def __StartMITAutoTrading(self, mit_num):
        """啟動MIT自動交易"""
        try:
            if not self.__validate_mit_auto_inputs(mit_num):
                return

            if self.__dOrder['boxAccount'] == '':
                messagebox.showerror("錯誤", '請選擇期貨帳號!')
                return

            self.mit_auto_trading[mit_num] = True
            chk_auto = getattr(self, f'chkMit{mit_num}AutoTrade')
            chk_auto.state(['selected'])

            messagebox.showinfo("MIT自動交易", f"MIT{mit_num}自動交易已啟動")

            # 啟動MIT自動交易監控
            self.__start_mit_auto_monitoring(mit_num)

        except Exception as e:
            messagebox.showerror("錯誤", f"啟動MIT{mit_num}自動交易失敗: {str(e)}")

    def __StopMITAutoTrading(self, mit_num):
        """停止MIT自動交易"""
        self.mit_auto_trading[mit_num] = False
        chk_auto = getattr(self, f'chkMit{mit_num}AutoTrade')
        chk_auto.state(['!selected'])
        messagebox.showinfo("MIT自動交易", f"MIT{mit_num}自動交易已停止")

    def __validate_mit_auto_inputs(self, mit_num):
        """驗證MIT自動交易輸入"""
        try:
            stock_no = getattr(self, f'txtMit{mit_num}StockNo').get()
            buy_price = getattr(self, f'txtMit{mit_num}BuyPrice').get()
            sell_price = getattr(self, f'txtMit{mit_num}SellPrice').get()
            qty = getattr(self, f'txtMit{mit_num}Qty').get()

            if not stock_no:
                messagebox.showerror("錯誤", f"請輸入MIT{mit_num}商品代碼!")
                return False

            if not buy_price:
                messagebox.showerror("錯誤", f"請輸入MIT{mit_num}買入價格!")
                return False

            if not sell_price:
                messagebox.showerror("錯誤", f"請輸入MIT{mit_num}賣出價格!")
                return False

            if not qty:
                messagebox.showerror("錯誤", f"請輸入MIT{mit_num}委託數量!")
                return False

            # 驗證數值格式
            buy_val = float(buy_price)
            sell_val = float(sell_price)
            qty_val = int(qty)

            if buy_val >= sell_val:
                messagebox.showerror("錯誤", f"MIT{mit_num}買入價格必須小於賣出價格!")
                return False

            if qty_val <= 0:
                messagebox.showerror("錯誤", f"MIT{mit_num}委託數量必須大於0!")
                return False

            return True

        except ValueError:
            messagebox.showerror("錯誤", f"MIT{mit_num}價格或數量格式錯誤!")
            return False

    def __start_mit_auto_monitoring(self, mit_num):
        """啟動MIT自動交易監控"""
        def mit_auto_monitor():
            while self.mit_auto_trading[mit_num]:
                try:
                    # 安全地獲取商品代碼
                    try:
                        stock_no = getattr(self, f'txtMit{mit_num}StockNo').get() if hasattr(self, f'txtMit{mit_num}StockNo') else ""
                    except:
                        stock_no = ""

                    if stock_no and QUOTE_AVAILABLE:
                        # 獲取當前價格
                        quote = QuoteModule.get_quote(stock_no)
                        if quote:
                            current_price = float(quote.get('price', 0))
                            self.mit_current_price[mit_num] = current_price

                            # 使用after方法安全地更新GUI
                            def update_price_gui():
                                try:
                                    price_label = getattr(self, f'lbMit{mit_num}CurrentPrice', None)
                                    if price_label:
                                        price_label.config(text=f"{current_price:.2f}")
                                except Exception as e:
                                    print(f"MIT{mit_num}價格GUI更新錯誤: {e}")

                            # 在主線程中執行GUI更新
                            self.after_idle(update_price_gui)

                            # 執行MIT自動交易邏輯
                            self.__execute_mit_auto_trading_logic(mit_num, current_price)

                    time.sleep(1)  # 每秒檢查一次

                except Exception as e:
                    print(f"MIT{mit_num}自動交易監控錯誤: {e}")
                    time.sleep(5)

        # 在背景執行緒中執行監控
        mit_monitor_thread = threading.Thread(target=mit_auto_monitor, daemon=True)
        mit_monitor_thread.start()

    def __execute_mit_auto_trading_logic(self, mit_num, current_price):
        """執行MIT自動交易邏輯 (增強版，加入安全檢查)"""
        try:
            with self.mit_trade_locks[mit_num]:  # 使用鎖防止重複下單
                # 安全檢查：確保價格有效
                if current_price <= 0:
                    return

                buy_price_str = getattr(self, f'txtMit{mit_num}BuyPrice').get().strip()
                sell_price_str = getattr(self, f'txtMit{mit_num}SellPrice').get().strip()

                # 安全檢查：確保價格輸入不為空
                if not buy_price_str or not sell_price_str:
                    return

                buy_price = float(buy_price_str)
                sell_price = float(sell_price_str)

                # 安全檢查：確保價格合理
                if buy_price <= 0 or sell_price <= 0:
                    return

                # 安全檢查：確保買入價格小於賣出價格
                if buy_price >= sell_price:
                    print(f"MIT{mit_num}警告: 買入價格({buy_price})應小於賣出價格({sell_price})")
                    return

                # 當前為空單且價格等於買入價格時，執行買入
                if self.mit_position_state[mit_num] == 0 and abs(current_price - buy_price) < 0.01:
                    print(f"MIT{mit_num}觸發買入條件 - 當前價格: {current_price}, 買入價格: {buy_price}")
                    if self.__execute_mit_auto_buy(mit_num):
                        self.mit_position_state[mit_num] = 1

                        # 安全地更新GUI
                        def update_position_buy():
                            try:
                                position_label = getattr(self, f'lbMit{mit_num}Position', None)
                                if position_label:
                                    position_label.config(text="多單", foreground="red")
                            except Exception as e:
                                print(f"MIT{mit_num}部位GUI更新錯誤: {e}")

                        self.after_idle(update_position_buy)
                        print(f"[SUCCESS] MIT{mit_num}自動買入執行 - 價格: {current_price}")

                # 當前為多單且價格等於賣出價格時，執行賣出
                elif self.mit_position_state[mit_num] == 1 and abs(current_price - sell_price) < 0.01:
                    print(f"MIT{mit_num}觸發賣出條件 - 當前價格: {current_price}, 賣出價格: {sell_price}")
                    if self.__execute_mit_auto_sell(mit_num):
                        self.mit_position_state[mit_num] = 0

                        # 安全地更新GUI
                        def update_position_sell():
                            try:
                                position_label = getattr(self, f'lbMit{mit_num}Position', None)
                                if position_label:
                                    position_label.config(text="空單", foreground="green")
                            except Exception as e:
                                print(f"MIT{mit_num}部位GUI更新錯誤: {e}")

                        self.after_idle(update_position_sell)
                        print(f"[SUCCESS] MIT{mit_num}自動賣出執行 - 價格: {current_price}")

        except ValueError as ve:
            print(f"MIT{mit_num}自動交易價格格式錯誤: {ve}")
        except Exception as e:
            print(f"MIT{mit_num}自動交易邏輯錯誤: {e}")

    def __execute_mit_auto_buy(self, mit_num):
        """執行MIT自動買入"""
        try:
            # 建立委託單物件
            oOrder = sk.FUTUREORDER()

            # 填入帳號資訊
            oOrder.bstrFullAccount = self.__dOrder['boxAccount']

            # 填入商品代號
            oOrder.bstrStockNo = getattr(self, f'txtMit{mit_num}StockNo').get()

            # 買進
            oOrder.sBuySell = 0

            # 委託條件 (IOC)
            oOrder.sTradeType = 1

            # 新倉
            oOrder.sNewClose = 0

            # 非當沖
            oOrder.sDayTrade = 0

            # 委託價 (市價單)
            oOrder.bstrPrice = "M"

            # 委託數量
            oOrder.nQty = int(getattr(self, f'txtMit{mit_num}Qty').get())

            # 發送委託單
            message, m_nCode = skO.SendFutureOrder(Global.Global_IID, True, oOrder)

            if m_nCode == 0:
                strMsg = f"MIT{mit_num}自動買入委託成功: {message}"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
                return True
            else:
                print(f"MIT{mit_num}自動買入委託失敗，錯誤代碼: {m_nCode}")
                return False

        except Exception as e:
            print(f"MIT{mit_num}自動買入錯誤: {e}")
            return False

    def __execute_mit_auto_sell(self, mit_num):
        """執行MIT自動賣出"""
        try:
            # 建立委託單物件
            oOrder = sk.FUTUREORDER()

            # 填入帳號資訊
            oOrder.bstrFullAccount = self.__dOrder['boxAccount']

            # 填入商品代號
            oOrder.bstrStockNo = getattr(self, f'txtMit{mit_num}StockNo').get()

            # 賣出
            oOrder.sBuySell = 1

            # 委託條件 (IOC)
            oOrder.sTradeType = 1

            # 平倉
            oOrder.sNewClose = 1

            # 非當沖
            oOrder.sDayTrade = 0

            # 委託價 (市價單)
            oOrder.bstrPrice = "M"

            # 委託數量
            oOrder.nQty = int(getattr(self, f'txtMit{mit_num}Qty').get())

            # 發送委託單
            message, m_nCode = skO.SendFutureOrder(Global.Global_IID, True, oOrder)

            if m_nCode == 0:
                strMsg = f"MIT{mit_num}自動賣出委託成功: {message}"
                self.__oMsg.WriteMessage(strMsg, self.__dOrder['listInformation'])
                return True
            else:
                print(f"MIT{mit_num}自動賣出委託失敗，錯誤代碼: {m_nCode}")
                return False

        except Exception as e:
            print(f"MIT{mit_num}自動賣出錯誤: {e}")
            return False


# 建立主要的策略視窗類別
class DoubleSellSpreadStrategyWindow:
    def __init__(self, information=None):
        self.root = Tk()
        self.root.title("雙賣價差策略交易系統")
        self.root.geometry("800x600")

        # 建立策略物件
        self.strategy = DoubleSellSpreadStrategy(information)
        self.strategy.pack(fill=BOTH, expand=True)

        # 建立選單
        self.__create_menu()

    def __create_menu(self):
        """建立選單列"""
        menubar = Menu(self.root)
        self.root.config(menu=menubar)

        # 檔案選單
        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="檔案", menu=file_menu)
        file_menu.add_command(label="載入策略參數", command=self.__load_strategy)
        file_menu.add_command(label="儲存策略參數", command=self.__save_strategy)
        file_menu.add_separator()
        file_menu.add_command(label="結束", command=self.root.quit)

        # 工具選單
        tools_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="工具", menu=tools_menu)
        tools_menu.add_command(label="策略回測", command=self.__backtest)
        tools_menu.add_command(label="風險評估", command=self.__risk_assessment)

        # 說明選單
        help_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="說明", menu=help_menu)
        help_menu.add_command(label="策略說明", command=self.__show_strategy_help)
        help_menu.add_command(label="關於", command=self.__show_about)

    def __load_strategy(self):
        """載入策略參數"""
        messagebox.showinfo("載入策略", "載入策略參數功能待實作")

    def __save_strategy(self):
        """儲存策略參數"""
        messagebox.showinfo("儲存策略", "儲存策略參數功能待實作")

    def __backtest(self):
        """策略回測"""
        messagebox.showinfo("策略回測", "策略回測功能待實作")

    def __risk_assessment(self):
        """風險評估"""
        messagebox.showinfo("風險評估", "風險評估功能待實作")

    def __show_strategy_help(self):
        """顯示策略說明"""
        help_text = """
雙賣價差策略說明:

1. 策略概念:
   - 同時賣出Call選擇權和Put選擇權
   - 收取權利金為主要收益來源
   - 適合預期標的物價格波動不大的情況

2. 風險特性:
   - 有限收益: 最大收益為收取的權利金
   - 無限風險: 標的物大幅上漲或下跌時虧損無上限
   - 需要足夠的保證金

3. 使用建議:
   - 適合有經驗的投資者
   - 需要密切監控部位
   - 建議設定停損點

4. MIT委託功能:
   - MIT委託1: 通常設定為停損單，當價格觸及設定價位時自動平倉
   - MIT委託2: 通常設定為停利單，當價格達到獲利目標時自動平倉
   - 可分別設定不同的期貨商品代碼、觸發價、委託價
   - 支援買進/賣出雙向操作，靈活管理風險

5. 自動交易功能:
   - 設定買入價格和賣出價格進行自動交易
   - 當前價格等於買入價格時自動買進(建立多單)
   - 當前價格等於賣出價格時自動賣出(平倉回到空單)
   - 具備部位管理功能，防止重複下單
   - 使用多執行緒監控，即時響應價格變化
   - 支援市價單快速成交
        """
        messagebox.showinfo("策略說明", help_text)

    def __show_about(self):
        """關於資訊"""
        messagebox.showinfo("關於", "雙賣價差策略交易系統 v1.0\n基於群益API開發")

    def set_account(self, account):
        """設定交易帳號"""
        self.strategy.SetAccount(account)

    def run(self):
        """執行主迴圈"""
        self.root.mainloop()


# 測試函數
def test_double_sell_spread():
    """測試雙賣價差策略"""
    try:
        # 建立測試視窗
        app = DoubleSellSpreadStrategyWindow()

        # 設定測試帳號 (實際使用時需要真實帳號)
        # app.set_account("YOUR_ACCOUNT_HERE")

        # 執行
        app.run()

    except Exception as e:
        print(f"測試錯誤: {e}")


# 主程式進入點
if __name__ == "__main__":
    test_double_sell_spread()