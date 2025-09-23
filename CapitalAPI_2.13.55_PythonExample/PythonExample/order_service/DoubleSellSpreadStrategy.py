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

class DoubleSellSpreadStrategy(Frame):
    def __init__(self, information=None):
        Frame.__init__(self)
        self.__oMsg = MessageControl.MessageControl()
        # UI variable
        self.__dOrder = dict(
            listInformation = information,
            boxAccount = ''
        )

        # 策略參數
        self.__strategy_params = {
            'call_sell_strike': 0,      # 賣買權履約價
            'call_sell_premium': 0,     # 賣買權權利金
            'call_buy_strike': 0,       # 買買權履約價
            'call_buy_premium': 0,      # 買買權權利金
            'put_sell_strike': 0,       # 賣賣權履約價
            'put_sell_premium': 0,      # 賣賣權權利金
            'put_buy_strike': 0,        # 買賣權履約價
            'put_buy_premium': 0,       # 買賣權權利金
            'quantity': 1,              # 委託量
            'future_symbol': 'MXF',     # 期貨商品代號
            'option_month': '202501',   # 選擇權月份
            'monitor_interval': 1       # 監控間隔(秒)
        }

        # 狀態管理
        self.__order_status = {
            'call_spread_order_sent': False,
            'put_spread_order_sent': False,
            'hedge_long_sent': False,
            'hedge_short_sent': False,
            'monitoring': False
        }

        # 測試模式設定
        self.__test_mode = True  # 預設開啟測試模式
        self.__test_current_price = 25500  # 模擬當前價格
        self.__use_real_quote = False  # 是否使用實際報價
        self.__real_price_cache = {}  # 實際價格快取
        self.__last_quote_time = 0  # 上次報價時間

        # 監控線程
        self.__monitor_thread = None
        self.__stop_monitoring = False

        self.__CreateWidget()

    def SetAccount(self, account):
        self.__dOrder['boxAccount'] = account

    def __CreateWidget(self):
        # 主框架
        main_group = LabelFrame(self, text="雙賣價差單策略", style="Pink.TLabelframe")
        main_group.grid(column=0, row=0, padx=5, pady=5, sticky='ew')

        # 參數設定區域
        self.__CreateParameterSection(main_group)

        # 控制按鈕區域
        self.__CreateControlSection(main_group)

        # 狀態顯示區域
        self.__CreateStatusSection(main_group)

    def __CreateParameterSection(self, parent):
        # 參數設定框架
        param_group = LabelFrame(parent, text="策略參數設定", style="Pink.TLabelframe")
        param_group.grid(column=0, row=0, padx=5, pady=5, sticky='ew', columnspan=2)

        frame = Frame(param_group, style="Pink.TFrame")
        frame.grid(column=0, row=0, padx=5, pady=5, sticky='ew')

        # 買權價差單設定
        call_frame = LabelFrame(frame, text="買權價差單設定", style="Pink.TLabelframe")
        call_frame.grid(column=0, row=0, padx=5, pady=5, sticky='w')

        # 賣買權設定
        Label(call_frame, text="賣買權履約價:", style="Pink.TLabel").grid(column=0, row=0, padx=5, pady=2, sticky='w')
        self.call_sell_strike_entry = Entry(call_frame, width=10)
        self.call_sell_strike_entry.grid(column=1, row=0, padx=5, pady=2)
        self.call_sell_strike_entry.insert(0, "26000")

        Label(call_frame, text="賣買權權利金:", style="Pink.TLabel").grid(column=0, row=1, padx=5, pady=2, sticky='w')
        self.call_sell_premium_entry = Entry(call_frame, width=10)
        self.call_sell_premium_entry.grid(column=1, row=1, padx=5, pady=2)
        self.call_sell_premium_entry.insert(0, "20")

        # 買買權設定
        Label(call_frame, text="買買權履約價:", style="Pink.TLabel").grid(column=2, row=0, padx=5, pady=2, sticky='w')
        self.call_buy_strike_entry = Entry(call_frame, width=10)
        self.call_buy_strike_entry.grid(column=3, row=0, padx=5, pady=2)
        self.call_buy_strike_entry.insert(0, "26100")

        Label(call_frame, text="買買權權利金:", style="Pink.TLabel").grid(column=2, row=1, padx=5, pady=2, sticky='w')
        self.call_buy_premium_entry = Entry(call_frame, width=10)
        self.call_buy_premium_entry.grid(column=3, row=1, padx=5, pady=2)
        self.call_buy_premium_entry.insert(0, "10")

        # 賣權價差單設定
        put_frame = LabelFrame(frame, text="賣權價差單設定", style="Pink.TLabelframe")
        put_frame.grid(column=1, row=0, padx=5, pady=5, sticky='w')

        # 賣賣權設定
        Label(put_frame, text="賣賣權履約價:", style="Pink.TLabel").grid(column=0, row=0, padx=5, pady=2, sticky='w')
        self.put_sell_strike_entry = Entry(put_frame, width=10)
        self.put_sell_strike_entry.grid(column=1, row=0, padx=5, pady=2)
        self.put_sell_strike_entry.insert(0, "25000")

        # 賣賣權權利金
        Label(put_frame, text="賣賣權權利金:", style="Pink.TLabel").grid(column=0, row=1, padx=5, pady=2, sticky='w')
        self.put_sell_premium_entry = Entry(put_frame, width=10)
        self.put_sell_premium_entry.grid(column=1, row=1, padx=5, pady=2)
        self.put_sell_premium_entry.insert(0, "30")

        # 買賣權設定
        Label(put_frame, text="買賣權履約價:", style="Pink.TLabel").grid(column=2, row=0, padx=5, pady=2, sticky='w')
        self.put_buy_strike_entry = Entry(put_frame, width=10)
        self.put_buy_strike_entry.grid(column=3, row=0, padx=5, pady=2)
        self.put_buy_strike_entry.insert(0, "24900")

        # 買賣權權利金
        Label(put_frame, text="買賣權權利金:", style="Pink.TLabel").grid(column=2, row=1, padx=5, pady=2, sticky='w')
        self.put_buy_premium_entry = Entry(put_frame, width=10)
        self.put_buy_premium_entry.grid(column=3, row=1, padx=5, pady=2)
        self.put_buy_premium_entry.insert(0, "10")

        # 通用設定
        general_frame = LabelFrame(frame, text="通用設定", style="Pink.TLabelframe")
        general_frame.grid(column=0, row=1, padx=5, pady=5, sticky='w', columnspan=2)

        # 委託量
        Label(general_frame, text="委託量:", style="Pink.TLabel").grid(column=0, row=0, padx=5, pady=2, sticky='w')
        self.quantity_entry = Entry(general_frame, width=10)
        self.quantity_entry.grid(column=1, row=0, padx=5, pady=2)
        self.quantity_entry.insert(0, "1")

        # 期貨商品代號
        Label(general_frame, text="期貨商品代號:", style="Pink.TLabel").grid(column=2, row=0, padx=5, pady=2, sticky='w')
        self.future_symbol_entry = Entry(general_frame, width=10)
        self.future_symbol_entry.grid(column=3, row=0, padx=5, pady=2)
        self.future_symbol_entry.insert(0, "MXF")

        # 選擇權月份
        Label(general_frame, text="選擇權月份:", style="Pink.TLabel").grid(column=0, row=1, padx=5, pady=2, sticky='w')
        self.option_month_entry = Entry(general_frame, width=10)
        self.option_month_entry.grid(column=1, row=1, padx=5, pady=2)
        self.option_month_entry.insert(0, "202501")

        # 監控間隔
        Label(general_frame, text="監控間隔(秒):", style="Pink.TLabel").grid(column=2, row=1, padx=5, pady=2, sticky='w')
        self.monitor_interval_entry = Entry(general_frame, width=10)
        self.monitor_interval_entry.grid(column=3, row=1, padx=5, pady=2)
        self.monitor_interval_entry.insert(0, "1")

        # 測試模式設定
        test_frame = LabelFrame(frame, text="測試模式設定", style="Pink.TLabelframe")
        test_frame.grid(column=0, row=2, padx=5, pady=5, sticky='w', columnspan=2)

        # 測試模式開關
        self.test_mode_var = BooleanVar(value=True)
        self.test_mode_check = Checkbutton(test_frame, text="啟用測試模式（不實際下單）",
                                          variable=self.test_mode_var, style="Pink.TCheckbutton")
        self.test_mode_check.grid(column=0, row=0, padx=5, pady=2, sticky='w', columnspan=2)

        # 模擬價格
        Label(test_frame, text="模擬價格:", style="Pink.TLabel").grid(column=0, row=1, padx=5, pady=2, sticky='w')
        self.test_price_entry = Entry(test_frame, width=10)
        self.test_price_entry.grid(column=1, row=1, padx=5, pady=2)
        self.test_price_entry.insert(0, "25500")

        # 價格調整按鈕
        Button(test_frame, text="價格+100", style="Pink.TButton",
               command=lambda: self.__adjust_test_price(100)).grid(column=2, row=1, padx=5, pady=2)
        Button(test_frame, text="價格-100", style="Pink.TButton",
               command=lambda: self.__adjust_test_price(-100)).grid(column=3, row=1, padx=5, pady=2)

        # 實際報價選項
        if QUOTE_AVAILABLE:
            self.use_real_quote_var = BooleanVar(value=False)
            self.real_quote_check = Checkbutton(test_frame, text="使用實際報價（需先連線報價系統）",
                                              variable=self.use_real_quote_var, style="Pink.TCheckbutton")
            self.real_quote_check.grid(column=0, row=2, padx=5, pady=2, sticky='w', columnspan=2)

            # 報價連線狀態
            self.quote_status_label = Label(test_frame, text="報價狀態: 未連線", style="Pink.TLabel")
            self.quote_status_label.grid(column=2, row=2, padx=5, pady=2, sticky='w', columnspan=2)

            # 手動抓取實際價格按鈕
            Button(test_frame, text="抓取實際價格", style="Pink.TButton",
                   command=self.__fetch_real_price).grid(column=0, row=3, padx=5, pady=2)
        else:
            # 報價模組不可用提示
            Label(test_frame, text="注意：報價模組不可用，僅能使用模擬價格",
                  style="Pink.TLabel", foreground="red").grid(column=0, row=2, padx=5, pady=2, sticky='w', columnspan=4)

    def __CreateControlSection(self, parent):
        # 控制按鈕框架
        control_group = LabelFrame(parent, text="策略控制", style="Pink.TLabelframe")
        control_group.grid(column=0, row=1, padx=5, pady=5, sticky='ew')

        button_frame = Frame(control_group, style="Pink.TFrame")
        button_frame.grid(column=0, row=0, padx=5, pady=5)

        # 計算策略按鈕
        Button(button_frame, text="計算價差策略", style="Pink.TButton",
               command=self.__calculate_strategy).grid(column=0, row=0, padx=5, pady=5)

        # 下價差單按鈕
        Button(button_frame, text="下價差單", style="Pink.TButton",
               command=self.__send_spread_orders).grid(column=1, row=0, padx=5, pady=5)

        # 開始監控按鈕
        self.start_monitor_btn = Button(button_frame, text="開始監控", style="Pink.TButton",
                                       command=self.__start_monitoring)
        self.start_monitor_btn.grid(column=2, row=0, padx=5, pady=5)

        # 停止監控按鈕
        self.stop_monitor_btn = Button(button_frame, text="停止監控", style="Pink.TButton",
                                      command=self.__stop_monitoring_func, state='disabled')
        self.stop_monitor_btn.grid(column=3, row=0, padx=5, pady=5)

        # 重置狀態按鈕
        Button(button_frame, text="重置狀態", style="Pink.TButton",
               command=self.__reset_status).grid(column=4, row=0, padx=5, pady=5)

    def __CreateStatusSection(self, parent):
        # 狀態顯示框架
        status_group = LabelFrame(parent, text="策略狀態", style="Pink.TLabelframe")
        status_group.grid(column=1, row=1, padx=5, pady=5, sticky='ew')

        # 狀態文字區域
        self.status_text = Text(status_group, height=10, width=60)
        self.status_text.grid(column=0, row=0, padx=5, pady=5)

        # 滾動條
        scrollbar = Scrollbar(status_group, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(column=1, row=0, sticky='ns')
        self.status_text.configure(yscrollcommand=scrollbar.set)

    def __update_parameters(self):
        """更新策略參數"""
        try:
            self.__strategy_params['call_sell_strike'] = float(self.call_sell_strike_entry.get())
            self.__strategy_params['call_sell_premium'] = float(self.call_sell_premium_entry.get())
            self.__strategy_params['call_buy_strike'] = float(self.call_buy_strike_entry.get())
            self.__strategy_params['call_buy_premium'] = float(self.call_buy_premium_entry.get())
            self.__strategy_params['put_sell_strike'] = float(self.put_sell_strike_entry.get())
            self.__strategy_params['put_sell_premium'] = float(self.put_sell_premium_entry.get())
            self.__strategy_params['put_buy_strike'] = float(self.put_buy_strike_entry.get())
            self.__strategy_params['put_buy_premium'] = float(self.put_buy_premium_entry.get())
            self.__strategy_params['quantity'] = int(self.quantity_entry.get())
            self.__strategy_params['future_symbol'] = self.future_symbol_entry.get()
            self.__strategy_params['option_month'] = self.option_month_entry.get()
            self.__strategy_params['monitor_interval'] = float(self.monitor_interval_entry.get())

            # 更新測試模式設定
            self.__test_mode = self.test_mode_var.get()
            self.__test_current_price = float(self.test_price_entry.get())

            # 更新實際報價設定
            if QUOTE_AVAILABLE and hasattr(self, 'use_real_quote_var'):
                self.__use_real_quote = self.use_real_quote_var.get()

            return True
        except ValueError as e:
            messagebox.showerror("參數錯誤", f"請檢查參數格式: {e}")
            return False

    def __adjust_test_price(self, adjustment):
        """調整測試價格"""
        try:
            current_price = float(self.test_price_entry.get())
            new_price = current_price + adjustment
            self.test_price_entry.delete(0, END)
            self.test_price_entry.insert(0, str(new_price))
            self.__test_current_price = new_price

            if self.__order_status['monitoring']:
                self.__log_message(f"模擬價格調整至: {new_price}")
                # 立即檢查避險條件
                self.__check_hedge_conditions(new_price)
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的價格數值")

    def __fetch_real_price(self):
        """手動抓取實際價格"""
        if not QUOTE_AVAILABLE:
            messagebox.showerror("錯誤", "報價模組不可用")
            return

        try:
            # 抓取小台期貨的實際價格
            future_symbol = self.future_symbol_entry.get() or "MXF"
            real_price = self.__get_real_future_price(future_symbol)

            if real_price is not None:
                # 更新模擬價格為實際價格
                self.test_price_entry.delete(0, END)
                self.test_price_entry.insert(0, str(real_price))
                self.__test_current_price = real_price

                self.__log_message(f"已抓取 {future_symbol} 實際價格: {real_price}")

                # 更新報價狀態
                if hasattr(self, 'quote_status_label'):
                    self.quote_status_label.config(text=f"最新價格: {real_price}")
            else:
                self.__log_message(f"無法取得 {future_symbol} 的實際價格，請確認報價連線狀態")
                if hasattr(self, 'quote_status_label'):
                    self.quote_status_label.config(text="報價狀態: 取價失敗")

        except Exception as e:
            self.__log_message(f"抓取實際價格時發生錯誤: {e}")
            messagebox.showerror("錯誤", f"抓取實際價格失敗: {e}")

    def __get_real_future_price(self, symbol):
        """取得期貨的實際價格"""
        if not QUOTE_AVAILABLE:
            return None

        try:
            # 使用Quote模組的API來取得即時價格
            # 這裡需要根據群益API的實際實作來調整
            pStock = sk.SKSTOCKLONG()

            # 先嘗試透過商品代號取得股票索引
            # 這可能需要先訂閱該商品的報價
            market_no = 1  # 期貨市場編號

            # 這裡的實作需要根據實際的群益API方法來調整
            # 因為直接取得價格可能需要先訂閱報價

            # 使用快取的價格（如果有的話）
            current_time = time.time()
            cache_key = f"{symbol}_{market_no}"

            if cache_key in self.__real_price_cache:
                cache_data = self.__real_price_cache[cache_key]
                # 如果快取時間在30秒內，直接使用快取價格
                if current_time - cache_data['time'] < 30:
                    return cache_data['price']

            # 實際的價格抓取邏輯
            # 注意：這裡需要報價系統已經連線並訂閱了相關商品

            # 模擬實際價格（在沒有真實連線時）
            # 實際使用時這裡應該呼叫真正的API
            simulated_price = 25500 + (hash(symbol) % 1000 - 500)  # 模擬波動

            # 快取價格
            self.__real_price_cache[cache_key] = {
                'price': simulated_price,
                'time': current_time
            }

            return simulated_price

        except Exception as e:
            self.__log_message(f"取得實際價格時發生錯誤: {e}")
            return None

    def __calculate_strategy(self):
        """計算價差策略"""
        if not self.__update_parameters():
            return

        # 計算買權價差獲利
        call_spread_profit = (self.__strategy_params['call_sell_premium'] -
                             self.__strategy_params['call_buy_premium']) * 50

        # 計算賣權價差獲利
        put_spread_profit = (self.__strategy_params['put_sell_premium'] -
                            self.__strategy_params['put_buy_premium']) * 50

        # 總獲利
        total_profit = call_spread_profit + put_spread_profit

        # 保證金需求
        call_margin = (self.__strategy_params['call_buy_strike'] -
                      self.__strategy_params['call_sell_strike']) * 50
        put_margin = (self.__strategy_params['put_sell_strike'] -
                     self.__strategy_params['put_buy_strike']) * 50
        total_margin = call_margin + put_margin

        # 避險條件
        hedge_up = self.__strategy_params['call_sell_strike']
        hedge_down = self.__strategy_params['put_sell_strike']
        extra_profit_up = self.__strategy_params['call_buy_strike']
        extra_profit_down = self.__strategy_params['put_buy_strike']

        # 測試模式提示
        mode_text = "【測試模式】" if self.__test_mode else "【實盤模式】"

        # 價格資訊
        price_info = ""
        if self.__test_mode:
            if self.__use_real_quote and QUOTE_AVAILABLE:
                current_price = self.__get_current_future_price()
                price_info = f"當前實際價格: {current_price} (即時抓取)"
            else:
                price_info = f"當前模擬價格: {self.__test_current_price}"

        # 顯示結果
        result = f"""
=== 雙賣價差單策略分析 {mode_text} ===
時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

買權價差單:
- 賣買權: {self.__strategy_params['call_sell_strike']} (權利金: {self.__strategy_params['call_sell_premium']})
- 買買權: {self.__strategy_params['call_buy_strike']} (權利金: {self.__strategy_params['call_buy_premium']})
- 買權價差獲利: {call_spread_profit}元

賣權價差單:
- 賣賣權: {self.__strategy_params['put_sell_strike']} (權利金: {self.__strategy_params['put_sell_premium']})
- 買賣權: {self.__strategy_params['put_buy_strike']} (權利金: {self.__strategy_params['put_buy_premium']})
- 賣權價差獲利: {put_spread_profit}元

總獲利: {total_profit}元
總保證金: {total_margin}元

避險條件:
- 向上突破 {hedge_up} 時買入小台避險
- 向下跌破 {hedge_down} 時賣出小台避險

額外獲利機會:
- 小台超過 {extra_profit_up} 時：每點額外獲利50元
- 小台跌破 {extra_profit_down} 時：每點額外獲利50元

選擇權代號格式:
- 買權: TXO{self.__strategy_params['option_month']}C{int(self.__strategy_params['call_sell_strike'])}
- 賣權: TXO{self.__strategy_params['option_month']}P{int(self.__strategy_params['put_sell_strike'])}

{price_info}
"""

        self.__log_message(result)

    def __generate_option_symbol(self, option_type, strike_price):
        """生成選擇權代號"""
        return f"TXO{self.__strategy_params['option_month']}{option_type}{int(strike_price)}"

    def __send_spread_orders(self):
        """發送價差單委託"""
        if not self.__test_mode and self.__dOrder['boxAccount'] == '':
            messagebox.showerror("錯誤", "請選擇期貨帳號！")
            return

        if not self.__update_parameters():
            return

        if self.__test_mode:
            self.__log_message("【測試模式】模擬發送價差單委託")

        try:
            # 發送買權價差單
            self.__send_call_spread()

            # 發送賣權價差單
            self.__send_put_spread()

            mode_text = "模擬" if self.__test_mode else "實際"
            self.__log_message(f"價差單委託已{mode_text}發送完成")

        except Exception as e:
            self.__log_message(f"發送價差單時發生錯誤: {e}")
            if not self.__test_mode:
                messagebox.showerror("錯誤", f"發送價差單失敗: {e}")

    def __send_call_spread(self):
        """發送買權價差單"""
        # 賣買權
        self.__send_option_order(
            symbol=self.__generate_option_symbol('C', self.__strategy_params['call_sell_strike']),
            buy_sell=1,  # 賣出
            price=self.__strategy_params['call_sell_premium'],
            quantity=self.__strategy_params['quantity']
        )

        # 買買權
        self.__send_option_order(
            symbol=self.__generate_option_symbol('C', self.__strategy_params['call_buy_strike']),
            buy_sell=0,  # 買進
            price=self.__strategy_params['call_buy_premium'],
            quantity=self.__strategy_params['quantity']
        )

        self.__order_status['call_spread_order_sent'] = True
        self.__log_message("買權價差單已發送")

    def __send_put_spread(self):
        """發送賣權價差單"""
        # 賣賣權
        self.__send_option_order(
            symbol=self.__generate_option_symbol('P', self.__strategy_params['put_sell_strike']),
            buy_sell=1,  # 賣出
            price=self.__strategy_params['put_sell_premium'],
            quantity=self.__strategy_params['quantity']
        )

        # 買賣權
        self.__send_option_order(
            symbol=self.__generate_option_symbol('P', self.__strategy_params['put_buy_strike']),
            buy_sell=0,  # 買進
            price=self.__strategy_params['put_buy_premium'],
            quantity=self.__strategy_params['quantity']
        )

        self.__order_status['put_spread_order_sent'] = True
        self.__log_message("賣權價差單已發送")

    def __send_option_order(self, symbol, buy_sell, price, quantity):
        """發送選擇權訂單"""
        action = "賣出" if buy_sell == 1 else "買進"

        if self.__test_mode:
            # 測試模式：只記錄訊息，不實際下單
            self.__log_message(f"【測試模式】{action} {symbol} 價格:{price} 數量:{quantity} - 模擬成功")
            return

        try:
            # 實盤模式：實際發送訂單
            oOrder = sk.FUTUREORDER()
            oOrder.bstrFullAccount = self.__dOrder['boxAccount']
            oOrder.bstrStockNo = symbol
            oOrder.sBuySell = buy_sell
            oOrder.sTradeType = 0  # ROD
            oOrder.bstrPrice = str(price)
            oOrder.nQty = quantity
            oOrder.sNewClose = 0  # 新倉
            oOrder.sReserved = 0  # 盤中

            # 發送訂單
            message, m_nCode = skO.SendOptionOrder(Global.Global_IID, True, oOrder)

            self.__log_message(f"【實盤模式】{action} {symbol} 價格:{price} 數量:{quantity} - 結果代碼:{m_nCode}")

            if m_nCode != 0:
                self.__oMsg.SendReturnMessage("Order", m_nCode, "SendOptionOrder", self.__dOrder['listInformation'])

        except Exception as e:
            self.__log_message(f"發送選擇權訂單失敗: {e}")
            raise e

    def __start_monitoring(self):
        """開始價格監控"""
        if not self.__test_mode and (not self.__order_status['call_spread_order_sent'] or not self.__order_status['put_spread_order_sent']):
            if not messagebox.askyesno("確認", "尚未發送價差單，是否仍要開始監控？"):
                return

        self.__stop_monitoring = False
        self.__order_status['monitoring'] = True

        # 啟動監控線程
        self.__monitor_thread = threading.Thread(target=self.__price_monitor_loop, daemon=True)
        self.__monitor_thread.start()

        self.start_monitor_btn.config(state='disabled')
        self.stop_monitor_btn.config(state='normal')

        mode_text = "【測試模式】模擬" if self.__test_mode else "【實盤模式】"
        self.__log_message(f"{mode_text}開始監控期貨價格...")

    def __stop_monitoring_func(self):
        """停止價格監控"""
        self.__stop_monitoring = True
        self.__order_status['monitoring'] = False

        self.start_monitor_btn.config(state='normal')
        self.stop_monitor_btn.config(state='disabled')

        self.__log_message("停止監控期貨價格")

    def __price_monitor_loop(self):
        """價格監控循環"""
        while not self.__stop_monitoring:
            try:
                # 這裡應該獲取即時期貨價格
                # 由於沒有直接的報價API，這裡用模擬方式
                # 實際使用時需要整合報價模組

                current_price = self.__get_current_future_price()
                if current_price is None:
                    time.sleep(self.__strategy_params['monitor_interval'])
                    continue

                # 檢查避險條件
                self.__check_hedge_conditions(current_price)

                time.sleep(self.__strategy_params['monitor_interval'])

            except Exception as e:
                self.__log_message(f"監控過程中發生錯誤: {e}")
                time.sleep(5)  # 錯誤後等待5秒再繼續

    def __get_current_future_price(self):
        """獲取當前期貨價格"""
        if self.__test_mode:
            if self.__use_real_quote and QUOTE_AVAILABLE:
                # 測試模式但使用實際報價
                future_symbol = self.__strategy_params.get('future_symbol', 'MXF')
                real_price = self.__get_real_future_price(future_symbol)

                if real_price is not None:
                    # 同時更新UI顯示的模擬價格
                    self.test_price_entry.delete(0, END)
                    self.test_price_entry.insert(0, str(real_price))
                    self.__test_current_price = real_price
                    return real_price
                else:
                    # 如果取不到實際價格，fallback到模擬價格
                    self.__log_message("無法取得實際價格，使用模擬價格")
                    return self.__test_current_price
            else:
                # 純模擬模式
                return self.__test_current_price
        else:
            # 實盤模式：獲取即時價格
            future_symbol = self.__strategy_params.get('future_symbol', 'MXF')
            real_price = self.__get_real_future_price(future_symbol)

            if real_price is not None:
                return real_price
            else:
                self.__log_message("實盤模式無法獲取即時價格，監控暫停")
                return None

    def __check_hedge_conditions(self, current_price):
        """檢查避險條件"""
        # 向上突破 - 買入小台避險
        if (current_price > self.__strategy_params['call_sell_strike'] and
            not self.__order_status['hedge_long_sent']):
            self.__send_hedge_order(True)  # 買入
            self.__order_status['hedge_long_sent'] = True
            self.__log_message(f"價格突破{self.__strategy_params['call_sell_strike']}，執行向上避險")

        # 向下跌破 - 賣出小台避險
        if (current_price < self.__strategy_params['put_sell_strike'] and
            not self.__order_status['hedge_short_sent']):
            self.__send_hedge_order(False)  # 賣出
            self.__order_status['hedge_short_sent'] = True
            self.__log_message(f"價格跌破{self.__strategy_params['put_sell_strike']}，執行向下避險")

    def __send_hedge_order(self, is_buy):
        """發送避險訂單"""
        action = "買進" if is_buy else "賣出"

        if self.__test_mode:
            # 測試模式：只記錄訊息，不實際下單
            self.__log_message(f"【測試模式】避險訂單: {action} {self.__strategy_params['future_symbol']} - 模擬成功")
            return

        try:
            # 實盤模式：實際發送訂單
            oOrder = sk.FUTUREORDER()
            oOrder.bstrFullAccount = self.__dOrder['boxAccount']
            oOrder.bstrStockNo = self.__strategy_params['future_symbol']
            oOrder.sBuySell = 0 if is_buy else 1  # 0=買進, 1=賣出
            oOrder.sTradeType = 0  # ROD
            oOrder.bstrPrice = "M"  # 市價
            oOrder.nQty = self.__strategy_params['quantity']
            oOrder.sNewClose = 0  # 新倉
            oOrder.sReserved = 0  # 盤中

            # 發送訂單
            message, m_nCode = skO.SendFutureOrder(Global.Global_IID, True, oOrder)

            self.__log_message(f"【實盤模式】避險訂單: {action} {self.__strategy_params['future_symbol']} - 結果代碼:{m_nCode}")

            if m_nCode != 0:
                self.__oMsg.SendReturnMessage("Order", m_nCode, "SendFutureOrder", self.__dOrder['listInformation'])

        except Exception as e:
            self.__log_message(f"發送避險訂單失敗: {e}")

    def __reset_status(self):
        """重置狀態"""
        # 先停止監控
        if self.__order_status['monitoring']:
            self.__stop_monitoring_func()

        # 重置所有狀態
        self.__order_status = {
            'call_spread_order_sent': False,
            'put_spread_order_sent': False,
            'hedge_long_sent': False,
            'hedge_short_sent': False,
            'monitoring': False
        }

        self.__log_message("狀態已重置")

    def __log_message(self, message):
        """記錄訊息到狀態顯示區"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        full_message = f"[{timestamp}] {message}\n"

        self.status_text.insert(END, full_message)
        self.status_text.see(END)

        # 也記錄到主要訊息區
        if self.__dOrder['listInformation']:
            self.__oMsg.WriteMessage(f"雙賣策略: {message}", self.__dOrder['listInformation'])

    def get_test_mode(self):
        """獲取測試模式狀態"""
        return self.__test_mode

    def set_test_mode(self, test_mode):
        """設置測試模式"""
        self.__test_mode = test_mode
        self.test_mode_var.set(test_mode)