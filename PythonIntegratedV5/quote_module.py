# 報價模組
import comtypes.client
comtypes.client.GetModule(r'SKCOM.dll')
import comtypes.gen.SKCOMLib as sk
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
import threading
import time

# 群益API元件導入Python code內用的物件宣告
m_pSKCenter = comtypes.client.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
m_pSKQuote = comtypes.client.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib)

class SKQuoteLibEvent():
    def OnConnection(self, nKind, nCode):
        msg = f"【OnConnection】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nKind)}_{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
        print(msg)

    def OnNotifyServerTime(self, sHour, sMinute, sSecond, nTotal):
        msg = f"【OnNotifyServerTime】{sHour}:{sMinute}:{sSecond} 總秒數:{nTotal}"
        print(msg)

    def OnNotifyQuote(self, sMarketNo, sStockidx, sPtr, sSimulate, nCode, nDate, nTimehms, nTimemillismicros, nBid, nAsk, nClose, nQty, nSimulate):
        # 即時報價事件
        stock_info = {
            'market': sMarketNo,
            'stock_id': sStockidx,
            'date': nDate,
            'time': nTimehms,
            'bid': nBid / 100.0,
            'ask': nAsk / 100.0,
            'close': nClose / 100.0,
            'qty': nQty
        }
        if hasattr(self, 'quote_callback'):
            self.quote_callback(stock_info)

    def OnNotifyHistoryTick(self, sMarketNo, sStockidx, nPtr, nDate, nTimehms, nTimemillismicros, nBid, nAsk, nClose, nQty, nSimulate):
        # 歷史Tick事件
        tick_info = {
            'market': sMarketNo,
            'stock_id': sStockidx,
            'date': nDate,
            'time': nTimehms,
            'bid': nBid / 100.0,
            'ask': nAsk / 100.0,
            'close': nClose / 100.0,
            'qty': nQty
        }
        if hasattr(self, 'history_callback'):
            self.history_callback(tick_info)

    def OnNotifyTicks(self, sMarketNo, sStockidx, nPtr, nDate, nTimehms, nTimemillismicros, nBid, nAsk, nClose, nQty, nSimulate):
        # Tick報價事件
        pass

    def OnNotifyBest5(self, sMarketNo, sStockidx, nPtr, nBestBid1, nBestBidQty1, nBestBid2, nBestBidQty2, nBestBid3, nBestBidQty3, nBestBid4, nBestBidQty4, nBestBid5, nBestBidQty5, nExtendBid, nExtendBidQty, nBestAsk1, nBestAskQty1, nBestAsk2, nBestAskQty2, nBestAsk3, nBestAskQty3, nBestAsk4, nBestAskQty4, nBestAsk5, nBestAskQty5, nExtendAsk, nExtendAskQty, nSimulate):
        # 五檔報價事件
        best5_info = {
            'market': sMarketNo,
            'stock_id': sStockidx,
            'bid_prices': [nBestBid1/100.0, nBestBid2/100.0, nBestBid3/100.0, nBestBid4/100.0, nBestBid5/100.0],
            'bid_qtys': [nBestBidQty1, nBestBidQty2, nBestBidQty3, nBestBidQty4, nBestBidQty5],
            'ask_prices': [nBestAsk1/100.0, nBestAsk2/100.0, nBestAsk3/100.0, nBestAsk4/100.0, nBestAsk5/100.0],
            'ask_qtys': [nBestAskQty1, nBestAskQty2, nBestAskQty3, nBestAskQty4, nBestAskQty5]
        }
        if hasattr(self, 'best5_callback'):
            self.best5_callback(best5_info)

# 初始化事件處理
SKQuoteEvent = SKQuoteLibEvent()
SKQuoteLibEventHandler = comtypes.client.GetEvents(m_pSKQuote, SKQuoteEvent)

class QuoteManager:
    def __init__(self):
        self.is_connected = False
        self.subscribed_stocks = set()
        self.quote_data = {}

    def connect_quote(self, user_id):
        """連線報價伺服器"""
        try:
            nCode = m_pSKQuote.SKQuoteLib_EnterMonitor()
            msg = f"【SKQuoteLib_EnterMonitor】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
            print(msg)

            if nCode == 0:
                self.is_connected = True
                return True
            return False
        except Exception as e:
            print(f"連線報價伺服器錯誤: {e}")
            return False

    def disconnect_quote(self):
        """中斷報價連線"""
        try:
            nCode = m_pSKQuote.SKQuoteLib_LeaveMonitor()
            msg = f"【SKQuoteLib_LeaveMonitor】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
            print(msg)
            self.is_connected = False
            return nCode == 0
        except Exception as e:
            print(f"中斷報價連線錯誤: {e}")
            return False

    def request_stocks(self, stock_list, request_type=1):
        """請求股票報價
        request_type: 1=即時報價, 2=tick, 3=最佳五檔, 4=歷史tick
        """
        try:
            # 將股票代碼列表轉換為字符串，用逗號分隔
            stock_str = ",".join(stock_list)

            if request_type == 1:  # 即時報價
                nCode = m_pSKQuote.SKQuoteLib_RequestStocks(stock_str)
            elif request_type == 2:  # tick報價
                nCode = m_pSKQuote.SKQuoteLib_RequestTicks(stock_str)
            elif request_type == 3:  # 最佳五檔
                nCode = m_pSKQuote.SKQuoteLib_RequestStockList(stock_str)

            msg = f"【RequestStocks】{m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)}"
            print(msg)

            if nCode == 0:
                for stock in stock_list:
                    self.subscribed_stocks.add(stock)
                return True
            return False
        except Exception as e:
            print(f"請求股票報價錯誤: {e}")
            return False

    def get_stock_by_no(self, market, stock_no):
        """根據市場別和股票代碼取得股票資訊"""
        try:
            nCode, stock_info = m_pSKQuote.SKQuoteLib_GetStockByNo(market, stock_no)
            if nCode == 0:
                return stock_info
            return None
        except Exception as e:
            print(f"取得股票資訊錯誤: {e}")
            return None

class QuoteWindow:
    def __init__(self, root, user_id=None):
        self.root = root
        self.user_id = user_id
        self.quote_manager = QuoteManager()
        self.quote_data = {}
        self.setup_ui()
        self.setup_callbacks()

    def setup_ui(self):
        self.root.title("Capital API 即時報價")
        self.root.geometry("800x600")

        # 連線狀態
        status_frame = tk.Frame(self.root)
        status_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(status_frame, text="連線狀態:").pack(side=tk.LEFT)
        self.label_status = tk.Label(status_frame, text="未連線", fg="red")
        self.label_status.pack(side=tk.LEFT, padx=5)

        self.btn_connect = tk.Button(status_frame, text="連線報價", command=self.connect_quote)
        self.btn_connect.pack(side=tk.LEFT, padx=5)

        self.btn_disconnect = tk.Button(status_frame, text="中斷連線", command=self.disconnect_quote, state="disabled")
        self.btn_disconnect.pack(side=tk.LEFT, padx=5)

        # 股票訂閱輸入
        input_frame = tk.Frame(self.root)
        input_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(input_frame, text="股票代碼:").pack(side=tk.LEFT)
        self.entry_stock = tk.Entry(input_frame, width=20)
        self.entry_stock.pack(side=tk.LEFT, padx=5)
        self.entry_stock.insert(0, "2330,0050")  # 預設台積電和0050

        # 報價類型選擇
        tk.Label(input_frame, text="報價類型:").pack(side=tk.LEFT, padx=(10,0))
        self.combo_type = ttk.Combobox(input_frame, values=["即時報價", "Tick報價", "最佳五檔"], state="readonly")
        self.combo_type.current(0)
        self.combo_type.pack(side=tk.LEFT, padx=5)

        self.btn_subscribe = tk.Button(input_frame, text="訂閱股票", command=self.subscribe_stocks, state="disabled")
        self.btn_subscribe.pack(side=tk.LEFT, padx=5)

        # 報價顯示區域
        self.setup_quote_display()

    def setup_quote_display(self):
        # 建立樹狀檢視來顯示報價資料
        columns = ('股票代碼', '股票名稱', '現價', '買價', '賣價', '成交量', '時間')
        self.tree = ttk.Treeview(self.root, columns=columns, show='tree headings', height=15)

        # 設定欄位寬度
        self.tree.column('#0', width=0, stretch=False)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor='center')

        # 捲軸
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # 打包
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=10)

        # 五檔顯示區域
        self.setup_best5_display()

    def setup_best5_display(self):
        # 五檔報價顯示視窗
        self.best5_window = None

    def setup_callbacks(self):
        """設定回調函數"""
        SKQuoteEvent.quote_callback = self.on_quote_update
        SKQuoteEvent.best5_callback = self.on_best5_update

    def connect_quote(self):
        """連線報價伺服器"""
        if self.quote_manager.connect_quote(self.user_id):
            self.label_status.config(text="已連線", fg="green")
            self.btn_connect.config(state="disabled")
            self.btn_disconnect.config(state="normal")
            self.btn_subscribe.config(state="normal")
        else:
            messagebox.showerror("錯誤", "連線報價伺服器失敗")

    def disconnect_quote(self):
        """中斷報價連線"""
        if self.quote_manager.disconnect_quote():
            self.label_status.config(text="已中斷", fg="red")
            self.btn_connect.config(state="normal")
            self.btn_disconnect.config(state="disabled")
            self.btn_subscribe.config(state="disabled")

    def subscribe_stocks(self):
        """訂閱股票報價"""
        stock_list = [s.strip() for s in self.entry_stock.get().split(',') if s.strip()]
        if not stock_list:
            messagebox.showwarning("警告", "請輸入股票代碼")
            return

        request_type = self.combo_type.current() + 1
        if self.quote_manager.request_stocks(stock_list, request_type):
            messagebox.showinfo("成功", f"已訂閱股票: {', '.join(stock_list)}")
        else:
            messagebox.showerror("錯誤", "訂閱股票失敗")

    def on_quote_update(self, stock_info):
        """即時報價更新回調"""
        stock_id = stock_info['stock_id']

        # 更新內部資料
        self.quote_data[stock_id] = stock_info

        # 更新樹狀檢視
        self.root.after(0, self.update_tree_display, stock_info)

    def on_best5_update(self, best5_info):
        """五檔報價更新回調"""
        # 可以開啟新視窗顯示五檔資料
        if not self.best5_window or not self.best5_window.winfo_exists():
            self.show_best5_window()

        # 更新五檔視窗資料
        self.root.after(0, self.update_best5_display, best5_info)

    def update_tree_display(self, stock_info):
        """更新樹狀檢視顯示"""
        stock_id = stock_info['stock_id']

        # 尋找是否已存在此股票
        item_id = None
        for item in self.tree.get_children():
            if self.tree.item(item)['values'][0] == stock_id:
                item_id = item
                break

        # 格式化時間
        time_str = f"{stock_info['time']:06d}"
        formatted_time = f"{time_str[:2]}:{time_str[2:4]}:{time_str[4:6]}"

        values = (
            stock_id,
            stock_id,  # 股票名稱暫時用代碼代替
            f"{stock_info['close']:.2f}",
            f"{stock_info['bid']:.2f}",
            f"{stock_info['ask']:.2f}",
            stock_info['qty'],
            formatted_time
        )

        if item_id:
            # 更新現有項目
            self.tree.item(item_id, values=values)
        else:
            # 新增項目
            self.tree.insert('', 'end', values=values)

    def show_best5_window(self):
        """顯示五檔報價視窗"""
        self.best5_window = tk.Toplevel(self.root)
        self.best5_window.title("最佳五檔報價")
        self.best5_window.geometry("400x300")

    def update_best5_display(self, best5_info):
        """更新五檔顯示"""
        # 在五檔視窗中更新資料
        pass

def main(user_id=None):
    """啟動報價視窗"""
    root = tk.Tk()
    app = QuoteWindow(root, user_id)
    root.mainloop()

if __name__ == "__main__":
    main()