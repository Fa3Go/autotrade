# 主程式 - 整合登入和報價功能
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
from login_module import LoginManager, LoginWindow
from quote_module import QuoteManager, QuoteWindow
import threading

class MainApplication:
    def __init__(self, root):
        self.root = root
        self.login_manager = LoginManager()
        self.quote_window = None
        self.current_user = ""
        self.setup_main_ui()

    def setup_main_ui(self):
        """設定主介面"""
        self.root.title("Capital API 整合系統")
        self.root.geometry("500x400")

        # 主要容器
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 標題
        title_label = tk.Label(main_frame, text="Capital API 整合交易系統",
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))

        # 登入區塊
        self.setup_login_section(main_frame)

        # 功能區塊
        self.setup_function_section(main_frame)

        # 狀態區塊
        self.setup_status_section(main_frame)

    def setup_login_section(self, parent):
        """登入區塊"""
        login_frame = tk.LabelFrame(parent, text="登入資訊", font=("Arial", 12))
        login_frame.pack(fill=tk.X, pady=10)

        # 使用者帳號
        tk.Label(login_frame, text="使用者帳號:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.entry_user = tk.Entry(login_frame, width=20)
        self.entry_user.grid(row=0, column=1, padx=10, pady=10)

        # 密碼
        tk.Label(login_frame, text="密碼:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        self.entry_password = tk.Entry(login_frame, show="*", width=20)
        self.entry_password.grid(row=1, column=1, padx=10, pady=10)

        # 連線環境
        tk.Label(login_frame, text="連線環境:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
        self.combo_env = ttk.Combobox(login_frame, values=["正式環境", "正式環境SGX", "測試環境", "測試環境SGX"],
                                     state="readonly", width=17)
        self.combo_env.current(2)  # 預設測試環境
        self.combo_env.grid(row=2, column=1, padx=10, pady=10)

        # 登入按鈕
        self.btn_login = tk.Button(login_frame, text="登入", command=self.do_login,
                                  bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        self.btn_login.grid(row=3, column=0, columnspan=2, pady=10)

        # 登出按鈕
        self.btn_logout = tk.Button(login_frame, text="登出", command=self.do_logout,
                                   state="disabled", bg="#f44336", fg="white", font=("Arial", 10, "bold"))
        self.btn_logout.grid(row=4, column=0, columnspan=2, pady=5)

    def setup_function_section(self, parent):
        """功能區塊"""
        function_frame = tk.LabelFrame(parent, text="系統功能", font=("Arial", 12))
        function_frame.pack(fill=tk.X, pady=10)

        # 報價功能按鈕
        self.btn_quote = tk.Button(function_frame, text="開啟即時報價",
                                  command=self.open_quote_window, state="disabled",
                                  bg="#2196F3", fg="white", font=("Arial", 10, "bold"),
                                  width=20, height=2)
        self.btn_quote.pack(pady=10)

        # 下單功能按鈕 (暫時未實現)
        self.btn_order = tk.Button(function_frame, text="開啟下單功能 (待開發)",
                                  state="disabled", bg="#FF9800", fg="white",
                                  font=("Arial", 10, "bold"), width=20, height=2)
        self.btn_order.pack(pady=5)

        # 帳戶查詢按鈕
        self.btn_account = tk.Button(function_frame, text="查詢交易帳號",
                                    command=self.show_accounts, state="disabled",
                                    bg="#9C27B0", fg="white", font=("Arial", 10, "bold"),
                                    width=20, height=1)
        self.btn_account.pack(pady=5)

    def setup_status_section(self, parent):
        """狀態區塊"""
        status_frame = tk.LabelFrame(parent, text="系統狀態", font=("Arial", 12))
        status_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 登入狀態
        tk.Label(status_frame, text="登入狀態:", font=("Arial", 10)).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.label_login_status = tk.Label(status_frame, text="未登入", fg="red", font=("Arial", 10, "bold"))
        self.label_login_status.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        # 目前使用者
        tk.Label(status_frame, text="目前使用者:", font=("Arial", 10)).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.label_current_user = tk.Label(status_frame, text="無", font=("Arial", 10))
        self.label_current_user.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # 交易帳號
        tk.Label(status_frame, text="交易帳號:", font=("Arial", 10)).grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.combo_account = ttk.Combobox(status_frame, state="readonly", width=20)
        self.combo_account.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        # 系統訊息文字框
        tk.Label(status_frame, text="系統訊息:", font=("Arial", 10)).grid(row=3, column=0, padx=10, pady=(10,5), sticky="nw")
        self.text_message = tk.Text(status_frame, height=5, width=50)
        scrollbar = tk.Scrollbar(status_frame, command=self.text_message.yview)
        self.text_message.config(yscrollcommand=scrollbar.set)
        self.text_message.grid(row=3, column=1, padx=10, pady=(10,5), sticky="ew")
        scrollbar.grid(row=3, column=2, pady=(10,5), sticky="ns")

        # 設定列權重
        status_frame.columnconfigure(1, weight=1)

    def add_message(self, message):
        """添加系統訊息"""
        self.text_message.insert(tk.END, message + "\n")
        self.text_message.see(tk.END)
        self.root.update()

    def do_login(self):
        """執行登入"""
        user_id = self.entry_user.get().strip()
        password = self.entry_password.get().strip()
        environment = self.combo_env.get()

        if not user_id or not password:
            messagebox.showerror("錯誤", "請輸入帳號和密碼")
            return

        self.add_message("開始登入...")

        # 設定環境
        self.login_manager.set_authority(environment)

        # 執行登入
        if self.login_manager.login(user_id, password):
            self.current_user = user_id
            self.label_login_status.config(text="已登入", fg="green")
            self.label_current_user.config(text=user_id)

            # 啟用功能按鈕
            self.btn_quote.config(state="normal")
            self.btn_account.config(state="normal")
            self.btn_login.config(state="disabled")
            self.btn_logout.config(state="normal")

            self.add_message("登入成功!")

            # 延遲更新帳號列表
            self.root.after(2000, self.update_account_list)
        else:
            self.add_message("登入失敗!")
            messagebox.showerror("錯誤", "登入失敗，請檢查帳號密碼")

    def do_logout(self):
        """執行登出"""
        self.login_manager.logout()
        self.current_user = ""

        # 更新UI狀態
        self.label_login_status.config(text="未登入", fg="red")
        self.label_current_user.config(text="無")
        self.combo_account['values'] = []

        # 停用功能按鈕
        self.btn_quote.config(state="disabled")
        self.btn_account.config(state="disabled")
        self.btn_login.config(state="normal")
        self.btn_logout.config(state="disabled")

        # 關閉報價視窗
        if self.quote_window and hasattr(self.quote_window, 'winfo_exists') and self.quote_window.winfo_exists():
            self.quote_window.destroy()

        self.add_message("已登出")

    def update_account_list(self):
        """更新交易帳號列表"""
        accounts = self.login_manager.get_accounts()
        if accounts and len(accounts) > 1:  # 排除"無"這個預設值
            account_list = [acc for acc in accounts if acc != "無"]
            self.combo_account['values'] = account_list
            if account_list:
                self.combo_account.current(0)
                self.add_message(f"取得交易帳號: {', '.join(account_list)}")

    def show_accounts(self):
        """顯示交易帳號資訊"""
        accounts = self.login_manager.get_accounts()
        if accounts and len(accounts) > 1:
            account_list = [acc for acc in accounts if acc != "無"]
            account_text = '\n'.join(account_list) if account_list else "無交易帳號"
            messagebox.showinfo("交易帳號", f"目前可用的交易帳號:\n\n{account_text}")
        else:
            messagebox.showinfo("交易帳號", "尚未取得交易帳號資訊")

    def open_quote_window(self):
        """開啟報價視窗"""
        if not self.login_manager.is_logged_in:
            messagebox.showwarning("警告", "請先登入")
            return

        # 檢查報價視窗是否已開啟
        if self.quote_window and hasattr(self.quote_window, 'winfo_exists') and self.quote_window.winfo_exists():
            self.quote_window.lift()  # 將視窗提到前面
            return

        try:
            # 創建新的報價視窗
            self.quote_window = tk.Toplevel(self.root)
            quote_app = QuoteWindow(self.quote_window, self.current_user)
            self.add_message("已開啟即時報價視窗")
        except Exception as e:
            self.add_message(f"開啟報價視窗錯誤: {e}")
            messagebox.showerror("錯誤", f"開啟報價視窗失敗: {e}")

def main():
    """主程式進入點"""
    root = tk.Tk()
    app = MainApplication(root)

    # 設定視窗關閉事件
    def on_closing():
        if app.login_manager.is_logged_in:
            app.do_logout()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()