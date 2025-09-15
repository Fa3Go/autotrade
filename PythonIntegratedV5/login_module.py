# 登入模組
import comtypes.client
comtypes.client.GetModule(r'SKCOM.dll')
import comtypes.gen.SKCOMLib as sk
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
import Config

# 群益API元件導入Python code內用的物件宣告
m_pSKCenter = comtypes.client.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
m_pSKReply = comtypes.client.CreateObject(sk.SKReplyLib, interface=sk.ISKReplyLib)
m_pSKOrder = comtypes.client.CreateObject(sk.SKOrderLib, interface=sk.ISKOrderLib)

# 全域變數
dictUserID = {}
dictUserID["更新帳號"] = ["無"]

class SKReplyLibEvent():
    def OnReplyMessage(self, bstrUserID, bstrMessages):
        nConfirmCode = -1
        msg = "【註冊公告OnReplyMessage】" + bstrUserID + "_" + bstrMessages
        print(msg)
        return nConfirmCode

class SKCenterLibEvent:
    def OnTimer(self, nTime):
        msg = "【OnTimer】" + str(nTime)
        print(msg)

    def OnShowAgreement(self, bstrData):
        msg = "【OnShowAgreement】" + bstrData
        print(msg)

    def OnNotifySGXAPIOrderStatus(self, nStatus, bstrOFAccount):
        msg = "【OnNotifySGXAPIOrderStatus】" + str(nStatus) + "_" + bstrOFAccount
        print(msg)

class SKOrderLibEvent():
    def OnAccount(self, bstrLogInID, bstrAccountData):
        msg = "【OnAccount】" + bstrLogInID + "_" + bstrAccountData
        print(msg)

        values = bstrAccountData.split(',')
        Account = values[1] + values[3]

        if bstrLogInID in dictUserID:
            accountExists = False
            for value in dictUserID[bstrLogInID]:
                if value == Account:
                    accountExists = True
                    break
            if accountExists == False:
                dictUserID[bstrLogInID].append(Account)
        else:
            dictUserID[bstrLogInID] = [Account]

    def OnProxyStatus(self, bstrUserId, nCode):
        msg = "【OnProxyStatus】" + bstrUserId + "_" + m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
        print(msg)

    def OnTelnetTest(self, bstrData):
        msg = "【OnTelnetTest】" + bstrData
        print(msg)

# 初始化事件處理
SKReplyEvent = SKReplyLibEvent()
SKReplyLibEventHandler = comtypes.client.GetEvents(m_pSKReply, SKReplyEvent)

SKCenterEvent = SKCenterLibEvent()
SKCenterEventHandler = comtypes.client.GetEvents(m_pSKCenter, SKCenterEvent)

SKOrderEvent = SKOrderLibEvent()
SKOrderLibEventHandler = comtypes.client.GetEvents(m_pSKOrder, SKOrderEvent)

class LoginManager:
    def __init__(self):
        self.is_logged_in = False
        self.current_user = ""
        self.accounts = []

    def set_authority(self, environment="測試環境"):
        """設定連線環境"""
        authority_map = {
            "正式環境": 0,
            "正式環境SGX": 1,
            "測試環境": 2,
            "測試環境SGX": 3
        }

        nAuthorityFlag = authority_map.get(environment, 2)  # 預設測試環境
        nCode = m_pSKCenter.SKCenterLib_SetAuthority(nAuthorityFlag)
        msg = "【SKCenterLib_SetAuthority】" + m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
        print(msg)
        return nCode == 0

    def login(self, user_id, password):
        """登入功能"""
        try:
            # 先設定環境
            self.set_authority()

            # 執行登入
            nCode = m_pSKCenter.SKCenterLib_Login(user_id, password)
            msg = "【SKCenterLib_Login】" + m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
            print(msg)

            if nCode == 0:
                self.is_logged_in = True
                self.current_user = user_id
                # 初始化下單模組
                self.initialize_order()
                return True
            else:
                return False
        except Exception as e:
            print(f"登入錯誤: {e}")
            return False

    def initialize_order(self):
        """初始化下單模組"""
        try:
            nCode = m_pSKOrder.SKOrderLib_Initialize()
            msg = "【SKOrderLib_Initialize】" + m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
            print(msg)

            # 取得交易帳號
            nCode = m_pSKOrder.GetUserAccount()
            msg = "【GetUserAccount】" + m_pSKCenter.SKCenterLib_GetReturnCodeMessage(nCode)
            print(msg)

            return nCode == 0
        except Exception as e:
            print(f"初始化下單模組錯誤: {e}")
            return False

    def get_accounts(self):
        """取得交易帳號列表"""
        if self.current_user in dictUserID:
            return dictUserID[self.current_user]
        return []

    def logout(self):
        """登出"""
        self.is_logged_in = False
        self.current_user = ""
        self.accounts = []
        print("已登出")

# 登入視窗類別
class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.login_manager = LoginManager()
        self.setup_ui()

    def setup_ui(self):
        self.root.title("Capital API 登入")
        self.root.geometry("400x300")

        # 使用者帳號
        tk.Label(self.root, text="使用者帳號:").grid(row=0, column=0, padx=10, pady=10)
        self.entry_user = tk.Entry(self.root, width=20)
        self.entry_user.grid(row=0, column=1, padx=10, pady=10)

        # 密碼
        tk.Label(self.root, text="密碼:").grid(row=1, column=0, padx=10, pady=10)
        self.entry_password = tk.Entry(self.root, show="*", width=20)
        self.entry_password.grid(row=1, column=1, padx=10, pady=10)

        # 連線環境選擇
        tk.Label(self.root, text="連線環境:").grid(row=2, column=0, padx=10, pady=10)
        self.combo_env = ttk.Combobox(self.root, values=Config.comboBoxSKCenterLib_SetAuthority, state="readonly")
        self.combo_env.current(2)  # 預設測試環境
        self.combo_env.grid(row=2, column=1, padx=10, pady=10)

        # 登入按鈕
        self.btn_login = tk.Button(self.root, text="登入", command=self.do_login)
        self.btn_login.grid(row=3, column=0, columnspan=2, pady=20)

        # 狀態標籤
        self.label_status = tk.Label(self.root, text="請輸入帳號密碼", fg="blue")
        self.label_status.grid(row=4, column=0, columnspan=2, pady=10)

        # 交易帳號顯示
        tk.Label(self.root, text="交易帳號:").grid(row=5, column=0, padx=10, pady=10)
        self.combo_account = ttk.Combobox(self.root, state="readonly")
        self.combo_account.grid(row=5, column=1, padx=10, pady=10)

        # API 版本顯示
        version = m_pSKCenter.SKCenterLib_GetSKAPIVersionAndBit("xxxxxxxxxx")
        tk.Label(self.root, text=f"API版本: {version}", font=("Arial", 8)).grid(row=6, column=0, columnspan=2, pady=10)

    def do_login(self):
        user_id = self.entry_user.get().strip()
        password = self.entry_password.get().strip()
        environment = self.combo_env.get()

        if not user_id or not password:
            messagebox.showerror("錯誤", "請輸入帳號和密碼")
            return

        self.label_status.config(text="登入中...", fg="orange")
        self.root.update()

        # 設定環境
        self.login_manager.set_authority(environment)

        # 執行登入
        if self.login_manager.login(user_id, password):
            self.label_status.config(text="登入成功!", fg="green")
            self.btn_login.config(text="已登入", state="disabled")

            # 更新交易帳號下拉選單
            self.root.after(2000, self.update_accounts)  # 延遲2秒更新帳號
        else:
            self.label_status.config(text="登入失敗!", fg="red")

    def update_accounts(self):
        """更新交易帳號列表"""
        accounts = self.login_manager.get_accounts()
        if accounts and len(accounts) > 1:  # 排除"無"這個預設值
            self.combo_account['values'] = accounts[1:]  # 排除第一個"無"
            if len(accounts) > 1:
                self.combo_account.current(0)

if __name__ == "__main__":
    root = tk.Tk()
    app = LoginWindow(root)
    root.mainloop()