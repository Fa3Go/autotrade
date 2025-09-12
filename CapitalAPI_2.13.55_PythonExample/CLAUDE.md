# CLAUDE.md

此檔案為 Claude Code (claude.ai/code) 在此程式庫中工作時的指導文件。

## 專案概述

這是群益證券 API Python 範例專案（版本 2.13.55），用於台灣股票/期貨交易。此專案提供透過 COM 元件與群益證券交易 API 互動的 Python 範例。

## 專案結構

程式庫分為兩個主要範例版本：

- **PythonExample/**: 原始單一檔案範例，所有元件整合在一起
- **PythonExampleV2/**: 模組化版本，依功能分離

### 核心架構

#### COM 元件整合
- 使用 `comtypes` 函式庫與群益證券的 SKCOM.dll 介接
- 全域初始化五個主要 COM 物件：
  - `skC`: SKCenterLib（登入、連線管理）
  - `skO`: SKOrderLib（下單與訂單管理）
  - `skQ`: SKQuoteLib（市場資料與報價）
  - `skR`: SKReplyLib（訊息處理）
  - `skOSQ`: SKOSQuoteLib（海外股票報價）
  - `skOOQ`: SKOOQuoteLib（海外選擇權報價）

#### 事件驅動架構
- 每個 COM 元件都有對應的事件處理器類別
- 事件回調處理即時資料更新、訂單確認和系統通知
- 透過事件處理器更新全域訊息顯示

### 主要元件

#### 身份驗證與連線
- 透過 SKCenterLib 使用用戶憑證處理登入
- 支援不同環境（正式、測試、SGX）
- 特定帳戶類型的憑證驗證
- 訂單路由的代理伺服器連線管理

#### 訂單管理（PythonExample/order_service/）
- **Order.py**: 主要訂單介面，包含 GUI（登入、帳戶管理、訂單分頁）
- **StockOrder.py**: 股票交易訂單
- **FutureOrder.py**: 期貨交易訂單
- **OptionOrder.py**: 選擇權交易訂單
- **SeaFutureOrder.py**: 海外期貨訂單
- **SeaOptionOrder.py**: 海外選擇權訂單
- **ForeignStockOrder.py**: 複委託訂單
- **StockSmartTrade.py**: 智慧交易策略
- **StopLossOrderGui.py**: 停損單管理
- **SendMITOrder.py**: MIT（觸價成交）訂單

#### 市場資料（PythonExample/Quote_Service/）
- **Quote.py**: 即時市場資料介面，包含 GUI
- 支援多種資料類型：報價、逐筆、K線資料、技術指標（MACD、布林通道）
- 市場資訊與統計資料
- 最佳五檔買賣資料顯示

#### 設定檔
- **Config.py**: 包含下拉選單選項和市場類型定義
- GUI 元件和 API 參數的共用設定

### 安裝與設定

#### 前置需求
1. 從 `元件/` 目錄安裝群益證券 COM 元件
2. 執行適當的安裝程式：
   - `元件/x64/install.bat` 適用於 64 位元系統
   - `元件/x86/install.bat` 適用於 32 位元系統

#### Python 相依性
- `comtypes`: COM 元件介面
- `tkinter`: GUI 框架（Python 內建）
- 標準函式庫模組：`os`、`time`、`math`

### 執行範例

#### 訂單管理
```bash
python PythonExample/order_service/Order.py
# 或
python PythonExampleV2/Login/LoginForm.py
```

#### 市場資料
```bash
python PythonExample/Quote_Service/Quote.py
```

### 重要注意事項

#### 開發環境
- 僅支援 Windows（因 COM 相依性）
- 需要安裝群益證券 API
- 測試需要有效的交易帳戶憑證
- 可使用不同環境：正式、測試、SGX

#### API 限制
- 速率限制：可設定每秒訂單數量與筆數限制
- 市場特定的解鎖需求
- 透過代理伺服器的連線管理
- 即時資料需要保持連線狀態

#### 檔案組織
- V2 範例較為模組化，分離關注點
- 原始範例為單體式但展示完整工作流程
- 兩個版本共享相同的底層 API 結構

### 錯誤處理
- API 呼叫的回傳代碼指示成功/失敗
- 透過 `SKCenterLib_GetReturnCodeMessage()` 取得錯誤訊息
- 透過 `SKCenterLib_GetLastLogInfo()` 存取最後的日誌資訊

### 日誌記錄
- 日誌檔案儲存在 `CapitalLog_Order/` 和 `CapitalLog_Quote/` 目錄
- 透過 `SKCenterLib_SetLogPath()` 設定日誌路徑
- 提供上傳功能用於除錯