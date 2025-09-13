# 統一設定檔 - 整合版本群益證券 API 配置
# Integrated Configuration for Capital Securities API

# SKCenterLib_SetAuthority 連線環境設定
comboBoxSKCenterLib_SetAuthority = [
    "正式環境",
    "正式環境SGX", 
    "測試環境",
    "測試環境SGX"
]

# Order Report 查詢種類
comboBoxGetOrderReport = [
    "1:全部",
    "2:有效", 
    "3:可消",
    "4:已消",
    "5:已成",
    "6:失敗",
    "7:合併同價格",
    "8:合併同商品",
    "9:預約"
]

# Fulfill Report 查詢種類
comboBoxGetFulfillReport = [
    "1:全部",
    "2:合併同書號",
    "3:合併同價格", 
    "4:合併同商品",
    "5:T+1成交回報"
]

# 市場種類 - 每秒委託量限制
comboBoxSetMaxQtynMarketType = [
    "0：TS(證券)",
    "1：TF(期貨)",
    "2：TO(選擇權)",
    "3：OS(複委託)",
    "4：OF(海外期貨)",
    "5：OO(海外選擇權)"
]

# 市場種類 - 每秒委託筆數限制
comboBoxSetMaxCountnMarketType = [
    "0：TS(證券)",
    "1：TF(期貨)",
    "2：TO(選擇權)",
    "3：OS(複委託)",
    "4：OF(海外期貨)", 
    "5：OO(海外選擇權)"
]

# 下單解鎖市場種類
comboBoxUnlockOrder = [
    "0：TS(證券)",
    "1：TF(期貨)",
    "2：TO(選擇權)",
    "3：OS(複委託)",
    "4：OF(海外期貨)",
    "5：OO(海外選擇權)"
]

# ========== 台股下單設定 ==========
# 買賣別
comboBoxTSBuySell = [
    "0:買進",
    "1:賣出"
]

# 倉別
comboBoxTSPeriod = [
    "0:ROD",
    "3:IOC",
    "4:FOK"
]

# 交易別
comboBoxTSTradeType = [
    "0:現股",
    "3:融資",
    "4:融券",
    "8:無券賣出"
]

# 價格旗標
comboBoxTSPriceFlag = [
    "0:限價", 
    "1:市價",
    "2:限價(LMT)"
]

# ========== 期貨下單設定 ==========
# 買賣別
comboBoxTFBuySell = [
    "0:買進",
    "1:賣出"
]

# 委託條件
comboBoxTFDayTrade = [
    "0:ROD",
    "3:IOC", 
    "4:FOK"
]

# 新平倉
comboBoxTFNewClose = [
    "0:新倉",
    "1:平倉",
    "2:自動"
]

# 價格旗標
comboBoxTFPriceFlag = [
    "0:限價",
    "1:市價"
]

# ========== 海外股票設定 ==========
# 買賣別
comboBoxOSBuySell = [
    "B:買進",
    "S:賣出"
]

# 交易別
comboBoxOSTradeType = [
    "0:普通股",
    "3:融資",
    "4:融券"
]

# 價格旗標
comboBoxOSPriceFlag = [
    "0:限價",
    "1:市價",
    "2:範圍市價",
    "3:停損限價",
    "4:停損市價"
]

# ========== 海外期貨設定 ==========
# 買賣別
comboBoxOFBuySell = [
    "B:買進",
    "S:賣出"
]

# 委託條件
comboBoxOFPeriod = [
    "R:ROD",
    "I:IOC",
    "F:FOK"
]

# 新平倉
comboBoxOFNewClose = [
    "0:新倉",
    "1:平倉"
]

# ========== 報價設定 ==========
# 報價請求類型
comboBoxQuoteRequestType = [
    "0:即時報價",
    "1:歷史報價"
]

# K線類型
comboBoxKLineType = [
    "0:1分K",
    "1:5分K", 
    "2:10分K",
    "3:15分K",
    "4:30分K",
    "5:60分K",
    "6:日K",
    "7:週K",
    "8:月K"
]

# 技術指標類型
comboBoxTechnicalIndex = [
    "0:KD",
    "1:MACD",
    "2:RSI",
    "3:布林通道"
]

# ========== 海外市場設定 ==========
# 海外市場別
comboBoxOverseaMarket = [
    "US:美股",
    "HK:港股",
    "JP:日股",
    "CN:陸股"
]

# 海外商品類型
comboBoxOverseaProductType = [
    "STK:股票",
    "FUT:期貨", 
    "OPT:選擇權",
    "IND:指數"
]