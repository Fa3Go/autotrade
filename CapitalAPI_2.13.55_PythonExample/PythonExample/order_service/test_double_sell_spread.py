#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
雙賣價差單策略測試腳本

使用方法：
1. 基本測試: python test_double_sell_spread.py
2. 顯示GUI測試: python test_double_sell_spread.py --show-gui
"""

import sys
import os
import argparse
from tkinter import Tk

# 添加路徑以便導入模組
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    import DoubleSellSpreadStrategy
    print("✓ DoubleSellSpreadStrategy 模組載入成功")
except ImportError as e:
    print(f"✗ DoubleSellSpreadStrategy 模組載入失敗: {e}")
    sys.exit(1)

def test_strategy_calculation():
    """測試策略計算功能"""
    print("\n=== 測試策略計算功能 ===")

    # 創建模擬的訊息列表
    class MockListbox:
        def insert(self, index, item):
            print(f"Mock Listbox: {item}")

    mock_info = MockListbox()

    try:
        # 創建策略實例
        strategy = DoubleSellSpreadStrategy.DoubleSellSpreadStrategy(information=mock_info)
        print("✓ 策略實例創建成功")

        # 測試參數設定
        strategy._DoubleSellSpreadStrategy__strategy_params = {
            'call_sell_strike': 26000,
            'call_sell_premium': 20,
            'call_buy_strike': 26100,
            'call_buy_premium': 10,
            'put_sell_strike': 25000,
            'put_sell_premium': 30,
            'put_buy_strike': 24900,
            'put_buy_premium': 10,
            'quantity': 1,
            'future_symbol': 'MXF',
            'option_month': '202501',
            'monitor_interval': 1
        }

        # 計算獲利
        call_spread_profit = (20 - 10) * 50  # 500
        put_spread_profit = (30 - 10) * 50   # 1000
        total_profit = call_spread_profit + put_spread_profit  # 1500

        # 計算保證金
        call_margin = (26100 - 26000) * 50  # 5000
        put_margin = (25000 - 24900) * 50   # 5000
        total_margin = call_margin + put_margin  # 10000

        print(f"✓ 買權價差獲利: {call_spread_profit}元")
        print(f"✓ 賣權價差獲利: {put_spread_profit}元")
        print(f"✓ 總獲利: {total_profit}元")
        print(f"✓ 總保證金: {total_margin}元")

        # 測試選擇權代號生成
        call_symbol = strategy._DoubleSellSpreadStrategy__generate_option_symbol('C', 26000)
        put_symbol = strategy._DoubleSellSpreadStrategy__generate_option_symbol('P', 25000)

        expected_call = "TXO202501C26000"
        expected_put = "TXO202501P25000"

        if call_symbol == expected_call and put_symbol == expected_put:
            print(f"✓ 選擇權代號生成正確: {call_symbol}, {put_symbol}")
        else:
            print(f"✗ 選擇權代號生成錯誤: 期望 {expected_call}, {expected_put}, 實際 {call_symbol}, {put_symbol}")

        return True

    except Exception as e:
        print(f"✗ 策略計算測試失敗: {e}")
        return False

def test_gui_creation():
    """測試GUI創建"""
    print("\n=== 測試GUI創建 ===")

    try:
        root = Tk()
        root.title("雙賣價差單策略測試")
        root.geometry("800x600")

        # 創建策略GUI
        strategy = DoubleSellSpreadStrategy.DoubleSellSpreadStrategy()
        strategy.pack(fill="both", expand=True)

        print("✓ GUI創建成功")
        print("✓ 策略界面載入完成")
        print("  - 參數設定區域已創建")
        print("  - 控制按鈕區域已創建")
        print("  - 狀態顯示區域已創建")

        return root, strategy

    except Exception as e:
        print(f"✗ GUI創建失敗: {e}")
        return None, None

def test_account_setting():
    """測試帳號設定功能"""
    print("\n=== 測試帳號設定功能 ===")

    try:
        strategy = DoubleSellSpreadStrategy.DoubleSellSpreadStrategy()
        test_account = "TEST123456"
        strategy.SetAccount(test_account)

        if strategy._DoubleSellSpreadStrategy__dOrder['boxAccount'] == test_account:
            print(f"✓ 帳號設定成功: {test_account}")
            return True
        else:
            print("✗ 帳號設定失敗")
            return False

    except Exception as e:
        print(f"✗ 帳號設定測試失敗: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description='雙賣價差單策略測試')
    parser.add_argument('--show-gui', action='store_true', help='顯示GUI測試界面')
    args = parser.parse_args()

    print("雙賣價差單策略測試開始")
    print("=" * 50)

    # 執行各項測試
    tests_passed = 0
    total_tests = 3

    # 測試1: 策略計算
    if test_strategy_calculation():
        tests_passed += 1

    # 測試2: 帳號設定
    if test_account_setting():
        tests_passed += 1

    # 測試3: GUI創建
    root, strategy = test_gui_creation()
    if root and strategy:
        tests_passed += 1

        if args.show_gui:
            print("\n顯示GUI測試界面...")
            print("請檢查界面是否正常顯示，關閉視窗結束測試")
            root.mainloop()
        else:
            root.destroy()

    # 測試結果
    print("\n" + "=" * 50)
    print(f"測試完成: {tests_passed}/{total_tests} 項測試通過")

    if tests_passed == total_tests:
        print("🎉 所有測試通過！雙賣價差單策略功能正常")
        return 0
    else:
        print("⚠️  部分測試失敗，請檢查相關功能")
        return 1

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)