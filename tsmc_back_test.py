import xlwings as xw
from xlwings.constants import Direction
import numpy as np

def hello_xlwings():
    wb = xw.Book.caller()
    wb.sheets[0].range("A1").value = "Hello xlwings!"

def run_back_test():
    wb = xw.Book.caller()
    tsmc_sheet = wb.sheets["2330"]
    # back testing day 1
    stock_shares = tsmc_sheet.cells(14, "K").value
    last_row = tsmc_sheet.cells(1, "A").end(Direction.xlDown).row

    tsmc_sheet.cells(4, "D").value = stock_shares
    tsmc_sheet.cells(4, "E").value = 0
    tsmc_sheet.cells(4, "F").value = tsmc_sheet.cells(4, "D").value
    tsmc_sheet.cells(4, "G").value = tsmc_sheet.cells(13, "K").value - stock_shares * tsmc_sheet.cells(4, "B").value
    tsmc_sheet.cells(4, "H").value = tsmc_sheet.cells(4, "G").value + tsmc_sheet.cells(4, "F").value * tsmc_sheet.cells(4, "B").value
    # 實作交易策略
    for i in range(5, last_row+1):
        # 截取當天的 3日移動平均
        sma_3d = tsmc_sheet.cells(i, "C").value
        # 截取當天收盤價
        price_today = tsmc_sheet.cells(i, 'B').value
        # 若 5日 > 10日，而且我有足夠買入以今日收盤價計價的 1000 股的現金，就買入 1000 股（在 E 欄顯示 1000）
        if (price_today > sma_3d) and (tsmc_sheet.cells(i-1, "G").value > price_today * stock_shares):
            tsmc_sheet.cells(i, "D").value = stock_shares
        else:
        # 若上述條件不符和，就買入 0 股，（在 E 欄顯示 0）
            tsmc_sheet.cells(i, "D").value = 0
        # 若 3日 > ，而且昨天的持有股數大於 1000 股，就賣出 1000 股
        if (price_today < sma_3d) and (tsmc_sheet.cells(i-1, "F").value >= stock_shares):
            tsmc_sheet.cells(i, "E").value = stock_shares
        else:
            tsmc_sheet.cells(i, "E").value = 0
        # 持有股數，算法是前一天的持有股數 + 今天的買入股數 - 今天的賣出股數
        tsmc_sheet.cells(i, "F").value = tsmc_sheet.cells(i-1, "F").value + tsmc_sheet.cells(i, "D").value - tsmc_sheet.cells(i, "E").value
        # 持有資金，算法是前一天的持有資金 + 今日收盤價 x (今天的賣出股數 - 今天的買入股數)
        tsmc_sheet.cells(i, "G").value = tsmc_sheet.cells(i-1, "G").value + price_today * (tsmc_sheet.cells(i, "E").value - tsmc_sheet.cells(i, "D").value)
        # 總資產則是持有股數 x 今日收盤價 + 今日持有資金
        tsmc_sheet.cells(i, "H").value = tsmc_sheet.cells(i, "G").value + tsmc_sheet.cells(i, "F").value * price_today

    # 計算并且將總收益顯示在 L20
    tsmc_sheet.cells(15, "K").value = tsmc_sheet.cells(last_row, "H").value - tsmc_sheet.cells(13, "K").value

@xw.func
def hello(name):
    return "hello {0}".format(name)

@xw.func
def double_sum(x, y):
    """Returns twice the sum of the two arguments"""
    return 2 * (x + y)

@xw.func
@xw.arg('prices', np.array, ndim=1)
def SMA(prices):
    return np.mean(prices)

@xw.func
@xw.arg('prices', np.array, ndim=1)
def WMA(prices):
    weights = np.arange(1, prices.size+1)
    return np.sum(prices * weights) / np.sum(weights)