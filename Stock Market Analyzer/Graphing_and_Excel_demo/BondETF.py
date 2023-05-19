import datetime
from datetime import date
from Bond import Bond
import bisect
from FixedQueue import FixedQueue
class BondETF(object):
    """description of class"""
    def __init__(self, startDate, minYData, maxYData, pDates, startValue=10000, cash=0, min_MRTY=20, max_MRTY=30):
        self.assets = FixedQueue(len(pDates)*(max_MRTY-min_MRTY))
        self.date = startDate
        self.cashReserves = cash
        self.min_MRTY = min_MRTY
        self.max_MRTY = max_MRTY
        self.minYieldData = minYData
        self.maxYieldData = maxYData
        self.purchaseDates = pDates
        self.initialize(startValue)

    def initialize(self, startValue):
        val = startValue/(1.*len(self.assets.getArray()))
        for i in range(self.max_MRTY-self.min_MRTY+1):
            for tup in self.purchaseDates:
                d = datetime.date(self.date.year-10+i,tup[0],tup[1])
                b = None
                if i == 0:
                    d2 = datetime.date(self.date.year-10+i, self.date.month, self.date.day)
                    if d > d2:
                        b = self.getInitialBond(val, d)
                elif i == self.max_MRTY-self.min_MRTY:
                    d2 = datetime.date(self.date.year-10+i, self.date.month, self.date.day)
                    if d < d2:
                        b = self.getInitialBond(val, d)
                else:
                    b = self.getInitialBond(val, d)
                if not b == None:
                    self.assets.enqueue(b)

    def getInitialBond(self, val, date):
        idx = bisect.bisect_left(self.maxYieldData.Date, date)
        couponRate = self.maxYieldData.Yield[idx]
        idx = bisect.bisect_left(self.maxYieldData.Date, self.date)
        marketRate = self.maxYieldData.Yield[idx]
        btemp = Bond(val, couponRate, date, self.max_MRTY)
        issuePrice = btemp.calculateTheoreticalIssuePrice(val, marketRate, self.date)
        return Bond(issuePrice, couponRate, date, self.max_MRTY)

    def getMarketValue(self, marketRate, date):
        total = 0
        for asset in self.assets.getArray():
            if not asset == None:
                total += asset.presentValue(marketRate, date)
        return total

