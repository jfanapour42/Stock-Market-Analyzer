import Asset
import datetime
from datetime import date
class Portfolio(object):
    """description of class"""
    def __init__(self, name, assets, weights, startingAmount = 10000, outMCF = 0, cashInLow = True):
        self.name = name
        #self.value = startingAmount
        self.outsideMonthlyCashFlow = outMCF
        self.investCashInLowestAsset = cashInLow
        self.assets = assets
        self.setWeights(weights)
        self.rebalance(startingAmount)
        self.accountValueData = []
        self.accountIncomeData = []

    def setWeights(self, weights):
        if(len(weights) == len(self.assets)):
            sum = 0.0
            for val in weights:
                sum += val
            if sum == 1.0:
                self.weights = weights
            else:
                print("Weights do not sum up to 1")
        else:
            print("Number of weights and number of assets don't match up")

    def setAsset(self, idx, asset):
        if(idx < len(self.assets)):
            self.assets[idx] = asset
        else:
            print("index out of range")

    def rebalance(self, startingAmount=None):
        val = startingAmount
        if startingAmount == None:
            val = self.getValue()
        for i in range(len(self.assets)):
            self.assets[i].setValue(val*self.weights[i])

    def getValue(self):
        sum = 0
        for asset in self.assets:
            sum += asset.getValue()
        #self.value = sum
        return sum

    def getMinAsset(self):
        minAsset = self.assets[0]
        for asset in self.assets:
            if(asset.getValue() < minAsset.getValue()):
                minAsset = asset
        #check if this returned object is the same instance as the one in the list.
        return minAsset

    def checkDataRange(self):
        length = len(self.assets[0].priceData)
        for i in range(1, len(self.assets)):
            if not length == len(self.assets[i].priceData):
                print("Asset 1 and " + str(i+1) + " have differing data ranges.")
                return False
        else:
            return True

    def getTupleOfAssetValues(self):
        tup = ()
        for asset in self.assets:
            t = (asset.getValue(),)
            tup += t
        return tup

    def getDividends(self, date):
        sum = 0
        for asset in self.assets:
            tempDate = asset.peekCashFlowDate()
            if not tempDate == None:
                if tempDate.year == date.year and tempDate.month == date.month:
                    sum += asset.getCashFlow()
        if sum != 0:
            self.accountIncomeData.append((date, sum))
        return sum

    def getIncomeData(self):
        return self.accountIncomeData

    def getAnnualReturnData(self):
        data = []
        idx = 0
        valIdx = len(self.accountValueData[idx])-1
        while idx < len(self.accountValueData):
            idx2 = idx        
            date = self.accountValueData[idx][0]
            while idx2 < len(self.accountValueData):
                if date.year + 1 == self.accountValueData[idx2][0].year:
                    break
                else:
                    idx2 += 1
            if idx2 >= len(self.accountValueData):
                growth = (self.accountValueData[idx2-1][valIdx]-self.accountValueData[idx][valIdx])/self.accountValueData[idx][valIdx]
            else:
                growth = (self.accountValueData[idx2][valIdx]-self.accountValueData[idx][valIdx])/self.accountValueData[idx][valIdx]
            year = datetime.date(self.accountValueData[idx][0].year, 1, 1)
            data.append((year, growth*100.))
            idx = idx2
        return data

    #Assumes price data range in each asset is the same
    def generateData(self):
        if not self.checkDataRange():
            return None
        else: 
            date = self.assets[0].peekPriceDate(0)
            data = [(date,) + self.getTupleOfAssetValues() + (self.getValue(),)]
            for i in range(len(self.assets[0].priceData)-1):
                date = self.assets[0].peekPriceDate(i+1)
                tup = (date,)
                for asset in self.assets:
                    asset.iteratePriceData(i, i+1)
                dividends = self.getDividends(date)
                minAsset = self.getMinAsset()
                minAsset.addCash(dividends)
                if date.month == 12:
                    self.rebalance()
                tup += self.getTupleOfAssetValues()
                tup += (self.getValue(),)
                data.append(tup)
            self.accountValueData = data
            return data

    #Assumes price data range in each asset is the same
    #endDate is the date that cash in flows are no longer added to the portfolio
    def generateData(self, monthlyInflows, endDate=None):
        if not self.checkDataRange():
            return None
        else: 
            date = self.assets[0].peekPriceDate(0)
            stopDate = None
            if endDate == None:
               stopDate = self.assets[0].peekPriceDate(-1)
            else:
                stopDate = endDate
            data = [(date,) + self.getTupleOfAssetValues() + (self.getValue(),)]
            for i in range(len(self.assets[0].priceData)-1):
                date = self.assets[0].peekPriceDate(i+1)
                tup = (date,)
                for asset in self.assets:
                    asset.iteratePriceData(i, i+1)

                if date < stopDate:
                    minAsset = self.getMinAsset()
                    minAsset.addCash(monthlyInflows)

                dividends = self.getDividends(date)
                minAsset = self.getMinAsset()
                minAsset.addCash(dividends)

                if date.month == 12:
                    self.rebalance()
                tup += self.getTupleOfAssetValues()
                tup += (self.getValue(),)
                data.append(tup)
            self.accountValueData = data
            return data
    
    #Assumes data range in each asset is the same
    def generateDataWithoutDividends(self):
        if not self.checkDataRange():
            return None
        else: 
            date = self.assets[0].peekPriceDate(0)
            data = [(date,) + self.getTupleOfAssetValues() + (self.getValue(),)]
            for i in range(len(self.assets[0].priceData)-1):
                date = self.assets[0].peekPriceDate(i+1)
                tup = (date,)
                for asset in self.assets:
                    asset.iteratePriceData(i, i+1)
                if date.month == 12:
                    self.rebalance()
                tup += self.getTupleOfAssetValues()
                tup += (self.getValue(),)
                data.append(tup)
            self.accountValueData = data
            return data

