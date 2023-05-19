from datetime import date
class Asset(object):
    """description of class"""
    def __init__(self, name, pData, yData=None, freq=4, value=1000, leverageMultiplier = 1):
        self.name = name
        self.value = value
        #self.hasCashFlows = hasCashFlows
        self.leverageMultiplier = leverageMultiplier
        #assumes data is in a 2 column matrix with the first column as dates and the second with prices
        self.priceData = pData
        self.yieldData = None
        self.frequency = None
        self.yIter = None
        self.setYieldData(yData, freq)

    #assumes data is in a 2 column matrix with the first column as dates and the second with yields
    def setYieldData(self, data, freq):
        if not data == None:
            self.yieldData = data
            self.frequency = freq
            self.yIter = 0
        else:
            print("Cannot set yield data for " + self.name)

    def getValue(self):
        return self.value

    def addCash(self, cash):
        self.value += cash

    def subtractCash(self, cash):
        self.value -= cash
        if(self.value < 0):
            self.value = 0
            print("There is 0 dollars in " + self.name)

    def setValue(self, val):
        if(val > 0):
            self.value = val
        else:
            print("Cannot set a negative value")

    def iteratePriceData(self, idx1, idx2):
        if(idx2 < len(self.priceData)):
            growth = (self.priceData[idx2][1]-self.priceData[idx1][1])/self.priceData[idx1][1]
            self.value *= (1 + self.leverageMultiplier*growth)
        else:
            print("No more price data to iterate through in " + self.name + " asset.")

    def getCashFlow(self):
        if not self.yieldData == None:
            if(self.yIter < len(self.yieldData)):
                yld = self.yieldData[self.yIter][1]
                cash = self.value * ((yld/100.)/self.frequency)
                #print(self.yieldData[self.yIter][0].strftime("%m/%d/%y") + ": Yield- " + str(yld) + "  $" + str(cash) + " in " + self.name + " asset.")
                self.yIter += 1
                return cash
            else:
                print("Index out of range for yield data for " + self.name + " asset.")
        else:
            return 0

    def peekPriceDate(self, idx):
        if idx == -1:
            return self.priceData[len(self.priceData)-1][0]
        elif idx < len(self.priceData):
            return self.priceData[idx][0]
        else:
            print("No more price data to iterate through in " + self.name + " asset.")
    
    def peekCashFlowDate(self):
        if not self.yieldData == None:
            if(self.yIter < len(self.yieldData)):
                return self.yieldData[self.yIter][0]
            else:
                print("No more yield data to iterate through in " + self.name + " asset.")
        else:
            None
    

