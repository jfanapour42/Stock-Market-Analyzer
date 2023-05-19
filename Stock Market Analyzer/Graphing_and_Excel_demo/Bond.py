import math
import datetime
from datetime import date
from dateutil.relativedelta import relativedelta
class Bond(object):
    """description of class"""
    def __init__(self, price, coupon, date, maturity, frequency=2):
        self.issuePrice = price
        self.couponRate = coupon/100.
        self.issueDate = date
        self.maturity = maturity
        self.frequency = frequency

    def getMaturityDate(self):
        #if self.maturity > 1:
        return datetime.date(self.issueDate.year + self.maturity,self.issueDate.month, self.issueDate.day)
        #else:

    def getRemainingMaturity(self, date):
        return relativedelta(self.getMaturityDate(), date).years

    def getPayment(self, date):
        if self.issueDate < date < self.getMaturityDate():
            if self.issueDate.day == date.day:
                if self.issueDate.month > 6:
                    if date.month == self.issueDate.month or date.month == self.issueDate.month - 6:
                        return self.getPaymentVal()
                else:
                    if date.month == self.issueDate.month or date.month == self.issueDate.month + 6:
                        return self.getPaymentVal()
        elif date == self.getMaturityDate():
            return self.getPaymentVal() + self.issuePrice
        return 0

    def getPaymentVal(self):
        return self.issuePrice * (self.couponRate/self.frequency)

    def getNumberOfRemainingPayments(self, date):
        if self.issueDate < date and date <= self.getMaturityDate():
            num = (self.getMaturityDate().year-date.year)*self.frequency
            if self.issueDate.month < 7:
                if self.issueDate.replace(year=date.year, month=self.issueDate.month+6) <= date:
                    num -= 1
                elif self.issueDate.replace(year=date.year) > date:
                    num += 1
            else:
                if self.issueDate.replace(year=date.year, month=self.issueDate.month-6) > date:
                    num += 2
                elif self.issueDate.replace(year=date.year) > date:
                    num += 1
            return num
        elif self.issueDate == date:
            return self.maturity*self.frequency
        else:
            return "Given date does not fall within range of issue date "

    
    def presentValue(self, marketRate, date):
        if self.issueDate < date <= self.getMaturityDate():
            #print(str(self.getCleanPrice(marketRate,date)) + "    " + str(self.getAccumulatedInterest(date)))
            return self.getCleanPrice(marketRate,date) #+ self.getAccumulatedInterest(date)
        elif self.issueDate == date:
            return self.issuePrice
        else:
            return "Given date falls outside maturity of bond"

    def getCleanPrice(self, marketRate, date):
        factor = pow((1+((marketRate/100.)/self.frequency)),-1*self.getNumberOfRemainingPayments(date))
        return self.getPaymentVal()*(1-factor)/((marketRate/100.)/self.frequency) + self.issuePrice*factor

    def getAccumulatedInterest(self, date):
        lastPaymentDate = None
        if self.issueDate.month < 7:
            if self.issueDate.replace(year=date.year, month=self.issueDate.month+6) <= date:
                lastPaymentDate = self.issueDate.replace(year=date.year, month=self.issueDate.month+6)
            elif self.issueDate.replace(year=date.year) > date:
                lastPaymentDate = self.issueDate.replace(year=date.year-1, month=self.issueDate.month+6)
            else:
                lastPaymentDate = self.issueDate.replace(year=date.year)
        else:
            if self.issueDate.replace(year=date.year, month=self.issueDate.month-6) > date:
                lastPaymentDate = self.issueDate.replace(year=date.year-1)
            elif self.issueDate.replace(year=date.year) <= date:
                lastPaymentDate = self.issueDate.replace(year=date.year)
            else:
                lastPaymentDate = self.issueDate.replace(year=date.year, month=self.issueDate.month-6)
        daysElapsed = (date - lastPaymentDate).days
        return self.issuePrice*(self.couponRate/self.frequency)*(daysElapsed/((360)/self.frequency))

    def calculateTheoreticalIssuePrice(self, presentVal, marketRate, date):
        r = (marketRate/100.)/self.frequency
        factor = pow(1+r,-1*self.getNumberOfRemainingPayments(date))
        return (presentVal*r)/(self.couponRate/2 + factor*(r - self.couponRate/2))

    def annualYield(self, marketRate, date):
        r = (marketRate/100.)/self.frequency
        factor = pow(1+r,-1*self.getNumberOfRemainingPayments(date))
        return self.frequency*((self.couponRate/2)*r)/(self.couponRate/2 + factor*(r - self.couponRate/2))

    def __str__(self):
        temp = ""
        temp += str(self.maturity) + " Year Bond:"
        temp += "\n     Principal - $" + str(self.issuePrice)
        temp += "\n     Coupon Rate - " + str(self.couponRate*100) + "%"
        temp += "\n     Issue Date - " + self.issueDate.strftime("%m/%d/%y")
        temp += "\n     Maturity Date - " + self.getMaturityDate().strftime("%m/%d/%y")
        return temp




