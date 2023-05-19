
'''
Created on Thursday July 30

@author: Jordan Fanapour
'''
import bs4
import requests
from bs4 import BeautifulSoup

def parsePrice():
    r=requests.get("https://finance.yahoo.com/quote/VTI?p=VTI&.tsrc=fin-srch")
    soup=bs4.BeautifulSoup(r.text, "xml")
    price=soup.find_all('div',{'class': 'D(ib) Mend(20px)'})[0].find('span').text
    return price

i = 0;
while i < 20:
    print("The price of VTI is $" + str(parsePrice))
    i+=1