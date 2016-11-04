# -*- coding: utf-8 -*-
"""
Created on Wed Apr 06 11:48:33 2016

@author: ima12
"""

from bs4 import BeautifulSoup

import urllib

def getYahooStockQuote(symbol): 
    url = "http://download.finance.yahoo.com/d/quotes.csv?s=%s&f=nl1t8p2m8k5j6j1rqy" % symbol 
    f = urllib.urlopen(url) 
    s = f.read() 
    f.close() 
    s = s.strip()  #remove blanks at front and back
        
    L = s.split(',') 
    
    D = {} 
    
    adj = s.count(',')-10  
    # "Normal tickers 10 commas, but for FB "Facebook, Inc" has 11 comma

    D['NAME'] = L[0] 
    D['PRICE'] = L[1+adj] 
    D['TARGET'] = L[2+adj] 
    
#    D['Potential']=  round((float(L[2]) - float(L[1])) / float(L[1]) * 100,2)
    D['CHG %'] = L[3+adj] 
    D['% 50D AVG'] = L[4+adj] 
    D['% 1YR HIGH'] = L[5+adj] 
    D['% 1YR LOW'] = L[6+adj] 
    D['MKT CAP'] = L[7+adj] 
    D['P/E'] = L[8+adj] 
    D['PRV EXDVD DT'] = L[9+adj] 
    D['DVD%'] = L[10+adj] 
    
    #=========================================
    #DVD estimate from dividend history.org
    #=========================================
    try:
        url = urllib.urlopen('http://dividendhistory.org/payout/' + symbol).read()
        soup = BeautifulSoup(url)
        dvd = soup.find(text='unconfirmed/estimated').findPrevious('i').contents[0]
        
        # To count how many [u'unconfirmed/estimated', u'unconfirmed/estimated']
        estcount = len(soup.findAll(text='unconfirmed/estimated'))
        
        #					<td><i>2017-05-05</i></td>
        #					<td><i>2017-06-01</i></td>
        #					<td><i>$1.6303**</i></td>
        #					<td><i>unconfirmed/estimated</i></td>
                 
        if estcount == 1:
            estdvd = soup.findAll(text='unconfirmed/estimated')[0].parent.parent.parent
        else:
            estdvd = soup.findAll(text='unconfirmed/estimated')[0].parent.parent.parent.nextSibling
        
        exdate =  estdvd.contents[1].string
        paydate = estdvd.contents[3].string
        dvdamt = estdvd.contents[5].string
        
        D['NXT EXDVD DT'] = exdate.encode('utf-8') 
        D['PAYDT'] = paydate.encode('utf-8') 
        D['DVD AMT'] = dvdamt.encode('utf-8') 
    except:
        D['NXT ExDVD DT'] = "N/A"
        D['PAYDT'] = "N/A"
        D['DVD AMT'] = "N/A" 
        
    return D 

#==========


symbol = "hfc"
print getYahooStockQuote(symbol)

import xlwt

workbook = xlwt.Workbook(encoding="utf-8")
worksheet = workbook.add_sheet("temp")

col = 0
for yfdata in getYahooStockQuote(symbol):
    col = col + 1
    worksheet.write(1,col,yfdata)
    workbook.save('temp.xlsx')

