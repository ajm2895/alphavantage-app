import requests
import json
import math
import xlsxwriter
import os, sys


##  function to make json easier to read  ##
def jprint(obj):
    text = json.dumps(obj, sort_keys=True, indent=4)
    print(text)
##

def mean(obj):
    total = sum(obj)
    mean = total / (len(obj)-1)
    return(mean)

def standard_deviation(values):
    total = sum(values)
    mean = total / (len(values))

    adj_values = []

    for val in values:
        adj_values.append((val-mean)**2)

    adj_total = sum(adj_values)
    variance = adj_total / (len(values))
    standard_dev = math.sqrt(variance)

    return(standard_dev)

## removes the earliest data point and adds the next for a moving 20-day average ##
def make_points():
    temp = []
    for i in range(len(date)-1):
        if i < 20:
            temp.append(float(close[i]))

    upper.append(mean(temp) + (standard_deviation(temp) *2))  ##upper band
    middle.append(mean(temp))                                 ##middle band  
    lower.append(mean(temp) - (standard_deviation(temp) *2))  ##upper band  
##

    

##  https://towardsdatascience.com/introduction-to-machine-learning-for-beginners-eed6024fdb08  ##
##  https://www.alphavantage.co/documentation/  ##
##  https://traderhq.com/ultimate-guide-to-bollinger-bands/  ##
##  https://school.stockcharts.com/doku.php?id=technical_indicators:bollinger_bands  ##
##  https://www.investopedia.com/terms/r/rsi.asp  ##

apikey = input('API Key: ')
symbol = input('Ticker: ')

base_url = 'https://www.alphavantage.co/query?'
parameters = {'function' : 'TIME_SERIES_DAILY_ADJUSTED',    #time series (endpoint) of your choice
              'symbol' : symbol,                            #equity of choice ('msft')
              'outputsize' : 'full',                        #compact/full --  last 100 data points/full-length time series
              'datatype' : 'json',                          #json/csv
              'apikey' : apikey                             #api key
              }

response = requests.get(base_url, params=parameters)
##jprint(response.json())

date = []
open_price = []
high = []
low = []
close = []
actual = []
RSI = []

for index, item in enumerate(response.json()["Time Series (Daily)"]):
    if index < 200:
        date.append(item)

for dates in reversed(date):
    open_price.append(response.json()["Time Series (Daily)"][dates]["1. open"])
    high.append(response.json()["Time Series (Daily)"][dates]["2. high"])
    low.append(response.json()["Time Series (Daily)"][dates]["3. low"])
    close.append(response.json()["Time Series (Daily)"][dates]["4. close"])
    actual.append(float(response.json()["Time Series (Daily)"][dates]["4. close"]))

##remove the earliest 20 data points##
for i in range(20):
    actual.pop(0)
    date.pop()
##

headings = ['Date', 'Upper Band', 'Middle Band', 'Lower Band', 'Actual']#, 'RSI']
upper = []
middle = []
lower = []

for i in range(180):
    make_points()
    close.pop(0)


##  write the results to an excel spreadsheet  ##
workbook = xlsxwriter.Workbook(symbol + '.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1})

worksheet.write_row('A1', headings, bold)

for row, item in enumerate(reversed(date), start=1):
    worksheet.write(row, 0, item)

worksheet.write_column('B2', upper)
worksheet.write_column('C2', middle)
worksheet.write_column('D2', lower)
worksheet.write_column('E2', actual)

worksheet.write(0, 7, symbol)

chart = workbook.add_chart({'type': 'line'})
chart.add_series({
    'name': '=Sheet1!$B$1',
    'values': '=Sheet1!$B$2:$B$181',
    'line': {
        'color': 'red',
        'transparency': '60'
        }
    })
chart.add_series({
    'name': '=Sheet1!$C$1',
    'values': '=Sheet1!$C$2:$C$181',
    'line': {
        'color': 'red',
        'dash_type': 'dash',
        'transparency': '60'
        }
    })
chart.add_series({
    'name': '=Sheet1!$D$1',
    'values': '=Sheet1!$D$2:$D$181',
    'line': {
        'color': 'red',
        'transparency': '60'
        }
    })
chart.add_series({
    'name': '=Sheet1!$E$1',
    'values': '=Sheet1!$E$2:$E$181',
    'line': {
        'color': 'purple'
        }
    })

lowest = min(lower + actual)
highest = max(upper + actual)

chart.set_x_axis({
    'label_position': 'none'
    })

chart.set_y_axis({
    'min': math.floor(lowest),
    'max': math.ceil(highest)
    })

worksheet.insert_chart('H2', chart)


## acquire the RSI indicator  ##
## This is now a premium endpoint ##

##base_url1 = 'https://www.alphavantage.co/query?'
##parameters1 = {'function' : 'RSI',    
##              'symbol' : symbol,                            
##              'interval' : 'daily',                        
##              'time_period' : 10,
##              'series_type': 'close',
##              'datatype': 'json',
##              'apikey' : apikey                             
##              }

##response1 = requests.get(base_url1, params=parameters1)
##jprint(response1.json())

##for dates in reversed(date):
##    try:
##        RSI.append(float(response1.json()["Technical Analysis: RSI"][dates]["RSI"]))
##    except:
##        pass
##
##worksheet.write_column('F2', RSI)
##
##chart1 = workbook.add_chart({'type': 'line'})
##chart1.add_series({
##    'name': '=Sheet1!$F$1',
##    'values': '=Sheet1!$F$2:$F$181',
##    'line': {
##        'color': 'purple'
##        }
##    })
##
##chart1.set_x_axis({
##    'label_position': 'none'
##    })
##
##chart1.set_y_axis({
##    'min': 0,
##    'max': 100
##    })
##
##worksheet.insert_chart('A9', chart1)
workbook.close()

os.startfile(symbol + '.xlsx')
