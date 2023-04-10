import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
# from robin_scrape import *
import time
import json
import robin_stocks as r
import config
from datetime import date
import calendar
import yfinance as yf

from bs4 import BeautifulSoup
import requests
from lxml import html
import pyotp

scope = ["https://spreadsheets.google.com/feeds",
"https://www.googleapis.com/auth/spreadsheets",
"https://www.googleapis.com/auth/drive.file",
"https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("Investment History").sheet1
data = sheet.get_all_records()
col = sheet.col_values(1)

#scrape the email for the ticker symbol
#email has subject "Your Order Has Been Executed!"
ticker = []
shares = []
blank = {}

#previous data
with open('stocks.json') as json_file:
    data = json.load(json_file)

previous_amount = int(len(data))


r.login(config.username, config.password, 100000)

my_stocks = r.build_holdings()

amount = int(len(my_stocks)) - previous_amount

if previous_amount < len(my_stocks):
    for i in range(amount):
        sheet.insert_row([], 5)

name_col = 2
share_col = 3
avg_cost_col = 4
equity_col = 7
port_pct = 8
equity_pct_change_col = 9
div_yield_col = 11
annual_div_col = 12



# my_stocks = robin_scrape.

tickers = []

#current data
for i in my_stocks:
    tickers.append(i)

yields = []
am = []


for j in range(len(tickers[0:])):
    url = 'https://finviz.com/quote.ashx?t=' + tickers[j]
    agent = {"User-Agent":"Mozilla/5.0"}

    website = requests.get(url, headers = agent).text
    soup = BeautifulSoup(website, 'lxml')

    # print(soup)

    table = soup.find('table', {'class': 'snapshot-table2'})

    dividend_yield = []
    dividend_amount = []

    for item in table.findAll('tr')[7]:
        # print(item.('td'))
        span = item.find('span')
        dividend_yield.append(span)

    if dividend_yield[2] == None:
        dividend_yield = []
        for j in table.findAll('tr')[7]:
            b = j.find('b')
            dividend_yield.append(b)

    for i in table.findAll('tr')[6]:
        price = i.find('b')
        dividend_amount.append(price)

    # print(dividend_yield)

    current_yield = dividend_yield[2].text
    annual_amount = dividend_amount[2].text

    yields.append(current_yield)
    am.append(annual_amount)





max_row = len(my_stocks) + 2

# print(yields)

# row range is from 2 to len(my_stocks) + 2

for i in range(2, max_row):

    sheet.update_cell(i, 1, tickers[i-2])
    sheet.update_cell(i, name_col, my_stocks[tickers[i-2]]['name'])
    time.sleep(2)
    sheet.update_cell(i, share_col, my_stocks[tickers[i-2]]['quantity'])
    time.sleep(5)
    sheet.update_cell(i, avg_cost_col, my_stocks[tickers[i-2]]['average_buy_price'])
    time.sleep(2)
    sheet.update_cell(i, equity_col, my_stocks[tickers[i-2]]['equity'])
    time.sleep(5)
    sheet.update_cell(i, port_pct, (float(my_stocks[tickers[i-2]]['percentage'])/100))
    time.sleep(2)
    sheet.update_cell(i, equity_pct_change_col, (float(my_stocks[tickers[i-2]]['equity_change'])/100))
    time.sleep(5)
    sheet.update_cell(i, div_yield_col, yields[i-2])
    time.sleep(5)
    sheet.update_cell(i, annual_div_col, am[i-2])


dollar_rows_format = str('E' + str(max_row))

sheet.format(('D2:' + dollar_rows_format), {'numberFormat' : {'type': 'CURRENCY'}})
sheet.format(('H2:I' + str(max_row)), {'numberFormat' : {'type': 'PERCENT'}})

outfile = open('stocks.json' , 'w')

outfile.write(json.dumps(my_stocks, indent = 2))

string = sheet.row_values(max_row)[12]
money, amount_ = string.split('$')
amount_ = float(amount_)

div_amount = {
'Total Annual Dividend': amount_
}


current_div = open('current_dividend.json', 'w')
current_div.write(json.dumps(div_amount, indent = 2))

my_date = date.today()
day_of_week = calendar.day_name[my_date.weekday()]

last_week = open('last_friday.json', 'w')
string = sheet.row_values(max_row)[12]
money, amount_ = string.split('$')
amount_ = float(amount_)

div_amount = {
'Total Annual Dividend': amount_
}

div = open('last_dividend.json', 'w')

if day_of_week == 'Friday':
    last_week.write(json.dumps(my_stocks, indent = 2))
    div.write(json.dumps(div_amount, indent = 2))
