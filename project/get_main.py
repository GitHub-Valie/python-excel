import investpy
import xlwings as xw
import pandas as pd

fund_sheet = 'funds'
index_sheet = 'indices'
stock_sheet= 'stocks'
bond_sheet = 'bonds'
etf_sheet = 'etfs'
crypto_sheet = 'cryptos'
home_sheet = 'Home'

fund_country = xw.sheets[home_sheet].range('F3').value
index_country = xw.sheets[home_sheet].range('F4').value
stock_country = xw.sheets[home_sheet].range('F5').value
bond_country = xw.sheets[home_sheet].range('F6').value
etf_country = xw.sheets[home_sheet].range('F7').value

req_funds = investpy.get_funds(
    country=fund_country
)

req_indices = investpy.get_indices(
    country=index_country
)

req_stocks = investpy.get_stocks(
    country=stock_country
)

req_bonds = investpy.get_bonds(
    country=bond_country
)

req_etfs = investpy.get_etfs(
    country=etf_country
)

def GetFunds():
    wb = xw.Book.caller()
    wb.sheets[fund_sheet].range('D1').value = req_funds

def GetIndices():
    wb = xw.Book.caller()
    wb.sheets[index_sheet].range('D1').value = req_indices

def GetStocks():
    wb = xw.Book.caller()
    wb.sheets[stock_sheet].range('D1').value = req_stocks

def GetBonds():
    wb = xw.Book.caller()
    wb.sheets[bond_sheet].range('D1').value = req_bonds

def GetETFs():
    wb = xw.Book.caller()
    wb.sheets[etf_sheet].range('D1').value = req_etfs