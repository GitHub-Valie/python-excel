import investpy
import xlwings as xw
import pandas as pd

fund_sheet = 'funds'
index_sheet = 'indices'
stock_sheet= 'stocks'
bond_sheet = 'bonds'
etf_sheet = 'etfs'
crypto_sheet = 'cryptos'

req_fund_countries = investpy.get_fund_countries()
df_fund_countries = pd.DataFrame(req_fund_countries)

req_bond_countries = investpy.get_bond_countries()
df_bond_countries = pd.DataFrame(req_bond_countries)

req_index_countries = investpy.get_index_countries()
df_index_countries = pd.DataFrame(req_index_countries) 

req_stock_countries = investpy.get_stock_countries()
df_stock_countries = pd.DataFrame(req_stock_countries)

req_etf_countries = investpy.get_etf_countries()
df_etf_countries = pd.DataFrame(req_etf_countries)

def GetCountries():
    wb = xw.Book.caller()
    wb.sheets[fund_sheet].range('A1').value = df_fund_countries
    wb.sheets[index_sheet].range('A1').value = df_index_countries
    wb.sheets[stock_sheet].range('A1').value = df_stock_countries
    wb.sheets[bond_sheet].range('A1').value = df_bond_countries
    wb.sheets[etf_sheet].range('A1').value = df_etf_countries
