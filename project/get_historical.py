import investpy
import xlwings as xw
import pandas as pd
from get_main import fund_country, index_country, stock_country, bond_country, etf_country

# Sheet names
fund_sheet = 'funds'
index_sheet = 'indices'
stock_sheet= 'stocks'
bond_sheet = 'bonds'
etf_sheet = 'etfs'
crypto_sheet = 'cryptos'
home_sheet = 'Home'

# Names / Symbols
fund = xw.sheets[home_sheet].range('G3').value
index = xw.sheets[home_sheet].range('G4').value
stock = xw.sheets[home_sheet].range('G5').value
bond = xw.sheets[home_sheet].range('G6').value
etf = xw.sheets[home_sheet].range('G7').value

# Start and End dates
fund_from_date = xw.sheets[home_sheet].range('H3').value
fund_from_date = fund_from_date.strftime("%d/%m/%Y")

fund_to_date = xw.sheets[home_sheet].range('I3').value
fund_to_date = fund_to_date.strftime("%d/%m/%Y")

index_from_date = xw.sheets[home_sheet].range('H4').value
index_from_date = index_from_date.strftime("%d/%m/%Y")

index_to_date = xw.sheets[home_sheet].range('I4').value
index_to_date = index_to_date.strftime("%d/%m/%Y")

stock_from_date = xw.sheets[home_sheet].range('H5').value
stock_from_date = stock_from_date.strftime("%d/%m/%Y")

stock_to_date = xw.sheets[home_sheet].range('I5').value
stock_to_date = stock_to_date.strftime("%d/%m/%Y")

bond_from_date = xw.sheets[home_sheet].range('H6').value
bond_from_date = bond_from_date.strftime("%d/%m/%Y")

bond_to_date = xw.sheets[home_sheet].range('I6').value
bond_to_date = bond_to_date.strftime("%d/%m/%Y")

if fund and fund_country and fund_from_date and fund_to_date != "":
    fund_data_historical = investpy.get_fund_historical_data(
    fund=fund,
    country=fund_country,
    from_date=fund_from_date,
    to_date=fund_to_date
)

if index and index_country and index_from_date and index_to_date != "":
    index_data_historical = investpy.get_index_historical_data(
    index=index,
    country=index_country,
    from_date=index_from_date,
    to_date=index_to_date
)

if stock and stock_country and stock_from_date and stock_to_date != "":
    stock_data_historical = investpy.get_stock_historical_data(
    stock=stock,
    country=stock_country,
    from_date=stock_from_date,
    to_date=stock_to_date
)

if bond and bond_country and bond_from_date and bond_to_date != "":
    bond_data_historical = investpy.get_bond_historical_data(
    bond=bond,
    from_date=bond_from_date,
    to_date=bond_to_date
)

def GetFundHistorical():
    wb = xw.Book.caller()
    wb.sheets[fund_sheet].range('R1').value = fund_data_historical

def GetIndexHistorical():
    wb = xw.Book.caller()
    wb.sheets[index_sheet].range('R1').value = index_data_historical

def GetStockHistorical():
    wb = xw.Book.caller()
    wb.sheets[stock_sheet].range('R1').value = stock_data_historical

def GetBondHistorical():
    wb = xw.Book.caller()
    wb.sheets[bond_sheet].range('R1').value = bond_data_historical
