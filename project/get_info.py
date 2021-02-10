import xlwings as xw
import investpy
from get_main import fund_country, index_country, stock_country, bond_country, etf_country

fund_sheet = 'funds'
index_sheet = 'indices'
stock_sheet= 'stocks'
bond_sheet = 'bonds'
etf_sheet = 'etfs'
crypto_sheet = 'cryptos'
home_sheet = 'Home'

fund = xw.sheets[home_sheet].range('G3').value
index = xw.sheets[home_sheet].range('G4').value
stock = xw.sheets[home_sheet].range('G5').value
bond = xw.sheets[home_sheet].range('G6').value
etf = xw.sheets[home_sheet].range('G7').value

if fund and fund_country != "":
    req_fund_info = investpy.get_fund_information(
    fund=fund,
    country=fund_country
)

if index and index_country != "":
    req_index_info = investpy.get_index_information(
    index=index,
    country=index_country
)

if stock and stock_country != "":
    req_stock_info = investpy.get_stock_information(
    stock=stock,
    country=stock_country
)

if bond and bond_country != "":
    req_bond_info = investpy.get_bond_information(
    bond=bond
)

if etf and etf_country != "":
    req_etf_info = investpy.get_etf_information(
    etf=etf,
    country=etf_country
)

def GetFundInformation():
    wb = xw.Book.caller()
    wb.sheets[fund_sheet].range('N1').options(transpose=True).value = req_fund_info

def GetIndexInformation():
    wb = xw.Book.caller()
    wb.sheets[index_sheet].range('N1').options(transpose=True).value = req_index_info

def GetStockInformation():
    wb = xw.Book.caller()
    wb.sheets[stock_sheet].range('N1').options(transpose=True).value = req_stock_info

def GetBondInformation():
    wb = xw.Book.caller()
    wb.sheets[bond_sheet].range('N1').options(transpose=True).value = req_bond_info

def GetETFInformation():
    wb = xw.Book.caller()
    wb.sheets[etf_sheet].range('N1').options(transpose=True).value = req_etf_info