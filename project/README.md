# World Market Analyzer V 1.0

## What is it?

The world market analyzer is a project that aims at centralizing as much market data as possible for different assets.

## Features

### Implemented

* Funds:

Fund information (country, name, symbol, issuer, isin, asset class, currency, inception date, 1-year % change, risk rating, ROE, ROA, total assets, expenses, minimum investment, market capitalization, category) and quantitative data (daily price data, OHLCV format)

* 

### Yet to implement

## Disclaimer

This is a work in progress. Therefore, workbooks' features are not fully implemented and you may encounter bugs and error messages. I did my best to describe the steps to use this workbook

## How to use the World Market Analyzer

The workbook is made of 9 sheets (funds, indices, stocks, bonds, etfs, cryptos, Home, stock analysis, bond analysis)


### Step 1 / Get country data

First step is to get country data by clicking on the "Get Countries" button on the "Home" sheet

### Step 2 / Select countries and assets

In the Home sheet, use the table to choose the country of your choice for the asset of your choice, and then click the appropriate button

Example: All funds in Luxembourg 

<img src="https://github.com/GitHub-Valie/python-excel/blob/main/images/ma1.png">

### Step 3 / Select asset name

Select the name of the asset you wish to get info for and click "Get 'asset' Info"

<img src="https://github.com/GitHub-Valie/python-excel/blob/main/images/ma2.png">

This will generate a range of informations in the respective asset sheet (in this example, 'funds' sheet)

<img src="https://github.com/GitHub-Valie/python-excel/blob/main/images/ma3.png">

### Step 4 / Get historical price data

With the same principle, click "Get Historical data" to get price data for the chosen asset.

<img src="https://github.com/GitHub-Valie/python-excel/blob/main/images/ma4.png">