# Stock Portfolio Tracker
A completely free stock portfolio tracker using Google Sheets that allows you to key in ticker symbols and transactions to get comprehensive summaries on your portfolio with automatically updating prices, dividends, splits and more. 
<br></br>
Spreadsheet link: [Make a copy](https://docs.google.com/spreadsheets/d/1FXQQ6PKt8o27YgWu_5Nh79tUKH_JTXjwdPPXvMvUFJY/edit?usp=sharing)

![Stock Portfolio Banner](https://i.imgur.com/MlSPvMJ.png?1)

## Features
- One click setup button that automatically sets up the portfolio for you, ready to use
- Automatically updating stock prices and other meta data (name, dividend yield, 200 day average etc.)
- Add new transactions easily and specify the type of transaction (buy/sell/dividend etc.)
- Automatic portfolio history logging for your portfolio's own historical data and summary graphs
- Automatic logging of dividends and splits and historical price/units correction
- Comprehensive portfolio summary with sector holdings, asset holdings, returns and gains
- Supports most exchanges globally, with stocks shown in their respective currencies
- Automatic conversion to your local currency for a more relatable view in your portfolio summary
- Locality support for both US and UK (for all those MM/DD versus DD/MM users)
- Custom menu and over 1000 lines of appscript code

## Install and Setup
1. Go the base spreadsheet: [LINK](https://docs.google.com/spreadsheets/d/1FXQQ6PKt8o27YgWu_5Nh79tUKH_JTXjwdPPXvMvUFJY/edit?usp=sharing)
2. Do not request edit access. Go to File > Make a copy to import a copy into your own Google Drive.
3. In the Setup page of the spreadsheet, click the blue setup button.
4. Change the user settings to your preferred locale and currency.
5. You're done! It's that easy. Read on for detailed instructions on usage.

## User Guide
The spreadsheet has a few main sheets: <b>Setup, Portfolio Summary, Cash Flows, Stock Summary and Transactions.</b> For explanations on each page, mouse over the header in each sheet within the spreadsheet itself. You can also mouse over most of the headers in the sheets for explanations on the data within that column.
<br></br>
In this user guide, we will cover the basic flow of how the spreadsheet is supposed to work. If you would like to see the inner workings and add your own things, you can find the source code in Tools > Script Editor.

### 1) Entering Transactions
The first step is to enter transactions. This can be done in the <b>Transactions</b> sheet. You will have to fill in all yellow cells.

1. In the top menu, go to <b>Portfolio Menu > Transactions > Add New Transaction/Add New Buy Transaction.</b> A new transaction row will appear for you to key in your transaction details.
> <b>New Transaction:</b> Adds a single new transaction row. Useful for transactions like selling and cash outs.
> <b>New Buy Transaction:</b> Adds two new transaction rows, one for cash in (to fund the buy transaction) and one for a buy transaction.

Note:

### 2) Adding Assets to Stock Summary
If you purchased a new asset you didn't own before, you will have to add it to the <b>Stock Summary</b> sheet. Just head on over and click an empty dropdown box to add your new asset. Only assets you have not added will be shown.

And honestly... that's about it! Everything else is automated. Mouse over the headers for detailed information if required.

Enjoy!


