// constants.gs - configuration and shared constants

// Name of the ledger sheet backing the portfolio tracker
var SHEET_NAME = 'Portfolio Ledger';

// Header row configuration for the ledger
// Stock | Con Note | Buy/Sell | Date | Number | Price | Brokerage | Total Cost | Dividend | Franking | Unfranked | DRP | Div/Share | Sell Price
var LEDGER_HEADERS = [
  'Stock',
  'Con Note',
  'Buy/Sell',
  'Date',
  'Number',
  'Price',
  'Brokerage',
  'Total Cost',
  'Dividend',
  'Franking',
  'Unfranked',
  'DRP',
  'Div/Share',
  'Sell Price'
];

// Time zone for date handling (adjust if needed)
var APP_TIMEZONE = 'Australia/Sydney';

// Sheet used to store live prices via GOOGLEFINANCE formulas.
// Layout (managed by Apps Script):
// Col A: Stock code (e.g. BHP)
// Col B: =GOOGLEFINANCE("ASX:"&Arow,"price")
var PRICE_SHEET_NAME = 'Prices';

var HISTORY_SHEET_NAME = 'History';