// Code.gs - main Apps Script backend

function doGet(e) {
  // Always serve the SPA shell (index.html). Client-side JS will
  // load partial views using getView(page).
  var template = HtmlService.createTemplateFromFile('index');
  return template
    .evaluate()
    .setTitle('DiviVue - AU Portfolio & Dividend Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Utility to get the main ledger sheet
function getLedgerSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(LEDGER_HEADERS);
  }
  return sheet;
}

// === Public backend methods exposed to client ===

/**
 * Append a new trade/dividend/DRP entry to the ledger.
 * Expects a plain object with properties matching LEDGER_HEADERS.
 */
function appendTrade(trade) {
  var sheet = getLedgerSheet();

  // Normalise fields in the correct column order
  var row = [
    trade.Stock || '',
    trade.ConNote || '',
    trade.BuySell || '',
    trade.Date || '',
    Number(trade.Number) || 0,
    Number(trade.Price) || 0,
    Number(trade.Brokerage) || 0,
    Number(trade.TotalCost) || 0,
    Number(trade.Dividend) || 0,
    Number(trade.Franking) || 0,
    Number(trade.Unfranked) || 0,
    trade.DRP || '',
    Number(trade.DivPerShare) || 0,
    Number(trade.SellPrice) || 0
  ];

  sheet.appendRow(row);

  // Placeholder: trigger recalculation of holdings/P&L/CGT helper columns
  recalculatePortfolioHelpers();

  return { success: true };
}

/**
 * Return all trades in the ledger as an array of objects.
 * This is used for the dashboard table and can be reused elsewhere.
 */
function getAllTrades() {
  var sheet = getLedgerSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return [];
  }

  var headers = data[0];
  var rows = data.slice(1);
  var result = [];

  rows.forEach(function (row) {
    var obj = {};
    headers.forEach(function (h, i) {
      obj[h] = row[i];
    });
    result.push(obj);
  });

  return result;
}

/**
 * Minimal dashboard data: total portfolio value, gain/loss placeholder, dividends YTD.
 * Real FIFO/CGT logic should be implemented in helper functions below.
 */
function getDashboardSummary() {
  // Use the existing investments summary to derive portfolio-level
  // totals for market value and capital gains.
  var investments = getInvestmentsSummary();

  var totalPortfolioValue = 0;
  var totalGainLoss = 0;

  (investments || []).forEach(function (entry) {
    totalPortfolioValue += Number(entry.value || 0);
    totalGainLoss += Number(entry.capitalGains || 0);
  });

  // Dividends: compute all-time, current AU financial year, and last
  // AU financial year (all 1 July to 30 June) using dividend trades
  // (dividend + franking).
  var trades = getAllTrades();
  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth(); // 0-based, July = 6
  var fyStartYear = month >= 6 ? year : (year - 1);
  var fyStart = new Date(fyStartYear, 6, 1); // 1 July
  var fyEnd = new Date(fyStartYear + 1, 5, 30); // 30 June

  // Last financial year range
  var lastFyStart = new Date(fyStartYear - 1, 6, 1);
  var lastFyEnd = new Date(fyStartYear, 5, 30);

  var dividendsAllTime = 0;
  var dividendsCurrentFy = 0;
  var dividendsLastFy = 0;

  trades.forEach(function (t) {
    if (t['Buy/Sell'] !== 'Dividend') return;
    var d = t['Date'] ? new Date(t['Date']) : null;
    if (!d) return;

    var dividend = Number(t['Dividend'] || 0);
    var franking = Number(t['Franking'] || 0);
    var amount = dividend + franking;

    dividendsAllTime += amount;

    if (d >= fyStart && d <= fyEnd) {
      dividendsCurrentFy += amount;
    } else if (d >= lastFyStart && d <= lastFyEnd) {
      dividendsLastFy += amount;
    }
  });

  return {
    totalPortfolioValue: totalPortfolioValue,
    totalGainLoss: totalGainLoss,
    // Backwards-compatible: treat YTD as current FY to date
    dividendsYTD: dividendsCurrentFy,
    dividendsAllTime: dividendsAllTime,
    dividendsCurrentFy: dividendsCurrentFy,
    dividendsLastFy: dividendsLastFy
  };
}

function getPortfolioNetInvestedSeries() {
  var trades = getAllTrades();
  if (!trades || trades.length === 0) {
    return [];
  }

  // Aggregate net cashflow per calendar date.
  var byDate = {};

  trades.forEach(function (t) {
    var rawDate = t['Date'] ? new Date(t['Date']) : null;
    if (!rawDate || isNaN(rawDate.getTime())) return;

    var key = Utilities.formatDate(rawDate, APP_TIMEZONE, 'yyyy-MM-dd');
    if (!byDate[key]) {
      byDate[key] = 0;
    }

    var type = (t['Buy/Sell'] || '').toString();
    var totalCost = Number(t['Total Cost'] || 0);
    var qty = Number(t['Number'] || 0);
    var brokerage = Number(t['Brokerage'] || 0);
    var sellPrice = Number(t['Sell Price'] || 0);

    if (type === 'Buy' || type === 'DRP') {
      // Cash out to buy shares increases net invested.
      byDate[key] += totalCost;
    } else if (type === 'Sell') {
      // Approximate proceeds using the same logic as in the
      // investments summary: treat Sell Price as per-share if
      // it looks like a price, otherwise as total proceeds.
      var proceeds;
      if (sellPrice > 0 && sellPrice <= 1000 && qty > 0) {
        proceeds = (sellPrice * qty) - brokerage;
      } else if (sellPrice > 0) {
        proceeds = sellPrice - brokerage;
      } else if (totalCost > 0) {
        proceeds = totalCost - brokerage;
      } else {
        proceeds = 0;
      }
      // Cash in from selling reduces net invested.
      byDate[key] -= proceeds;
    }
  });

  // Build a cumulative series sorted by date.
  var keys = Object.keys(byDate).sort();
  if (keys.length === 0) {
    return [];
  }

  var series = [['Date', 'Net Invested']];
  var running = 0;
  keys.forEach(function (key) {
    running += byDate[key];
    series.push([key, running]);
  });

  return series;
}

function getPortfolioMarketValueSeries() {
  var trades = getAllTrades();
  if (!trades || trades.length === 0) {
    return [];
  }

  // Sort trades by date ascending so we can apply them as we walk History.
  trades.sort(function (a, b) {
    var da = a['Date'] ? new Date(a['Date']) : new Date(0);
    var db = b['Date'] ? new Date(b['Date']) : new Date(0);
    return da - db;
  });

  var sheet = getHistorySheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) {
    return [];
  }

  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  // Build a list of stock columns from History, skipping the first column (Date)
  // and any IOZ/index column.
  var stockCols = [];
  for (var c = 2; c <= lastCol; c++) {
    var hRaw = header[c - 1];
    if (!hRaw) continue;
    var code = String(hRaw).trim();
    if (!code) continue;
    var upper = code.toUpperCase();
    if (upper === 'IOZ') continue;
    stockCols.push({ col: c, code: code });
  }
  if (stockCols.length === 0) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Track cumulative quantity per stock code as we move forward in time.
  var qtyByCode = {};
  var tradeIndex = 0;

  var series = [['Date', 'Market Value']];

  values.forEach(function (row) {
    var dtRaw = row[0];
    var dt;
    if (Object.prototype.toString.call(dtRaw) === '[object Date]') {
      dt = dtRaw;
    } else if (dtRaw) {
      dt = new Date(dtRaw);
    }
    if (!dt || isNaN(dt.getTime())) {
      return;
    }

    // Apply all trades up to and including this date.
    while (tradeIndex < trades.length) {
      var t = trades[tradeIndex];
      var tDate = t['Date'] ? new Date(t['Date']) : null;
      if (!tDate || isNaN(tDate.getTime())) {
        tradeIndex++;
        continue;
      }
      if (tDate > dt) {
        break;
      }

      var code = (t['Stock'] || '').toString().trim();
      if (code) {
        if (!qtyByCode[code]) {
          qtyByCode[code] = 0;
        }

        var type = (t['Buy/Sell'] || '').toString();
        var q = Number(t['Number'] || 0);
        if (type === 'Buy' || type === 'DRP') {
          qtyByCode[code] += q;
        } else if (type === 'Sell') {
          qtyByCode[code] -= q;
        }
      }

      tradeIndex++;
    }

    // Compute market value on this date using History prices.
    var totalValue = 0;
    stockCols.forEach(function (info) {
      var code = info.code;
      var qty = qtyByCode[code] || 0;
      if (!qty) return;

      var priceRaw = row[info.col - 1];
      var price = Number(priceRaw);
      if (isNaN(price) || price <= 0) return;

      totalValue += qty * price;
    });

    if (totalValue > 0) {
      var key = Utilities.formatDate(dt, APP_TIMEZONE, 'yyyy-MM-dd');
      series.push([key, totalValue]);
    }
  });

  if (series.length <= 1) {
    return [];
  }

  return series;
}

/**
 * Per-stock detailed summary for the Stock Details page.
 * Uses the same moving-average cost logic as getInvestmentsSummary, but
 * scoped to a single stock, and exposes:
 * - currentValue
 * - totalQty
 * - costBase (sum of all buys/DRP)
 * - costBasePerShare
 * - capitalGain (realised + unrealised)
 * - capitalGainPct (vs cost base)
 * - dividendsTotal (dividends + franking)
 * - totalReturn (capitalGain + dividends)
 * - totalReturnPct (vs cost base)
 * - latestPrice
 * - peRatio (Price / EPS, both from GOOGLEFINANCE via helper sheet)
 */
function getStockSummary(stockCode) {
  if (!stockCode) {
    return null;
  }

  var trades = getAllTrades();
  var filtered = trades.filter(function (t) {
    return t['Stock'] === stockCode;
  });

  if (filtered.length === 0) {
    return null;
  }

  // Ensure trades are processed in date order
  filtered.sort(function (a, b) {
    var da = a['Date'] ? new Date(a['Date']) : new Date(0);
    var db = b['Date'] ? new Date(b['Date']) : new Date(0);
    return da - db;
  });

  var qty = 0;                 // remaining quantity
  var remainingCost = 0;       // cost of remaining quantity
  var realisedGain = 0;        // realised gains/losses from sells
  var dividendsTotal = 0;      // dividends + franking
  var costBaseTotalBuys = 0;   // sum of all buy/DRP cost
  var firstTradeDate = null;   // earliest date we see for this stock

  filtered.forEach(function (t) {
    var type = t['Buy/Sell'];
    var tQty = Number(t['Number'] || 0);
    var total = Number(t['Total Cost'] || 0);
    var dividend = Number(t['Dividend'] || 0);
    var franking = Number(t['Franking'] || 0);
    var brokerage = Number(t['Brokerage'] || 0);
    var sellPrice = Number(t['Sell Price'] || 0);

    // Track earliest trade date for annualised calculations
    var rawDate = t['Date'] ? new Date(t['Date']) : null;
    if (rawDate && !firstTradeDate) {
      firstTradeDate = rawDate;
    }

    if (type === 'Buy' || type === 'DRP') {
      costBaseTotalBuys += total;
      remainingCost += total;
      qty += tQty;
    } else if (type === 'Sell') {
      if (qty > 0 && tQty > 0) {
        var avgCost = remainingCost / qty;
        var costOfSold = avgCost * tQty;

        var proceeds;
        if (sellPrice > 0 && sellPrice <= 1000 && tQty > 0) {
          proceeds = (sellPrice * tQty) - brokerage;
        } else if (sellPrice > 0) {
          proceeds = sellPrice - brokerage;
        } else if (total > 0) {
          proceeds = total - brokerage;
        } else {
          proceeds = 0;
        }

        realisedGain += (proceeds - costOfSold);
        qty -= tQty;
        remainingCost -= costOfSold;
      }
    } else if (type === 'Dividend') {
      dividendsTotal += (dividend + franking);
    }
  });

  var latestPrice = fetchLiveAsxPricePlaceholder(stockCode) || 0;
  var eps = fetchAsxEpsPlaceholder(stockCode) || 0;

  var currentValue = latestPrice * qty;
  var avgCostPerShare = qty > 0 ? (remainingCost / qty) : 0;

  // Unrealised gain on remaining position
  var unrealised = currentValue - remainingCost;
  var capitalGain = realisedGain + unrealised;

  var baseForPct = costBaseTotalBuys > 0 ? costBaseTotalBuys : 0;
  var capitalGainPct = baseForPct > 0 ? (capitalGain / baseForPct) : 0;
  var dividendsPct = baseForPct > 0 ? (dividendsTotal / baseForPct) : 0;
  var totalReturn = capitalGain + dividendsTotal;
  var totalReturnPct = baseForPct > 0 ? (totalReturn / baseForPct) : 0;

  // Simple annualised (p.a.) approximation based on earliest trade date
  var capitalGainPctAnnual = 0;
  var dividendsPctAnnual = 0;
  var totalReturnPctAnnual = 0;
  if (firstTradeDate && baseForPct > 0) {
    var now = new Date();
    var msPerDay = 1000 * 60 * 60 * 24;
    var daysHeld = Math.max(1, Math.round((now.getTime() - firstTradeDate.getTime()) / msPerDay));
    var yearsHeld = daysHeld / 365;
    if (yearsHeld > 0) {
      var cgBase = 1 + capitalGainPct;
      var divBase = 1 + dividendsPct;
      var trBase = 1 + totalReturnPct;

      capitalGainPctAnnual = cgBase > 0 ? Math.pow(cgBase, 1 / yearsHeld) - 1 : 0;
      dividendsPctAnnual = divBase > 0 ? Math.pow(divBase, 1 / yearsHeld) - 1 : 0;
      totalReturnPctAnnual = trBase > 0 ? Math.pow(trBase, 1 / yearsHeld) - 1 : 0;
    }
  }

  var peRatio = (latestPrice > 0 && eps > 0) ? (latestPrice / eps) : 0;

  return {
    stockCode: stockCode,
    // Core figures for Stock Details summary
    currentValue: currentValue,
    totalQty: qty,
    costBase: costBaseTotalBuys,
    costBasePerShare: avgCostPerShare,
    capitalGain: capitalGain,
    capitalGainPct: capitalGainPct,
    dividendsTotal: dividendsTotal,
    dividendsPct: dividendsPct,
    totalReturn: totalReturn,
    totalReturnPct: totalReturnPct,
    capitalGainPctAnnual: capitalGainPctAnnual,
    dividendsPctAnnual: dividendsPctAnnual,
    totalReturnPctAnnual: totalReturnPctAnnual,
    firstTradeDate: firstTradeDate
      ? Utilities.formatDate(firstTradeDate, APP_TIMEZONE, 'yyyy-MM-dd')
      : '',
    latestPrice: latestPrice,
    peRatio: peRatio,

    // Backwards-compatible fields in case other views still read them
    totalCost: costBaseTotalBuys,
    totalDividends: dividendsTotal,
    unrealisedPnl: unrealised,
    realisedPnl: realisedGain,
    cgtPosition: 0
  };
}

/**
 * Return a distinct list of stock codes in the ledger.
 */
function getAllStockCodes() {
  var trades = getAllTrades();
  var set = {};
  trades.forEach(function (t) {
    if (t['Stock']) {
      set[t['Stock']] = true;
    }
  });
  return Object.keys(set).sort();
}

/**
 * Return all trades in a JSON-safe form (strings/numbers only) so that
 * google.script.run can always serialise them back to the client.
 * This is used by the Stock Details view which then filters by stock
 * code on the client.
 */
function getAllTradesSafe() {
  var trades = getAllTrades();
  return trades.map(function (t) {
    return {
      Stock: (t['Stock'] || '').toString(),
      'Buy/Sell': (t['Buy/Sell'] || '').toString(),
      Date: t['Date']
        ? Utilities.formatDate(new Date(t['Date']), APP_TIMEZONE, 'yyyy-MM-dd')
        : '',
      Number: Number(t['Number'] || 0),
      Price: Number(t['Price'] || 0),
      Brokerage: Number(t['Brokerage'] || 0),
      'Total Cost': Number(t['Total Cost'] || 0),
      Dividend: Number(t['Dividend'] || 0),
      Franking: Number(t['Franking'] || 0),
      Unfranked: Number(t['Unfranked'] || 0),
      DRP: (t['DRP'] || '').toString(),
      'Div/Share': Number(t['Div/Share'] || 0),
      'Sell Price': Number(t['Sell Price'] || 0)
    };
  });
}

/**
 * Return all trades for a specific stock code, sorted by date descending,
 * wrapped in a simple object for safe serialisation to the client.
 * { marker, requestedCode, totalCount, filteredCount, trades }
 */
function getTradesForStock(stockCode) {
  var trades = getAllTrades();
  var total = trades.length;

  var codeNorm = stockCode ? String(stockCode).trim().toUpperCase() : '';
  Logger.log('getTradesForStock FINAL: total trades=' + total + ', requestedCode=' + stockCode + ', codeNorm=' + codeNorm);

  var filtered = trades.filter(function (t) {
    if (!codeNorm) return true; // if no code, return all
    var stock = (t['Stock'] || '').toString().trim().toUpperCase();
    return stock === codeNorm;
  });

  filtered.sort(function (a, b) {
    var da = a['Date'] ? new Date(a['Date']) : new Date(0);
    var db = b['Date'] ? new Date(b['Date']) : new Date(0);
    return db - da;
  });

  Logger.log('getTradesForStock FINAL: filteredCount=' + filtered.length + ' for codeNorm=' + codeNorm);

  return {
    marker: 'TRADES_PAYLOAD',
    requestedCode: stockCode,
    totalCount: total,
    filteredCount: filtered.length,
    trades: filtered
  };
}

/**
 * Per-stock investment summary for dashboard "My Investments" table.
 * Uses a simple moving-average cost method per stock:
 * - Tracks remaining quantity and cost as you buy/DRP/sell.
 * - Each sell realises gain based on current average cost.
 * - Unrealised gain is based on remaining quantity and latest price.
 */
function getInvestmentsSummary() {
  var trades = getAllTrades();

  // Ensure trades are processed in date order for each stock
  trades.sort(function (a, b) {
    var da = a['Date'] ? new Date(a['Date']) : new Date(0);
    var db = b['Date'] ? new Date(b['Date']) : new Date(0);
    return da - db;
  });

  var byCode = {};

  trades.forEach(function (t) {
    var code = t['Stock'];
    if (!code) return;
    if (!byCode[code]) {
      byCode[code] = {
        code: code,
        qty: 0,           // remaining quantity
        remainingCost: 0, // total cost of remaining quantity
        realisedGain: 0,  // realised gains/losses from sells
        dividends: 0,
        latestPrice: 0,
        capitalGains: 0,
        returnTotal: 0
      };
    }

    var entry = byCode[code];
    var type = t['Buy/Sell'];
    var qty = Number(t['Number'] || 0);
    var total = Number(t['Total Cost'] || 0);
    var dividend = Number(t['Dividend'] || 0);
    var franking = Number(t['Franking'] || 0);
    var brokerage = Number(t['Brokerage'] || 0);
    var sellPrice = Number(t['Sell Price'] || 0);

    if (type === 'Buy' || type === 'DRP') {
      // Increase remaining quantity and cost by the buy/DRP
      entry.remainingCost += total;
      entry.qty += qty;
    } else if (type === 'Sell') {
      if (entry.qty > 0 && qty > 0) {
        var avgCost = entry.remainingCost / entry.qty;
        var costOfSold = avgCost * qty;

        // Determine proceeds from the sell. In your ledger layout, the
        // "Sell Price" column is currently used as the *total* sale
        // value (e.g. 22,253.28) rather than price per share. To make
        // this robust, we treat reasonable values (<= 1000) as a per-
        // share price, and larger values as total proceeds.
        var proceeds;
        if (sellPrice > 0 && sellPrice <= 1000 && qty > 0) {
          proceeds = (sellPrice * qty) - brokerage;
        } else if (sellPrice > 0) {
          proceeds = sellPrice - brokerage;
        } else if (total > 0) {
          proceeds = total - brokerage;
        } else {
          proceeds = 0;
        }

        var realised = proceeds - costOfSold;
        entry.realisedGain += realised;

        // Reduce remaining position
        entry.qty -= qty;
        entry.remainingCost -= costOfSold;
      }
    } else if (type === 'Dividend') {
      entry.dividends += (dividend + franking);
    }
  });

  // Enrich with latest price, avg price, capital gains and total return
  Object.keys(byCode).forEach(function (code) {
    var entry = byCode[code];
    var latestPrice = fetchLiveAsxPricePlaceholder(code) || 0;
    entry.latestPrice = latestPrice;

    var remainingQty = entry.qty;
    var remainingCost = entry.remainingCost;
    var avgCost = remainingQty > 0 ? (remainingCost / remainingQty) : 0;
    entry.avgPrice = avgCost;

    // Market value and unrealised gain on remaining position
    var marketValue = latestPrice * remainingQty;
    entry.value = marketValue;
    var unrealised = marketValue - remainingCost;

    var totalCapitalGains = entry.realisedGain + unrealised;
    entry.capitalGains = totalCapitalGains;
    entry.returnTotal = totalCapitalGains + entry.dividends;
  });

  // Return sorted list by code
  return Object.keys(byCode).sort().map(function (code) { return byCode[code]; });
}

/**
 * Placeholder FIFO/CGT helper recalculation.
 *
 * This function is where you would:
 * - Walk through the ledger in date order
 * - Maintain FIFO parcels for each stock
 * - Track realised gains on sells and apply 12-month 50% CGT discount where applicable
 * - Write helper columns back to the sheet if desired (e.g. remaining parcel qty, cost base, realised CGT, discount, etc.)
 */
function recalculatePortfolioHelpers() {
  var sheet = getLedgerSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  // TODO: Implement FIFO parcel tracking and CGT helper columns.
  // This skeleton intentionally leaves the implementation to you.
}

/**
 * Placeholder taxable income report (dividends and franking) for a given date range.
 */
function getTaxReport(startDate, endDate) {
  var trades = getAllTrades();
  var start = startDate ? new Date(startDate) : null;
  var end = endDate ? new Date(endDate) : null;

  var totalDividends = 0;
  var totalFranking = 0;
  var totalUnfranked = 0;

  trades.forEach(function (t) {
    if (t['Buy/Sell'] !== 'Dividend') return;
    var d = t['Date'] ? new Date(t['Date']) : null;
    if (!d) return;
    if (start && d < start) return;
    if (end && d > end) return;

    totalDividends += Number(t['Dividend'] || 0);
    totalFranking += Number(t['Franking'] || 0);
    totalUnfranked += Number(t['Unfranked'] || 0);
  });

  return {
    totalDividends: totalDividends,
    totalFranking: totalFranking,
    totalUnfranked: totalUnfranked
  };
}

/**
 * Placeholder capital gains report using FIFO and 12-month discount.
 * Currently only returns an empty structure ready to be populated.
 */
function getCgtReport(startDate, endDate) {
  var trades = getAllTrades();
  if (!trades || trades.length === 0) {
    return {
      realisedGains: 0,
      discountedGains: 0,
      shortTermGains: 0,
      longTermGains: 0,
      entries: []
    };
  }

  var start = startDate ? new Date(startDate) : null;
  var end = endDate ? new Date(endDate) : null;

  // Sort trades by date ascending so FIFO parcels behave correctly.
  trades.sort(function (a, b) {
    var da = a['Date'] ? new Date(a['Date']) : new Date(0);
    var db = b['Date'] ? new Date(b['Date']) : new Date(0);
    return da - db;
  });

  var parcelsByCode = {};
  var entries = [];

  trades.forEach(function (t) {
    var type = (t['Buy/Sell'] || '').toString();
    var code = (t['Stock'] || '').toString().trim();
    var qty = Number(t['Number'] || 0);
    var totalCost = Number(t['Total Cost'] || 0);
    var brokerage = Number(t['Brokerage'] || 0);
    var sellPrice = Number(t['Sell Price'] || 0);
    var d = t['Date'] ? new Date(t['Date']) : null;
    if (!code || !d || isNaN(d.getTime())) return;

    if (!parcelsByCode[code]) {
      parcelsByCode[code] = [];
    }

    if (type === 'Buy' || type === 'DRP') {
      if (qty > 0 && totalCost > 0) {
        parcelsByCode[code].push({
          acqDate: d,
          qtyRemaining: qty,
          costBaseRemaining: totalCost
        });
      }
    } else if (type === 'Sell' && qty > 0) {
      // Work out total proceeds using the same heuristic as elsewhere.
      var proceeds;
      if (sellPrice > 0 && sellPrice <= 1000 && qty > 0) {
        proceeds = (sellPrice * qty) - brokerage;
      } else if (sellPrice > 0) {
        proceeds = sellPrice - brokerage;
      } else if (totalCost > 0) {
        proceeds = totalCost - brokerage;
      } else {
        proceeds = 0;
      }

      var remainingToSell = qty;
      var saleParcels = parcelsByCode[code];
      if (!saleParcels || saleParcels.length === 0) {
        return;
      }

      // We will allocate proceeds across slices proportionally to quantity.
      var totalSaleQty = qty;

      for (var i = 0; i < saleParcels.length && remainingToSell > 0; i++) {
        var parcel = saleParcels[i];
        if (parcel.qtyRemaining <= 0) continue;

        var sliceQty = Math.min(parcel.qtyRemaining, remainingToSell);
        if (sliceQty <= 0) continue;

        // Cost base slice proportional to quantity remaining in this parcel.
        var costPerShare = parcel.costBaseRemaining / parcel.qtyRemaining;
        var costSlice = costPerShare * sliceQty;

        // Proceeds slice proportional to quantity sold in this slice.
        var proceedsSlice = proceeds * (sliceQty / totalSaleQty);

        // Update parcel remaining balance.
        parcel.qtyRemaining -= sliceQty;
        parcel.costBaseRemaining -= costSlice;
        remainingToSell -= sliceQty;

        var gain = proceedsSlice - costSlice;

        // Determine holding period in days for discount eligibility.
        var msPerDay = 1000 * 60 * 60 * 24;
        var daysHeld = Math.round((d.getTime() - parcel.acqDate.getTime()) / msPerDay);
        var isLongTerm = daysHeld >= 365;

        var discountApplied = isLongTerm && gain > 0;
        var discountedGain = discountApplied ? (gain * 0.5) : gain;

        // Only include entries whose SELL date falls within the requested range.
        if ((!start || d >= start) && (!end || d <= end)) {
          entries.push({
            date: Utilities.formatDate(d, APP_TIMEZONE, 'yyyy-MM-dd'),
            stock: code,
            quantity: sliceQty,
            proceeds: proceedsSlice,
            costBase: costSlice,
            gain: gain,
            discountApplied: discountApplied,
            discountedGain: discountedGain,
            longTerm: isLongTerm
          });
        }
      }
    }
  });

  var realisedGains = 0;
  var discountedGains = 0;
  var shortTermGains = 0;
  var longTermGains = 0;

  entries.forEach(function (e) {
    realisedGains += e.gain;
    discountedGains += e.discountedGain;
    if (e.longTerm) {
      longTermGains += e.gain;
    } else {
      shortTermGains += e.gain;
    }
  });

  return {
    realisedGains: realisedGains,
    discountedGains: discountedGains,
    shortTermGains: shortTermGains,
    longTermGains: longTermGains,
    entries: entries
  };
}

function getHistorySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET_NAME);
  }
  return sheet;
}

function isBusinessDay(d) {
  var day = d.getDay();
  return day !== 0 && day !== 6;
}

function fetchDailyCloseForHistory(ticker, date) {
  if (!ticker || !date) {
    return 0;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('HistoryTemp');
  if (!sheet) {
    sheet = ss.insertSheet('HistoryTemp');
  }

  sheet.getRange(1, 1).setValue('ASX:' + ticker);
  sheet.getRange(1, 2).setValue(date);

  var valueCell = sheet.getRange(1, 3);
  // Use historical close price: GOOGLEFINANCE(ticker, "close", startDate, 1)
  // This returns a header row + one data row for the requested date.
  valueCell.setFormula('=INDEX(GOOGLEFINANCE(A1,"close",B1,1),2,2)');

  Utilities.sleep(1000);
  var value = Number(valueCell.getValue() || 0);
  sheet.getRange(1, 1, 1, 3).clearContent();

  if (!isNaN(value) && value > 0) {
    Logger.log('fetchDailyCloseForHistory: ' + ticker + ' @ ' + date + ' -> ' + value);
    return value;
  }

  Logger.log('fetchDailyCloseForHistory: ' + ticker + ' @ ' + date + ' returned 0/NaN, attempting fallback to previous History value');

  // Fallback: use the last known non-zero close for this ticker from the
  // main History sheet, so we avoid dropping to 0 on days where
  // GOOGLEFINANCE has no data (e.g. ETFs/index funds on some dates).
  try {
    var historySheet = getHistorySheet();
    var lastRow = historySheet.getLastRow();
    var lastCol = historySheet.getLastColumn();
    if (lastRow >= 2 && lastCol >= 2) {
      var header = historySheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var targetCol = -1;
      var tUpper = String(ticker).trim().toUpperCase();
      for (var c = 1; c < header.length; c++) { // skip col A (date)
        var hRaw = header[c];
        if (!hRaw) continue;
        var hUpper = String(hRaw).trim().toUpperCase();
        if (hUpper === tUpper) {
          targetCol = c + 1; // header index -> 1-based column
          break;
        }
      }

      if (targetCol !== -1) {
        for (var r = lastRow; r >= 2; r--) {
          var cellVal = Number(historySheet.getRange(r, targetCol).getValue() || 0);
          if (!isNaN(cellVal) && cellVal > 0) {
            Logger.log('fetchDailyCloseForHistory: fallback for ' + ticker + ' using previous History value ' + cellVal + ' from row ' + r);
            return cellVal;
          }
        }
      }
    }
  } catch (e) {
    Logger.log('fetchDailyCloseForHistory: error during fallback for ' + ticker + ' -> ' + (e && e.message));
  }

  Logger.log('fetchDailyCloseForHistory: no fallback found for ' + ticker + ', returning 0');
  return 0;
}

function updateHistorySheetDaily() {
  var sheet = getHistorySheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) {
    Logger.log('updateHistorySheetDaily: exiting, insufficient data (lastRow=' + lastRow + ', lastCol=' + lastCol + ')');
    return;
  }

  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var tickers = [];
  for (var i = 1; i < header.length; i++) {
    if (header[i]) {
      tickers.push(String(header[i]).trim().toUpperCase());
    }
  }
  if (tickers.length === 0) {
    Logger.log('updateHistorySheetDaily: exiting, no tickers found in header row');
    return;
  }

  var lastDateVal = sheet.getRange(lastRow, 1).getValue();
  if (!lastDateVal) {
    Logger.log('updateHistorySheetDaily: exiting, last date cell empty at row ' + lastRow);
    return;
  }

  var lastDate = new Date(lastDateVal);
  if (isNaN(lastDate.getTime())) {
    Logger.log('updateHistorySheetDaily: exiting, last date is NaN -> ' + lastDateVal);
    return;
  }

  var today = new Date();
  var targetEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);

  var lastDay = new Date(lastDate.getFullYear(), lastDate.getMonth(), lastDate.getDate());
  if (lastDay >= targetEnd) {
    Logger.log('updateHistorySheetDaily: nothing to do (lastDay=' + lastDay.toDateString() + ', targetEnd=' + targetEnd.toDateString() + ')');
    return;
  }

  Logger.log('updateHistorySheetDaily: starting fill from ' + lastDay.toDateString() + ' to ' + targetEnd.toDateString() + ' for tickers: ' + tickers.join(','));

  var rowsToAppend = [];
  var d = new Date(lastDay.getTime());
  d.setDate(d.getDate() + 1);

  while (d <= targetEnd) {
    if (isBusinessDay(d)) {
      var prices = [];
      for (var j = 0; j < tickers.length; j++) {
        prices.push(fetchDailyCloseForHistory(tickers[j], d));
      }
      rowsToAppend.push([new Date(d.getFullYear(), d.getMonth(), d.getDate())].concat(prices));
    }
    d.setDate(d.getDate() + 1);
  }

  if (rowsToAppend.length > 0) {
    Logger.log('updateHistorySheetDaily: appending ' + rowsToAppend.length + ' new row(s) starting at row ' + (lastRow + 1));
    sheet.getRange(lastRow + 1, 1, rowsToAppend.length, 1 + tickers.length).setValues(rowsToAppend);
  } else {
    Logger.log('updateHistorySheetDaily: no business days found between lastDay and targetEnd, nothing to append');
  }
}

/**
 * Build an indexed performance series for a single stock vs ASX200 proxy (IOZ)
 * using static history data stored in the History sheet.
 *
 * Expected History layout:
 *   Col A: Date
 *   Col B: IOZ (benchmark)
 *   Col C+: One column per stock, with header equal to the stock code
 *
 * Returns an array-of-arrays suitable for Google Charts:
 *   [ ['Date', 'CODE', 'ASX 200'], [date, stockIndex, asxIndex], ... ]
 */
function getStockVsIndexSeries(stockCode) {
  if (!stockCode) return [];

  var sheet = getHistorySheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) {
    return [];
  }

  // Read header row to find IOZ column and the stock column
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var iozCol = -1;
  var stockCol = -1;
  var targetStock = String(stockCode).trim().toUpperCase();

  for (var c = 0; c < header.length; c++) {
    var hRaw = header[c];
    if (!hRaw) continue;
    var h = String(hRaw).trim().toUpperCase();
    if (h === 'IOZ') {
      iozCol = c + 1; // 1-based
    }
    if (h === targetStock) {
      stockCol = c + 1; // 1-based
    }
  }

  // Require both IOZ and stock columns to exist
  if (iozCol === -1 || stockCol === -1) {
    return [];
  }

  // Read all rows from row 2 downwards and stream-build the series.
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var baseStock = null;
  var baseIndex = null;
  var series = [];
  series.push(['Date', stockCode, 'ASX 200']);

  values.forEach(function (row) {
    var dtRaw = row[0];
    var stockPriceRaw = row[stockCol - 1];
    var indexPriceRaw = row[iozCol - 1];

    // Accept either Date objects or date-like strings
    var dt;
    if (Object.prototype.toString.call(dtRaw) === '[object Date]') {
      dt = dtRaw;
    } else if (dtRaw) {
      dt = new Date(dtRaw);
    }
    if (!dt || isNaN(dt.getTime())) {
      return;
    }

    var stockPrice = Number(stockPriceRaw);
    var indexPrice = Number(indexPriceRaw);
    if (isNaN(stockPrice) || isNaN(indexPrice)) {
      return;
    }

    if (baseStock === null || baseIndex === null) {
      baseStock = stockPrice;
      baseIndex = indexPrice;
      if (!baseStock || !baseIndex) {
        return;
      }
    }

    var key = Utilities.formatDate(dt, APP_TIMEZONE, 'yyyy-MM-dd');
    var stockIndexed = (stockPrice / baseStock) * 100;
    var indexIndexed = (indexPrice / baseIndex) * 100;
    // Use a simple string key to keep the payload JSON-safe and easy to debug
    series.push([key, stockIndexed, indexIndexed]);
  });

  if (series.length <= 1) {
    return [];
  }

  return series;
}

/**
 * Placeholder for live ASX price lookup.
 * Replace with real API endpoint and parsing logic.
 */
function fetchLiveAsxPricePlaceholder(asxCode) {
  if (!asxCode) return 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PRICE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PRICE_SHEET_NAME);
    sheet.getRange(1, 1).setValue('Stock');
    sheet.getRange(1, 2).setValue('Price');
  }

  // Find existing row for this code in column A
  var lastRow = sheet.getLastRow();
  var values = [];
  if (lastRow > 1) {
    var range = sheet.getRange(2, 1, lastRow - 1, 1);
    values = range.getValues();
  }
  var rowIndex = null;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === asxCode) {
      rowIndex = i + 2; // offset for header
      break;
    }
  }

  if (!rowIndex) {
    // Append new row for this code
    rowIndex = lastRow + 1;
    sheet.getRange(rowIndex, 1).setValue(asxCode);
    var formula = '=GOOGLEFINANCE("ASX:"&A' + rowIndex + ',"price")';
    sheet.getRange(rowIndex, 2).setFormula(formula);
  }

  // Read the current price value from column B
  var priceCell = sheet.getRange(rowIndex, 2);
  var price = Number(priceCell.getValue() || 0);
  if (!isNaN(price) && price > 0) {
    return price;
  }

  // If the formula just got set, it may not have calculated yet.
  // Give Sheets a brief moment and try once more.
  Utilities.sleep(500);
  price = Number(priceCell.getValue() || 0);
  if (!isNaN(price) && price > 0) {
    return price;
  }

  return 0;
}

/**
 * Placeholder for ASX EPS lookup using GOOGLEFINANCE.
 * Stores EPS in the same Prices sheet as prices, in column C.
 */
function fetchAsxEpsPlaceholder(asxCode) {
  if (!asxCode) return 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PRICE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PRICE_SHEET_NAME);
    sheet.getRange(1, 1).setValue('Stock');
    sheet.getRange(1, 2).setValue('Price');
    sheet.getRange(1, 3).setValue('EPS');
  }

  // Find existing row for this code in column A
  var lastRow = sheet.getLastRow();
  var values = [];
  if (lastRow > 1) {
    var range = sheet.getRange(2, 1, lastRow - 1, 1);
    values = range.getValues();
  }
  var rowIndex = null;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === asxCode) {
      rowIndex = i + 2; // offset for header
      break;
    }
  }

  if (!rowIndex) {
    // Append new row for this code
    rowIndex = lastRow + 1;
    sheet.getRange(rowIndex, 1).setValue(asxCode);
    // Ensure headers exist
    if (sheet.getRange(1, 2).getValue() !== 'Price') {
      sheet.getRange(1, 2).setValue('Price');
    }
    if (sheet.getRange(1, 3).getValue() !== 'EPS') {
      sheet.getRange(1, 3).setValue('EPS');
    }
  } else {
    // Ensure EPS header exists for existing sheet
    if (sheet.getRange(1, 3).getValue() !== 'EPS') {
      sheet.getRange(1, 3).setValue('EPS');
    }
  }

  // Ensure there is a GOOGLEFINANCE eps formula for this row
  var epsCell = sheet.getRange(rowIndex, 3);
  var eps = Number(epsCell.getValue() || 0);
  if ((epsCell.getFormula() || '') === '' || eps === 0) {
    var formula = '=GOOGLEFINANCE("ASX:"&A' + rowIndex + ',"eps")';
    epsCell.setFormula(formula);
  }
  if (!isNaN(eps) && eps !== 0) {
    return eps;
  }

  // Give Sheets a brief moment and try once more.
  Utilities.sleep(500);
  eps = Number(epsCell.getValue() || 0);
  if (!isNaN(eps) && eps !== 0) {
    return eps;
  }

  return 0;
}

// Helper for HTML templating to include partials if you later add them.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Return the HTML for a named view partial. The client-side SPA
// uses this to swap views without reloading the whole page.
function getView(page) {
  var allowed = {
    index: 'view_index',
    addTrade: 'view_addTrade',
    stockDetails: 'view_stockDetails',
    report_tax: 'view_report_tax',
    report_cgt: 'view_report_cgt',
    about: 'view_about'
  };

  var file = allowed[page] || allowed.index;
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

function getHistoryStatus() {
  var sheet = getHistorySheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return {
      hasHistory: false,
      lastDate: '',
      lastRow: lastRow
    };
  }

  var lastDateVal = sheet.getRange(lastRow, 1).getValue();
  var lastDateStr = '';
  if (lastDateVal) {
    var d = new Date(lastDateVal);
    if (!isNaN(d.getTime())) {
      lastDateStr = Utilities.formatDate(d, APP_TIMEZONE, 'yyyy-MM-dd');
    }
  }

  return {
    hasHistory: !!lastDateStr,
    lastDate: lastDateStr,
    lastRow: lastRow
  };
}
