// ==================== Global Variables ==================== //

// Get current spreadsheet
const currSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// Get enhanced cache and properties on top of Google's Cache service. Credit: https://github.com/yinonavraham/GoogleAppsScripts/blob/master/EnhancedCacheService
const cache = wrap(CacheService.getDocumentCache()); // For caching data

// Get the sheets within current spreadsheet
const setupSheet = currSpreadsheet.getSheetByName("Setup");
const portfolioSummarySheet = currSpreadsheet.getSheetByName(
  "Portfolio Summary"
);
const portfolioHistorySheet = currSpreadsheet.getSheetByName(
  "Portfolio History"
);
const stockSummarySheet = currSpreadsheet.getSheetByName("Stock Summary");
const transactionsSheet = currSpreadsheet.getSheetByName("Transactions");
const divSplitRefSheet = currSpreadsheet.getSheetByName("Div/Split Ref");
const refSheet = currSpreadsheet.getSheetByName("Ref");

// ==================== Custom Functions ==================== //

// ==================== Setup Functions and Custom Menu ==================== //
function runSetup() {
  const setupButtonRange = setupSheet.getRange("A9:C12");
  const ui = SpreadsheetApp.getUi();

  // Check if setup has finished before, if yes, display error, else run setup
  if (
    getAdvancedUserSetting("setupComplete") == 1 ||
    getAdvancedUserSetting("setupComplete") == "1"
  ) {
    // Flash message to alert user that setup had already been done before
    setupButtonRange.setValue("SETUP COMPLETE. DO NOT CLICK AGAIN!     ");

    let result = ui.alert(
      "Couldn't resist, huh? :D",
      "This spreadsheet has already been setup. Clicking the setup button will do nothing.",
      ui.ButtonSet.OK
    );
  }
  // Set up not done before
  else if (
    getAdvancedUserSetting("setupComplete") == 0 ||
    getAdvancedUserSetting("setupComplete") == "0"
  ) {
    // Update button UI
    // Update user setting so the setup button wont run twice
    setupButtonRange.setValue("SETUP COMPLETE. DO NOT CLICK AGAIN!     ");
    setAdvancedUserSetting("setupComplete", 1);

    // Set tab color to green to show setup complete
    setupSheet.setTabColor("#6bcf3e");

    Utilities.sleep(100);

    // Set all triggers

    // 1. Create time-driven trigger(s) to check for and add any dividends and splits for the stocks the user owns, repeat X times a day
    const divSplitCheckTimes = [10, 17, 23]; // Check at 10am, 5pm and 11pm (market open, market close, end of day)

    for (
      let checkTimeIndex = 0;
      checkTimeIndex < divSplitCheckTimes.length;
      checkTimeIndex++
    ) {
      ScriptApp.newTrigger("addDividendSplit")
        .timeBased()
        .atHour(divSplitCheckTimes[checkTimeIndex])
        .everyDays(1)
        .create();

      Utilities.sleep(100);
    }

    // 2. Install time-driven trigger to copy portfolio summary every night to history at midnight
    ScriptApp.newTrigger("copyPortfolioSummaryToHistory")
      .timeBased()
      .atHour(0) // Trigger at midnight to copy portfolio summary to history
      .everyDays(1)
      .create();

    Utilities.sleep(100);

    // 3. Create time-driven trigger to delete all rows in stock summary with no more stocks owned AFTER copying to portfolio history is complete
    ScriptApp.newTrigger("cleanUpStockSummary")
      .timeBased()
      .atHour(1) // Trigger at 1am clean up stock summary after copying to portfolio history
      .everyDays(1)
      .create();

    Utilities.sleep(100);

    // 4. Clear checked dividends and splits after every day at 2am before first dividend check for new day
    ScriptApp.newTrigger("clearCheckedDivSplit")
      .timeBased()
      .atHour(2) // Trigger at 2am to clear today's dividends and splits
      .everyDays(1)
      .create();

    Utilities.sleep(100);

    // Confirmation that setup is complete
    let result = ui.alert(
      "Setup Complete",
      "The setup has been completed and the spreadsheet is operational. Go to the Github link to read detailed instructions on how to use the spreadsheet. Do not click the setup button again.",
      ui.ButtonSet.OK
    );
  } // Error
  else {
    let result = ui.alert(
      "Error",
      "Please check that the user setting for setupComplete is either 0 or 1.",
      ui.ButtonSet.OK
    );
  }
}

// Resets setup and deletes all triggers
function resetSetup() {
  const setupButtonRange = setupSheet.getRange("A9:C12");
  const ui = SpreadsheetApp.getUi();

  // Set tab color to red and reset setup button words
  setAdvancedUserSetting("setupComplete", 0);
  setupSheet.setTabColor("#fd4141");
  setupButtonRange.clearContent();

  // Delete all triggers
  deleteAllTriggers();

  let result = ui.alert(
    "Setup Reset",
    "The setup has been reset. All triggers have been removed. Please click the setup button again.",
    ui.ButtonSet.OK
  );
}

// Deletes all triggers
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
    Utilities.sleep(200);
  }
}

// When spreadsheet opens, create a custom menu to add new transaction row
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Portfolio Menu")
    .addSubMenu(ui.createMenu("Setup").addItem("Reset Setup", "resetSetup"))
    .addSubMenu(
      ui
        .createMenu("Stock Summary")
        .addItem("Sort by Category", "sortStockSummaryByCategory")
        .addItem("Sort by Sector", "sortStockSummaryBySector")
        .addItem("Sort by Stock Name", "sortStockSummaryByStockName")
    )
    .addSubMenu(
      ui
        .createMenu("Transactions")
        .addItem("Add New Transaction", "addNewTransactionRow")
        .addItem("Add New Buy Transaction", "addNewBuyTransactionRow")
    )
    .addToUi();
}

// ==================== ADMINSTRATIVE FUNCTIONS ==================== //

// ==================== Read/Write User Settings ==================== //

// Returns a user setting's value
function getUserSetting(settingName) {
  // Get positions of the settings range
  const settingsPos = {
    startRow: "16",
    endRow: "18",
    colNames: "A",
    colVals: "B",
  };

  // Get values
  const settingsVals = setupSheet
    .getRange(
      settingsPos.colNames +
        settingsPos.startRow +
        ":" +
        settingsPos.colVals +
        settingsPos.endRow
    )
    .getValues();

  // Assign each setting value as a property under the settings object
  for (let row = 0; row < settingsVals.length; row++) {
    if (settingsVals[row][0] == settingName) {
      return settingsVals[row][1];
    }
  }

  throw new Error("Setting not found!");
}

// Gets an advanced user setting's value
function getAdvancedUserSetting(settingName) {
  // Get positions of the settings range
  const settingsPos = {
    startRow: "23",
    endRow: "24",
    colNames: "A",
    colVals: "B",
  };

  // Get values
  const settingsVals = setupSheet
    .getRange(
      settingsPos.colNames +
        settingsPos.startRow +
        ":" +
        settingsPos.colVals +
        settingsPos.endRow
    )
    .getValues();

  // Assign each setting value as a property under the settings object
  for (let row = 0; row < settingsVals.length; row++) {
    if (settingsVals[row][0] == settingName) {
      return settingsVals[row][1];
    }
  }

  throw new Error("Setting not found!");
}

// Sets an advanced user setting to a new value
function setAdvancedUserSetting(settingName, newValue) {
  // Get positions of the settings range
  const settingsPos = {
    startRow: "23",
    endRow: "24",
    colNames: "A",
    colVals: "B",
  };

  // Get values
  const settingsVals = setupSheet
    .getRange(
      settingsPos.colNames +
        settingsPos.startRow +
        ":" +
        settingsPos.colVals +
        settingsPos.endRow
    )
    .getValues();

  // Match the setting names and set the value
  for (let row = 0; row < settingsVals.length; row++) {
    if (settingsVals[row][0] == settingName) {
      setupSheet
        .getRange(
          settingsPos.colVals + (eval(settingsPos.startRow) + row).toString()
        )
        .setValue(newValue);
      return;
    }
  }
}

// Check user settings if updated to update spreadsheet
function updateUserSettings() {
  // Set timezone
  const chosenTimezone = getUserSetting("timezone");
  currSpreadsheet.setSpreadsheetTimeZone(chosenTimezone);

  Utilities.sleep(400);

  // Set locale (requires refresh sometimes)
  const localeMap = { US: "en_US", UK: "en_GB" };
  const chosenLocale = getUserSetting("locale");
  currSpreadsheet.setSpreadsheetLocale(localeMap[chosenLocale]);
}

// ==================== PORTFOLIO HISTORY FUNCTIONS ==================== //

// Function to move a copy of today's Portfolio Summary into Portfolio History at a specified time (default 12am-1am)
function copyPortfolioSummaryToHistory() {
  return;

  // Positions of cells where data can be found in portfolio summary
  var cashPos = { row: "2", col: "B" };
  var totalsPos = {
    row: "7",
    portfolioCostCol: "B",
    portfolioValueCol: "D",
    unrealizedGainLossCol: "F",
    realizedGainLossCol: "H",
    dividendsCollectedCol: "I",
    expectedDividendsCol: "J",
  };

  var portfolioHistoryPos = { startCol: "A", endCol: "H" };

  // Get values from portfolio summary table
  var cashVal = portfolioSummarySheet
    .getRange(cashPos.col + cashPos.row)
    .getValue();
  var portfolioCostVal = portfolioSummarySheet
    .getRange(totalsPos.portfolioCostCol + totalsPos.row)
    .getValue();
  var portfolioValueVal = portfolioSummarySheet
    .getRange(totalsPos.portfolioValueCol + totalsPos.row)
    .getValue();
  var unrealizedGainLossVal = portfolioSummarySheet
    .getRange(totalsPos.unrealizedGainLossCol + totalsPos.row)
    .getValue();
  var realizedGainLossVal = portfolioSummarySheet
    .getRange(totalsPos.realizedGainLossCol + totalsPos.row)
    .getValue();
  var dividendsCollectedVal = portfolioSummarySheet
    .getRange(totalsPos.dividendsCollectedCol + totalsPos.row)
    .getValue();
  var expectedDividendsVal = portfolioSummarySheet
    .getRange(totalsPos.expectedDividendsCol + totalsPos.row)
    .getValue();

  // Create new row of values to copy into portfolio history
  var today = new Date();
  var arrNewRow = [
    today,
    portfolioCostVal,
    portfolioValueVal,
    cashVal,
    unrealizedGainLossVal,
    realizedGainLossVal,
    dividendsCollectedVal,
    expectedDividendsVal,
  ];

  // Add this new row below the last row in Portfolio History sheet
  portfolioHistorySheet.appendRow(arrNewRow);

  // Format the row's number and date formats

  // We need to find the last row with values (aka the row we just appended), and then edit it
  var portfolioHistoryStartColVals = portfolioHistorySheet
    .getRange(
      portfolioHistoryPos.startCol + "1:" + portfolioHistoryPos.startCol
    )
    .getValues();
  var tempSS = currSpreadsheet.insertSheet("temporarySheet"); // Create a temporary sheet
  tempSS
    .getRange(
      1,
      1,
      portfolioHistoryStartColVals.length,
      portfolioHistoryStartColVals[0].length
    )
    .setValues(portfolioHistoryStartColVals); // Copy first col values into temporary sheet
  var lastRowIndex = tempSS.getLastRow(); // Get the last non-null row index in the temporary sheet
  currSpreadsheet.deleteSheet(currSpreadsheet.getSheetByName("temporarySheet")); // Delete the temporary sheet

  // Specify the formats for date time and other cells
  var rowFormat = [
    [
      "ddd, dd MMM yyyy, HH:mm:ss",
      "$#,###,###.00",
      "$#,###,###.00",
      "$#,###,###.00",
      "$#,###,###.00",
      "$#,###,###.00",
      "$#,###,###.00",
      "$#,###,###.00",
    ],
  ];

  // Get the last row and format it
  var lastRow = portfolioHistorySheet.getRange(
    portfolioHistoryPos.startCol +
      lastRowIndex +
      ":" +
      portfolioHistoryPos.endCol +
      lastRowIndex
  );

  // Set the formats for the last row
  lastRow.setNumberFormats(rowFormat);
  lastRow.setBackground("#ccffff");
  lastRow.setVerticalAlignment("middle"); // “top” or “middle” or “bottom”
  lastRow.setFontFamily("Arial");
  lastRow.setFontSize(9);
  lastRow.setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    "#cccccc",
    SpreadsheetApp.BorderStyle.SOLID
  ); // setBorder(top, left, bottom, right, vertical, horizontal, color, SpreadsheetApp.BorderStyle.DASHED/DOTTED/SOLID)
  lastRow.setHorizontalAlignment("center");
}

// ==================== STOCK SUMMARY FUNCTIONS ==================== //

// ==================== Stock Summary Sorting Functions ==================== //

// Columns in stock summary to sort by and range to sort
const stockSummarySortPos = {
  category: 1,
  sector: 2,
  stockName: 4,
  sortRange: "A3:AC",
  lastRowColRange: "Stock Summary!C:C",
};

// Get range of filled values, update the range below if changes are made to stock summary
// We find last filled row of values under the ticker symbols
function sortStockSummaryByCategory() {
  let stockSummaryRange = stockSummarySheet.getRange(
    stockSummarySortPos.sortRange +
      getLastRowIndex(stockSummarySortPos.lastRowColRange).toString()
  );
  stockSummaryRange.sort({
    column: stockSummarySortCols.category,
    ascending: true,
  });
}

function sortStockSummaryBySector() {
  let stockSummaryRange = stockSummarySheet.getRange(
    stockSummarySortPos.sortRange +
      getLastRowIndex(stockSummarySortPos.lastRowColRange).toString()
  );
  stockSummaryRange.sort({
    column: stockSummarySortCols.sector,
    ascending: true,
  });
}

function sortStockSummaryByStockName() {
  let stockSummaryRange = stockSummarySheet.getRange(
    stockSummarySortPos.sortRange +
      getLastRowIndex(stockSummarySortPos.lastRowColRange).toString()
  );
  stockSummaryRange.sort({
    column: stockSummarySortCols.stockName,
    ascending: true,
  });
}

// ==================== Stock Summary Remove Stocks No Longer Owned Function ==================== //

// Removes all stocks in stock summary that no longer have any units owned. This functions runs AFTER portfolio history has been logged.
function cleanUpStockSummary() {
  const pos = {
    startRow: 3,
    unitsOwnedColIndex: 9,
    startCol: "A",
    endCol: "AC",
    lastRowCheck: "Stock Summary!C:C",
  };

  const lastRow = getLastRowIndex(pos.lastRowCheck);

  // If nothing to delete, do nothing
  if (lastRow < pos.startRow) {
    return;
  }

  const data = stockSummarySheet
    .getRange(
      pos.startCol +
        pos.startRow.toString() +
        ":" +
        pos.endCol +
        lastRow.toString()
    )
    .getValues();

  // Iterate from the back so indexes stay constant for all preceding rows
  let row = lastRow - pos.startRow;

  for (row; row >= 0; row--) {
    // Check units owned column's value. If 0, no units owned for that stock so we delete the row
    if (
      data[row][pos.unitsOwnedColIndex] == 0 ||
      data[row][pos.unitsOwnedColIndex] == "0"
    ) {
      // Add one back to get row number from array index
      let rowNumber = pos.startRow + row;
      stockSummarySheet.deleteRow(rowNumber);
      Utilities.sleep(100);
    }
  }
}

// ==================== Alpha Vantage API Functions (For EOD prices, only if realtime stops working) ==================== //

// Returns stock data from Alpha Vantage and populates the row
function getAVStockData() {
  // Alpha vantage key can be anything
  var avKey = "somerandomkey";

  // Positions to update and arbitrary key token variables
  avPos = {
    startRow: "3",
    eodCol: "AA",
    priceCol: "E",
    endRow: getLastRowIndex("Stock Summary!A:A").toString(),
  };

  // Check which row we need to update now. If last row, we update then reset the count.
  var test;

  // Get a JSON object with stock data from AV API
  var url =
    "https://www.alphavantage.co/query?function=TIME_SERIES_DAILY_ADJUSTED&outputsize=full&symbol=" +
    JNJ +
    "&apikey=" +
    avKey;
  var response = UrlFetchApp.fetch(url);
  var rawData = JSON.parse(response);

  // Now we have a JSON object, parse it to get details
  var stockData = {};

  return stockData[dataType];
}

// ==================== Yahoo Finance Historical Data Functions ==================== //

// Gets stock data from Yahoo
function getYahooSimpleData(symbol, exchange, dataType) {
  // Check cache for existing values
  var cacheKey = symbol.toUpperCase() + exchange.toUpperCase() + "YAHOOSIMPLE"; // e.g. ES3SGXYAHOOSIMPLE for simple data
  var cacheExpiry = 3600; // Number of seconds to expire cached data (default: one hour)

  // If data exists in cache already, we just use it
  if (cache.getObject(cacheKey) != null) {
    // If data is not null we return it. Else we continue and try to get new data.
    if (cache.getObject(cacheKey)[dataType] != null) {
      return cache.getObject(cacheKey)[dataType];
    }
  }

  // No data found, we continue and cache results later

  // Get corresponding Yahoo exchange code
  var yahooExchangeCode = getYahooExchangeCode(exchange);
  if (yahooExchangeCode != "") {
    yahooExchangeCode = "." + yahooExchangeCode;
  }

  var yahooSymbol = encodeURIComponent(symbol + yahooExchangeCode);

  // Create object to store return values
  var stockData = {};

  // Fetch meta data and other fundamental info (price, week52 high low, moving average, name)
  var url =
    "https://query1.finance.yahoo.com/v6/finance/quote/?symbols=" + yahooSymbol;
  var response = UrlFetchApp.fetch(url);
  var rawDataFundamental = JSON.parse(response)["quoteResponse"]["result"][0];

  // Now we have a JSON object, parse it to get details
  // Define stock properties
  var stockName;

  if (rawDataFundamental["longName"] != null) {
    stockName = rawDataFundamental["longName"];
  } else {
    stockName = rawDataFundamental["shortName"];
  }

  // Define rest of the properties
  Object.defineProperty(stockData, "name", {
    value: stockName,
    writable: true,
    enumerable: true,
  }); // Stock name
  Object.defineProperty(stockData, "price", {
    value: rawDataFundamental["regularMarketPrice"],
    writable: true,
    enumerable: true,
  }); // Stock current price
  Object.defineProperty(stockData, "priceChange", {
    value: rawDataFundamental["regularMarketChange"],
    writable: true,
    enumerable: true,
  }); // Change in price vs ystd's closing
  Object.defineProperty(stockData, "percentChange", {
    value: rawDataFundamental["regularMarketChangePercent"],
    writable: true,
    enumerable: true,
  }); // Percent change in stock price vs ystd close
  Object.defineProperty(stockData, "currency", {
    value: rawDataFundamental["currency"],
    writable: true,
    enumerable: true,
  }); // Currency
  Object.defineProperty(stockData, "fiftyTwoWeekLow", {
    value: rawDataFundamental["fiftyTwoWeekLow"],
    writable: true,
    enumerable: true,
  }); // 52 week low
  Object.defineProperty(stockData, "fiftyTwoWeekHigh", {
    value: rawDataFundamental["fiftyTwoWeekHigh"],
    writable: true,
    enumerable: true,
  }); // 52 week high
  Object.defineProperty(stockData, "fiftyDayAverage", {
    value: rawDataFundamental["fiftyDayAverage"],
    writable: true,
    enumerable: true,
  }); // 50 day moving average
  Object.defineProperty(stockData, "twoHundredDayAverage", {
    value: rawDataFundamental["twoHundredDayAverage"],
    writable: true,
    enumerable: true,
  }); // 200 day moving average

  // Check if return values are null
  if (stockData[dataType] == null) {
    throw "Error! No data found.";
  } else {
    // Store results in cache for future retrieval and return the requested value
    cache.putObject(cacheKey, stockData, cacheExpiry);
    return stockData[dataType];
  }
}

// Function to return yahoo time-series data
function getYahooAdvancedData(symbol, exchange, dataType) {
  // Check cache for existing values
  const cacheKey =
    symbol.toUpperCase() + exchange.toUpperCase() + "YAHOOADVANCED"; // e.g. ES3SGXYAHOOADVANCED for advanced data
  const cacheExpiry = 14400; // Number of seconds to expire cached data (default: 4 hours, as data doesn't change much)

  // If other data exists in cache already, we just use it
  if (cache.getObject(cacheKey) != null) {
    // If data is not null we return it. Else we continue and try to get new data.
    if (cache.getObject(cacheKey)[dataType] != null) {
      // If a date is requested, we need to parse the JSON string into a date object before returning
      if (dataType == "lastDividendDate" || dataType == "lastSplitDate") {
        return parseDate(cache.getObject(cacheKey)[dataType]);
      }

      // If not just return cached value
      return cache.getObject(cacheKey)[dataType];
    }
  }

  // No data found, we continue and cache results later

  // Get corresponding Yahoo exchange code
  var yahooExchangeCode = getYahooExchangeCode(exchange);
  if (yahooExchangeCode != "") {
    yahooExchangeCode = "." + yahooExchangeCode;
  }

  var yahooSymbol = encodeURIComponent(symbol + yahooExchangeCode);

  // Create object to store return values
  let stockData = {};

  // Set defaults
  Object.defineProperty(stockData, "annualDividendRate", {
    value: "-",
    writable: true,
    enumerable: true,
  });
  Object.defineProperty(stockData, "lastDividendAmnt", {
    value: "-",
    writable: true,
    enumerable: true,
  });
  Object.defineProperty(stockData, "lastSplitRatio", {
    value: "-",
    writable: true,
    enumerable: true,
  });
  Object.defineProperty(stockData, "lastDividendDate", {
    value: "-",
    writable: true,
    enumerable: true,
  });
  Object.defineProperty(stockData, "lastSplitDate", {
    value: "-",
    writable: true,
    enumerable: true,
  });

  // Fetch time-series data (prices for the past up to one year)
  var lastYear = new Date().getFullYear() - 1;
  var oneYearAgoDate = new Date(lastYear, 0, 1); // Set to start of last year, months start from 0
  var epochOneYearAgo = dateToEpoch(oneYearAgoDate).toString();

  var url =
    "https://query2.finance.yahoo.com/v8/finance/chart/" +
    yahooSymbol +
    "?symbol=" +
    yahooSymbol +
    "&period1=" +
    epochOneYearAgo +
    "&period2=9999999999&interval=1d&includePrePost=true&events=div%2Csplit";
  var response = UrlFetchApp.fetch(url);
  var rawDataTimeSeries = JSON.parse(response)["chart"]["result"][0]; // Results

  // Now we have a JSON object, parse it to get details

  // Get last dividend and split if any
  if (rawDataTimeSeries["events"] != null) {
    if (rawDataTimeSeries["events"]["dividends"] != null) {
      var dividends = rawDataTimeSeries["events"]["dividends"];

      // Get total of last year's annual dividends
      var divSum = 0; // Annual dividend rate
      var lastDividendDate = -9999; // Epoch date of dividend to find latest date
      var lastDividendAmnt = "-";

      // Loop through dividend data and get total of all dividends paid last year, as well as the last dividend
      var i = 0;
      for (i; i < Object.keys(dividends).length; i++) {
        var div = dividends[Object.keys(dividends)[i]];
        var divYear = epochToDate(div["date"]).getFullYear();

        // Find latest dividend
        if (eval(div["date"]) >= lastDividendDate) {
          lastDividendDate = eval(div["date"]);
          lastDividendAmnt = div["amount"];
        }

        // If div is from last year, add it to our annual div yield
        if (divYear == lastYear) {
          divSum += div["amount"];
        }
      }

      Object.defineProperty(stockData, "annualDividendRate", {
        value: divSum,
        writable: true,
        enumerable: true,
      }); // Sum of dividend payouts in a year

      // Change to normal date
      lastDividendDate = epochToDate(lastDividendDate);

      Object.defineProperty(stockData, "lastDividendDate", {
        value: lastDividendDate,
        writable: true,
        enumerable: true,
      });
      Object.defineProperty(stockData, "lastDividendAmnt", {
        value: lastDividendAmnt,
        writable: true,
        enumerable: true,
      }); // Last dividend amount and max date
    }

    // If splits exist, get latest split
    if (rawDataTimeSeries["events"]["splits"] != null) {
      // Loop through all splits to get last stock split
      var lastStockSplitDate = -9999;
      var lastStockSplitRatio;
      var stockSplits = rawDataTimeSeries["events"]["splits"];

      for (var j = 0; j < Object.keys(stockSplits).length; j++) {
        var split = stockSplits[Object.keys(stockSplits)[j]];
        if (eval(split["date"]) >= lastStockSplitDate) {
          lastStockSplitDate = eval(split["date"]);
          lastStockSplitRatio = split["splitRatio"];
        }
      }

      lastStockSplitDate = epochToDate(lastStockSplitDate);

      Object.defineProperty(stockData, "lastSplitDate", {
        value: lastStockSplitDate,
        writable: true,
        enumerable: true,
      });
      Object.defineProperty(stockData, "lastSplitRatio", {
        value: lastStockSplitRatio,
        writable: true,
        enumerable: true,
      }); // Last split ratio fraction and date
    }
  }

  // Store results in cache for future retrieval and return the requested value
  cache.putObject(cacheKey, stockData, cacheExpiry);

  // We do not cache historical data, so we just return it when called, and cache the rest.
  if (dataType == "oneYearClosingPrices") {
    let oneYearClosingPrices =
      rawDataTimeSeries["indicators"]["quote"][0]["close"];
    return oneYearClosingPrices;
  }

  // Else return other data
  return stockData[dataType];
}

// Function to return yahoo time-series data for currency exchange rates or today's rate
function getYahooCurrencyData(baseCurrency, newCurrency, startDate, dataType) {
  // We do not cache currency data as it's too big (historical)
  // Check cache for existing values only for non historical data
  var cacheKey =
    baseCurrency.toUpperCase() + newCurrency.toUpperCase() + "YAHOOCURRENCY"; // e.g. USDSGDYAHOOCURRENCY for currency data
  var cacheExpiry = 21600; // Number of seconds to expire cached data (default: 6 hours max for google cache as data doesn't change much)

  // If other data exists in cache already, we just use it
  if (cache.getObject(cacheKey) != null) {
    // If data is not null we return it. Else we continue and try to get new data.
    if (cache.getObject(cacheKey)[dataType] != null) {
      // If requested data is a date, we parse it from JSON to a date object
      if (dataType == "lastRateDate") {
        return parseDate(cache.getObject(cacheKey)[dataType]);
      }

      // Else just return cached value
      return cache.getObject(cacheKey)[dataType];
    }
  }

  // Get corresponding Yahoo currency code
  var yahooSymbol =
    baseCurrency.toUpperCase() + newCurrency.toUpperCase() + "=X";

  // Create object to store return values
  var stockData = {};

  // Fetch time-series data (currency rates from start date until now)
  var epochStartDate = dateToEpoch(startDate).toString();

  var url =
    "https://query2.finance.yahoo.com/v8/finance/chart/" +
    yahooSymbol +
    "?symbol=" +
    yahooSymbol +
    "&period1=" +
    epochStartDate +
    "&period2=9999999999&interval=1d";
  var response = UrlFetchApp.fetch(url);
  var rawDataTimeSeries = JSON.parse(response)["chart"]["result"][0]; // Results

  // Now we have a JSON object, parse it to get details

  // Get currency rates
  var dates = rawDataTimeSeries["timestamp"];
  var rates = rawDataTimeSeries["indicators"]["quote"][0]["close"];

  Object.defineProperty(stockData, "lastRateDate", {
    value: epochToDate(dates[dates.length - 1]),
    writable: true,
    enumerable: true,
  }); // Latest currency rate date
  Object.defineProperty(stockData, "lastRate", {
    value: rates[rates.length - 1],
    writable: true,
    enumerable: true,
  }); // Lastest currency rate

  Object.defineProperty(stockData, "specificDate", {
    value: startDate,
    writable: true,
    enumerable: true,
  }); // Specific currency rate date
  Object.defineProperty(stockData, "specificRate", {
    value: rates[0],
    writable: true,
    enumerable: true,
  }); // Specific currency rate on specified date

  // Store results in cache for future retrieval and return the requested value
  cache.putObject(cacheKey, stockData, cacheExpiry);

  // We do not cache historical data, so we just return it when called, and cache the rest.
  if (dataType == "rates") {
    return rates;
  } else if (dataType == "dates") {
    for (var i = 0; i < dates.length; i++) {
      dates[i] = epochToDate(dates[i]);
    }
    return dates;
  }

  return stockData[dataType];
}

// Takes in a google exchange code and returns a yahoo exchange code
function getYahooExchangeCode(googleExchangeCode) {
  // Records positions of exchange codes in reference sheet
  var exchangeCodePos = { startRow: "5", endRow: "57", col: "G" };

  // Get values of exchange codes from reference sheet
  var exchangeCodeRange = refSheet.getRange(
    exchangeCodePos.col +
      exchangeCodePos.startRow +
      ":" +
      exchangeCodePos.col +
      exchangeCodePos.endRow
  );
  var exchangeCodeValues = exchangeCodeRange.getValues();

  // Loop through values to find exchange code match, then get the Yahoo exchange code to the right
  var row = 0;
  for (row; row <= exchangeCodePos.endRow - exchangeCodePos.startRow; row++) {
    if (exchangeCodeValues[row] == googleExchangeCode) {
      // Get the corresponding Yahoo code on the right cell
      var yahooExchangeCode = refSheet
        .getRange(
          exchangeCodePos.col +
            (parseInt(exchangeCodePos.startRow) + row).toString()
        )
        .offset(0, 1)
        .getValue();
      return yahooExchangeCode;
    }
  }
}

// ==================== Dividends And Splits (TRIGGER) for Transactions ==================== //
// Functions here will run on trigger to check for splits or dividends with ex-date equal to today, and add any unadded ones to the transactions list.

// Dividend and split transaction updating functions, go through all stocks in stock summary, and find their dividends and splits
function addDividendSplit() {
  // Positions within divsplit ref sheet to find values in
  const pos = {
    startRow: 3,
    startCol: "A",
    lastCol: "E",
    divCheckCol: "F",
    splitCheckCol: "G",
    transactionStart: 3,
    lastRowRange: "Div/Split Ref!C:C",
  };

  const lastDataRow = getLastRowIndex(pos.lastRowRange);

  // If no stocks, dont need to run function as no stocks recorded
  if (lastDataRow < pos.startRow) {
    return;
  }

  // Get data range of stocks to check from div split ref
  const dataRange = divSplitRefSheet.getRange(
    pos.startCol +
      pos.startRow.toString() +
      ":" +
      pos.lastCol +
      lastDataRow.toString()
  );
  const data = dataRange.getValues();

  // Get today's date to check against dividends

  const todayDate = new Date();

  // Get all values and loop through each row (each stock) to check dividends and splits
  // We need to make sure that the ex-dividend date and the split date matches yesterday's date
  // Loop through each row in the data range [][] for columns, 0: exchange index, 1: stock name, 2: stock ticker, 3: pre ex-Div units owned, 4: pre ex-Split units owned
  for (let row = 0; row < data.length; row++) {
    // Get variables from columns in row
    const exchange = data[row][0];
    const stockName = data[row][1];
    const stockSymbol = data[row][2];
    const preExDivUnits = data[row][3];
    const preExSplitUnits = data[row][4];

    // Check if stock has already been checked before
    const divSplitCheck = divSplitChecked(
      stockSymbol,
      exchange,
      preExDivUnits,
      preExSplitUnits
    );
    let divChecked = divSplitCheck[0];
    let splitChecked = divSplitCheck[1];

    // Set "locks" so that we do not need to check again when setting the checked list.
    // If the lock is true, we should NOT add this stock to the checked list again.
    let divCheckLock = divSplitCheck[2];
    let splitCheckLock = divSplitCheck[3];

    // If dividend not checked before, check it
    if (divChecked == false) {
      // Check if this stock has a dividend
      const lastDividendDate = getYahooAdvancedData(
        stockSymbol,
        exchange,
        "lastDividendDate"
      );

      // If dividend exists then get the dividend and update transactions
      if (lastDividendDate != "-" && lastDividendDate != null) {
        // Check if last dividend date is the same as today's date. If it is, dividend happened today (matches to local time)
        if (sameDate(lastDividendDate, todayDate)) {
          // There was a dividend today. We add it to Transactions based on how many units there is in the stock.
          const lastDividendAmnt = getYahooAdvancedData(
            stockSymbol,
            exchange,
            "lastDividendAmnt"
          );

          addNewDivTransactionRow(
            stockSymbol,
            exchange,
            lastDividendDate,
            lastDividendAmnt,
            preExDivUnits
          );

          Utilities.sleep(100);

          // Set div checked to true
          divChecked = true;
        }
      }
    }

    // If split not checked before, check it
    if (splitChecked == false) {
      // Check if this stock has a split
      const lastSplitDate = getYahooAdvancedData(
        stockSymbol,
        exchange,
        "lastSplitDate"
      );

      // If splits exist, check if it happened today
      if (lastSplitDate != "-" && lastSplitDate != null) {
        // Check if last split date is the same as today's date, if yes, then last split occured today, we add it
        if (sameDate(lastSplitDate, todayDate)) {
          // Get split ratio
          var lastSplitRatio = getYahooAdvancedData(
            stockSymbol,
            exchange,
            "lastSplitRatio"
          );

          // Add split to transactions and modify all previous entries to accomodate split.
          addNewSplitTransactionRow(
            stockSymbol,
            exchange,
            lastSplitDate,
            lastSplitRatio,
            preExSplitUnits
          );

          Utilities.sleep(100);

          // Set split checked to true
          splitChecked = true;
        }
      }
    }

    // If div checked, split checked or both, we add this stock to the list of "checked" stocks
    setCheckedDivSplit(
      stockSymbol,
      exchange,
      divChecked,
      splitChecked,
      divCheckLock,
      splitCheckLock
    );
  }
}

// This function checks if the current stock has already been checked for dividends and splits today.
// Returns a four element array of booleans, representing: Dividend checked today, split checked today, dividend check lock, split check lock
// If a lock is true, we prevent the stock from being written to that particular checked list
function divSplitChecked(symbol, exchange, divUnits, splitUnits) {
  const checkPos = {
    startRow: 3,
    divCheckCol: "F",
    splitCheckCol: "G",
    lastRowRange: "Div/Split Ref!F:G",
  };

  // Initialize check variables
  let divChecked = false;
  let splitChecked = false;
  let divCheckLock = false;
  let splitCheckLock = false;

  // If units are 0 then we don't need to waste time checking, so we just say we have checked it already, don't allow writing to check range though as we might need to check again
  if (divUnits <= 0 && splitUnits <= 0) {
    return [true, true, true, true];
  }

  // Else we get last filled data row for the check range
  const lastRow = getLastRowIndex(checkPos.lastRowRange);

  // If no data exists, means all stocks haven't been checked, means we need to check this stock, allow writing to checked range
  if (lastRow < checkPos.startRow) {
    return [false, false, false, false];
  }

  // Else get check values
  const checkVals = divSplitRefSheet
    .getRange(
      checkPos.divCheckCol +
        checkPos.startRow.toString() +
        ":" +
        checkPos.splitCheckCol +
        lastRow.toString()
    )
    .getValues();

  // Check if for this stock, dividends and/or splits have been checked
  for (let row = 0; row < checkVals.length; row++) {
    // Check if dividends checked today
    if (checkVals[row][0] == exchange + ":" + symbol) {
      divChecked = true;
      divCheckLock = true;
    }

    // Check if splits checked today
    if (checkVals[row][1] == exchange + ":" + symbol) {
      splitChecked = true;
      splitCheckLock = true;
    }

    // If both checked already, terminate
    if (divChecked && splitChecked) {
      return [true, true, true, true];
    }
  }

  // After checking finish, return if we should check div/split or not. If we found an existing check, then we don't allow writing. Else allow.
  return [divChecked, splitChecked, divCheckLock, splitCheckLock];
}

// This function sets this stock status to checked for either dividend or split based on parameters given
function setCheckedDivSplit(
  symbol,
  exchange,
  divChecked,
  splitChecked,
  divCheckLock,
  splitCheckLock
) {
  const pos = { startRow: 3, divCheckCol: "F", splitCheckCol: "G" };

  let newLastRow;
  let setRange;

  // If dividend and/or split was checked, then add it to their respective checked lists
  if (divChecked && divCheckLock == false) {
    newLastRow =
      getLastRowIndex(
        "Div/Split Ref!" + pos.divCheckCol + ":" + pos.divCheckCol
      ) + 1;
    setRange = divSplitRefSheet.getRange(
      pos.divCheckCol + newLastRow.toString()
    );
    setRange.setValue(exchange + ":" + symbol);
  }

  if (splitChecked && splitCheckLock == false) {
    newLastRow =
      getLastRowIndex(
        "Div/Split Ref!" + pos.splitCheckCol + ":" + pos.splitCheckCol
      ) + 1;
    setRange = divSplitRefSheet.getRange(
      pos.splitCheckCol + newLastRow.toString()
    );
    setRange.setValue(exchange + ":" + symbol);
  }
}

// This function clears the checked dividends and splits list after every day
function clearCheckedDivSplit() {
  // Positions within divsplit ref sheet to find values in
  const pos = { startRow: 3, divCheckCol: "F", splitCheckCol: "G" };

  // Clear all values in the two columns
  const clearRange = divSplitRefSheet.getRange(
    pos.divCheckCol + pos.startRow.toString() + ":" + pos.splitCheckCol
  );
  clearRange.clearContent();
}

// ==================== TRANSACTION FUNCTIONS ==================== //

// Function for button to add new row at the top of the sheet
function addNewTransactionRow() {
  var firstRow = 3;
  var lastCol = transactionsSheet.getLastColumn();
  var lastRow = transactionsSheet.getLastRow();
  var newLastRow = lastRow + 1;
  var newLastRowString = newLastRow.toString();

  var manualRangeList = transactionsSheet.getRangeList([
    "A" + newLastRowString + ":B" + newLastRowString,
    "D" + newLastRowString + ":I" + newLastRowString,
    "U" + newLastRowString,
  ]); // Range of all manual values that should be reset to null

  // Insert a new row at the top of the transactions sheet
  transactionsSheet.insertRowAfter(lastRow);

  // Copy all formulas from below row over
  Utilities.sleep(200); // Pause to let it detect new row
  transactionsSheet
    .getRange(lastRow, 1, 1, lastCol)
    .copyTo(transactionsSheet.getRange(newLastRow, 1, 1, lastCol), {
      contentsOnly: false,
    });
  manualRangeList.clearContent(); // Reset all manual values so the row is blank

  // Scroll down to new row
  var activeRange = transactionsSheet.getRange(newLastRow, 1);
  transactionsSheet.setActiveRange(activeRange);
}

// Adds a new buy row with an extra cashIn row
function addNewBuyTransactionRow() {
  // Positions of values to fill in for new buy transaction
  const posToFill = { typeCol: "B", startCol: "D", endCol: "I" };

  // Make a new row for CashIn
  addNewTransactionRow();
  Utilities.sleep(100);

  let lastRow = transactionsSheet.getLastRow();

  // Populate the last row (new transaction with the new buy values)
  let typeRange = transactionsSheet.getRange(
    posToFill.typeCol + lastRow.toString()
  );
  let manualRange = transactionsSheet.getRange(
    posToFill.startCol +
      lastRow.toString() +
      ":" +
      posToFill.endCol +
      lastRow.toString()
  );

  // Make a new cash in row and set values for new cash in row
  // Manual range: Stock symbol, Exchange, Transacted Units, Transacted Price (blank for user entry), Fees, Stock Split Ratio
  typeRange.setValue("CashIn");
  manualRange.setValues([["$$$", "-", 1, "", 0, 1]]);

  Utilities.sleep(100);

  // Now make a new row for buy and set values for new buy row
  addNewTransactionRow();

  Utilities.sleep(100);

  lastRow += 1;
  typeRange = transactionsSheet.getRange(
    posToFill.typeCol + lastRow.toString()
  );
  manualRange = transactionsSheet.getRange(
    posToFill.startCol +
      lastRow.toString() +
      ":" +
      posToFill.endCol +
      lastRow.toString()
  );

  // Set values for buy row, blank for user entry
  typeRange.setValue("Buy");
  manualRange.setValues([["", "", "", "", "", 1]]);
}

// Adds a new dividend transaction (for internal script use only)
function addNewDivTransactionRow(
  stockSymbol,
  exchange,
  lastDividendDate,
  lastDividendAmnt,
  preExDivUnits
) {
  // Range within transaction row to fill up with manual values, as well as range in which if split occurs, we have to update for all previous transactions, while checking the symbol matches
  const transactionRow = {
    dateTypeStart: "A",
    dateTypeEnd: "B",
    symbolFeesStart: "D",
    symbolFeesEnd: "I",
  };

  // Add new transaction row for dividend
  addNewTransactionRow();

  Utilities.sleep(100);

  // Get dividend values
  const lastRow = transactionsSheet.getLastRow();

  // Populate the last row (new transaction with the new dividend values)
  // Date and Div transaction type
  const dateTypeValues = [[lastDividendDate, "Div"]];
  const dateTypeRange = transactionsSheet.getRange(
    transactionRow.dateTypeStart +
      lastRow.toString() +
      ":" +
      transactionRow.dateTypeEnd +
      lastRow.toString()
  );
  dateTypeRange.setValues(dateTypeValues);

  Utilities.sleep(100);

  // Stock symbol, stock exchange, dividend units, dividend price per unit, fees = 0, stocksplit = 1
  const symbolFeesValues = [
    [stockSymbol, exchange, preExDivUnits, lastDividendAmnt, 0, 1],
  ];
  const symbolFeesRange = transactionsSheet.getRange(
    transactionRow.symbolFeesStart +
      lastRow.toString() +
      ":" +
      transactionRow.symbolFeesEnd +
      lastRow.toString()
  );
  symbolFeesRange.setValues(symbolFeesValues);
}

// Adds a new split transaction, and modifies all previous price and units owned for this stock (for internal script use only)
function addNewSplitTransactionRow(
  stockSymbol,
  exchange,
  lastSplitDate,
  lastSplitRatio,
  preExSplitUnits
) {
  const transactionPos = {
    startRow: 3,
    dateTypeStart: "A",
    dateTypeEnd: "B",
    symbolFeesStart: "D",
    symbolFeesEnd: "I",
    splitCheckStart: "D",
    splitCheckEnd: "G",
  };

  // Add new transaction row for split
  addNewTransactionRow();

  Utilities.sleep(100);

  const lastRow = transactionsSheet.getLastRow();

  // Populate the last row (new transaction with the new split values)
  // Date and Div transaction type
  const dateTypeValues = [[lastSplitDate, "Split"]];
  const dateTypeRange = transactionsSheet.getRange(
    transactionPos.dateTypeStart +
      lastRow.toString() +
      ":" +
      transactionPos.dateTypeEnd +
      lastRow.toString()
  );
  dateTypeRange.setValues(dateTypeValues);

  Utilities.sleep(100);

  // Stock symbol, stock exchange, original units, price per unit, fees = 0, stocksplit ratio
  const symbolFeesValues = [
    [stockSymbol, exchange, preExSplitUnits, 0, 0, eval(lastSplitRatio)],
  ];
  const symbolFeesRange = transactionsSheet.getRange(
    transactionPos.symbolFeesStart +
      lastRow.toString() +
      ":" +
      transactionPos.symbolFeesEnd +
      lastRow.toString()
  );
  symbolFeesRange.setValues(symbolFeesValues);

  Utilities.sleep(100);

  // Go through all previous entries and change all their stock prices and buy units to be affected by the split
  // Buy/sell price is MULTIPLIED by split ratio (e.g. 1 stock split to 7, multiply by 1/7), and units transacted is DIVIDED by split ratio (200 / (1/7) = 200 * 7)
  const checkRange = transactionsSheet.getRange(
    transactionPos.splitCheckStart +
      transactionPos.startRow.toString() +
      ":" +
      transactionPos.splitCheckEnd +
      (lastRow - 1).toString()
  );
  const checkRangeVals = checkRange.getValues();

  // We now have a 2D array of values. Each row is STOCK SYMBOL, STOCK EXCHANGE, TRANSACTED UNITS, TRANSACTED PRICE
  // We need to check stock symbol AND stock exchange matches, then update the transacted units and price with the split price and units
  for (let row = 0; row < checkRangeVals.length; row++) {
    // If match our current stock split, then we need to edit its transacted units and transacted price (index position 2 and 3 in inner array)
    if (
      checkRangeVals[row][0] == stockSymbol &&
      checkRangeVals[row][1] == exchange
    ) {
      const originalUnits = checkRangeVals[checkRow][2];
      const originalPrice = checkRangeVals[checkRow][3];
      const newUnits = eval(originalUnits) / eval(lastSplitRatio);
      const newPrice = eval(originalPrice) * eval(lastSplitRatio);

      const rowNo = transactionPos.startRow + row;

      // Update values in transaction row
      const rangeToUpdate = transactionsSheet.getRange(
        transactionPos.splitCheckStart +
          rowNo.toString() +
          ":" +
          transactionPos.splitCheckEnd +
          rowNo.toString()
      );
      const newValues = [[stockSymbol, exchange, newUnits, newPrice]];
      rangeToUpdate.setValues(newValues);

      Utilities.sleep(100);
    }
  }
}

// ==================== MISCELLANEOUS FUNCTIONS ==================== //

// Function that extracts the alphabet notation of column, given a range. Returns a string.
function getCol(range) {
  return currSpreadsheet
    .getRange(range)
    .getA1Notation()
    .match(/([A-Z]+)/)[0];
}

// Function that extracts the string notation of row, given a range.
function getRow(range) {
  return currSpreadsheet
    .getRange(range)
    .getA1Notation()
    .match(/[0-9]+/)[0];
}

// Function that returns the last NON-NULL row number in a range index
// Note that this return value is RELATIVE to the range (e.g. A3:A10, if 2 is the return, A4 is last row). 0 = all empty, else returns row number of last data found
function getLastRowIndex(range) {
  var rangeVals = currSpreadsheet.getRange(range).getValues();
  var noOfRows = rangeVals.length;

  // Loop from the bottom of the range. Once we hit a non-null cell, we return that row's relative index.
  for (let row = noOfRows - 1; row >= 0; row--) {
    // Loop through each column in row to see if a value exists
    for (let col = 0; col < rangeVals[row].length; col++) {
      if (rangeVals[row][col] != "") {
        return row + 1; // Add one to get row number from index notation
      }
    }
  }

  // If we reach here, it means the entire range is empty, we return 0.
  return 0;
}

// Function to compare two date objects to check if they are the same day
function sameDate(date1, date2) {
  // Check if either date is not a valid date
  if (
    Object.prototype.toString.call(date1) != "[object Date]" ||
    Object.prototype.toString.call(date2) != "[object Date]"
  ) {
    return false;
  }

  // Check day, month, year to be same, ignores time
  return (
    date1.getDate() == date2.getDate() &&
    date1.getMonth() == date2.getMonth() &&
    date1.getFullYear() == date2.getFullYear()
  );
}

// Takes in a date and returns the epoch time (seconds since 1970-1-1 0:00:00)
function dateToEpoch(date) {
  return Math.floor(date.getTime() / 1000.0);
}

// Takes in an epoch date and returns javascript date object in local timezone
function epochToDate(epochDate) {
  const date = new Date(epochDate * 1000);
  return date;
}

// Parse function for date objects in cache JSON string > Date object
function parseDate(dateStr) {
  const dateFormat = /^\d{4}-(0[1-9]|1[0-2])-([12]\d|0[1-9]|3[01])([T\s](([01]\d|2[0-3])\:[0-5]\d|24\:00)(\:[0-5]\d([\.,]\d+)?)?([zZ]|([\+-])([01]\d|2[0-3])\:?([0-5]\d)?)?)?$/;

  // If string is recognized as an ISO date, parse it into a date object
  if (typeof dateStr === "string" && dateFormat.test(dateStr)) {
    return new Date(dateStr);
  }

  // Else just return the string
  return dateStr;
}

// ==================== ENHANCED CACHE FUNCTIONS (COURTESY OF https://github.com/yinonavraham/) ==================== //

// Wraps an existing cache as an improved cache
function wrap(cache) {
  return new EnhancedCache(cache);
}

/**
 * Enhanced cache - wraps a native Cache object and provides additional features.
 * @param {Cache} cache the cache to enhance
 * @constructor
 */
function EnhancedCache(cache) {
  var cache_ = cache;

  //### PUBLIC Cache methods ###

  /**
   * Put a string value in the cache
   * @param {string} key
   * @param {string} value
   * @param {number} ttl (optional) time-to-live in seconds for the key:value pair in the cache
   */
  this.put = function (key, value, ttl) {
    this.putString(key, value, ttl);
  };

  /**
   * Get a string value from the cache
   * @param {string} key
   * @return {string} The string value, or null if none is found
   */
  this.get = function (key) {
    return this.getString(key);
  };

  /**
   * Removes an entry from the cache using the given key.
   * @param {string} key
   */
  this.remove = function (key) {
    var valueDescriptor = getValueDescriptor(key);
    if (valueDescriptor.keys) {
      for (var i = 0; i < valueDescriptor.keys.length; i++) {
        var k = valueDescriptor.keys[i];
        remove_(k);
      }
    }
    remove_(key);
  };

  //### PUBLIC EnhancedCache methods ###

  /**
   * Put a string value in the cache
   * @param {string} key
   * @param {string} value
   * @param {number} ttl (optional) time-to-live in seconds for the key:value pair in the cache
   */
  this.putString = function (key, value, ttl) {
    var type = "string";
    ensureValueType(value, type);
    putValue(key, value, type, ttl);
  };

  /**
   * Get a string value from the cache
   * @param {string} key
   * @return {string} The string value, or null if none is found
   */
  this.getString = function (key) {
    var value = getValue(key, "string");
    return value;
  };

  /**
   * Put a numeric value in the cache
   * @param {string} key
   * @param {number} value
   * @param {number} ttl (optional) time-to-live in seconds for the key:value pair in the cache
   */
  this.putNumber = function (key, value, ttl) {
    var type = "number";
    ensureValueType(value, type);
    putValue(key, value, type, ttl);
  };

  /**
   * Get a numeric value from the cache
   * @param {string} key
   * @return {number} The numeric value, or null if none is found
   */
  this.getNumber = function (key) {
    var value = getValue(key, "number");
    return value;
  };

  /**
   * Put a boolean value in the cache
   * @param {string} key
   * @param {boolean} value
   * @param {number} ttl (optional) time-to-live in seconds for the key:value pair in the cache
   */
  this.putBoolean = function (key, value, ttl) {
    var type = "boolean";
    ensureValueType(value, type);
    putValue(key, value, type, ttl);
  };

  /**
   * Get a boolean value from the cache
   * @param {string} key
   * @return {boolean} The boolean value, or null if none is found
   */
  this.getBoolean = function (key) {
    var value = getValue(key, "boolean");
    return value;
  };

  /**
   * Put an object in the cache
   * @param {string} key
   * @param {string} value
   * @param {number} ttl (optional) time-to-live in seconds for the key:value pair in the cache
   * @param {function(object)} stringify (optional) function to use for converting the object to string. If not specified, JSON's stringify function is used:
   * <pre>stringify = function(obj) { return JSON.stringify(obj); };</pre>
   */
  this.putObject = function (key, value, ttl, stringify) {
    stringify = stringify || JSON.stringify;
    var type = "object";
    ensureValueType(value, type);
    var sValue = value === null ? null : stringify(value);
    putValue(key, sValue, type, ttl);
  };

  /**
   * Get an object from the cache
   * @param {string} key
   * @param {function(string)} parse (optional) function to use for converting the string to an object. If not specified, JSON's parse function is used:
   * <pre>parse = function(text) { return JSON.parse(text); };</pre>
   * @return {object} The object, or null if none is found
   */
  this.getObject = function (key, parse) {
    parse = parse || JSON.parse;
    var sValue = getValue(key, "object");
    var value = sValue === null ? null : parse(sValue);
    return value;
  };

  /**
   * Get the date an entry was last updated
   * @param {string} key
   * @return {Date} the date the entry was last updated, or null if no such key exists
   */
  this.getLastUpdated = function (key) {
    var valueDescriptor = getValueDescriptor(key);
    return valueDescriptor === null ? null : new Date(valueDescriptor.time);
  };

  // ### PRIVATE ###

  function ensureValueType(value, type) {
    if (value !== null) {
      var actualType = typeof value;
      if (actualType !== type) {
        throw new Error(
          Utilities.formatString(
            "Value type mismatch. Expected: %s, Actual: %s",
            type,
            actualType
          )
        );
      }
    }
  }

  function ensureKeyType(key) {
    if (typeof key !== "string") {
      throw new Error("Key must be a string value");
    }
  }

  function createValueDescriptor(value, type, ttl) {
    return {
      value: value,
      type: type,
      ttl: ttl,
      time: new Date().getTime(),
    };
  }

  function putValue(key, value, type, ttl) {
    ensureKeyType(key);
    var valueDescriptor = createValueDescriptor(value, type, ttl);
    splitLargeValue(key, valueDescriptor);
    var sValueDescriptor = JSON.stringify(valueDescriptor);
    put_(key, sValueDescriptor, ttl);
  }

  function put_(key, value, ttl) {
    if (ttl) {
      cache_.put(key, value, ttl);
    } else {
      cache_.put(key, value);
    }
  }

  function get_(key) {
    return cache_.get(key);
  }

  function remove_(key) {
    return cache_.remove(key);
  }

  function getValueDescriptor(key) {
    ensureKeyType(key);
    var sValueDescriptor = get_(key);
    var valueDescriptor =
      sValueDescriptor === null ? null : JSON.parse(sValueDescriptor);
    return valueDescriptor;
  }

  function getValue(key, type) {
    var valueDescriptor = getValueDescriptor(key);
    if (valueDescriptor === null) {
      return null;
    }
    if (type !== valueDescriptor.type) {
      throw new Error(
        Utilities.formatString(
          "Value type mismatch. Expected: %s, Actual: %s",
          type,
          valueDescriptor.type
        )
      );
    }
    mergeLargeValue(valueDescriptor);
    return valueDescriptor.value;
  }

  function mergeLargeValue(valueDescriptor) {
    //If the value descriptor has 'keys' instead of 'value' - collect the values from the keys and populate the value
    if (valueDescriptor.keys) {
      var value = "";
      for (var i = 0; i < valueDescriptor.keys.length; i++) {
        var k = valueDescriptor.keys[i];
        var v = get_(k);
        value += v;
      }
      valueDescriptor.value = value;
      valueDescriptor.keys = undefined;
    }
  }

  function splitLargeValue(key, valueDescriptor) {
    //Max cached value size: 128KB
    //According the ECMA-262 3rd Edition Specification, each character represents a single 16-bit unit of UTF-16 text
    var DESCRIPTOR_MARGIN = 2000;
    var MAX_STR_LENGTH = (128 * 1024) / 2 - DESCRIPTOR_MARGIN;
    //If the 'value' in the descriptor is a long string - split it and put in different keys, add the 'keys' to the descriptor
    var value = valueDescriptor.value;
    if (
      value !== null &&
      typeof value === "string" &&
      value.length > MAX_STR_LENGTH
    ) {
      Logger.log("Splitting string value of length: " + value.length);
      var keys = [];
      do {
        var k = "$$$" + key + keys.length;
        var v = value.substring(0, MAX_STR_LENGTH);
        value = value.substring(MAX_STR_LENGTH);
        keys.push(k);
        put_(k, v, valueDescriptor.ttl);
      } while (value.length > 0);
      valueDescriptor.value = undefined;
      valueDescriptor.keys = keys;
    }
    //TODO Maintain previous split values when putting new value in an existing key
  }
}
