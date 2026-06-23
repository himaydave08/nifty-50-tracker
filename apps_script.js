/**
 * Google Apps Script for NIFTY 50 Live Price Tracker
 * Paste this script into your Google Sheet's Apps Script editor (Extensions -> Apps Script).
 * Deploy it as a Web App with access for "Anyone" to receive tracking updates.
 */

var STOCK_LIST_SHEET = "Stock List";

// Fallback Nifty 50 constituents list in case NSE is unreachable
var FALLBACK_STOCKS = [
  ["ADANIENT.NS", "Adani Enterprises Ltd."],
  ["ADANIPORTS.NS", "Adani Ports and Special Economic Zone Ltd."],
  ["APOLLOHOSP.NS", "Apollo Hospitals Enterprise Ltd."],
  ["ASIANPAINT.NS", "Asian Paints Ltd."],
  ["AXISBANK.NS", "Axis Bank Ltd."],
  ["BAJAJ-AUTO.NS", "Bajaj Auto Ltd."],
  ["BAJFINANCE.NS", "Bajaj Finance Ltd."],
  ["BAJAJFINSV.NS", "Bajaj Finserv Ltd."],
  ["BEL.NS", "Bharat Electronics Ltd."],
  ["BPCL.NS", "Bharat Petroleum Corporation Ltd."],
  ["BHARTIARTL.NS", "Bharti Airtel Ltd."],
  ["BRITANNIA.NS", "Britannia Industries Ltd."],
  ["CIPLA.NS", "Cipla Ltd."],
  ["COALINDIA.NS", "Coal India Ltd."],
  ["DIVISLAB.NS", "Divi's Laboratories Ltd."],
  ["DRREDDY.NS", "Dr. Reddy's Laboratories Ltd."],
  ["EICHERMOT.NS", "Eicher Motors Ltd."],
  ["GRASIM.NS", "Grasim Industries Ltd."],
  ["HCLTECH.NS", "HCL Technologies Ltd."],
  ["HDFCBANK.NS", "HDFC Bank Ltd."],
  ["HDFCLIFE.NS", "HDFC Life Insurance Company Ltd."],
  ["HEROMOTOCO.NS", "Hero MotoCorp Ltd."],
  ["HINDALCO.NS", "Hindalco Industries Ltd."],
  ["HINDUNILVR.NS", "Hindustan Unilever Ltd."],
  ["ICICIBANK.NS", "ICICI Bank Ltd."],
  ["ITC.NS", "ITC Ltd."],
  ["INDUSINDBK.NS", "IndusInd Bank Ltd."],
  ["INFY.NS", "Infosys Ltd."],
  ["JSWSTEEL.NS", "JSW Steel Ltd."],
  ["KOTAKBANK.NS", "Kotak Mahindra Bank Ltd."],
  ["LT.NS", "Larsen & Toubro Ltd."],
  ["LTM.NS", "LTM Limited"],
  ["M&M.NS", "Mahindra & Mahindra Ltd."],
  ["MARUTI.NS", "Maruti Suzuki India Ltd."],
  ["NTPC.NS", "NTPC Ltd."],
  ["NESTLEIND.NS", "Nestle India Ltd."],
  ["ONGC.NS", "Oil & Natural Gas Corporation Ltd."],
  ["POWERGRID.NS", "Power Grid Corporation of India Ltd."],
  ["RELIANCE.NS", "Reliance Industries Ltd."],
  ["SBILIFE.NS", "SBI Life Insurance Company Ltd."],
  ["SHRIRAMFIN.NS", "Shriram Finance Ltd."],
  ["SBIN.NS", "State Bank of India"],
  ["SUNPHARMA.NS", "Sun Pharmaceutical Industries Ltd."],
  ["TCS.NS", "Tata Consultancy Services Ltd."],
  ["TATACONSUM.NS", "Tata Consumer Products Ltd."],
  ["TMCV.NS", "Tata Motors Ltd. (CV)"],
  ["TATASTEEL.NS", "Tata Steel Ltd."],
  ["TECHM.NS", "Tech Mahindra Ltd."],
  ["TITAN.NS", "Titan Company Ltd."],
  ["ULTRACEMCO.NS", "UltraTech Cement Ltd."],
  ["WIPRO.NS", "Wipro Ltd."],
  ["TRENT.NS", "Trent Ltd."]
];

/**
 * Handles GET requests from Python.
 * Returns the stock symbols list from the "Stock List" config tab.
 */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Remote diagnostics endpoint to check sheet content
    if (e && e.parameter && e.parameter.sheet) {
      var targetSheet = ss.getSheetByName(e.parameter.sheet);
      if (!targetSheet) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Sheet not found: " + e.parameter.sheet }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      var range = targetSheet.getDataRange();
      var data = range.getValues();
      var formulas = range.getFormulas();
      var combined = data.map(function(row, rIdx) {
        return row.map(function(val, cIdx) {
          var f = formulas[rIdx][cIdx];
          return f ? f : val;
        });
      });
      return ContentService.createTextOutput(JSON.stringify({ status: "success", data: combined }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    var stockListSheet = ss.getSheetByName(STOCK_LIST_SHEET);
    
    // Create config sheet if it doesn't exist or is empty
    if (!stockListSheet) {
      stockListSheet = ss.insertSheet(STOCK_LIST_SHEET);
      initializeStockListSheet(stockListSheet);
    } else if (stockListSheet.getLastRow() <= 1) {
      initializeStockListSheet(stockListSheet);
    } else {
      // Defensive check: Ensure email config cells are filled if sheet already exists
      var emailCell = stockListSheet.getRange("D2");
      if (!emailCell.getValue()) {
        stockListSheet.getRange("D1").setValue("Alert Email").setFontWeight("bold").setBackground("#1F4E79").setFontColor("white").setFontFamily("Segoe UI");
        var email = "";
        try {
          email = Session.getActiveUser().getEmail();
        } catch (err) {}
        emailCell.setValue(email).setFontFamily("Segoe UI");
        stockListSheet.autoResizeColumns(1, 4);
      }
    }
    
    // Read stocks
    var values = stockListSheet.getRange(2, 1, stockListSheet.getLastRow() - 1, 2).getValues();
    var stocks = [];
    for (var j = 0; j < values.length; j++) {
      if (values[j][0]) {
        stocks.push({
          ticker: values[j][0].toString().trim(),
          name: values[j][1] ? values[j][1].toString().trim() : ""
        });
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", stocks: stocks }))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles POST requests from Python to update daily checkpoints.
 */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var dateStr = payload.date; // e.g. "19-Jun-2026"
    var pricesData = payload.data; // Object: ticker -> { "09:15": price, ... }
    
    var cellUpdates = writePricesToSheet(dateStr, pricesData);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", updatedCells: cellUpdates }))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 100% Serverless Cloud Fetcher: Downloads prices directly from Yahoo Finance
 * and updates the Google Sheet daily. Can be scheduled using a time-driven trigger.
 */
function fetchNiftyPrices() {
  try {
    // 1. Check if today is a trading day
    if (!isTradingDay()) {
      Logger.log("Today is not an NSE trading day. Skipping.");
      return;
    }
    
    // 2. Check if current time is within market tracking hours (9:10 AM - 4:00 PM IST)
    var now = new Date();
    var timeStr = Utilities.formatDate(now, "GMT+5:30", "HH:mm");
    if (timeStr < "09:10" || timeStr > "18:00") {
      Logger.log("Outside market hours (" + timeStr + "). Skipping.");
      return;
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var stockListSheet = ss.getSheetByName(STOCK_LIST_SHEET);
    if (!stockListSheet) {
      stockListSheet = ss.insertSheet(STOCK_LIST_SHEET);
      initializeStockListSheet(stockListSheet);
    } else if (stockListSheet.getLastRow() <= 1) {
      initializeStockListSheet(stockListSheet);
    }
    
    // 3. Read tickers from configuration tab
    var lastRow = stockListSheet.getLastRow();
    var tickers = stockListSheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function(row) {
      return row[0].toString().trim();
    }).filter(function(t) {
      return t !== "";
    });
    
    if (tickers.length === 0) {
      Logger.log("No tickers found in Stock List.");
      return;
    }
    
    Logger.log("Fetching bulk prices for " + tickers.length + " tickers...");
    
    // 4. Fetch spark data in batches of 20 (Yahoo Finance API limit)
    var batchSize = 20;
    var result = {};
    
    for (var bIdx = 0; bIdx < tickers.length; bIdx += batchSize) {
      var batchTickers = tickers.slice(bIdx, bIdx + batchSize);
      var url = "https://query1.finance.yahoo.com/v8/finance/spark?symbols=" + batchTickers.map(encodeURIComponent).join(",") + "&range=1d&interval=15m";
      
      var response = UrlFetchApp.fetch(url, {
        "headers": {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        },
        "muteHttpExceptions": true
      });
      
      if (response.getResponseCode() !== 200) {
        Logger.log("Yahoo Finance API HTTP Error for batch: " + response.getResponseCode() + " - " + response.getContentText());
        return;
      }
      
      var batchResult = JSON.parse(response.getContentText());
      // Merge batchResult into the final result object
      for (var key in batchResult) {
        result[key] = batchResult[key];
      }
    }
    
    // 5. Parse Yahoo Finance spark data into standard checkpoints format
    var pricesData = {};
    for (var k = 0; k < tickers.length; k++) {
      var symbol = tickers[k];
      var stockObj = result[symbol];
      if (!stockObj || !stockObj.timestamp || !stockObj.close) {
        continue;
      }
      
      var prevClose = stockObj.previousClose;
      var timestamps = stockObj.timestamp;
      var closes = stockObj.close;
      
      // Map Unix timestamps to GMT+5:30 time strings (HH:mm)
      var priceMap = {};
      for (var i = 0; i < timestamps.length; i++) {
        var candleTimeStr = Utilities.formatDate(new Date(timestamps[i] * 1000), "GMT+5:30", "HH:mm");
        priceMap[candleTimeStr] = closes[i];
      }
      
      var stockPrices = {};
      var lastKnownPrice = null;
      
      // 1. Initial 09:15 price using yesterday's close as baseline, or first price of the day
      if (prevClose !== undefined && prevClose !== null) {
        lastKnownPrice = prevClose;
      } else if (priceMap["09:15"] !== undefined && priceMap["09:15"] !== null) {
        lastKnownPrice = priceMap["09:15"];
      }
      
      // If 09:15 is still missing, find the first non-null close price of the day
      if (lastKnownPrice === null && timestamps.length > 0) {
        for (var idx = 0; idx < closes.length; idx++) {
          if (closes[idx] !== null && closes[idx] !== undefined) {
            lastKnownPrice = closes[idx];
            break;
          }
        }
      }
      
      if (lastKnownPrice !== null) {
        stockPrices["09:15"] = lastKnownPrice;
      }
      
      // 2. Match 15m candle close prices to checkpoint targets with forward-fill fallback
      var checkpointsConfig = [
        { name: "09:45", time: "09:30" },
        { name: "10:15", time: "10:00" },
        { name: "10:45", time: "10:30" },
        { name: "11:15", time: "11:00" },
        { name: "11:45", time: "11:30" },
        { name: "12:15", time: "12:00" },
        { name: "12:45", time: "12:30" },
        { name: "13:15", time: "13:00" },
        { name: "13:45", time: "13:30" },
        { name: "14:15", time: "14:00" },
        { name: "14:45", time: "14:30" },
        { name: "15:15", time: "15:00" },
        { name: "15:30", time: "15:15" }
      ];
      
      checkpointsConfig.forEach(function(cfg) {
        var p = priceMap[cfg.time];
        if (p !== undefined && p !== null) {
          lastKnownPrice = p;
          stockPrices[cfg.name] = p;
        } else if (lastKnownPrice !== null) {
          // Forward-fill: Use the last known price to prevent empty cell gaps
          stockPrices[cfg.name] = lastKnownPrice;
        }
      });
      
      pricesData[symbol] = stockPrices;
    }
    
    // 6. Write parsed prices to sheet and trigger alerts
    var dateStr = Utilities.formatDate(now, "GMT+5:30", "dd-MMM-yyyy");
    var cellUpdates = writePricesToSheet(dateStr, pricesData);
    Logger.log("Prices successfully updated. Total cells changed: " + cellUpdates);
    
  } catch (err) {
    Logger.log("Error in fetchNiftyPrices: " + err.toString());
  }
}

/**
 * Checks if today is an NSE trading day.
 */
function isTradingDay() {
  var today = new Date();
  var dateStr = Utilities.formatDate(today, "GMT+5:30", "yyyy-MM-dd");
  var dayOfWeek = Utilities.formatDate(today, "GMT+5:30", "E"); // "Mon", "Tue", etc.
  
  if (dayOfWeek === "Sat" || dayOfWeek === "Sun") {
    return false;
  }
  
  var nseHolidays = [
    "2026-01-26", "2026-03-03", "2026-03-26", "2026-03-31", "2026-04-03",
    "2026-04-14", "2026-05-01", "2026-05-28", "2026-06-26", "2026-09-14",
    "2026-10-02", "2026-10-20", "2026-11-10", "2026-11-24", "2026-12-25"
  ];
  
  if (nseHolidays.indexOf(dateStr) !== -1) {
    return false;
  }
  return true;
}

/**
 * Shared function to write prices, format styles, and trigger email alerts.
 */
function writePricesToSheet(dateStr, pricesData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(dateStr);
  var isNew = false;
  
  if (!sheet) {
    sheet = ss.insertSheet(dateStr);
    isNew = true;
  }
  
  var stockListSheet = ss.getSheetByName(STOCK_LIST_SHEET);
  if (!stockListSheet) {
    stockListSheet = ss.insertSheet(STOCK_LIST_SHEET);
    initializeStockListSheet(stockListSheet);
  } else if (stockListSheet.getLastRow() <= 1) {
    initializeStockListSheet(stockListSheet);
  }
  var lastRow = stockListSheet.getLastRow();
  var masterStocks = stockListSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var checkpoints = ["09:15", "09:45", "10:15", "10:45", "11:15", "11:45", "12:15", "12:45", "13:15", "13:45", "14:15", "14:45", "15:15", "15:30"];
  
  var existingRows = sheet.getLastRow();
  if (isNew || existingRows <= 1) {
    if (!isNew) {
      sheet.clear();
      sheet.clearConditionalFormatRules();
    }
    sheet.setHiddenGridlines(false); // ensure gridlines are visible
    var headers = ["Symbol", "Company Name"].concat(checkpoints).concat(["Greens", "Reds", "Verdict"]);
    sheet.appendRow(headers);
    
    var numStocks = masterStocks.length;
    for (var i = 0; i < numStocks; i++) {
      var rowNum = 2 + i;
      var ticker = masterStocks[i][0].toString().trim();
      var name = masterStocks[i][1] ? masterStocks[i][1].toString().trim() : "";
      sheet.getRange(rowNum, 1).setValue(ticker);
      sheet.getRange(rowNum, 2).setValue(name);
      
      // Greens: Col 17 (Q) - counts when a checkpoint is greater than previous and not empty
      sheet.getRange(rowNum, 17).setValue("=IF(C" + rowNum + "=\"\",\"\",SUMPRODUCT((D" + rowNum + ":P" + rowNum + ">C" + rowNum + ":O" + rowNum + ")*(D" + rowNum + ":P" + rowNum + "<>\"\")))");
      
      // Reds: Col 18 (R) - counts when a checkpoint is less than previous and not empty
      sheet.getRange(rowNum, 18).setValue("=IF(C" + rowNum + "=\"\",\"\",SUMPRODUCT((D" + rowNum + ":P" + rowNum + "<C" + rowNum + ":O" + rowNum + ")*(D" + rowNum + ":P" + rowNum + "<>\"\")))");
      
      // Verdict: Col 19 (S) - Hold if Greens > Reds, else Review
      sheet.getRange(rowNum, 19).setValue("=IF(C" + rowNum + "=\"\",\"No Data\",IF(Q" + rowNum + ">R" + rowNum + ",\"Hold\",\"Review\"))");
    }
    
    var tableRange = sheet.getRange(1, 1, numStocks + 1, 19);
    tableRange.setFontFamily("Segoe UI").setFontSize(11);
    
    var headerRange = sheet.getRange(1, 1, 1, 19);
    headerRange.setBackground("#1F4E79").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    
    var dataRange = sheet.getRange(2, 1, numStocks, 19);
    dataRange.setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(2, 1, numStocks, 2).setHorizontalAlignment("left");
    sheet.getRange(2, 3, numStocks, 14).setHorizontalAlignment("center").setNumberFormat("#,##0.00");
    sheet.getRange(2, 17, numStocks, 2).setHorizontalAlignment("center").setNumberFormat("0");
    sheet.getRange(2, 19, numStocks, 1).setHorizontalAlignment("center").setFontWeight("bold");
    
    for (var r = 2; r <= numStocks + 1; r++) {
      if (r % 2 === 0) {
        sheet.getRange(r, 1, 1, 19).setBackground("#F2F4F7");
      }
    }
    
    var sumStartRow = numStocks + 4;
    sheet.getRange(sumStartRow, 1, 1, 4).merge().setValue("DAILY VERDICT SUMMARY")
         .setBackground("#1F4E79").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    
    sheet.getRange(sumStartRow + 1, 1).setValue("Hold Signals").setFontWeight("bold").setFontColor("#006100").setBackground("#E2EFDA");
    sheet.getRange(sumStartRow + 1, 2).setValue("=COUNTIF(S2:S" + (numStocks + 1) + ', "Hold")')
         .setFontWeight("bold").setFontColor("#006100").setBackground("#E2EFDA").setHorizontalAlignment("center");
         
    sheet.getRange(sumStartRow + 2, 1).setValue("Review Signals").setFontWeight("bold").setFontColor("#9C0006").setBackground("#FCE4D6");
    sheet.getRange(sumStartRow + 2, 2).setValue("=COUNTIF(S2:S" + (numStocks + 1) + ', "Review")')
         .setFontWeight("bold").setFontColor("#9C0006").setBackground("#FCE4D6").setHorizontalAlignment("center");
         
    sheet.getRange(sumStartRow + 3, 1).setValue("Total Checked").setFontWeight("bold").setFontColor("#595959").setBackground("#F2F2F2");
    sheet.getRange(sumStartRow + 3, 2).setValue("=COUNTA(A2:A" + (numStocks + 1) + ")")
         .setFontWeight("bold").setFontColor("#595959").setBackground("#F2F2F2").setHorizontalAlignment("center");
         
    var summaryBlockRange = sheet.getRange(sumStartRow, 1, 4, 4);
    summaryBlockRange.setFontFamily("Segoe UI").setFontSize(11).setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
    
    var rules = sheet.getConditionalFormatRules();
    var holdRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Hold")
      .setBackground("#C6EFCE")
      .setFontColor("#006100")
      .setBold(true)
      .setRanges([sheet.getRange("S2:S" + (numStocks + 1))])
      .build();
      
    var reviewRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Review")
      .setBackground("#FFC7CE")
      .setFontColor("#9C0006")
      .setBold(true)
      .setRanges([sheet.getRange("S2:S" + (numStocks + 1))])
      .build();
      
    var greenRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=D2>C2")
      .setBackground("#C6EFCE")
      .setFontColor("#006100")
      .setRanges([sheet.getRange("D2:P" + (numStocks + 1))])
      .build();
      
    var redRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(D2<>"",D2<C2)')
      .setBackground("#FFC7CE")
      .setFontColor("#9C0006")
      .setRanges([sheet.getRange("D2:P" + (numStocks + 1))])
      .build();
      
    var greyRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(D2<>"",D2=C2)')
      .setBackground("#EAEAEA")
      .setFontColor("#595959")
      .setRanges([sheet.getRange("D2:P" + (numStocks + 1))])
      .build();
      
    rules.push(holdRule, reviewRule, greenRule, redRule, greyRule);
    sheet.setConditionalFormatRules(rules);
  }
  
  var alertEmail = stockListSheet.getRange("D2").getValue();
  if (!alertEmail) {
    alertEmail = Session.getActiveUser().getEmail();
    stockListSheet.getRange("D1").setValue("Alert Email").setFontWeight("bold").setBackground("#1F4E79").setFontColor("white").setFontFamily("Segoe UI");
    stockListSheet.getRange("D2").setValue(alertEmail).setFontFamily("Segoe UI");
  }
  
  var cellUpdates = 0;
  var lastRow = sheet.getLastRow();
  var tickerRange = sheet.getRange(2, 1, lastRow - 1, 1);
  var tickerValues = tickerRange.getValues();
  
  var prevVerdicts = {};
  for (var k = 0; k < tickerValues.length; k++) {
    var symbol = tickerValues[k][0].toString().trim();
    var r = 2 + k;
    prevVerdicts[symbol] = sheet.getRange(r, 19).getValue();
  }
  
  var isCloseCheckpoint = false;
  
  for (var k = 0; k < tickerValues.length; k++) {
    var symbol = tickerValues[k][0].toString().trim();
    var stockData = pricesData[symbol];
    if (stockData) {
      var r = 2 + k;
      for (var cIdx = 0; cIdx < checkpoints.length; cIdx++) {
        var chkName = checkpoints[cIdx];
        var price = stockData[chkName];
        if (price !== undefined && price !== null) {
          var cellCol = 3 + cIdx;
          sheet.getRange(r, cellCol).setValue(price);
          cellUpdates++;
          if (chkName === "15:30") {
            isCloseCheckpoint = true;
          }
        }
      }
    }
  }
  
  SpreadsheetApp.flush();
  
  var newlyFlagged = [];
  var allReviewStocks = [];
  for (var k = 0; k < tickerValues.length; k++) {
    var symbol = tickerValues[k][0].toString().trim();
    var r = 2 + k;
    var newVerdict = sheet.getRange(r, 19).getValue();
    var prevVerdict = prevVerdicts[symbol];
    
    if (newVerdict === "Review") {
      allReviewStocks.push(symbol);
      if (prevVerdict !== "Review" && prevVerdict !== "No Data" && prevVerdict !== "") {
        newlyFlagged.push(symbol);
      }
    }
  }
  
  if (newlyFlagged.length > 0) {
    var sheetUrl = ss.getUrl();
    var subject = "🔴 NIFTY 50 Tracker Alert: " + newlyFlagged.length + " stock(s) flagged for Review";
    var body = "The following NIFTY 50 stock(s) have turned RED (marked for Review) at the latest checkpoint on " + dateStr + ":\n\n" +
               newlyFlagged.map(function(s) { return " - " + s; }).join("\n") + "\n\n" +
               "Spreadsheet Link: " + sheetUrl + "\n\n" +
               "Best regards,\nNifty 50 Price Tracker";
    sendEmailAlert(alertEmail, subject, body);
  }
  
  if (isCloseCheckpoint && cellUpdates > 0) {
    var sheetUrl = ss.getUrl();
    var numStocks = tickerValues.length;
    var numReview = allReviewStocks.length;
    var numHold = numStocks - numReview;
    
    var subject = "📊 NIFTY 50 Tracker: End-of-Day Report (" + dateStr + ")";
    var body = "The trading day for " + dateStr + " is complete.\n\n" +
               "SUMMARY STATS:\n" +
               " - Total Stocks Tracked: " + numStocks + "\n" +
               " - Hold Signals: " + numHold + "\n" +
               " - Review Signals: " + numReview + "\n\n" +
               (numReview > 0 ? "STOCKS FLAGGED FOR REVIEW:\n" + allReviewStocks.map(function(s) { return " - " + s; }).join("\n") : "All stocks remained GREEN/Neutral. Excellent day!") + "\n\n" +
               "Spreadsheet Link: " + sheetUrl + "\n\n" +
               "Best regards,\nNifty 50 Price Tracker";
    sendEmailAlert(alertEmail, subject, body);
  }
  
  sheet.autoResizeColumns(1, 19);
  return cellUpdates;
}

/**
 * Helper to convert column index to column letter (1 -> A, 2 -> B, etc.)
 */
function getColumnLetter(colIdx) {
  var temp, letter = "";
  while (colIdx > 0) {
    temp = (colIdx - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colIdx = (colIdx - temp - 1) / 26;
  }
  return letter;
}

/**
 * Sends a Gmail alert message.
 */
function sendEmailAlert(email, subject, body) {
  if (email && email.indexOf("@") !== -1) {
    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body
      });
      Logger.log("Email sent successfully to: " + email);
    } catch (e) {
      Logger.log("Error sending email: " + e.toString());
    }
  }
}

/**
 * Self-healing helper to initialize the Stock List configuration sheet
 */
function initializeStockListSheet(stockListSheet) {
  stockListSheet.clear();
  stockListSheet.appendRow(["Symbol", "Company Name"]);
  
  // Style headers
  var headerRange = stockListSheet.getRange(1, 1, 1, 2);
  headerRange.setBackground("#1F4E79").setFontColor("white").setFontWeight("bold").setFontFamily("Segoe UI").setHorizontalAlignment("left");
  
  // Write fallback stocks
  for (var i = 0; i < FALLBACK_STOCKS.length; i++) {
    stockListSheet.appendRow(FALLBACK_STOCKS[i]);
  }
  
  // Add Alert Email Config
  stockListSheet.getRange("D1").setValue("Alert Email").setFontWeight("bold").setBackground("#1F4E79").setFontColor("white").setFontFamily("Segoe UI");
  var email = "";
  try {
    email = Session.getActiveUser().getEmail();
  } catch (e) {
    email = "";
  }
  stockListSheet.getRange("D2").setValue(email).setFontFamily("Segoe UI");
  
  // Format borders
  var dataRange = stockListSheet.getRange(2, 1, FALLBACK_STOCKS.length, 2);
  dataRange.setFontFamily("Segoe UI").setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
  
  // Apply zebra striping
  for (var r = 2; r <= FALLBACK_STOCKS.length + 1; r++) {
    if (r % 2 === 0) {
      stockListSheet.getRange(r, 1, 1, 2).setBackground("#F2F4F7");
    }
  }
  stockListSheet.autoResizeColumns(1, 4);
}

/**
 * Diagnostic function to test Yahoo Finance API response in Apps Script
 */
function testYahooApi() {
  var url = "https://query1.finance.yahoo.com/v8/finance/spark?symbols=LTM.NS,M%26M.NS,MARUTI.NS&range=1d&interval=15m";
  var response = UrlFetchApp.fetch(url, {
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    },
    "muteHttpExceptions": true
  });
  Logger.log("Response Code: " + response.getResponseCode());
  try {
    var keys = Object.keys(JSON.parse(response.getContentText()));
    Logger.log("Keys returned: " + keys.join(", "));
  } catch (err) {
    Logger.log("Error parsing JSON: " + err.toString());
    Logger.log("Response body: " + response.getContentText());
  }
}
