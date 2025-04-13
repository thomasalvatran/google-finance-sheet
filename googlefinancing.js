function getMinLast3Months(rangeDate, rangeValue) {
  const today = new Date();
  const threeMonthsAgo = new Date(today);
  threeMonthsAgo.setMonth(today.getMonth() - 3);

  let minValue = null;
  let minDate = null;

  for (let i = 0; i < rangeDate.length; i++) {
    if (!rangeDate[i][0] || !rangeValue[i][0]) continue;  // B? qua ô tr?ng
   
    let date = new Date(rangeDate[i][0]);
    let value = parseFloat(rangeValue[i][0]);

    if (isNaN(date) || isNaN(value)) continue;  // B? qua d? li?u không h?p l?

    if (date >= threeMonthsAgo && date <= today) {
      if (minValue === null || value < minValue) {
        minValue = value;
        minDate = date;
      }
    }
  }

  if (minValue === null) return "No Data";

  return "Min: " + minValue + " @ " + Utilities.formatDate(minDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getTSLAMinLast3Months() {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getRangeByName("TSLA_DATA")
    .getValues();

  let minPrice = null;
  let minDate = null;

  data.forEach(row => {
    let date = row[0];
    let price = parseFloat(row[1]);

    if (!date || isNaN(price)) return;

    if (minPrice === null || price < minPrice) {
      minPrice = price;
      minDate = date;
    }
  });

  if (minPrice === null) return "No Data";

  return "TSLA: $" + minPrice + " @ " + Utilities.formatDate(new Date(minDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getTSLAMin3MonthsYahoo(symbol) {
  //const url = "https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?range=3mo&interval=1d";
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?range=3mo&interval=1d`;

  const response = UrlFetchApp.fetch(url);
  const json = JSON.parse(response.getContentText());
 
  const timestamps = json.chart.result[0].timestamp;
  const prices = json.chart.result[0].indicators.adjclose[0].adjclose;
 
  let minPrice = null;
  let minDate = null;
 
  for (let i = 0; i < prices.length; i++) {
    const price = prices[i];
    const date = new Date(timestamps[i] * 1000);
   
    if (price === null) continue;
   
    if (minPrice === null || price < minPrice) {
      minPrice = price;
      minDate = date;
    }
  }

 
  if (minPrice === null) return "No Data";
 
  return " $" + minPrice.toFixed(2) + " @ " + Utilities.formatDate(minDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function GetMinPrice(symbol, range) {
 const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?range=${range}&interval=1d`; 
  const quoteUrl = `https://query1.finance.yahoo.com/v7/finance/quote?symbols=${symbol}`;
 
  try {
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());
   
    if (!json.chart.result) return "Invalid Symbol or No Data";
   
    const timestamps = json.chart.result[0].timestamp;
    const prices = json.chart.result[0].indicators.adjclose[0].adjclose;
   
    let minPrice = null;
    let minDate = null;
   
    for (let i = 0; i < prices.length; i++) {
      const price = prices[i];
      const date = new Date(timestamps[i] * 1000);
     
      if (price === null) continue;
     
      if (minPrice === null || price < minPrice) {
        minPrice = price;
        minDate = date;
      }
    }
   
    // Fetch giá realtime hôm nay
    const quoteResp = UrlFetchApp.fetch(quoteUrl);
    const quoteJson = JSON.parse(quoteResp.getContentText());
    const priceNow = quoteJson.quoteResponse.result[0].regularMarketPrice;
    const now = new Date();
   
    if (priceNow < minPrice) {
      minPrice = priceNow;
      minDate = now;
    }
   
    if (minPrice === null) return "No Data";
   
    return symbol + " th?p nh?t: $" + minPrice.toFixed(2) + " @ " + Utilities.formatDate(minDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
   
  } catch(e) {
    return "Error: " + e.message;
  }
}

function getTSLAMin3MonthsYahoo1(symbol) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?range=3mo&interval=1d`;

  try {
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());

    const result = json.chart.result[0];

    const timestamps = result.timestamp;
    const adjclose = result.indicators.adjclose[0].adjclose;

    // Combine prices and timestamps into one array to avoid mismatches
    const priceData = [];

    for (let i = 0; i < timestamps.length; i++) {
      const price = adjclose[i];
      const time = timestamps[i];

      if (price !== null && time !== null) {
        priceData.push({
          price: price,
          date: new Date(time * 1000)
        });
      }
    }

    if (priceData.length === 0) return "No valid data";

    // Find the real lowest price (not the first one)
    let minEntry = priceData[0];

    for (let i = 1; i < priceData.length; i++) {
      const entry = priceData[i];

      if (entry.price < minEntry.price) {
        minEntry = entry;
      }
    }

    return " $" + minEntry.price.toFixed(2) + " @ " +
      Utilities.formatDate(minEntry.date, Session.getScriptTimeZone(), "yyyy-MM-dd");

  } catch (e) {
    Logger.log("Error: " + e);
    return "Error fetching data";
  }
}

function getTSLAMinLast90Days(symbol) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?range=6mo&interval=1d`;

  try {
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());

    const result = json.chart.result[0];

    const timestamps = result.timestamp;
    const prices = result.indicators.adjclose[0].adjclose;

    const now = new Date();
    const ninetyDaysAgo = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);

    let minPrice = null;
    let minDate = null;

    for (let i = 0; i < timestamps.length; i++) {
      const ts = timestamps[i];
      const price = prices[i];
      if (price === null) continue;

      const date = new Date(ts * 1000);
      if (date < ninetyDaysAgo) continue; // skip older than 90 days

      if (minPrice === null || price < minPrice) {
        minPrice = price;
        minDate = date;
      }
    }

    if (minPrice === null) return "No valid data in last 90 days";

    return " $" + minPrice.toFixed(2) + " @ " +
      Utilities.formatDate(minDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

  } catch (e) {
    Logger.log("Error fetching or parsing data: " + e);
    return "Error fetching data";
  }
}

function getTSLAMin3MonthsYahoo1(symbol) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?range=3mo&interval=1d`;

  try {
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());

    const result = json.chart.result[0];

    const timestamps = result.timestamp;
    const lowPrices = result.indicators.quote[0].low;

    const priceData = [];

    for (let i = 0; i < timestamps.length; i++) {
      const price = lowPrices[i];
      const time = timestamps[i];

      if (price !== null && time !== null) {
        priceData.push({
          price: price,
          date: new Date(time * 1000)
        });
      }
    }

    if (priceData.length === 0) return "No valid data";

    let minEntry = priceData[0];

    for (let i = 1; i < priceData.length; i++) {
      const entry = priceData[i];
      if (entry.price < minEntry.price) {
        minEntry = entry;
      }
    }

    return " $" + minEntry.price.toFixed(2) + " @ " +
      Utilities.formatDate(minEntry.date, Session.getScriptTimeZone(), "yyyy-MM-dd");

  } catch (e) {
    Logger.log("Error: " + e);
    return "Error fetching data";
  }
}

function getTSLAMin3MonthsYahoo2(symbol) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?range=3mo&interval=1d`;

  try {
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());

    const result = json.chart.result[0];

    const timestamps = result.timestamp;
    const adjclose = result.indicators.adjclose[0].adjclose;

    const priceData = [];

    for (let i = 0; i < timestamps.length; i++) {
      const price = adjclose[i];
      const time = timestamps[i];

      if (price !== null && time !== null) {
        priceData.push({
          price: price,
          date: new Date(time * 1000)
        });
      }
    }

    if (priceData.length === 0) return [["No valid data", ""]];

    let minEntry = priceData[0];

    for (let i = 1; i < priceData.length; i++) {
      const entry = priceData[i];
      if (entry.price < minEntry.price) {
        minEntry = entry;
      }
    }

    // Return 2 columns directly: price, date
    return [[
      minEntry.price.toFixed(2),
      Utilities.formatDate(minEntry.date, Session.getScriptTimeZone(), "yyyy-MM-dd")
    ]];

  } catch (e) {
    Logger.log("Error: " + e);
    return [["Error fetching data", ""]];
  }
}
function getTSLAMin3MonthsYahooLow(symbol) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?range=3mo&interval=1d`;

  try {
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());

    const result = json.chart.result[0];

    const timestamps = result.timestamp;
    const lowPrices = result.indicators.quote[0].low;

    const priceData = [];

    for (let i = 0; i < timestamps.length; i++) {
      const price = lowPrices[i];
      const time = timestamps[i];

      if (price !== null && time !== null) {
        priceData.push({
          price: price,
          date: new Date(time * 1000)
        });
      }
    }

    if (priceData.length === 0) return [["No valid data", ""]];

    let minEntry = priceData[0];

    for (let i = 1; i < priceData.length; i++) {
      const entry = priceData[i];
      if (entry.price < minEntry.price) {
        minEntry = entry;
      }
    }

    // Return 2 columns directly: low price, date
    return [[
      minEntry.price.toFixed(2),
      Utilities.formatDate(minEntry.date, Session.getScriptTimeZone(), "yyyy-MM-dd")
    ]];

  } catch (e) {
    Logger.log("Error: " + e);
    return [["Error fetching data", ""]];
  }
}


function getLowestIntraday1Month(symbol) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?range=1mo&interval=5m`;

  try {
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());

    const result = json.chart.result[0];
    const timestamps = result.timestamp;
    const quotes = result.indicators.quote[0];
    const lowPrices = quotes.low;

    if (!timestamps || !lowPrices || timestamps.length !== lowPrices.length) {
      return [["Data mismatch or missing"]];
    }

    let minPrice = null;
    let minTimestamp = null;

    for (let i = 0; i < timestamps.length; i++) {
      const price = lowPrices[i];
      const time = timestamps[i];

      if (price != null) {
        if (minPrice === null || price < minPrice) {
          minPrice = price;
          minTimestamp = time;
        }
      }
    }

    if (minPrice === null) {
      return [["No valid low price"]];
    }

    const date = new Date(minTimestamp * 1000);
    return [[
      minPrice.toFixed(2),
      Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
    ]];

  } catch (e) {
    Logger.log("Error: " + e.toString());
    return [["Fetch error", e.toString()]];
  }
}


function getLowestIntradayLast7Days(symbol) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?range=7d&interval=5m`;

  try {
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());

    const result = json.chart.result?.[0];
    if (!result) return [["No data", ""]];

    const timestamps = result.timestamp;
    const lows = result.indicators.quote[0].low;

    let minPrice = null;
    let minTime = null;

    for (let i = 0; i < lows.length; i++) {
      const price = lows[i];
      if (price != null && (minPrice === null || price < minPrice)) {
        minPrice = price;
        minTime = timestamps[i];
      }
    }

    if (minPrice === null) return [["No valid price", ""]];

    const date = new Date(minTime * 1000);
    return [[
      minPrice.toFixed(2),
      Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
    ]];

  } catch (e) {
    Logger.log("Error: " + e);
    return [["Fetch error", e.toString()]];
  }
}
function getLowestDailyLow1Month(symbol) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?range=1mo&interval=1d`;

  try {
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());
    const result = json.chart.result[0];
    const timestamps = result.timestamp;
    const lowPrices = result.indicators.quote[0].low;

    let minEntry = { price: null, date: null };

    for (let i = 0; i < timestamps.length; i++) {
      const price = lowPrices[i];
      if (price !== null && (minEntry.price === null || price < minEntry.price)) {
        minEntry.price = price;
        minEntry.date = new Date(timestamps[i] * 1000);
      }
    }

    if (minEntry.price === null) return [["No valid data", ""]];
    
    return [[
      minEntry.price.toFixed(2),
      Utilities.formatDate(minEntry.date, Session.getScriptTimeZone(), "yyyy-MM-dd")
    ]];
    
  } catch (e) {
    return [["Error fetching data", ""]];
  }
}