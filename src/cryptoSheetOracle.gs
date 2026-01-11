/**
 * Returns the current price of a cryptocurrency in USD.
 * Uses CoinMarketCap API and Script Cache to reduce API calls.
 *
 * @param {string} cryptoSymbol - Cryptocurrency symbol (e.g. BTC, ETH)
 * @return {number|string} Price rounded to 4 decimals or error string
 */
function getCryptoPrice(cryptoSymbol) {
  // Get script cache
  var cache = CacheService.getScriptCache();
  var cachedPrice = cache.get(cryptoSymbol);

  // Return cached value if available
  if (cachedPrice !== null) {
    return parseFloat(cachedPrice);
  }

  // CoinMarketCap API endpoint
  var url = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest";
  var apiKey = CMC_API_KEY; // Stored as Script Property

  var headers = {
    "X-CMC_PRO_API_KEY": apiKey,
    "Accept": "application/json"
  };

  // Request parameters
  var parameters = {
    symbol: cryptoSymbol,
    convert: "USD"
  };

  var queryString = Object.keys(parameters)
    .map(key => key + "=" + encodeURIComponent(parameters[key]))
    .join("&");

  try {
    // API request
    var response = UrlFetchApp.fetch(url + "?" + queryString, { headers });
    var json = JSON.parse(response.getContentText());

    // Validate response and extract price
    if (
      json.data &&
      json.data[cryptoSymbol] &&
      json.data[cryptoSymbol].quote &&
      json.data[cryptoSymbol].quote.USD
    ) {
      var price = json.data[cryptoSymbol].quote.USD.price;
      var roundedPrice = price.toFixed(4);

      // Cache price for 5 minutes
      cache.put(cryptoSymbol, roundedPrice, 300);

      return roundedPrice;
    }

    return "#ERROR!";
  } catch (e) {
    return "#ERROR!";
  }
}

/**
 * Re-applies formulas in cells that currently show errors.
 * Useful when API temporarily fails.
 */
function updateErrorCells() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();

  var values = range.getValues();
  var formulas = range.getFormulas();

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (formulas[i][j] !== "") {
        var cellValue = values[i][j];

        if (
          cellValue === null ||
          (typeof cellValue === "string" && cellValue.startsWith("#"))
        ) {
          try {
            sheet.getRange(i + 1, j + 1).setFormula(formulas[i][j]);
          } catch (e) {
            Logger.log(
              "Failed to update cell [" + (i + 1) + "," + (j + 1) + "]: " +
              e.message
            );
          }
        }
      }
    }
  }
}

/**
 * Returns crypto price using symbol or CoinMarketCap ID.
 *
 * @param {string} cryptoSymbol
 * @param {number} cryptoId
 * @return {string|number}
 */
function getCryptoPriceId(cryptoSymbol, cryptoId) {
  var cache = CacheService.getScriptCache();
  var cacheKey = cryptoId || cryptoSymbol;
  var cachedPrice = cache.get(cacheKey);

  if (cachedPrice !== null) {
    return parseFloat(cachedPrice);
  }

  var url = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest";
  var apiKey = CMC_API_KEY;

  var headers = {
    "X-CMC_PRO_API_KEY": apiKey,
    "Accept": "application/json"
  };

  var parameters = cryptoId
    ? { id: cryptoId, convert: "USD" }
    : { symbol: cryptoSymbol, convert: "USD" };

  var queryString = Object.keys(parameters)
    .map(key => key + "=" + encodeURIComponent(parameters[key]))
    .join("&");

  var response = UrlFetchApp.fetch(url + "?" + queryString, { headers });
  var json = JSON.parse(response.getContentText());

  var dataKey = cryptoId || cryptoSymbol;

  if (json.data && json.data[dataKey]?.quote?.USD) {
    var price = json.data[dataKey].quote.USD.price;
    var roundedPrice = price.toFixed(4);

    cache.put(cacheKey, roundedPrice, 300);
    return roundedPrice;
  }

  return "Error: Data not found";
}

/**
 * Writes full cryptocurrency list (ID, name, symbol) to active sheet.
 */
function writeCryptocurrenciesToSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cryptos = getAllCryptocurrencies();

  sheet.getRange(1, 1, 1, 3).setValues([["ID", "Name", "Symbol"]]);

  cryptos.forEach((crypto, index) => {
    sheet
      .getRange(index + 2, 1, 1, 3)
      .setValues([[crypto.id, crypto.name, crypto.symbol]]);
  });
}

/**
 * Fetches full cryptocurrency list from CoinMarketCap.
 *
 * @return {Array}
 */
function getAllCryptocurrencies() {
  var url = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/map";
  var apiKey = CMC_API_KEY;

  var headers = {
    "X-CMC_PRO_API_KEY": apiKey,
    "Accept": "application/json"
  };

  try {
    var response = UrlFetchApp.fetch(url, { headers });
    var json = JSON.parse(response.getContentText());

    if (json.status.error_code === 0) {
      return json.data.map(crypto => ({
        id: crypto.id,
        name: crypto.name,
        symbol: crypto.symbol
      }));
    }
  } catch (e) {
    Logger.log("API Error: " + e.message);
  }

  return [];
}

/**
 * Updates prices in a predefined range.
 * Column A: crypto symbol
 * Column B: price
 */
function updatePrices() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange("A1:B10");
  var data = range.getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) {
      data[i][1] = getCryptoPrice(data[i][0]);
    }
  }

  range.setValues(data);
  SpreadsheetApp.flush();
}
