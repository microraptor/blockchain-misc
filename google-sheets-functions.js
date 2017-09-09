/**
 * This file has several functions for crypto market tools which can be used with Google Sheets.
 *
 * In Google Sheets simple click on 'Tools -> Script Editor...' in the menu and save the code there.
 * To get the value of a function simply enter '=functionName(parameter1, parameter2)' in a cell.
 * For example: '=getTicker("GDAX", "last", "ETH-USD")'
 * If the Spreadsheet is in a locale which use a comma as decimal mark you, need to use a semicolon
 * as a parameter seperator. Like this: '=getTicker("GDAX"; "last"; "ETH-USD")'
 *
 * It is very helpful to put a cell with a current date as an extra parameter to the functions
 * like this: '=getTicker("GDAX", "last", "ETH-USD", A1)'. Every time the date is updated all the
 * functions will update themselves as well. A refresh button in the menu fo this is implemented in
 * the 'onOpen' and the 'refreshLastUpdate' functions.
 */

/**
 * Adds menu entries and opening commands.
 */
function onOpen() {
    //refreshLastUpdate();
    
    var entries = [{
      name : 'Refresh',
      functionName : 'refreshLastUpdate'
    },
    {
      name : 'Update Hyperlinks',
      functionName : 'updateHyperlinks'
    }];
    SpreadsheetApp.getActiveSpreadsheet().addMenu('Custom', entries);
  };
  
  
  /**
   * Refreshes the time of last update.
   */
  function refreshLastUpdate() {
    SpreadsheetApp.getActiveSpreadsheet().getRange('B1').setValue(new Date().toTimeString());
  }
  
  /**
   * Simple function to get the last price of ETH/USD at gdax.com.
   * 
   * @customfunction
   */
  function gdaxEthUsd() {
    return parseFloat(JSON.parse(UrlFetchApp.fetch('https://api.gdax.com/products/ETH-USD/ticker').getContentText()).price);
  }
  
  /**
   * Gets the ticker of a specified exchange fpr a specific currency pair.
   * 
   * @param {string} exchange The exchange to get ticker from Options: 'GDAX', 'Poloniex', 'Kraken', 'Bittrex', 'Liqui', 'Coinmarketcap', 'Etherscan'.
   * @param {string} type The type of ticker. EG: 'Last', 'Ask', 'Bid', 'High', 'Low', 'Average', 'Volume', 'QuoteVolume'.
   * @param {string} pair The currency pair (eg: 'ETH-USD' for GDAX, 'BTC_ETH' for Poloniex, ''XXBTZEUR' for Kraken, 'wings_btc' for Liqui, 'BTC-WINGS' for Bittrex, 'ethusd' for Etherscan and 'ETHEREUM' for Coinmarketcap). 
   * @return The current ticker.
   * @customfunction
   */
  function getTicker(exchange, type, pair) {
    switch(exchange.toUpperCase()) {
      
      case 'GDAX':      
        // Gets the ticker of the input currency pair at gdax.com.
        var class;
        switch(type.toUpperCase()) {
          case 'LAST':
          case 'PRICE':
            var key = 'price';
            class = 'ticker';
            break;
          case 'ASK':
          case 'SELL':
            var key = 'ask';
            class = 'ticker';
            break;
          case 'BID':
          case 'BUY':
            var key = 'bid';
            class = 'ticker';
            break;
          case 'HIGH': 
            var key = 'high';
            class = 'stats';
            break;
          case 'LOW':
            var key = 'low';
            class = 'stats';
            break;
          case 'AVG':
          case 'AVERAGE': // not volume weighted
            var key = 'average';
            class = 'stats';
            break;
          case 'VOL':
          case 'VOLUME':
          case 'BASEVOLUME':
            var key = 'volume';
            class = 'ticker';
            break;
          case 'VOL_30D':
          case 'VOLUME_30DAY':
            var key = 'volume_30day';
            class = 'stats';
            break;
          default:
            throw new Error('Type "' + type + '" not found!');
        }
        var apiPath =  'https://api.gdax.com/products/' + pair + '/' + class;
        var response = UrlFetchApp.fetch(apiPath);
        var data = JSON.parse(response.getContentText());
        if(key !== 'average') {
          return parseFloat(data[key]);
        } else {
          // Calculate average
          var high = parseFloat(data['high']);
          var low = parseFloat(data['low']);
          var result = (high + low) / 2;
          return +result.toFixed(8); // Only 8 decimal points
        }
      
      case 'POLONIEX':      
        // Gets the ticker of the input currency pair at poloniex.com.
        switch(type.toUpperCase()) {
          case 'LAST':
          case 'PRICE':
            var key = 'last';
            break;
          case 'ASK':
          case 'SELL':
            var key = 'lowestAsk';
            break;
          case 'BID':
          case 'BUY':
            var key = 'highestBid';
            break;
          case 'HIGH': 
            var key = 'high24hr';
            break;
          case 'LOW':
            var key = 'low24hr';
            break;
          case 'AVG':
          case 'AVERAGE': // not volume weighted
            var key = 'average';
            break;
          case 'VOL':
          case 'VOLUME':
          case 'BASEVOLUME':
            var key = 'baseVolume';
            break;
          case 'QUOTEVOLUME':
          case 'VOLUME_CURRENCY':
            var key = 'quoteVolume';
            break;
          case 'QUOTEVOLUME':
          case 'VOLUME_CURRENCY':
            var key = 'quoteVolume';
            break;
          case 'PERCENTCHANGE':
            var key = 'percentChange';
            break;
          default:
            throw new Error('Type "' + type + '" not found!');
        }
        var apiPath =  'https://poloniex.com/public?command=returnTicker';
        var response = UrlFetchApp.fetch(apiPath);
        var data = JSON.parse(response.getContentText());
        if(key !== 'average') {
          return parseFloat(data[pair][key]);
        } else {
          // Calculate average
          var high = parseFloat(data[pair]['high24hr']);
          var low = parseFloat(data[pair]['low24hr']);
          var result = (high + low) / 2;
          return +result.toFixed(8); // Only 8 decimal points
        }
        
      case 'KRAKEN':      
        // Gets the ticker of the input currency pair at kraken.com.
        switch(type.toUpperCase()) {
          case 'LAST':
          case 'PRICE':
            var key = 'c';
            var index = 0; // current
            break;
          case 'ASK':
          case 'SELL':
            var key = 'a';
            var index = 0; // current
            break;
          case 'BID':
          case 'BUY':
            var key = 'b';
            var index = 0; // current
            break;
          case 'HIGH': 
            var key = 'h';
            var index = 1; // last 24h
            break;
          case 'LOW':
            var key = 'l';
            var index = 1; // last 24h
            break;
          case 'AVG':
          case 'AVERAGE': // volume weighted
            var key = 'p';
            var index = 1; // last 24h
            break;
          case 'VOL':
          case 'VOLUME':
          case 'BASEVOLUME':
            var key = 'v';
            var index = 1; // last 24h
            break;
          default:
            throw new Error('Type "' + type + '" not found!');
        }
        var apiPath =  'https://api.kraken.com/0/public/Ticker?pair=' + pair;
        var response = UrlFetchApp.fetch(apiPath);
        var data = JSON.parse(response.getContentText());
        return parseFloat(data.result[pair][key][index]);
     
      case 'LIQUI':
        // Gets the ticker of the input currency pair at liqui.io over the last 24h.
        switch(type.toUpperCase()) {
          case 'LAST':
          case 'PRICE':
            var key = 'last';
            break;
          case 'ASK':
          case 'SELL':
            var key = 'sell';
            break;
          case 'BID':
          case 'BUY':
            var key = 'buy';
            break;
          case 'HIGH': 
            var key = 'high';
            break;
          case 'LOW':
            var key = 'low';
            break;
          case 'AVG':
          case 'AVERAGE': // not volume weighted
            var key = 'avg';
            break;
          case 'VOL':
          case 'VOLUME':
          case 'BASEVOLUME':
            var key = 'vol';
            break;
          case 'QUOTEVOLUME':
          case 'VOLUME_CURRENCY':
            var key = 'vol_cur';
            break;
          default:
            throw new Error('Type "' + type + '" not found!');
        }      
        var apiPath =  'https://api.liqui.io/api/3/ticker/' + pair;
        var response = UrlFetchApp.fetch(apiPath);
        var data = JSON.parse(response.getContentText());
        return parseFloat(data[pair][key]);
        
      case 'BITTREX':
        // Gets the ticker of the input currency pair at bittrex.com.
        switch(type.toUpperCase()) {
          case 'LAST':
          case 'PRICE':
            var key = 'Last';
            break;
          case 'ASK':
          case 'SELL':
            var key = 'Ask';
            break;
          case 'BID':
          case 'BUY':
            var key = 'Bid';
            break;
          case 'HIGH': 
            var key = 'High';
            break;
          case 'LOW':
            var key = 'Low';
            break;
          case 'AVG':
          case 'AVERAGE': // not volume weighted
            var key = 'Average';
            break;
          case 'VOL':
          case 'VOLUME':
          case 'BASEVOLUME':
            var key = 'BaseVolume';
            break;
          case 'VOL_CUR':
          case 'QUOTEVOLUME':
          case 'VOLUME_CURRENCY':
            var key = 'Volume';
            break;
          default:
            throw new Error('Type "' + type + '" not found!');
        } 
        var apiPath =  'https://bittrex.com/api/v1.1/public/getmarketsummary?market=' + pair;
        var response = UrlFetchApp.fetch(apiPath);
        var data = JSON.parse(response.getContentText());
        if(key !== 'Average') {
          return parseFloat(data['result'][0][key]);
        } else {
          // Calculate average
          var high = parseFloat(data['result'][0]['High']);
          var low = parseFloat(data['result'][0]['Low']);
          var result = (high + low) / 2;
          return +result.toFixed(8); // Only 8 decimal points
        }
        
      case 'ETHERSCAN':      
        // Gets the ticker of the input currency pair at etherscan.io.
        if(type.toUpperCase() != 'LAST') {
            throw new Error('Type "' + type + '" not found!');
        }
        var apiPath =  'https://api.etherscan.io/api?module=stats&action=ethprice';
        var response = UrlFetchApp.fetch(apiPath);
        var data = JSON.parse(response.getContentText());
        return parseFloat(data.result[pair.toLowerCase()]);
        
      case 'COINMARKETCAP':
      case 'CMC':
        // See https://coinmarketcap.com/api/ for types
        var apiPath = 'https://api.coinmarketcap.com/v1/ticker/' + pair; // actually just one currency
        var response = UrlFetchApp.fetch(apiPath);
        var data = JSON.parse(response.getContentText());
        return parseFloat(data[0][type]);
        
      default:
        throw new Error('Exchange "' + exchange + '" not found!');
    }
  }
  
  
  /**
   * Looks up the not volume weighted average price of the currency against USD on the specified day at gdax.com.
   *
   * @param {date} date The date.
   * @param {string} currency The currency (eg: 'Bitcoin').
   * @return The not volume weighted average price of that day.
   * @customfunction
   */
  function getGdaxAverageUsd(date, currency) {
    switch(currency) {
      case 'Bitcoin':
        currency = 'BTC-USD';
        break;
      case 'Ether':
        currency = 'ETH-USD';
        break;
      case 'EtherBitcoin':
        currency = 'ETH-BTC';
        break;
      case 'USD':
        return 1;
        break;
      default:
        return 'Error, invalid currency!';
    }
    
    date.setTime(date.getTime() - date.getTimezoneOffset() * 1000 * 60); // Remove timezone differences
    var dateTimestamp = Math.floor(date.getTime() / 1000);
    var nextDay = new Date();
    nextDay.setTime(date.getTime() + 1000 * 60 * 60 * (24 + 2)); // 26h to make sure with daylight savings
    
    var apiPath =  'https://api.gdax.com/products/' + currency + '/candles?granularity=86400&start=' + date.toISOString().substr(0,10) + '&end=' + nextDay.toISOString().substr(0,10);
    var response = UrlFetchApp.fetch(apiPath);
    var data = JSON.parse(response.getContentText());
    var dayBucket = data[0]; // [ time, low, high, open, close, volume ]
    
    if(dayBucket[0] != dateTimestamp) {
      return 'Error in internal function';
    }
    var result = (parseFloat(dayBucket[3]) + parseFloat(dayBucket[4])) / 2; // (open + close) / 2
    return result;
  }
  
  
  /**
   * Looks up the volume weighted average price of the currency against euro on the specified day at kraken.com.
   *
   * @param {date} date The date.
   * @param {string} currency The currency (eg: 'Bitcoin').
   * @return The volume weighted average price of that day.
   * @customfunction
   */
  function getKrakenAverageEuro(date, currency) {
    switch(currency) {
      case 'Bitcoin':
        currency = 'XXBTZEUR';
        break;
      case 'Ether':
        currency = 'XETHZEUR';
        break;
      case 'EtherBitcoin':
        currency = 'XETHXXBT';
        break;
      case 'Tether':
        // To convert to EUR, you need a cell with the formuala '=GOOGLEFINANCE("CURRENCY:EURUSD")'
        // and make it a named range called 'EURUSD'
        currency = 'USDTZUSD';
        break;
      case 'Euro':
        return 1;
        break;
      default:
        return 'Error, invalid currency!';
    }
    
    date.setTime(date.getTime() - date.getTimezoneOffset() * 1000 * 60); // Remove timezone differences
    var dateTimestamp = Math.floor(date.getTime() / 1000);
    
    var apiPath =  'https://api.kraken.com/0/public/OHLC?pair=' + currency + '&interval=1440';
    var response = UrlFetchApp.fetch(apiPath);
    var data = JSON.parse(response.getContentText());
    data = data.result[currency];
    var dayBucket = data[data.map(function(e) {return e[0];}).indexOf(dateTimestamp)];
    
    if(dayBucket[0] != dateTimestamp) {
      return 'Error in internal function';
    }
    var result = parseFloat(dayBucket[5]);  // VWAP
    
    // If it needs to be converted from USD (see comment above)
    if(currency === 'USDTZUSD') {
      result /= SpreadsheetApp.getActiveSpreadsheet().getRangeByName('EURUSD').getValue();
    }
    return result;
  }
  
  
  /**
   * Makes blockchain explorer links out of IDs in the selected range.
   */
  function updateHyperlinks() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var range = spreadsheet.getActiveRange();
    var seperator = spreadsheet.getSpreadsheetLocale() === 'de_DE' ? ';' : ',';
    
    for(var i = 1; i <= range.getNumRows(); i++) {
      for(var j = 1; j <= range.getNumColumns(); j++) {
        if(!range.getCell(i,j).isBlank()) {
          var currentValue = range.getCell(i,j).getValue();
          if(currentValue.slice(0,2) === '0x') {
            if(currentValue.length === 66) {
              range.getCell(i,j).setFormula('=HYPERLINK("https://etherscan.io/tx/' + currentValue + '"' + seperator + '"ETH-TXID: ' + currentValue + '")');
            } else if(currentValue.length === 42) {
              range.getCell(i,j).setFormula('=HYPERLINK("https://etherscan.io/address/' + currentValue + '"' + seperator + '"ETH-ADDR: ' + currentValue + '")');
            }
          } else if(currentValue.length === 64) {
              range.getCell(i,j).setFormula('=HYPERLINK("https://tradeblock.com/bitcoin/tx/' + currentValue + '"' + seperator + '"BTC-TXID: ' + currentValue + '")');
          }
        }
      }
    }
  }
