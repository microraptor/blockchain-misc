/**
 * This file has several functions as crypto market tools which can be used with Google Sheets.
 *
 * In Google Sheets click on 'Tools -> Script Editor...' in the menu and save the code there.
 * To get the value of a function simply enter '=functionName(parameter1, parameter2)' in a cell.
 * For example: '=getTicker("GDAX", "last", "ETH-USD")'.
 * If the Spreadsheet is in a locale which uses a comma as decimal mark you, need to use a semicolon
 * as a parameter seperator. Like this: '=getTicker("GDAX"; "last"; "ETH-USD")'.
 *
 * It is very helpful to put a cell with a current date as an extra parameter to the functions
 * like this: '=getTicker("GDAX", "last", "ETH-USD", A1)'. Every time the date is updated all the
 * functions will update themselves as well. A refresh button in the menu for this is implemented with
 * the 'onOpen' and the 'refreshLastUpdate' functions.
 */

/* globals SpreadsheetApp, UrlFetchApp */
/* exported refreshLastUpdate, onOpen, gdaxEthUsd, getTicker, getGdaxAverageUsd, getKrakenAverageEuro, updateHyperlinks */

/**
 * Refreshes the time of last update.
 */
function refreshLastUpdate() {
    SpreadsheetApp.getActiveSpreadsheet().getRange('B1').setValue(new Date().toTimeString());
}


/**
 * Adds menu entries and opening commands.
 */
function onOpen() {
    refreshLastUpdate();

    var entries = [{
            name: 'Refresh',
            functionName: 'refreshLastUpdate'
        },
        {
            name: 'Update Hyperlinks',
            functionName: 'updateHyperlinks'
        }
    ];
    SpreadsheetApp.getActiveSpreadsheet().addMenu('Custom', entries);
}


/**
 * Simple function to get the last price of ETH/USD at gdax.com.
 *
 * @customfunction
 */
function gdaxEthUsd() {
    return parseFloat(JSON.parse(UrlFetchApp.fetch('https://api.gdax.com/products/ETH-USD/ticker').getContentText())
        .price);
}

/**
 * Not so simple function to get the ticker of a specified exchange for a specific currency pair.
 *
 * @param {string} exchange The exchange to get ticker from. Options: 'GDAX', 'Poloniex', 'Kraken', 'Bittrex', 'Liqui', 'Coinmarketcap', 'Etherscan'.
 * @param {string} type The type of ticker. For example: 'Last', 'Ask', 'Bid', 'High', 'Low', 'Average', 'Volume', 'QuoteVolume'.
 * @param {string} pair The currency pair. For example: 'ETH-USD' for GDAX, 'BTC_ETH' for Poloniex, ''XXBTZEUR' for Kraken, 'wings_btc' for Liqui, 'BTC-WINGS' for Bittrex, 'ethusd' for Etherscan and 'ETHEREUM' for Coinmarketcap.
 * @return The current ticker.
 * @customfunction
 */
function getTicker(exchange, type, pair) {
    var apiPath;
    var response;
    var data;
    var high;
    var low;
    var result;

    switch (exchange.toUpperCase()) {

        case 'GDAX':
            // Gets the ticker of the input currency pair at gdax.com.
            var apiclass;
            switch (type.toUpperCase()) {
                case 'LAST':
                case 'PRICE':
                    type = 'price';
                    apiclass = 'ticker';
                    break;
                case 'ASK':
                case 'SELL':
                    type = 'ask';
                    apiclass = 'ticker';
                    break;
                case 'BID':
                case 'BUY':
                    type = 'bid';
                    apiclass = 'ticker';
                    break;
                case 'HIGH':
                    type = 'high';
                    apiclass = 'stats';
                    break;
                case 'LOW':
                    type = 'low';
                    apiclass = 'stats';
                    break;
                case 'AVG':
                case 'AVERAGE': // not volume weighted
                    type = 'average';
                    apiclass = 'stats';
                    break;
                case 'VOL':
                case 'VOLUME':
                case 'BASEVOLUME':
                    type = 'volume';
                    apiclass = 'ticker';
                    break;
                case 'VOL_30D':
                case 'VOLUME_30DAY':
                    type = 'volume_30day';
                    apiclass = 'stats';
                    break;
                default:
                    throw new Error('Type "' + type + '" not found!');
            }
            apiPath = 'https://api.gdax.com/products/' + pair + '/' + apiclass;
            response = UrlFetchApp.fetch(apiPath);
            data = JSON.parse(response.getContentText());
            if (type !== 'average') {
                return parseFloat(data[type]);
            } else {
                // Calculate average
                high = parseFloat(data.high);
                low = parseFloat(data.low);
                result = (high + low) / 2;
                return +result.toFixed(8); // Only 8 decimal points
            }
            break;

        case 'POLONIEX':
            // Gets the ticker of the input currency pair at poloniex.com.
            switch (type.toUpperCase()) {
                case 'LAST':
                case 'PRICE':
                    type = 'last';
                    break;
                case 'ASK':
                case 'SELL':
                    type = 'lowestAsk';
                    break;
                case 'BID':
                case 'BUY':
                    type = 'highestBid';
                    break;
                case 'HIGH':
                    type = 'high24hr';
                    break;
                case 'LOW':
                    type = 'low24hr';
                    break;
                case 'AVG':
                case 'AVERAGE': // not volume weighted
                    type = 'average';
                    break;
                case 'VOL':
                case 'VOLUME':
                case 'BASEVOLUME':
                    type = 'baseVolume';
                    break;
                case 'QUOTEVOLUME':
                case 'VOLUME_CURRENCY':
                    type = 'quoteVolume';
                    break;
                case 'QUOTEVOLUME':
                case 'VOLUME_CURRENCY':
                    type = 'quoteVolume';
                    break;
                case 'PERCENTCHANGE':
                    type = 'percentChange';
                    break;
                default:
                    throw new Error('Type "' + type + '" not found!');
            }
            apiPath = 'https://poloniex.com/public?command=returnTicker';
            response = UrlFetchApp.fetch(apiPath);
            data = JSON.parse(response.getContentText());
            if (type !== 'average') {
                return parseFloat(data[pair][type]);
            } else {
                // Calculate average
                high = parseFloat(data[pair].high24hr);
                low = parseFloat(data[pair].low24hr);
                result = (high + low) / 2;
                return +result.toFixed(8); // Only 8 decimal points
            }
            break;

        case 'KRAKEN':
            // Gets the ticker of the input currency pair at kraken.com.
            var index = -1;
            switch (type.toUpperCase()) {
                case 'LAST':
                case 'PRICE':
                    type = 'c';
                    index = 0; // current
                    break;
                case 'ASK':
                case 'SELL':
                    type = 'a';
                    index = 0; // current
                    break;
                case 'BID':
                case 'BUY':
                    type = 'b';
                    index = 0; // current
                    break;
                case 'HIGH':
                    type = 'h';
                    index = 1; // last 24h
                    break;
                case 'LOW':
                    type = 'l';
                    index = 1; // last 24h
                    break;
                case 'AVG':
                case 'AVERAGE': // volume weighted
                    type = 'p';
                    index = 1; // last 24h
                    break;
                case 'VOL':
                case 'VOLUME':
                case 'BASEVOLUME':
                    type = 'v';
                    index = 1; // last 24h
                    break;
                default:
                    throw new Error('Type "' + type + '" not found!');
            }
            apiPath = 'https://api.kraken.com/0/public/Ticker?pair=' + pair;
            response = UrlFetchApp.fetch(apiPath);
            data = JSON.parse(response.getContentText());
            return parseFloat(data.result[pair][type][index]);

        case 'LIQUI':
            // Gets the ticker of the input currency pair at liqui.io over the last 24h.
            switch (type.toUpperCase()) {
                case 'LAST':
                case 'PRICE':
                    type = 'last';
                    break;
                case 'ASK':
                case 'SELL':
                    type = 'sell';
                    break;
                case 'BID':
                case 'BUY':
                    type = 'buy';
                    break;
                case 'HIGH':
                    type = 'high';
                    break;
                case 'LOW':
                    type = 'low';
                    break;
                case 'AVG':
                case 'AVERAGE': // not volume weighted
                    type = 'avg';
                    break;
                case 'VOL':
                case 'VOLUME':
                case 'BASEVOLUME':
                    type = 'vol';
                    break;
                case 'QUOTEVOLUME':
                case 'VOLUME_CURRENCY':
                    type = 'vol_cur';
                    break;
                default:
                    throw new Error('Type "' + type + '" not found!');
            }
            apiPath = 'https://api.liqui.io/api/3/ticker/' + pair;
            response = UrlFetchApp.fetch(apiPath);
            data = JSON.parse(response.getContentText());
            return parseFloat(data[pair][type]);

        case 'BITTREX':
            // Gets the ticker of the input currency pair at bittrex.com.
            switch (type.toUpperCase()) {
                case 'LAST':
                case 'PRICE':
                    type = 'Last';
                    break;
                case 'ASK':
                case 'SELL':
                    type = 'Ask';
                    break;
                case 'BID':
                case 'BUY':
                    type = 'Bid';
                    break;
                case 'HIGH':
                    type = 'High';
                    break;
                case 'LOW':
                    type = 'Low';
                    break;
                case 'AVG':
                case 'AVERAGE': // not volume weighted
                    type = 'Average';
                    break;
                case 'VOL':
                case 'VOLUME':
                case 'BASEVOLUME':
                    type = 'BaseVolume';
                    break;
                case 'VOL_CUR':
                case 'QUOTEVOLUME':
                case 'VOLUME_CURRENCY':
                    type = 'Volume';
                    break;
                default:
                    throw new Error('Type "' + type + '" not found!');
            }
            apiPath = 'https://bittrex.com/api/v1.1/public/getmarketsummary?market=' + pair;
            response = UrlFetchApp.fetch(apiPath);
            data = JSON.parse(response.getContentText());
            if (type !== 'Average') {
                return parseFloat(data.result[0][type]);
            } else {
                // Calculate average
                high = parseFloat(data.result[0].High);
                low = parseFloat(data.result[0].Low);
                result = (high + low) / 2;
                return +result.toFixed(8); // Only 8 decimal points
            }
            break;

        case 'ETHERSCAN':
            // Gets the ticker of the input currency pair at etherscan.io.
            if (type.toUpperCase() !== 'LAST') {
                throw new Error('Type "' + type + '" not found!');
            }
            apiPath = 'https://api.etherscan.io/api?module=stats&action=ethprice';
            response = UrlFetchApp.fetch(apiPath);
            data = JSON.parse(response.getContentText());
            return parseFloat(data.result[pair.toLowerCase()]);

        case 'COINMARKETCAP':
        case 'CMC':
            // See https://coinmarketcap.com/api/ for types
            apiPath = 'https://api.coinmarketcap.com/v1/ticker/' + pair; // actually just one currency
            response = UrlFetchApp.fetch(apiPath);
            data = JSON.parse(response.getContentText());
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
    switch (currency) {
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
        default:
            return 'Error, invalid currency!';
    }

    date.setTime(date.getTime() - date.getTimezoneOffset() * 1000 * 60); // Remove timezone differences
    var dateTimestamp = Math.floor(date.getTime() / 1000);
    var nextDay = new Date();
    nextDay.setTime(date.getTime() + 1000 * 60 * 60 * (24 + 2)); // 26h to make sure with daylight savings

    var apiPath = 'https://api.gdax.com/products/' + currency + '/candles?granularity=86400&start=' +
        date.toISOString().substr(0, 10) + '&end=' + nextDay.toISOString().substr(0, 10);
    var response = UrlFetchApp.fetch(apiPath);
    var data = JSON.parse(response.getContentText());
    var dayBucket = data[0]; // [ time, low, high, open, close, volume ]

    if (dayBucket[0] !== dateTimestamp) {
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
    switch (currency) {
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
        default:
            return 'Error, invalid currency!';
    }

    date.setTime(date.getTime() - date.getTimezoneOffset() * 1000 * 60); // Remove timezone differences
    var dateTimestamp = Math.floor(date.getTime() / 1000);

    var apiPath = 'https://api.kraken.com/0/public/OHLC?pair=' + currency + '&interval=1440';
    var response = UrlFetchApp.fetch(apiPath);
    var data = JSON.parse(response.getContentText());
    data = data.result[currency];
    var dayBucket = data[data.map(function(e) {
        return e[0];
    }).indexOf(dateTimestamp)];

    if (dayBucket[0] !== dateTimestamp) {
        return 'Error in internal function';
    }
    var result = parseFloat(dayBucket[5]); // VWAP

    // If it needs to be converted from USD (see comment above)
    if (currency === 'USDTZUSD') {
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

    for (var i = 1; i <= range.getNumRows(); i++) {
        for (var j = 1; j <= range.getNumColumns(); j++) {
            if (!range.getCell(i, j).isBlank()) {
                var currentValue = range.getCell(i, j).getValue();
                if (currentValue.slice(0, 2) === '0x') {
                    if (currentValue.length === 66) {
                        range.getCell(i, j).setFormula('=HYPERLINK("https://etherscan.io/tx/' +
                            currentValue + '"' + seperator + '"ETH-TXID: ' + currentValue +
                            '")');
                    } else if (currentValue.length === 42) {
                        range.getCell(i, j).setFormula('=HYPERLINK("https://etherscan.io/address/' +
                            currentValue + '"' + seperator + '"ETH-ADDR: ' + currentValue +
                            '")');
                    }
                } else if (currentValue.length === 64) {
                    range.getCell(i, j).setFormula('=HYPERLINK("https://tradeblock.com/bitcoin/tx/' +
                        currentValue + '"' + seperator + '"BTC-TXID: ' + currentValue + '")');
                }
            }
        }
    }
}
