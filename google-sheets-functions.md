# Google Sheets Crypto Market Functions
Google Apps Script code with several functions as crypto market tools which can be used with Google Sheets.

## [google-sheets-functions.js](https://github.com/microraptor/blockchain-misc/blob/master/google-sheets-functions.js)

In Google Sheets click on 'Tools -> Script Editor...' in the menu and save the code there. To get the value of a function simply enter `=functionName(parameter1, parameter2)` in a cell. For example: `=getTicker("GDAX", "last", "ETH-USD")`. If the Spreadsheet is in a locale which uses a comma as decimal mark you, need to use a semicolon as a parameter seperator. Like this: `=getTicker("GDAX"; "last"; "ETH-USD")`.

It is pretty helpful to put a cell with a current date as an extra parameter to the functions like this: `=getTicker("GDAX", "last", "ETH-USD", A1)`. Every time the date is updated all the functions will update themselves as well. A refresh button in the menu for this is implemented with the 'onOpen' and the 'refreshLastUpdate' functions.

The other functions are:

* gdaxEthUsd - Simple function to get the last price of ETH/USD at gdax.com

    `=gdaxEthUsd()`

* getTicker - Not so simple function to get the ticker of a specified exchange for a specific currency pair

    * Parameter 1: The exchange to get ticker from. Options: 'GDAX', 'Poloniex', 'Kraken', 'Bittrex', 'Liqui', 'Coinmarketcap', 'Etherscan'.
    * Parameter 2: The type of ticker. For example: 'Last', 'Ask', 'Bid', 'High', 'Low', 'Average', 'Volume', 'QuoteVolume'.
    * Parameter 3: The currency pair. For example: 'ETH-USD' for GDAX, 'BTC_ETH' for Poloniex, ''XXBTZEUR' for Kraken, 'wings_btc' for Liqui, 'BTC-WINGS' for Bittrex, 'ethusd' for Etherscan and 'ETHEREUM' for Coinmarketcap).

    `=getTicker("GDAX", "last", "ETH-USD")`

* getGdaxAverageUsd - Looks up the not volume weighted average price of the currency against USD on the specified day at gdax.com

    `=getGdaxAverageUsd(A1, 'Bitcoin')` where A1 is a cell with a date

* getKrakenAverageEuro - Looks up the volume weighted average price of the currency against euro on the specified day at kraken.com

    `=getKrakenAverageEuro(A1, 'Ethereum')` where A1 is a cell with a date

* updateHyperlinks - Makes blockchain explorer links out of IDs in the selected range

    Select the range where there are the IDs written in the cell and the select 'Custom -> Update Hyperlinks' in the menu.
