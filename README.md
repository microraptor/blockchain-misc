# blockchain-misc
[microraptor.github.io/blockchain-misc](https://microraptor.github.io/blockchain-misc/) - Miscellaneous blockchain and crypto stuff

## [Google Sheets Crypto Market Functions](https://github.com/microraptor/blockchain-misc/blob/master/google-sheets-functions.js)
google-sheets-functions.js is Google Apps Script code with several functions as crypto market tools which can be used with Google Sheets.

In Google Sheets click on 'Tools -> Script Editor...' in the menu and save the code there. To get the value of a function simply enter `=functionName(parameter1, parameter2)` in a cell. For example: `=getTicker("GDAX", "last", "ETH-USD")`. If the Spreadsheet is in a locale which uses a comma as decimal mark you, need to use a semicolon as a parameter seperator. Like this: `=getTicker("GDAX"; "last"; "ETH-USD")`.
 
It is very helpful to put a cell with a current date as an extra parameter to the functions like this: `=getTicker("GDAX", "last", "ETH-USD", A1)`. Every time the date is updated all the functions will update themselves as well. A refresh button in the menu for this is implemented with the 'onOpen' and the 'refreshLastUpdate' functions.

The other functions are:
* gdaxEthUsd - Simple function to get the last price of ETH/USD at gdax.com
* getTicker - Not so simple function to get the ticker of a specified exchange for a specific currency pair
* getGdaxAverageUsd - Looks up the not volume weighted average price of the currency against USD on the specified day at gdax.com
* getKrakenAverageEuro - Looks up the volume weighted average price of the currency against euro on the specified day at kraken.com
* updateHyperlinks - Makes blockchain explorer links out of IDs in the selected range
