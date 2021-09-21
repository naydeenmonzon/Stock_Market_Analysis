# VBA scripting to analyze real stock market data


Script that loops through all the stocks for one year and output the following information:
* The ticker symbol.
* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
* The total stock volume of the stock.

Conditional formatting will highlight positive change in green and negative change in red.

It will also return the stock with the “Greatest % increase”, “Greatest % decrease” and “Greatest total volume”.

The VBA script will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.
