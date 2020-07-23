## The VBA of Wall Street
VBA script to analyze real stock market data

**Input Data**

Stock data for three years (one year data on one worksheet)

![] (https://github.com/Aastha-Arora/VBA-challenge/blob/master/Images/Input%20Data.png)

**VBA Script**

A VBA script was designed to analyze the the data. The script loops through all the worksheets
and runs the code on each worksheet to analyse the stock data for each year.

Two summary tables are created for each year. 

The first table aggregates the stock data for each ticker
and provides the **Yearly Change** from opening price at the beginning of a given year to the closing price at the end of that year;
the **Percent Change** from opening price at the beginning of a given year to the closing price at the end of that year;
and the **Total Stock Volume** of the stock. Positive change in stock is higlighed in green and negative change in red using `Conditional Formatting.`

![] (https://github.com/Aastha-Arora/VBA-challenge/blob/master/Images/Output%201.png)

The second table displays the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".

![] (https://github.com/Aastha-Arora/VBA-challenge/blob/master/Images/Output%202.png)
