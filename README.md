# Stock Analysis

## Overview of Project
The purpose of this analysis is to find the total daily volume for 12 different stocks (Tickers) as a measurement of trade activity and the yearly return to find the most successful ticker.

## Results

### Best Investment
The most successful ticker was DQ and RUN in 2017 and 2018, respectively. However, DQ suffered a loss of 62.6% the following year whereas RUN had seen a small profit of 5.5% the previous year. Although RUN experienced the most success in the latter year, DQ's high return in 2017 is still a better return despite the 62.6% decrease. For example, if $100 were invested at the beginning of 2017, by then end of 2018 DQ would have a surplus of $136.80 whereas RUN would have a surplus of about $94.12.
![Alt text](https://github.com/ftercero/stock-analysis/blob/main/2017%20Ticker%20Results.png?raw=true "2017 Ticker Results")

![Alt text](https://github.com/ftercero/stock-analysis/blob/main/2018%20Ticker%20Results.png?raw=true "2018 Ticker Results")


### Code differences
There were two codes used to obtain the same results. The refactor code used "tickerIndex" consistently to as opposed to calling out the cells in the excel sheet. the refactored code used a tickerIndex to access the stock ticker index for tickers, ticker volumes, starting and ending prices. In the refactored code, the two lines of script below;
"If Cells(j, 1).Value = ticker Then
   totalVolume = totalVolume + Cells(j, 8).Value"
was "simplified" to one line, "tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value". In the script mentioned, the used of tickerIndex is clearly visible.

Additionally, the refactored code below does not use nested loops, 
    For i = 0 To 11
        tickerVolumes(i) = 0  
    Next i
    For i = 2 To RowCount
whereas the original code does.
 For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
    
    Worksheets(yearValue).Activate
        For j = 2 To RowCount

### Output time
The refracted code was created to find a more efficient code to run stock analysis. As a result the refracted code was about 0.2 seconds faster than the initial code. If data increased 5 fold, the refracted code would be 1 second faster. To save 1 minute, data would need to be 30 times the size of the current data.

![Alt text](https://github.com/ftercero/stock-analysis/blob/main/2017_Original..png?raw=true "2017 Original Execution Time")

![Alt text](https://github.com/ftercero/stock-analysis/blob/main/2017_Refactored..png?raw=true "2017 Refactored Execution Time")


![Alt text](https://github.com/ftercero/stock-analysis/blob/main/2018_Original..png?raw=true "2018 Original Execution Time")

![Alt text](https://github.com/ftercero/stock-analysis/blob/main/2018_Refactored..png?raw=true "2018 Refactored Execution Time")


## Summary
A potential advantage for refactoring code is time. However, it is possible that the time refactoring the code may outweigh the efficiency. Another advantage is creating a clean and easy to read script over time which in turn can make it easier to maintain. Nevertheless, the script may be simplified to the extent where only a few individuals can understand the script.

From my perspective, the original code was more simple to write and comprehend. The script was longer but was easier to read as it was a step by step script. The refactored code used the tickerIndex frequently and was easy to get lost. For comparison, reading the refactored was like reading a medical article. A biology or chemistry student may be able to understand the overall view, the output, but will need to research the details, the references to "tickerIndex". Whereas the orginal code is like a medical article having all the details in the article so additional researching is not necessary.

