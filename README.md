# Green Stocks Analysis

## Project Overview
Client, Steve, is looking for analysis various green energy stocks to diversify his clients portfolio.  Green Energy is a passion of said clients and there are numerous forms of energy to invest in.  Today, the client has invested solely in DAQO New Energy Corp(DQ), that makes silicon wafers for solar panels.  Steve is concerned his clients have put all of their money into one stock and wants to suggest they diversify by investing in other green energy companies.  Steve has identified twielve other green companies that could be potential investments but needs further analytics on the health of the companies before he can suggest to his clients.

## Results
Steve was interested in finding total daily volume and yearly return for each stock.  Since Steve's clients have invest only in Daqo, we started with an analysis of Daq only. We discovered that the stock is traded often, 107,873,900 times in 2018.  Steve's clients believe if a stock is traded often then the price will accurately reflect the value of the stock 
###[Contribution for the project](green_stocks (2).xlsm) DQ Analysis Tab

In performing the analysis of Daqo only, we wanted to understand the yearly return in addition to how often the stock was traded.  Through our analysis we were able to uncover the stock dropped over 63% 
###[Contribution for the project](vbachallenge.xlsm) DQ Analysis Tab

Based on some of the data, it appears Daqo may not be the best stock for sole investment.  We wanted to analyze multiple stocks to find better choices for Steve's clients.
###[Contribution for the project] (vbachallenge.xlsm) All Stocks Analysis Tab

To see how stocks performed at a glance, we colored positive returns in green and negative returns in red.  We provided this information for all stocks in 2017 and 2018.
https://github.com/kcpals/stock-analysis/blob/main/2017.png
https://github.com/kcpals/stock-analysis/blob/main/2018.png
Additionally, to make running the analysis easier, we provided a Run Analysis for All stocks button within the document

In the case Steve would like to to perform analysis on larger datasets and wants to know how fast his VBA code will compile results, we've added a script to calculate how long the code takes to execute and output the elapsed time in a message box.
Referencing the following, will show you the elapsed time for outputs as they are today
https://github.com/kcpals/stock-analysis/tree/main/Resources

SunRun(Run) and Enphase Energy(ENPH) showed the most positive returns in 2018 and 2017.  In fact all stocks with the exception of Terra Form Power(TERP) showed a positive return in 2017

## Summary
In summary, refactoring code appears to be a worth-while exercise.  The advantage is to give the client a cleaner document to build upon with larger data sets without prolonging run times.  If maintaing the document over a long period of time is known, one would definitely want to refactor the code.   The only disadvantage I can find is having to comb through the code in order to get to this state.  It can be a complicated if one if not well versed in VBA code. 

Advantages of the original code appear to be for simplicity sake, for example if the data set is small, will not be changed or built upon any time soon or if time to complete the project is limited.  The disadvantages of the original code of course is if larger data will be added to the raw data later for larger analysis there could be long analysis times as well potential bugs as the data set grows.


