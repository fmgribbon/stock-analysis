# Stocks Analysis

## Overview of Project
This project is an analysis of the 2017 and 2018 annual stock performance using the given Excel worksheet stock data and a VBA code. Refactoring the VBA code is also done to improve the processing time. 
     
## Purpose
- To compare the stock performance for the years 2017 and 2018 based on the given stock data and VBA code. 
- To improve processing time of  the given VBA code.  

## Results
      
### Stock Analysis Results 

#### Comparing stocks for the years 2017  & 2018 

[Summary of 2017 Stocks Annual Daily Volume and Return](https://github.com/fmgribbon/stock-analysis/blob/main/Resources/AllStocks2017.png)

[Summary of 2018 Stocks Annual Daily Volume and Return](https://github.com/fmgribbon/stock-analysis/blob/main/Resources/AllStocks2018.png)

##### Significant changes noted.

######     Stock Return
- The RUN stocks have the greatest increase in return (up by 78% from 6% in 2017 to 84% in 2018). 
- The DQ stocks have the greatest decrease in return (down  by 262% from 199% in 2017 to -63% in 2018). 

######     Stock Volumes

- The ENPH stocks have the greatest increase in volume up by 385 million from 222 million in 2017 to 607 million in 2018.
- The SPWR stocks have the greatest decrease in volume down by 244 million from 782 million in 2017 to 538 million in 2018.

### VBA Code Refactoring Results
- 2017 Runtime Results

[Original VBA Code](https://github.com/fmgribbon/stock-analysis/blob/main/Resources/OriginalVBACode2017.PNG)

[Refactored VBA Code](https://github.com/fmgribbon/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

- 2018 Runtime Results
- 
[Original VBA Code]([Original VBA Code](https://github.com/fmgribbon/stock-analysis/blob/main/Resources/OriginalVBACode2017.PNG)

[Refactored VBA Code](https://github.com/fmgribbon/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

Through refactoring, the following decrease runtime was seen 

- In the 2017 Stock Analysis VBA Code, there was  a 74.85% decrease in runtime from 0.6523438 to 0.1640625 seconds.  
- The 2018 Stock Analysis VBA Code, there was a 75% decrease in runtime from 0.640625 to 0.1601563 seconds. 

## Summary

### Stock Analysis Summary
  
An analysis usedthe given stock data and the VBA code. The VBA code automated the calculations and output the annual performance table. 
The year, stock ticker, daily stock closing price and stock volume data were used to calculate the annual stock return and total stock volume. 
The year that was used in the analysis was provided by the user through a popup message box. 
It prompts the user to enter a year (2017 or 2018). Based on the year entered by the user, the VBA code calculated the annual stock total volume and return. The annual stock volume is the sum of the daily stock volume for the year (user input) . The stock return is the ratio (in percent) of the ending stock price for the year and the starting stock price minus 1. A new All Stock Analysis EXCEL worksheet, contains the annual stock volume and return of the given data, was created.   
  
### VBA Code Refactoring Summary
  
#### Changes made to the given VBA Code 

- The original code had the output of calculation results in a new worksheet inside the loop that does the calculation. 
  In the refactored code the output of the calculation was created outside of the calculation loop. 
  The decrease in runtime could be attributed to the reduced number of  active worksheet,  from 2 active worksheet to 1, inside the calculation loop. 

- One of the "if" statement in the original code was changed to a nested "if" statement to simplify the compound conditions.
   
      
 [Snapshot of the original code](https://github.com/fmgribbon/stock-analysis/blob/main/Resources/SnipitOriginalVBACode.txt)
 
 [Snapshot of the refactored code](https://github.com/fmgribbon/stock-analysis/blob/main/Resources/SnipitRefactoredVBACode.txt)
