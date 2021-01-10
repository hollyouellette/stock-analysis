# Analysis of Green Energy Stocks
  VBA analysis of Green Energy Stocks for 2017 & 2018.


## Introduction & Project Overview

An original script for VBA analysis of Green Energy stock was created to analyze the marketplace performance of 12 different Green Energy stocks in the years 2017 and 2018. The goal of this project was to refactor the orginal script to make it gather the same information, only faster. This project was taken on to ensure that this code can run efficiently and accurately when analyzing the entire stock market, versus the dozen stocks that the original code was written to analyze.

## Results

### Analysis of 2017 and 2018 Return
  
  This analysis was performed to assess the performance of each stock based on it's **Yearly Volume**, how often the stock get's traded, and it's **Yearly Return**, the percentage increase (or decreate) in stock price from the beginning of the year to the end of the year.
  
  Below are screenshots of the VBA Analyis generated with the refactored code:
  
  <img align="left" src="Additional_Resources/Ticker_Analysis_2017.png">
 
 ![](Additional_Resources/Ticker_Analysis_2018.png)
 
 Based on these analyses, we can make the following comparisons to the stock performance between 2017 and 2018:
 
  - Green Energy Stocks as a stock sub-category saw a significant decrease in performance in 2018.
  - The amount that a stock is traded, it's yearly Volume, does not necessarily influence it's Yearly Return.
  - DQ's stock plummeted in 2018, going from haveing the highest yearly reutn in 2017 to the lowest Yearly Return in 2018.
  - If we were to advise someone on where best to invest their money, based on 2017 and 2018 data, ENPH and RUN have continued to see positive returns accross the last two yers.
  
### Analysis of Execution Times
#### Refactored Script vs. Original Script

When refactoring this VBA script, two considerations were taken into account:

 1. Improving the processing speed of the code.
 2. Updating Variable names to more specifically delineate what data they represent.
 
 
##### Improving the Processing Speed 

   The orignial script was using two indipendent loops; one loop to calculate the Valume and a separate loop to calculate the Return. 
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
   In order to make this code more run more efficient, the re-factored code included the tickerVolumes as an output array in the same nested loop as the tickerStartingPrices and tickerEndingPrices.
   
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single
  
  As a result of the refactoring, we saw an increase in the processing spead for both the 2018 worksheet and an identical processing spead for the 2017 worksheet (shown in the screenshots below):
  
  ###### Processing Speed of Refactored Script
  
  <img align="left" src="Resources/VBA_Challenge_2017.png" width="450">
  
  <img src="Resources/VBA_Challenge_2018.png" width ="475">
  
  ###### Processing Speed of Original Script
  
  <img align="left" src="Additional_Resources/OriginalScript_2017.png" width="450">
  
  <img src="Additional_Resources/OriginalScript_2018.png" width ="450">
 
 
##### Updating Variable Names
  
   This taylors the code more specifically to the dataset that we are analyzing and makes the code easier to read and understand.
   
   Original Script:
    
      Cells(4 + i, 1).Value = ticker
      Cells(4 + i, 2).Value = totalVolume
      Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
  
  Refactored Code:
  
      Cells(4 + i, 1).Value = tickerIndex
      Cells(4 + i, 2).Value = tickerVolumes
      Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
 
 ## Summary
 ### Advantages and Disadvantages of Refactoring Code
 
 **For refactoring code in general:**
  
  Generally speaking, the first attempt at writing a code to accomplish a specific task might not be the best solution. By refactoring, we can update the code to take fewer steps, run faster and use less memory. In addition to this, refactoring code can make it easier for users to read and understand. 
  
  However, refactoring code also has disadvantages. There is a potential risk of introducing bugs and stop a previously functional code from running. In addition to this, refactoring code requires an additional time investment layered on top of the time taken to write the orignal script. This additional time invested might only yield an improved processing time of merely factors of a second and ultimately produces the same end product as the original code. 

**For the refactored VBA Script:**
  
  In this assignment specifically, the refactored VBA script did make the code run faster when looping through the 2018 Worksheet. In addition to this, the refactored VBA Script is both easier for users to read and to understand.
  
  For disadvantages related to this assignment, while the refactored code did slightly improve the processing time for the 2018 Worksheet, this improvement was very minimal in relation to the time that it took to refactor the code. In addition to this, in the process of refactoring the code, there were several instances where the code did not run. This required additional time investment to de-bugging the updates to the code so that it could run correctly.
  
