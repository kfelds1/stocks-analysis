# Stocks Analysis

## Background

Steve is looking to analyze the stock market but needs help finding a way to do it. Excel's Visual Basic Application is the perfect tool to analyze the performance of the 12 green stocks to show that DAQO is not the only stock worth investing in the way his parents want him to.

## Purpose

The purpose of this analysis was to help Steve show his parents different green stocks for the years 2017 and 2018 and whether or not they were worth investing in. To do this, I showed the total volume for one particular stock for either 2017 or 2018, and I also showed the return to that stock's investment using its beginning and ending prices. This return value was either positive or negative, giving a general idea on the stock's overall performance that year. After successfully completing the analysis with nested for loops, it was my job to refactor the code to find a quicker and more efficient way to achieve the same outcome.

## Results

One of the biggest differences between the two versions of code and the perhaps the driving factor behind the increase in efficiency is the creation of the arrays for 'tickerVolume,' 'tickerStartingPrices,' and 'tickerEndingPrices,' These arrays, each with a size of 12 for the 12 stocks, allowed for the storage of data for each ticker (stock). This way, we could compile data for each ticker rather than each row one by one, tremendously shortening the running time of the analysis, as seen below.

Th


## Summary

### Advantages and Disadvantages of Refactoring Code

The most evident advantage to refactoring code is that it will decrease the run-time and work that a program will have to do in order to complete the analysis. The increase in efficiency is significant and is helpful when there are thousands of rows in a dataset. Refactoring code might not help just the software programming, but it also might be more concise and readable for the viewer of the code.

A disadvantage to refactoring code that comes to mind is the time and effort it might take to do so. Almost always there is more than one way to complete a task while coding, so if the route that the coder decides to go is not the most efficient, it will take more time for them to determine the most efficient way. Another disadvantage of refactoring is the risk of creating errors with already working and useful code, especially in large datasets. It might be easy to accidentally delete or mess up perfectly working code when trying to decide the most logical program.

### Advantages and Disadvantages of Refactoring the Stocks Analysis

The main advantage to refactoring this code is the decrease in run time between the original and refactored programs. Another advantage to refactoring this particular code is the elimination of the nested for loops, which made the original code a little bit difficult to read and follow. The refactored code has for loops, but there are separate from one another rather than within one another. Through this process, I feel that I have developed a deeper understanding of writing code on VBA, especially with for loops and arrays.

The disadvantage to the refactored code is the creation of another variable, 'tickerIndex.' This might pose as a source of confusion because it has a similar name to the array 'ticker.' It is important to remember that 'tickerIndex' is a variable while 'ticker' is an array, and it is helpful to think of 'tickerIndex' like the 'ticker' array but in variable form, so each position in the array is assigned a number 0-11.


