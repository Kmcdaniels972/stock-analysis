# Stock Analysis Refactored Vs Original 

## Overview of Project 

Using VBA Marcos, I create one macro to calculate the differnces of the daily starting and ending prices of all 12 present Company Stocks for the year. 
This marco takes the year input, generates a report that displays the total daily volume and the return for the year of each stock, and displays the process time in another dialog box.
The second marco optimized or refactors the amount of process time spent looping through all 3013 rows of data to find the ticker value and calculate the daily volume. In this macro there is more definite variables using arrays and less nested loops.

## Results 

*Compare Stock performance
The original marco with many variable defined by nested loops code ran about 0.87 seconds for 2017 and 0.84 seconds for 2018.
After refactoring the macro and defining all the stocks with arrays and separating the loops the code ran about 0.09 seconds for both 2017 and 2018 stocks.
Separating the 

*Execution times original vs refactored

## Summary 
*What are the advantages of refactoring code
*How do these pros and cons apply to refactoring the orignial VBA script
