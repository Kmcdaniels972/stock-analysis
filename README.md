# Stock Analysis Refactored Vs Original 

## Overview of Project 

Using VBA Marcos, I create one macro to calculate the differnces of the daily starting and ending prices of all 12 present Company Stocks for the year. 
This marco takes the year input, generates a report that displays the total daily volume and the return for the year of each stock, and displays the process time in another dialog box.
The second marco optimized or refactors the amount of process time spent looping through all 3013 rows of data to find the ticker value and calculate the daily volume. In this macro there is more definite variables using arrays and less nested loops.

## Results 

### Comparing Stock performance
The original marco with many variable defined by nested loops code ran about 0.87 seconds for 2017 and 0.84 seconds for 2018.
After refactoring the macro and defining all the stocks with arrays and separating the loops the code ran about 0.09 seconds for both 2017 and 2018 stocks.
Separating the nesting looks and using arrays to have a more defined variables for the stock tickers

#### The Refactor Code Run Time
![VBA_Challenge_2017](https://user-images.githubusercontent.com/106792451/176613172-fa135b5e-d667-4b6d-8f96-307f38f0af9a.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/106792451/176613209-608338d4-53d2-466f-9bd3-30f9ca11512c.PNG)

#### The Original Code Run Time
![Original Run time 2017](https://user-images.githubusercontent.com/106792451/176614913-e2ec6912-d6ed-4414-807e-e0e7f425fda4.PNG)
![Original Run time 2018](https://user-images.githubusercontent.com/106792451/176614949-597cc3be-9327-48a1-a146-a222c4b503e3.PNG)


## Summary 
The main advantages of refactoring are the less store data VBA is handling at once while running the code initial, the code run time is obviously better, and the code is more scalable. The disadvantage of refactoring the amount of code that is there since the nested loops are separated instead of one piece of code. Refactoring may take longer to build but this process of opitimization is ideal for longevity.
