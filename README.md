# Stock Analysis with VBA

## Overview

In this challenge, I helped a friend analyze stock data in Excel. Through VBA, I was able to analyze the total traded volume and total returns for 12 different stocks.
My initial code analyzed the data quickly, but it would probably not be fast enough to analyze thousands of stocks options. Therefore, I was tasked with refactoring my initial code, so that it would run faster. 

## Analysis

Originally, my code yielded the following stock analysis results for the years 2017 and 2018, respectively:

![Stock Analysis Results 2017](https://github.com/CSoldo1/stock-analysis/blob/main/All_Stocks_2017.PNG)

![Stock Analysis Results 2018](https://github.com/CSoldo1/stock-analysis/blob/main/All_Stocks_2018.PNG)

To make the code run faster, I set the tickerIndex to zero before looping over all the rows.

 '1a) Create a ticker Index
    tickerIndex = 0
    
 I created arrays fot the volume, starting price, and ending price outputs.
 
 '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
  The a programmed the script to loop through begin all the stock volumes at 0.
  For i = 0 To 11
        tickerVolumes(i) = 0
    Next i

## Results

After refactoring the code, I was able to get the following results for each year. 

![Refactored Stock Analysis Results 2017](https://github.com/CSoldo1/stock-analysis/blob/main/All_Stocks_2017_Refactored.PNG)

![Refactored Stock Analysis Results 2018](https://github.com/CSoldo1/stock-analysis/blob/main/All_Stocks_2018_Refactored.PNG)


## Discussion

My refactored code sped up the computational time by over 1 second for the year 2017, and almost 0.7 seconds for the year 2018. While this is a relatively small difference, the results would be more exaggerated if I was analyzing much larger data sets. 
I think the refactored code ran faster because I was able to remove the nested loops from the original code. But honestly, the time that I saved running the new code was nothing in comparison to the time it took me to figure out how to refactor it. I did not understand how the original code analyzed the data in the Excel spreadsheet, so it was very hard to figure out what the new code should run. I think that's one of the main disadvantages with trying to refactor code, especially if it was created by someone else. Tryign to refactor code seems like it would introduce new errors into an otherwise functional script, thereby increaseing overall effort and time. 

The only advantage I can see to refactoring code is in the world of algorithmic trading. When milliseconds could make the difference on whether or not a transaction is executed (and whether or not money is made or lost), getting the code to run as fast as possible would be worth the time and effort. 
