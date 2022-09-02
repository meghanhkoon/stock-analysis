# Stock-Analysis With VBA

## Overview of the Project 

### Purpose and Background

***Background***

Steve wants us to analyze twelve different green energy stocks to see if they are worth investing in. He wants to be able to easily analyze an entire dataset of stocks in both years 2017 and 2018. The key information he wants summarized for each type of stock is the Total Daily Volume and Return. 

Instead of manually using formulas in excel each time there is new data added, we wrote different VBA Scripts for Excel to do the work for us! Since he doesn't know how to use VBA in Excel, it was our job to create a VBA script to make the analysis happen at the click of a button. 

***Purpose***

The purpose of this new VBA script is to refactor the code to make it run more efficiently (faster) to output each Ticker's name, Total Daily Volume, and Return. See the screenshot below for the Analysis for All Stocks in 2018.

![AllStocks](Resources/AllStocks.png)


## Results 
In order to refactor the code that was already written, I needed to create separate for loops instead of a nested loop to run through the data one time to collect all the needed information. To do this, I needed to first create a ticker Index, three different output arrays (for Volumes, Starting and Ending Prices), and set tickerVolumes to zero. 


  'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Single
        
        'Set equal to zero
        tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    Next i
    
![Timer_2018](Resources/Timer_2018.png)
![Timer_2018_Refactored](Resources/Timer_2018_Refactored.png)


## Summary 
1. What are the advantages or disadvantages of refactoring code?
2. How do these pros and cons apply to refactoring the original VBA script?
