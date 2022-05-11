# VBA of Wall Street

## Overview of Project

### Purpose
The purpose of this project was to perform data analysis on green energy stocks with respect to their Total Daily Volume and Return. Refactoring the code will improve its efficiency when the dataset is expanded to include the entire stock market. 

## Results

### Planning the code 
Before the refactored script can be designed, a comprehensive analysis of the initial script was completed to determine what can be kept and what should be changed. The steps for the refactoring process were clearly documented as comments to maintain the organization and readability of the original script. The refactored script can be seen below. 

``` VBA
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    
    'Ask user what year to perform analysis on
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'Start Timer
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

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
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
        '3c) Check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d) Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        Else
            Cells(i, 3).Interior.Color = vbRed
        End If
    Next i
    
    'End Timer
    endTime = Timer
    
    'Display run time for code
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

### Analysis of Green Energy Stocks
Comparing the stock performance between 2017 and 2018, the return in 2017 was considerably better than 2018. Specifically, 11 of the 12 stocks in 2017 had a positive return while only 2 of the 12 stocks in 2018 had a positive return, as shown in the screenshots below. Positive returns are indicated by green cells and negative returns are shown in red. 

![VBA_Challenge_2017.png](https://github.com/daniel-sh-au/UofT_DataBC_Module02_stocks-analysis/blob/main/resources/VBA_Challenge_2017.PNG)

![VBA_Challenge_2018.png](https://github.com/daniel-sh-au/UofT_DataBC_Module02_stocks-analysis/blob/main/resources/VBA_Challenge_2018.PNG)

### Original vs. Refactored Script
When comparing the execution times between the original and refactored script, refactoring the code decreased the runtime by around 1.6 seconds (1.648 seconds from original – 0.074 seconds from refactored). The execution time for the original script is provided in the screenshot below. The main modification in the refactoring process was the removal of the nested for loop and the addition of a ticker index which was incremented for each ticker symbol. 

![VBA_Runtime_original_script.png](https://github.com/daniel-sh-au/UofT_DataBC_Module02_stocks-analysis/blob/main/resources/VBA_Runtime_original_script.PNG)

## Summary

### Advantages and Disadvantages of Refactoring Code
pros: reduce runtime of code
cons: might reduce readability

### Application to Refactored Script
runtime of code 

