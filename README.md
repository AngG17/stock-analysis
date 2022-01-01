# An Analysis of Stocks Using Refactor
## Overview of Project

Analyze stocks over a two year time period by refactoring code to see if new code is faster

### Purpose

To help Steve give advice to his parents on which stocks to invest in.

## Results
Before starting the assignment, I needed to figure out how to open the VBA Challenge Code.  With help from classmates, I was able to open the code by 'Open With' using Notepad.  Once I was able to view the code, I copied the code from Notepad and pasted it in the Visual Basics editor.  As I had named the worksheet tab 'AllStocksAnalysis' instead of 'All Stocks Analysis', I had to edit the code to match my naming convention.  

I made the below changes to the code to get the results shown in the screenshots.

'1a) Create a ticker Index
    Dim tickerIndex As Integer
    
        tickerIndex = 0
        

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        
        tickerVolumes(12) = 0
        
    
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If

            '3d Increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("AllStocksAnalysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

![VBA_Challenge_2017](https://user-images.githubusercontent.com/95720986/147861568-24bf731a-54dc-4e4a-bf92-1d9ede9e19da.png)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/95720986/147861571-1e07888b-0e96-4629-80b6-8f6173c0ca32.png)


## Summary
### Advantages of Refactoring Code


### Disadvantages of Refactoring Code


### Advantages of the Original and Refactored VBA Script


### Disadvantages of the Original and Refactored VBA Script
