# Stocks Analysis with VBA in Excel

## Overview of Project

To help Steve and his parents to analysis the Green energy stocks for investment. Using the Microsoft VBA 

### Purpose
The Purpose of the stocks analysis is to collect stock information from 2017 to 2018. By refactoring the code on Microsoft VBA, we use these collected information to determind whether or not the stocks are good for investing. We also collecting the run time for the code to see how effienct that VBA brings to us for analyzing data.   

## Results

I followed the steps by implementing the refactored code on VBA. I created a new module in the VBA project to compared the run time between the original code and the refactored code. Below is my written code followed by the instruction. 
>     '1a) Create a ticker Index
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As LongLong
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
         
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
               tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
             If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then

               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
               
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
                    
            tickerIndex = tickerIndex + 1
            
            End If
                
        'End If
    
    Next i
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

  For i = 0 To 11
  
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
 Next i
## Summary

### The advantages of the refactoring code

The advantages for refactored code is it increased the run time on macro when analyze the stocks data. The original macro it took about 0.6 seconds to run the macro when I was analyzing the stocks. The refactored code run time is 0.093 seconds to run the macro. It increased the speed to analysis data about 85%. Attached is the picture of the run time screenshot in 2017 and 2018. 

 ! [2018 Run time] (/Users/veronica/Desktop/Module 2 Challenge/Resources)
 ! [2017 Run time] (/Users/veronica/Desktop/Module 2 Challenge/Resources/2017-Run-time.png)

### The pros and cons apply to refactoring the original VBA script
The pros apply to refactored codes is well organized. It improves the programming time for the dataset. Overall, it helps the code more visulized. The cons apply to refactored code is subscript out of range often. The application may be to large to have the right informations to test the result. Attached is the stocks analysis for Steve's parent. These are some better stocks to invest which is "ENPH" and "Run". 

! [2018 All Stocks Analysis] (/Users/veronica/Desktop/Module 2 Challenge/Resources)
 ! [2017 All Stocks Analysis] (/Users/veronica/Desktop/Module 2 Challenge/Resources)
