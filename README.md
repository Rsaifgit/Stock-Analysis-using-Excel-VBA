# Stock-Analysis-using-Excel-VBA

## Project Overview

**Purpose**

The objective of this project is to refactor code in Excel VBA to collect information about certain stocks for the year 2017 and 2018 which may help to determine whether the stocks are worth investing or not. Originially the process followed similar format, however, this time around the primary objective was to increase efficiency of the original code. 

**Data set**

Overall dataset  includes two tables with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.

## Results

Some preliminary tasks were already in place that was needed before refactoring the code, such as: creating the input box, chart headers, ticker array, and activating the appropriate worksheet. The steps were then listed out in order to set the structure for the refactoring. Below is the code as written in the file.

Sub AllStockAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

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
    
    Worksheets(yearValue).Activate
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    tickerIndex = 0
    
    Dim tickervolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    For i = 0 To 11
        tickervolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
    
    For i = 2 To RowCount
    
        tickervolumes(tickerIndex) = tickervolumes(tickerIndex) + Cells(i, 8).Value
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
            
        
    Next i
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickervolumes(i)
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
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


## Summary

**Pros & Cons of refactoring the code**

Refactoring faclitating in making the code cleaner and more organized. A few advantages of a cleaner code include design and software improvement, debugging, and faster programming. On the other hand, in terms of disadvantages applications could be too large to have any proper test cases for the existing codes posing some risks in refactoring. 

**Key advantages**

The key advantage of the refactored code is in significant reduction of time in running the overall analysis. For example, for the year 2017, it took around 548 seconds to run the code while it took only 3 seconds to run the analysis for 2018. The original versions run time was much higher than the 1st one (2017) when ran for the first time. 

![2017 analysis](https://github.com/Rsaifgit/Stock-Analysis-using-Excel-VBA/blob/main/vba_challenge_2017.png)
![2018 analysis](https://github.com/Rsaifgit/Stock-Analysis-using-Excel-VBA/blob/main/vba_challenge_2018.png)











