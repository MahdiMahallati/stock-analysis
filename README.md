# Stock Analysis With Excel VBA
https://github.com/MahdiMahallati/stock-analysis

## Overview of Project
### Purpose
The purpose of this analysis is to expand the pre-existing dataset to include the entire stock market over the last few years and decipher which stocks have performed best based on volume of trades and price fluctuation. The project utilised refactoring Microsoft Excel VBA code to collect information on 12 stocks in the years 2017 and 2018 to determine what stocks are worth investing in. 

## Results
### Analysis
The steps to completing the analysis were listed out in order to set the structure for the refactoring of the code. The code is written below. 

Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer

'1)Format the output sheet on the "All Stocks Analysis" worksheet.

    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

'2)Initialize an array of all tickers.

    Dim tickers(12) As String
    'Creates and array with 12 elements
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
  
'3a)Initialize variables for the starting price and ending price.

    Dim startingPrice As Single
    Dim endingPrice As Single
    
'3b)Activate the data worksheet.

    Worksheets(yearValue).Activate

'3c)Find the number of rows to loop over.

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'4)Loop through the tickers.

    For i = 0 To 11
        ticker = tickers(i)
        'Do stuff with ticker
        totalVolume = 0
    
'5)Loop through rows in the data.

    Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
    '5a)Find the total volume for the current ticker.
    
        If Cells(j, 1).Value = ticker Then
        totalVolume = totalVolume + Cells(j, 8).Value
        
    End If
    
    '5b)Find the starting price for the current ticker.
    
        If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
            startingPrice = Cells(j, 6).Value
            
        End If
        'Determines the beginning of the ticker section
        
    '5c)Find the ending price for the current ticker.
    
        If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
            endingPrice = Cells(j, 6).Value
        End If
        'Determines the end of the ticker section
        
    Next j
    
'6)Output the data for the current ticker.

    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = ticker
    
    Cells(4 + i, 2).Value = totalVolume
    
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
 Next i


Worksheets("All Stocks Analysis").Activate

Range("A3:C3").Font.Bold = True

Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous

Range("B4:B15").NumberFormat = "#,##0"

Range("C4:C15").NumberFormat = "0.0%"

Columns("B").AutoFit

'Color Formatting
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
    
    
    If Cells(i, 3) > 0 Then
        Cells(i, 3).Interior.Color = vbGreen
        'Color the cell green
    ElseIf Cells(i, 3) < 0 Then
        Cells(i, 3).Interior.Color = vbRed
        'Color the cell red
    Else
        Cells(i, 3).Interior.Color = xlNone
        'Clear the cell color
    
    End If
    
Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

## Summary
### Pros and Cons of Refactoring Code
Code refactoring helps make code more organised, cleaner, improve design and software, help debugging and lead to faster programming. However some shortcomings include having an application that is too long to have proper test cases for existing code, code of this nature can be easily broken while attempting to refactor it. 

### The Advantages of Refactoring Stock Analysis
The biggest advantage of running refactored code was a decrease in macro run time. Attached below are screenshots that indicate the run time for our new analysis. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/100375726/159191676-afd030e7-911f-4f7b-b5ac-39c141d60c22.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/100375726/159191682-f01035c7-c732-4859-b26a-4065962106a2.PNG)


