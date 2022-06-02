# Stock Analysis

## Overview of Project:
- Refactor the Module 2 solution code to loop through all the data one time in order to collect the same information that I did in this module. Then, I was able to determine whether refactoring the code successfully made the VBA script run faster.

## Results: 
#### Bellow the refectored code

- Sub AllStocksAnalysis()
   '1) Format the output sheet on All Stocks Analysis worksheet
   Sheets("All Stocks Analysis").Activate
   
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer
   
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
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
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
   Sheets(yearValue).Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Sheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
           Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
         
   Next i
   
   'Formatting all
   Worksheets("All Stocks Analysis").Activate
   Range("A3:C3").Font.FontStyle = "Bold"
   Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
   Range("B4:B15").NumberFormat = "#,##0"
   Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
   
     Worksheets("All Stocks Analysis").Activate
     
   'Coloring the cell that needed
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3) < 0 Then
            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed
        Else
            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

  Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)


End Sub

Sub formatAllStocksAnalysisTable()

'Formatting
    Worksheets("All Stocks Analysis").Activate
    
    'Select the header range and make the text bold
         Range("A3:C3").Font.Bold = True
         
      'Select the same range and add a border to the bottom edge
         Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
         
      'Formatting
   Worksheets("All Stocks Analysis").Activate
   
   Range("A3:C3").Font.FontStyle = "Bold"
   Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
   Range("B4:B15").NumberFormat = "#,##0"
   Range("C4:C15").NumberFormat = "0.0%"
     Columns("B").AutoFit
     
      

End Sub

Sub ClearWorksheet()

 Worksheets("All Stocks Analysis").Activate
   
    Cells.Clear

End Sub

Sub yearValueAnalysis()

yearValue = InputBox("What year would you like to run the analysis on?")



End Sub

#### Here are the imagens of the time for the analysis:

- *VBA_Challenge_2017.png*

![This is an image](https://github.com/KandiJayana/stock-analysis/blob/1672d1e1d9ceccb29b9cee095956d21e49b32107/Resources/VBA_Challenge_2017.png)

- *VBA_Challenge_2018.png*

![This is an image](https://github.com/KandiJayana/stock-analysis/blob/1672d1e1d9ceccb29b9cee095956d21e49b32107/Resources/VBA_Challenge_2018.png)


## Summary:

#### What are the advantages or disadvantages of refactoring code?
- An advantage is that after refactoring a code, you can make it a better quality with detailed comments that make it easier to understand, maintaining and to run it faster.

- On the other hand, one disadvantage of refactoring a code can be the time consuming. Also, If you don't remember or don't understand the previous code, you may not figure out which steps to take next. 


#### How do these pros and cons apply to refactoring the original VBA script?

- From the pros we can see that a clean and well-organized code is better to understand and to make changes. 
- From the cons I would say that the time consumed to change the original code was a lot, but worth to maintain it later as we can see the results on the images above.
