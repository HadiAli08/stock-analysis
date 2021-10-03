# Stocks Analysis with VBA 

## Overview of Project
### Purpose
Originally we created a workbook for Steve to help him and his parents analyze "Green Stocks" which his parents were looking to invest in. Although this original code works, we are looking to refactor this code as it may not run smoothly with larger datasets or as efficiently as we would like it to. We will be writing code to retrieve the ticker, the total daily volume and lastly the return on each stock and make it easily readible. 
## Results

### Analysis
After opening the starting VBA code I could see that the outline of what needed to be done was in there. We already had the basis for what needed to be done from the original code, and now we also had the steps laid out in comments for what we needed to do to refactor the code. Below is the code file which should be simple to view. Looking at it, there are the added for loops and If-Then statements which were crucial to make the original code much more efficient. 
    
	Sub AllStocksAnalysisRefactored()
        Dim startTime As Single
        Dim endTime  As Single
        
        yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer
        tickerIndex = 0
        
        
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
                tickerStartingPrices(i) = 0
                tickerEndingPrices(i) = 0
            Next i
            
            
            
        ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker
                
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
            'If  Then
                
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                End If
                    
                
                
                
            'End If
            
            '3c) check if the current row is the last row with the selected ticker
             'If the next row’s ticker doesn’t match, increase the tickerIndex.
            'If  Then
            
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                End If
                
                
    
                '3d Increase the tickerIndex.
                
                    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
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
    


##Summary
### Advantages and Disadvantages of Refactoring
Right off the bat there are lots more advantages to refactoring. Some advantages include, efficiency (as seen in this case), cleanliness of the code, better design, and possibly much better ease of use. Although there aren't many disadvantages, there are some. Some disadvantages of refactoring could include: time as it takes lots of time to refactor code that does the same thing, and possibly the struggle of having to go line through line trying to understand what each line does and having to rewrite it. 
### Pros and Cons Original vs Refactored code
Right away we can bring up the numerous advantages of the Refactored code. The biggest advantages for the refactored code would have to be the efficiency as it is able to run much quicker compared to the original. For the 2017 data, the original code ran much slower taking .625 seconds to run vs the refactored code for 2017 taking .0859 seconds to run which is kmore than 7x faster. A con of the refactored code could be that the headings are not as nicely formatted, but the original code had only visual changes that could slow performance.
[![Old code 2017](https://raw.githubusercontent.com/HadiAli08/stock-analysis/main/Resources/Old%20code_2017.PNG "Old code 2017")](http://https://raw.githubusercontent.com/HadiAli08/stock-analysis/main/Resources/Old%20code_2017.PNG "Old code 2017")
[![new code 2017](https://github.com/HadiAli08/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG?raw=true "new code 2017")](https://github.com/HadiAli08/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG?raw=true "new code 2017")