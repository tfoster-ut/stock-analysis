Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    
    'YearValue Input
    yearValue = InputBox("What year would like to run the analysis on?")
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Naming headers
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'set initial volume to zero
    totalVolume = 0
    
    
    Worksheets(yearValue).Activate
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    
    'Creating a loop
    For i = 2 To RowCount
    
        'increase totalVolume if ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
        End If
        
        'Find first closing price
        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        'Set starting price
        startingPrice = Cells(i, 6).Value
        End If
        
        'Find ending price
        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        'Set ending price
        endingPrice = Cells(i, 6).Value
        End If
        
    Next i
    
    'Adding data
    Worksheets("DQ Analysis").Activate
    
    Cells(4, 1).Value = yearValue
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = endingPrice / startingPrice - 1
    
    'Formatting
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:C3").Interior.ColorIndex = 23
    Range("A3:C3").Font.ColorIndex = 2
    Cells(4, 2).NumberFormat = "#,##0.00"
    Cells(4, 2).Style = "Currency"
    Cells(4, 3).NumberFormat = "0.00%"
    
    Columns("A:C").AutoFit
    
    'Condtional Formatting
    If Cells(4, 3) > 0 Then
        Cells(4, 3).Interior.Color = vbGreen
    ElseIf Cells(4, 3) < 0 Then
        Cells(4, 3).Interior.Color = vbRed
    Else: Cells(4, 3).Interior.Color = xlNone
    
    End If
    
    
End Sub


Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime As Single
    
    'value InputBox
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'Timer
    startTime = Timer
    
    Cells(1, 1).Value = "All Stocks (" + yearValue + ")"

    '1) Format the output sheet on All Stocks Analysis worksheet
    Worksheets("AllStocksAnalysis").Activate
    
    
    
    
    'headers
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
    Worksheets(yearValue).Activate
    
    '3c) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '4) Loop through tickers
    For i = 0 To 11
        
            ticker = tickers(i)
            totalVolume = 0
            
            '5) loop through rows in the data
            Worksheets(yearValue).Activate
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
            
            Worksheets("AllStocksAnalysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
            
            
    Next i
    
    'Formatting
    
    Worksheets("AllStocksAnalysis").Activate
    
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:C3").Interior.ColorIndex = 23
    Range("A3:C3").Font.ColorIndex = 2
    Range("B4:B15").NumberFormat = "#,##0.00"
    Range("B4:B15").Style = "Currency"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("A:C").AutoFit
    
    'Conditional Formatting
    
    Worksheets("AllStocksAnalysis").Activate
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = vbRed
        Else: Cells(i, 3).Interior.Color = xlNone
        
        End If
    
    
    Next i
    
    'EndTimer
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    

End Sub
