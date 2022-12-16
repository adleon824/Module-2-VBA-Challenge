# Module-2-VBA-Challenge
Overview of Project
The purpose of this module challenge was to refactor an old stock analysis VBA code in order to gather stock information from 2017 and 2018 to determine if the stocks are worth investing any money in. We used the similar format of the old VBA code to increase its efficiency when it runs on a larger dataset. The green_stocks macro Excel sheet included two sheets, one for 2018 and 2017, with information on 12 stocks. On the charts you will find.  The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock.  We are trying to refactor code to retrice the ticker, total daily volume, and the return for each of the stocks.

Analysis
After opening the challenge starter code, I followed the guidelines to write code that would result in creating ticker array, chart headers, input box and activate the proper worksheet to make sure I am collecting data for the proper year.  I have attached the refactored VBA challenge code.
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("2017").Activate    
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
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 7).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If       
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If
            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If    
    Next i    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11        
        Worksheets("2017").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1        
    Next i    
    'Formatting
    Worksheets("2017").Activate
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

Summary
Refactoring code is an amazingly helpful tool that helps us to display clear and straightforward code that is easy to read when viewed by others that may want access it. Refactoring code helps improve software and design, faster programming, and debugging. One problem I had with refactoring this code was having applications that were too large.  One observation I noticed was as a result of refactoring was the decrease of the run time for the macro.
