Dim ResultRow As Long


Sub runScanner()
    ' This function will scan through the current worksheet and exract the data to the results page.
    Dim curRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim symbol As String
    Dim change As Double
    Dim yDate As String
    Dim fStop As Boolean
    Dim ResultRow As Long

    For Each Sheet In Sheets
        Sheet.Select
    
        'initialize for the page.
        openPrice = 0
        curRow = 2
        fStop = False
        ResultRow = 2
    
        'continue until reaching the end of the rows.
        While Not fStop
            'get the ticker symbol if the symbol is "" then this is the start of a new ticker symbol
            If symbol = "" Then
                symbol = Cells(curRow, 1).Value
                openPrice = Cells(curRow, 3).Value 'Open price from the first date
                volume = 0 'Clear the volume
            End If
    
            'sum the volume for each line in the data
            volume = volume + Cells(curRow, 7).Value
            
    
            'Check the next cell to see if it is blank if yes then stop the scan
            If Cells(curRow + 1, 1).Value = "" Then
                fStop = True
            End If
    
            ' if the next cell is not the same symbol as the current cell we are done with this ticker
            If Cells(curRow + 1, 1) <> symbol Then
                yDate = Left(Cells(curRow, 2).Value, 4) 'get the year from the date field
                'finished the symbol report results.
                closePrice = Cells(curRow, 6).Value
                change = closePrice - openPrice
                'if the open price was zero the %change is inf w
                If openPrice <> 0 Then
                    Percent = change / openPrice
                Else
                    Percent = 0
                End If
                
                startCol = 10 ' start in column J
                'Record the data for this ticker symbol into the summary table
                Cells(ResultRow, startCol + 0) = symbol
                Cells(ResultRow, startCol + 1) = yDate
                Cells(ResultRow, startCol + 2) = openPrice
                Cells(ResultRow, startCol + 3) = closePrice
                Cells(ResultRow, startCol + 4) = change
                Cells(ResultRow, startCol + 5) = Percent
                Cells(ResultRow, startCol + 6) = volume
                ResultRow = ResultRow + 1
    
                symbol = "" 'done with this symbol so clear it
            End If
    
            curRow = curRow + 1
            DoEvents
        Wend
        Call sheetFormatter
        Call findSummary
    Next Sheet
End Sub

Sub sheetFormatter()
    'This function is used to format the summary table
    
    'add header Row
    Range("J1:P1") = VBA.Array("Ticker", "Year", "Open", "Close", "Change", "Percent Change", "Volume")
    Range("J1:P1").Font.Bold = True
    Range("J1:P1").Interior.ColorIndex = 1
    Range("J1:P1").Font.ColorIndex = 2
    
    'format the Percentage column
    With Range("O2:O" & Rows.Count)
        .NumberFormat = "% 0.0"
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(1).Interior.Color = RGB(230, 50, 50) 'RED
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(2).Interior.Color = RGB(50, 230, 50) 'GREEN
    End With

End Sub


Sub findSummary()
    'find the greatest values

    Dim rowCount As Long
    Dim max As Double
    Dim min As Double
    Dim maxVolume As Double
    Dim maxTicker As String
    Dim mimTicker As String
    Dim volumeTicker As String

    rowCount = Cells(Rows.Count, 1).End(xlUp).Row


    'initialize the check values
    max = 0
    min = 0
    maxVolume = 0

    maxTicker = ""
    minTicker = ""
    volumeTicker = ""

    'Step through each row in the table
    For i = 2 To rowCount
        If Cells(i, 15) > max Then
            max = Cells(i, 15)
            maxTicker = Cells(i, 10)
        End If
        If Cells(i, 15) < min Then
            min = Cells(i, 15)
            minTicker = Cells(i, 10)
        End If
        If Cells(i, 16) > maxVolume Then
            maxVolume = Cells(i, 16)
            volumeTicker = Cells(i, 10)
        End If
    Next i
    
    Range("S2").Value = "Greatest % Increase"
    Range("S3").Value = "Greatest % Decrease"
    Range("S4").Value = "Greatest Total Volume"

    Range("T1").Value = "Ticker"
    Range("T2").Value = maxTicker
    Range("T3").Value = minTicker
    Range("T4").Value = volumeTicker
    
    Range("U1").Value = "Value"
    Range("U2").Value = max
    Range("U2").NumberFormat = "% #0.0"
    Range("U3").Value = min
    Range("U3").NumberFormat = "% #0.0"
    Range("U4").Value = maxVolume

End Sub




