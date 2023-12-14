Attribute VB_Name = "Module1"
Sub Stock_Data_All_Sheets()

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim ws As Worksheet
    
    For Each ws In wb.Sheets
        Call Stock_Data(ws)
    Next ws

End Sub

Sub Stock_Data(ws As Worksheet)

    Dim greatestIncrease As Double: greatestIncrease = 0
    Dim greatestDecrease As Double: greatestDecrease = 0
    Dim greatestVolume As Double: greatestVolume = 0
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String

    
    Dim Ticker_Name As String
    Dim yearStart As Double
    Dim yearEnd As Double
    Dim percentageChange As Double
    Dim Vol_Total As Double
    Vol_Total = 0
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Headers for the summary table
    With ws
        .Cells(1, 9).Value = "Ticker"
        .Cells(1, 10).Value = "Yearly Change"
        .Cells(1, 11).Value = "Percentage Change"
        .Cells(1, 12).Value = "Total Stock Volume"
        .Cells(2, 15).Value = "Greatest % Increase"
        .Cells(3, 15).Value = "Greatest % Decrease"
        .Cells(4, 15).Value = "Greatest Total Volume"
        .Cells(1, 16).Value = "Ticker"
        .Cells(1, 17).Value = "Value"
        
    End With

    ' Initialize summary table row
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    ' Loop through all rows
    For i = 2 To lastRow

        ' If we have reached a new ticker or the last row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = lastRow Then

            ' Set the Ticker name
            Ticker_Name = ws.Cells(i, 1).Value

            ' Add to the volume total
            Vol_Total = Vol_Total + ws.Cells(i, 7).Value

            ' Print the Ticker Name in the Summary Table
            ws.Cells(Summary_Table_Row, 9).Value = Ticker_Name

            ' Print the Yearly Change
            ws.Cells(Summary_Table_Row, 10).Value = yearEnd - yearStart

            ' Calculate and print the Percentage Change
            If yearStart <> 0 Then ' Avoid division by zero
                percentageChange = (yearEnd - yearStart) / yearStart
            Else
                percentageChange = 0
            End If

            ws.Cells(Summary_Table_Row, 11).Value = percentageChange

            ' Print the Total Volume to the Summary Table
            ws.Cells(Summary_Table_Row, 12).Value = Vol_Total

            ' Check for greatest increase, decrease, and volume
            If percentageChange > greatestIncrease Then
                greatestIncrease = percentageChange
                tickerGreatestIncrease = Ticker_Name
            ElseIf percentageChange < greatestDecrease Then
                greatestDecrease = percentageChange
                tickerGreatestDecrease = Ticker_Name
            End If
            
            If Vol_Total > greatestVolume Then
                greatestVolume = Vol_Total
                tickerGreatestVolume = Ticker_Name
            End If

            ' Increment the summary table row
            Summary_Table_Row = Summary_Table_Row + 1

            ' Reset the Volume Total and yearStart for the next ticker
            Vol_Total = 0
            yearStart = 0
        Else
            ' Accumulate the volume total
            Vol_Total = Vol_Total + ws.Cells(i, 7).Value
        End If

        ' Update currentTicker
        Dim currentTicker As String
        currentTicker = ws.Cells(i, 1).Value

        If currentTicker <> Ticker_Name Or yearStart = 0 Then
            Ticker_Name = currentTicker
            yearStart = ws.Cells(i, 3).Value ' Opening prices are in column C
        End If

        ' Always update the yearEnd price to the current row's close price
        yearEnd = ws.Cells(i, 6).Value ' Closing prices are in column F

    Next i

    ' Output the results for the greatest increase, decrease, and total volume
    With ws
        .Cells(2, 16).Value = tickerGreatestIncrease
        .Cells(3, 16).Value = tickerGreatestDecrease
        .Cells(4, 16).Value = tickerGreatestVolume
        .Cells(2, 17).Value = greatestIncrease
        .Cells(3, 17).Value = greatestDecrease
        .Cells(4, 17).Value = greatestVolume
    End With

End Sub


