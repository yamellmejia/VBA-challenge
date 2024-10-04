Sub SummarizeQuarterlyStockData()

    ' Define the sheets representing the quarters
    Dim sheetsArray As Variant
    sheetsArray = Array("Q1", "Q2", "Q3", "Q4")
    
    ' Define variables for looping through sheets and rows
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim i As Long
    Dim j As Long
    Dim summaryStartRow As Long
    
    ' Variables to track greatest values
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim currentPercentChange As Double
    Dim currentTotalVolume As Double
    
    ' Loop through each quarter sheet
    For Each sheetName In sheetsArray
        Set ws = ThisWorkbook.Sheets("Q1")
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Add headers for the summary
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Price Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' Add headers for greatest % increase, decrease, and total volume in columns
        ws.Cells(1, 15).Value = ""
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        summaryStartRow = 2
        totalVolume = 0
        
        ' Initialize variables for tracking greatest values
        greatestIncrease = -999999
        greatestDecrease = 999999
        greatestVolume = 0
        
        ' Process each ticker individually
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 3).Value
            
            ' Loop through all rows of the same ticker
            j = i
            Do While ws.Cells(j, 1).Value = ticker And j <= lastRow
                totalVolume = totalVolume + ws.Cells(j, 7).Value
                j = j + 1
            Loop
            
            ' Closing price for the last row of the current ticker
            closePrice = ws.Cells(j - 1, 6).Value
            
            ' Calculate the change and percentage change
            Dim priceChange As Double
            priceChange = closePrice - openPrice
            
            If openPrice <> 0 Then
                currentPercentChange = (priceChange / openPrice) * 100
            Else
                currentPercentChange = 0
            End If
            
            ' Output the results in the next available columns
            ws.Cells(summaryStartRow, 9).Value = ticker
            ws.Cells(summaryStartRow, 10).Value = priceChange
            ws.Cells(summaryStartRow, 11).Value = currentPercentChange
            ws.Cells(summaryStartRow, 12).Value = totalVolume
            
            ' Track the greatest % increase, % decrease, and total volume
            If currentPercentChange > greatestIncrease Then
                greatestIncrease = currentPercentChange
                greatestIncreaseTicker = ticker
            End If
            
            If currentPercentChange < greatestDecrease Then
                greatestDecrease = currentPercentChange
                greatestDecreaseTicker = ticker
            End If
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If
            
            ' Move to the next row for summary output
            summaryStartRow = summaryStartRow + 1
            
            ' Reset total volume for the next ticker
            totalVolume = 0
            
            ' Move the main loop counter to the last processed row
            i = j - 1
        Next i
        
        ' Output the greatest values in columns O, P, and Q
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncrease
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolume
        
    Next sheetName

End Sub
