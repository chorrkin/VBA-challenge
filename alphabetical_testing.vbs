Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim outputRow As Long
    
    ' Variables for summary calculations
    Dim maxPercentChange As Double
    Dim minPercentChange As Double
    Dim maxVolume As Double
    Dim maxPercentTicker As String
    Dim minPercentTicker As String
    Dim maxVolumeTicker As String
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Determine the last row with data
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Set up headers for output columns
        ws.Cells(1, 9).Value = "Ticker"            ' Column I
        ws.Cells(1, 10).Value = "Quarterly Change" ' Column J
        ws.Cells(1, 11).Value = "Percent Change"   ' Column K
        ws.Cells(1, 12).Value = "Total Volume"     ' Column L
        
        ' Initialize summary values
        maxPercentChange = 0
        minPercentChange = 0
        maxVolume = 0
        maxPercentTicker = ""
        minPercentTicker = ""
        maxVolumeTicker = ""
        
        ' Start output from the second row in the output columns
        outputRow = 2
        
        ' Loop through each row in the worksheet
        For i = 2 To lastRow ' Assuming headers are in the first row
            ticker = ws.Cells(i, 1).Value ' Ticker is in Column A
            openPrice = ws.Cells(i, 3).Value ' Open price is in Column C
            closePrice = ws.Cells(i, 6).Value ' Close price is in Column F
            volume = ws.Cells(i, 7).Value ' Volume is in Column G
            
            ' Calculate quarterly change and percentage change
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentChange = (quarterlyChange / openPrice) * 100
            Else
                percentChange = 0 ' Handle division by zero
            End If
            
            ' Output the values to columns I, J, K, and L
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = quarterlyChange
            ws.Cells(outputRow, 11).Value = percentChange
            ws.Cells(outputRow, 12).Value = volume
            
            ' Check for greatest percent increase
            If percentChange > maxPercentChange Then
                maxPercentChange = percentChange
                maxPercentTicker = ticker
            End If
            
            ' Check for greatest percent decrease
            If percentChange < minPercentChange Then
                minPercentChange = percentChange
                minPercentTicker = ticker
            End If
            
            ' Check for greatest total volume
            If volume > maxVolume Then
                maxVolume = volume
                maxVolumeTicker = ticker
            End If
            
            outputRow = outputRow + 1 ' Move to the next row for output
        Next i
        
        ' Apply conditional formatting to the Quarterly Change column (Column J)
        With ws.Range("J2:J" & outputRow - 1)
            .FormatConditions.Delete ' Clear any existing conditional formatting
            
            ' Green for positive change
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0)
            
            ' Red for negative change
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0)
        End With
        
        ' Format Percent Change column as percentage with two decimal places
        ws.Range("K2:K" & outputRow - 1).NumberFormat = "0.00%"
        
        ' Output summary results starting in Column O
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = maxPercentTicker
        ws.Cells(2, 17).Value = maxPercentChange & "%"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = minPercentTicker
        ws.Cells(3, 17).Value = minPercentChange & "%"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = maxVolumeTicker
        ws.Cells(4, 17).Value = maxVolume
    Next ws
End Sub

