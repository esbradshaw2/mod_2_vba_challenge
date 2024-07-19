Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Create headers for new columns
        ws.Cells(1, 6).Value = "Ticker Symbol"
        ws.Cells(1, 7).Value = "Total Stock Volume"
        ws.Cells(1, 8).Value = "Quarterly Change"
        ws.Cells(1, 9).Value = "Percent Change"
        
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            volume = ws.Cells(i, 7).Value
            openPrice = ws.Cells(i, 3).Value
            closePrice = ws.Cells(i, 6).Value
            
            ' Calculate quarterly change and percent change
            quarterlyChange = closePrice - openPrice
            percentChange = (closePrice - openPrice) / openPrice * 100
            
            ' Output data to the corresponding columns
            ws.Cells(i, 6).Value = ticker
            ws.Cells(i, 7).Value = volume
            ws.Cells(i, 8).Value = quarterlyChange
            ws.Cells(i, 9).Value = percentChange
            
            ' Check for greatest increase, decrease, and volume
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If
            
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            
            If volume > greatestVolume Then
                greatestVolume = volume
                greatestVolumeTicker = ticker
            End If
        Next i
        
        ' Apply conditional formatting
        ws.Range("I2:I" & lastRow).FormatConditions.Delete
        ws.Range("I2:I" & lastRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        ws.Range("I2:I" & lastRow).FormatConditions(1).Interior.Color = RGB(0, 255, 0)
        ws.Range("I2:I" & lastRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        ws.Range("I2:I" & lastRow).FormatConditions(2).Interior.Color = RGB(255, 0, 0)
        
        ws.Range("J2:J" & lastRow).FormatConditions.Delete
        ws.Range("J2:J" & lastRow).FormatConditions.AddColorScale ColorScaleType:=3
    Next ws
    
    ' Output the stocks with the greatest increase, decrease, and volume
    MsgBox "Greatest % Increase: " & greatestIncreaseTicker & " (" & greatestIncrease & "%)" & vbCrLf & _
           "Greatest % Decrease: " & greatestDecreaseTicker & " (" & greatestDecrease & "%)" & vbCrLf & _
           "Greatest Total Volume: " & greatestVolumeTicker & " (" & greatestVolume & ")"
End Sub

