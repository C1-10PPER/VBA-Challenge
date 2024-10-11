Attribute VB_Name = "Module3MarketMovers"
Sub MarketMovers()

Dim ws As Worksheet
Dim i As Long

    For Each ws In Worksheets

        Dim lastrow As Long
        Dim maxValue As Double
        Dim maxTicker As String
        Dim minValue As Double
        Dim minTicker As String
        Dim MaxVolume As Double
        Dim MaxVTicker As String
        
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        WorksheetName = ws.Name
         
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"
        
        
        maxTicker = ws.Cells(2, "I").Value
        maxValue = ws.Cells(2, "K").Value
        minTicker = ws.Cells(2, "I").Value
        minValue = ws.Cells(2, "K").Value
        MaxVTicker = ws.Cells(2, "I").Value
        MaxVolume = ws.Cells(2, "L").Value
        
        
            For i = 2 To lastrow
                    If ws.Cells(i, "K").Value > maxValue Then
                    maxTicker = ws.Cells(i, "I").Value
                    maxValue = ws.Cells(i, "K").Value
                    
                End If
            Next i
        
            For i = 2 To lastrow
                    If ws.Cells(i, "K").Value < minValue Then
                    minTicker = ws.Cells(i, "I").Value
                    minValue = ws.Cells(i, "K").Value
                    
                End If
            Next i
            
           For i = 2 To lastrow
                    If ws.Cells(i, "L").Value > MaxVolume Then
                    MaxVTicker = ws.Cells(i, "I").Value
                    MaxVolume = ws.Cells(i, "L").Value
                End If
            Next i
            ws.Cells(2, "P").Value = maxTicker
            ws.Cells(2, "Q").Value = maxValue
            ws.Cells(2, "Q").NumberFormat = "0.00%"
            ws.Cells(3, "P").Value = minTicker
            ws.Cells(3, "Q").Value = minValue
            ws.Cells(3, "Q").NumberFormat = "0.00%"
            ws.Cells(4, "P").Value = MaxVTicker
            ws.Cells(4, "Q").Value = MaxVolume
            
    Next ws
End Sub


