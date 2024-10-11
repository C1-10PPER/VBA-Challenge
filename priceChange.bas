Attribute VB_Name = "Module1priceChange"
Sub priceChange()
Dim i As Long
Dim ws As Worksheet

    For Each ws In Worksheets
    
    Dim lastrow As Long
    Dim VolumeTotal As Double
    Dim ticker As String
    Dim currentTicker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim SummaryTableRow As Long
    Dim firstRow As Long
    Dim WorksheetName As String
    

    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    WorksheetName = ws.Name
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    SummaryTableRow = 2
    
    firstRow = 2
    ticker = ws.Cells(firstRow, 1).Value
    
    VolumeTotal = 0

 
        For i = 2 To lastrow
        
                currentTicker = ws.Cells(i, 1).Value
                VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
                
                If i = firstRow Then
                    openingPrice = ws.Cells(i, 3).Value
                End If
                If i = lastrow Or ws.Cells(i + 1, 1).Value <> currentTicker Then
                    openingPrice = ws.Cells(firstRow, 3).Value
                    closingPrice = ws.Cells(i, 6).Value
                    quarterlyChange = closingPrice - openingPrice
                        If openingPrice <> 0 Then
                            percentChange = (quarterlyChange / openingPrice)
                        Else
                            percentChange = 0
                        End If
                        ws.Cells(SummaryTableRow, 9).Value = ticker
                        ws.Cells(SummaryTableRow, 10).Value = quarterlyChange
                        ws.Cells(SummaryTableRow, 11).Value = percentChange
                        ws.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
                        ws.Cells(SummaryTableRow, 12).Value = VolumeTotal
                            
                        SummaryTableRow = SummaryTableRow + 1
                        If i < lastrow Then
                            ticker = ws.Cells(i + 1, 1).Value
                            firstRow = i + 1
                        End If
                        VolumeTotal = 0
        
            End If
        Next i
    Next ws
End Sub


