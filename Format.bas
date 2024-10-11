Attribute VB_Name = "Module2Format"
Sub Format()

    Dim i As Long
    Dim lastrow As Long
    Dim ws As Worksheet
    
    For Each ws In Worksheets ' A loop to process the VBA for all worksheets in the workbook
        
            lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' To determine the last row with a stored value
        
            For i = 2 To lastrow
                    If IsEmpty(ws.Cells(i, "J").Value) Or IsEmpty(ws.Cells(i, "K").Value) Then  ' Exit Loop if there is no Value in Column J or Column K
                        Exit For
                    End If
                    
                    If ws.Cells(i, "J") >= 0 Then
                    ws.Cells(i, "J").Interior.ColorIndex = 4  ' Color cell green if positive
                    ws.Cells(i, "K").Interior.ColorIndex = 4 ' Color cell green if positive
                    Else
                    ws.Cells(i, "J").Interior.ColorIndex = 3 ' Color cell red if negative
                    ws.Cells(i, "K").Interior.ColorIndex = 3 ' Color cell red if negative
                    End If
            Next i ' Cycle through the loop
     Next ws ' Process the next worksheet
End Sub



