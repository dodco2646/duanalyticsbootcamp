Sub countticker() 'This code works!!!!!
Dim ws As Worksheet
Dim v As Variant
Dim w As Variant
Dim GIV As Variant
Dim GIN As Integer
Dim GDV As Variant
Dim GDN As Integer
Dim GVV As Variant
Dim GVN As Integer

For Each ws In Worksheets
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

v = 0
w = 0

lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow

If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value And ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1) Then
    w = ws.Cells(i, 3).Value
    v = v + ws.Cells(i, 7).Value
    
        ElseIf ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        v = v + ws.Cells(i, 7).Value
    
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value And w = 0 Then
            ws.Cells(i, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(i, 12).Value = v + ws.Cells(i, 7).Value
            ws.Cells(i, 10).Value = ws.Cells(i, 6) - w
            ws.Cells(i, 11).Value = w
            
            Else:
            ws.Cells(i, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(i, 12).Value = v + ws.Cells(i, 7).Value
            ws.Cells(i, 10).Value = ws.Cells(i, 6) - w
            ws.Cells(i, 11).Value = ws.Cells(i, 6) / w - 1
            w = 0
            v = 0
 
 End If
 
 If ws.Cells(i, 11).Value > 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, 11).Value < 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 3
    End If
    

Next i
    ws.Range("I2:L" & lastRow).Sort key1:=ws.Range("I2")
    ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"
        
 lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

GIV = 0
GIN = 0

For i = 2 To lastrow2

    If ws.Cells(i, 11).Value > GIV Then
    GIV = ws.Cells(i, 11).Value
    GIN = i
    Else
    GIV = GIV
    GIN = GIN
    
End If

Next i

ws.Cells(2, 15).Value = ws.Cells(GIN, 9)
ws.Cells(2, 16).Value = GIV
     
GDV = 0
GDN = 0

For i = 2 To lastrow2

    If ws.Cells(i, 11).Value < GDV Then
    GDV = ws.Cells(i, 11).Value
    GDN = i
    Else
    GDV = GDV
    GDN = GDN
    
End If

Next i

ws.Cells(3, 15).Value = ws.Cells(GDN, 9)
ws.Cells(3, 16).Value = GDV
     
GVV = 0
GVN = 0

For i = 2 To lastrow2

    If ws.Cells(i, 12).Value > GVV Then
    GVV = ws.Cells(i, 12).Value
    GVN = i
    Else
    GVV = GVV
    GVN = GVN
    
End If

Next i

ws.Cells(4, 15).Value = ws.Cells(GVN, 9)
ws.Cells(4, 16).Value = GVV

ws.Range("P2:P3").NumberFormat = "0.00%"

Next

End Sub
