Attribute VB_Name = "Module1"
Sub testCode():

Dim lastRow As Long
Dim tickerLabel As String
Dim initPrice As Double
Dim volumeTot As Double
Dim closingPrice As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim pos As Integer

lastRow = Cells(Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

tickerLabel = Cells(2, 1).Value
initPrice = Cells(2, 3).Value
volumeTot = 0
pos = 1

For i = 2 To lastRow + 1

    If Cells(i, 1).Value = tickerLabel Then
        'volumeTot = 0
        volumeTot = volumeTot + Cells(i, 7).Value
    
    ElseIf Cells(i, 1).Value <> tickerLabel Then
        'MsgBox (Cells(i - 1, 6).Value)
        closingPrice = Cells(i - 1, 6).Value
        'MsgBox (closingPrice)
        'MsgBox (initPrice)
        yearlyChange = closingPrice - initPrice
        'MsgBox (yearlyChange)
        percentChange = yearlyChange / (initPrice + 0.0001)
        pos = pos + 1
        
        'fill ticker
        Cells(pos, 10).Value = tickerLabel
        'fill yearly change
        Cells(pos, 11).Value = yearlyChange
        If Cells(pos, 11).Value < 0 Then
            Cells(pos, 11).Interior.ColorIndex = 3
        Else
            Cells(pos, 11).Interior.ColorIndex = 4
        End If
    
    'fill percent
    Cells(pos, 12).Value = percentChange
    'fill volumeTot
    Cells(pos, 13).Value = volumeTot
    
    tickerLabel = Cells(i, 1).Value
    initPrice = Cells(i, 3).Value
    volumeTot = 0
    End If

Next i

End Sub
