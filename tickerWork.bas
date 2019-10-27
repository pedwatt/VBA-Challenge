Attribute VB_Name = "Module1"
Sub Ticker()
Dim i As Double
Dim j As Double
Dim tick(0 To 1000000) As String
Dim closing(0 To 1000000) As Double
Dim opening(0 To 1000000) As Double
Dim rowcount As Double
Dim tickSym As String
Dim tickStart As Double
Dim tickEnd As Double
Dim startValue As Double
Dim endValue As Double
Dim vol As Double
Dim tickerCount As Integer
Dim positiveGain As FormatCondition
Dim negativeGain As FormatCondition
Dim ws As Worksheet

For Each ws In Worksheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Volume"
    tickStart = 2
    tickerCount = 2
    tick(2) = ws.Range("A2").Value
    opening(2) = ws.Range("C2").Value
    closing(2) = ws.Range("F2").Value
    rowcount = Application.CountA(ws.Range("A:A"))
    For i = 3 To rowcount + 1
        tick(i) = ws.Range("A" & i).Value
        opening(i) = ws.Range("C" & i).Value
        closing(i) = ws.Range("F" & i).Value
        
        If ws.Range("a" & i).Value = ws.Range("a" & i - 1).Value Then
        
        ElseIf ws.Range("a" & i).Value <> ws.Range("a" & i - 1).Value Then
            tickEnd = i - 1
            For j = tickStart To tickEnd
                vol = vol + ws.Range("G" & j).Value
            Next j
            
            tickSym = tick(i - 1)
            startValue = opening(tickStart)
            endValue = closing(tickEnd)
            
            ws.Range("I" & tickerCount).Value = tickSym
            ws.Range("J" & tickerCount).Value = endValue - startValue
            If startValue = 0 Then
                ws.Range("K" & tickerCount).Value = 0
            Else
                ws.Range("K" & tickerCount).Value = endValue / startValue - 1
            End If
            ws.Range("L" & tickerCount).Value = vol
            tickerCount = tickerCount + 1
            tickStart = i
            vol = 0
        End If
           
    Next i
    ws.Range("K2:K" & tickerCount).NumberFormat = "0.00%"
    Set positiveGain = ws.Range("K2:K" & tickerCount).FormatConditions.Add(xlCellValue, xlGreater, "=0")
    Set negativeGain = ws.Range("K2:K" & tickerCount).FormatConditions.Add(xlCellValue, xlLess, "=0")
    With positiveGain
        .Interior.Color = vbGreen
    End With
    With negativeGain
        .Interior.Color = vbRed
    End With
Next

End Sub


