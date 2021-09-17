Attribute VB_Name = "Module1"
Sub stocks()

For Each ws In Worksheets

'Add tally headers
ws.Range("I1,O1").Value = "Ticker"
ws.Range("j1").Value = "Yearly Change"
ws.Range("k1").Value = "Percent Change"
ws.Range("l1").Value = "Total Stock Volume"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Value"

'Set row looping variable
Dim row As Long

'Set counter for tally recording, starting at row 2
Dim count As Integer
count = 2

'Set volume accumulating variable, starting at 0
Dim vol As Double
vol = 0

'Set opening and closing price variable, starting at C2
Dim opening As Double
Dim closing As Double
opening = ws.Range("C2")

'Set lastrow variable and obtain the last row number
Dim lastrow As Long
lastrow = ws.Cells(Rows.count, 1).End(xlUp).row

'Start loop from row 2
For row = 2 To lastrow

'Accumulate vol for ticker at every row
vol = vol + ws.Cells(row, 7).Value

    If ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value Then
    
    'Add ticker name and volume to tally
    ws.Cells(count, 9).Value = ws.Cells(row, 1).Value
    ws.Cells(count, 12).Value = vol
    
    'Obtain closing
    closing = ws.Cells(row, 6).Value
    
    'Calculate yearly change
    ws.Cells(count, 10).Value = closing - opening
    
    'Conditional formatting of yearly change cell
    If ws.Cells(count, 10).Value > 0 Then
    ws.Cells(count, 10).Interior.ColorIndex = 4
    Else
    ws.Cells(count, 10).Interior.ColorIndex = 3
    End If
    
    'Calculate percent change
    If opening = 0 Then
        If closing = opening Then
        ws.Cells(count, 11).Value = FormatPercent(0, 2)
        ElseIf closing > opening Then
        ws.Cells(count, 11).Value = FormatPercent(1, 2)
        Else
        ws.Cells(count, 11).Value = FormatPercent(-1, 2)
        End If
    Else
    ws.Cells(count, 11).Value = FormatPercent(((closing - opening) / opening), 2)
    End If
    
    'Go to next row in tally
    count = count + 1
    
    'Reset volume as 0 for next ticker
    vol = 0
    
    'Set opening price for next ticker
    opening = ws.Cells(row + 1, 3).Value
    
    End If
    
Next row

'Set lastrow variable and obtain the last row number of second tally
Dim lastrow2 As Long
lastrow2 = ws.Cells(Rows.count, 9).End(xlUp).row

For row = 2 To lastrow2

    'Update greatest increase percent if it's greater than the existing one
    If ws.Cells(row, 11).Value > ws.Range("P2").Value Then
    ws.Range("P2").Value = FormatPercent((ws.Cells(row, 11).Value), 2)
    ws.Range("o2").Value = ws.Cells(row, 9).Value
    End If
    
    'Update greatest decrease percent if it's greater than the existing one
    If ws.Cells(row, 11).Value < ws.Range("P3").Value Then
    ws.Range("P3").Value = FormatPercent((ws.Cells(row, 11).Value), 2)
    ws.Range("o3").Value = ws.Cells(row, 9).Value
    End If
    
    'Update greatest total volume if it's greater than the existing one
    If ws.Cells(row, 12).Value > ws.Range("P4").Value Then
    ws.Range("P4").Value = ws.Cells(row, 12).Value
    ws.Range("o4").Value = ws.Cells(row, 9).Value
    End If
    
Next row

Next ws

End Sub


