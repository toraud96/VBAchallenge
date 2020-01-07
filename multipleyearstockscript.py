Sub stocks()
'declaring the variables'
Dim Ticker As String
Dim TotalVolume As Double
TotalVolume = 0
'declaring ws so the sub routine can run through the whole workbook'
Dim ws As Worksheet

For Each ws In Worksheets
j = 0
Dim start As Double
start = 2
Dim openingValue As Double
openingValue = ws.Cells(2, 3).Value
Dim change As Double
change = 0
Dim dailyChange As Double
dailyChange = 0
Dim percentChange As Double
Dim LastRow As Long
LastRow = Cells(Rows.Count, "A").End(xlUp).Row

'Naming the headers for the summary table'
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
'declaring a variable to start off the summary table'
Dim SummTableRow As Integer
SummTableRow = 2
'for loop to find total volume, yearly change, and percent change'
For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        nextValue = ws.Cells(i + 1, 3).Value
        Ticker = ws.Cells(i, 1).Value
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        'if statement to skip zeros in the data'
        If TotalVolume = 0 Then
            ws.Range("J" & 2 + j).Value = 0
            ws.Range("I" & 2 + j).Value = Ticker
            ws.Range("K" & 2 + j).Value = 0
            ws.Range("L" & 2 + j).Value = 0
            Else
            If ws.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                openingValue = ws.Cells(find_value, 3).Value
                                Exit For
            End If
                        Next find_value
                    End If
                    'finding yearly change and percent change'
                    change = (ws.Cells(i, 6).Value - openingValue)
                    percentChange = Round(((change) / (openingValue) * 100), 2)
                    'assigning the range for the results in the summary table'
                    ws.Range("I" & SummTableRow).Value = Ticker
                    ws.Range("L" & SummTableRow).Value = TotalVolume
                    ws.Range("J" & SummTableRow).Value = change
                    ws.Range("K" & SummTableRow).Value = percentChange
                    'reset'
                    start = i + 1
                    openingValue = nextValue
        End If
        'moving to the next row on the summary table'
        SummTableRow = SummTableRow + 1
        'reset'
        TotalVolume = 0
        j = j + 1
        change = 0
    Else
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    End If
Next i
'loop to color format the cells in row J'
    For i = 2 To LastRow
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i
Next ws
End Sub