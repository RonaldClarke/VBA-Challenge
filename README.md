# VBA-Challenge
Sub StockData():
Dim TickerNum As String
Dim YearChange As Double
Dim PercentChange As Double
Dim TotalVolume As Double
Dim StartPrice As Double
Dim EndPrice As Double
Dim ChartRow As Integer
Dim lastrow As Long
Dim ws As Worksheet
Dim w As Integer
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestTotal As Double
Dim TickerIncrease As String
Dim TickerDecrease As String
Dim TickerTotal As String
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    ChartRow = 2
    TotalVolume = 0
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotal = 0
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastrow
            If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
                TotalVolume = TotalVolume + Cells(i, 7).Value
            End If
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                StartPrice = Cells(i, 3).Value
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                TickerNum = Cells(i, 1).Value
                EndPrice = Cells(i, 6).Value
                TotalVolume = TotalVolume + Cells(i, 7).Value
                YearChange = EndPrice - StartPrice
                Range("I" & ChartRow).Value = TickerNum
                Range("J" & ChartRow).Value = YearChange
                Range("L" & ChartRow).Value = TotalVolume
                    If YearChange > 0 Then
                        Range("J" & ChartRow).Interior.ColorIndex = 4
                    ElseIf YearChange < 0 Then
                        Range("J" & ChartRow).Interior.ColorIndex = 3
                    End If
                    If 0 <> StartPrice Then
                        PercentChange = YearChange / StartPrice
                        Range("K" & ChartRow).Value = FormatPercent(PercentChange)
                    End If
                ChartRow = ChartRow + 1
                TotalVolume = 0
            End If
        Next i
        For i = 2 To lastrow
            If Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = Cells(i, 11).Value
                TickerIncrease = Cells(i, 9).Value
            ElseIf Cells(i, 11).Value < GreatestDecrease Then
                GreatestDecrease = Cells(i, 11).Value
                TickerDecrease = Cells(i, 9).Value
            ElseIf Cells(i, 12).Value > GreatestTotal Then
                GreatestTotal = Cells(i, 12).Value
                TickerTotal = Cells(i, 9).Value
            End If
            Next i
        Cells(2, 16).Value = TickerIncrease
        Cells(2, 17).Value = FormatPercent(GreatestIncrease)
        Cells(3, 16).Value = TickerDecrease
        Cells(3, 17).Value = FormatPercent(GreatestDecrease)
        Cells(4, 16).Value = TickerTotal
        Cells(4, 17).Value = GreatestTotal
Next ws
End Sub
