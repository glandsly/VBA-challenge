Sub stock_analyzer()
    Dim count As Integer
    Dim openPrice, closePrice As Double
    Dim LastRow As Long
    Dim totalVolume As Double
    Dim percentChange, greatestIncrease, greatestDecrease, greatestTotalVol As Double
    Dim greatestIncreaseTicker, greatestDecreaseTicker, greatestVolTicker As String

    count = 2
    Dim minDate, maxDate As Long
    minDate = 0
    openPrice = 0
    closePrice = 0
    totalVolume = 0
    percentChange = 0
    greatestIncrease = 0
    greatestDecrease = 0
    greatestTotalVol = 0

    LastRow = Cells(Rows.count, "A").End(xlUp).Row

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"

    For i = 2 To LastRow
        totalVolume = totalVolume + Cells(i, 7).Value

        If minDate = 0 Then
            minDate = Cells(i, 2).Value
            openPrice = Cells(i, 3).Value
        ElseIf Cells(i, 2).Value < minDate Then
            minDate = Cells(i, 2).Value
            openPrice = Cells(i, 3).Value
        End If

        If maxDate = 0 Then
            maxDate = Cells(i, 2).Value
            closePrice = Cells(i, 6).Value
        ElseIf Cells(i, 2).Value > maxDate Then
            maxDate = Cells(i, 2).Value
            closePrice = Cells(i, 6).Value
        End If

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Cells(count, 9).Value = Cells(i, 1).Value
            Cells(count, 10).Value = closePrice - openPrice

            If closePrice - openPrice > 0 Then
                Cells(count, 10).Interior.ColorIndex = 4
            ElseIf closePrice - openPrice < 0 Then
                Cells(count, 10).Interior.ColorIndex = 3
            End If

            If openPrice > 0 Then
                percentChange = (closePrice - openPrice) / openPrice
            Else
                percentChange = 0
            End If

            Cells(count, 11).Value = percentChange
            Cells(count, 12).Value = totalVolume

            If percentChange > greatestIncrease Then
                greatestIncreaseTicker = Cells(i, 1).Value
                greatestIncrease = percentChange

                Cells(2, 15).Value = greatestIncreaseTicker
                Cells(2, 16).Value = greatestIncrease
            End If

            If percentChange < greatestDecrease Then
                greatestDecreaseTicker = Cells(i, 1).Value
                greatestDecrease = percentChange

                Cells(3, 15).Value = greatestDecreaseTicker
                Cells(3, 16).Value = greatestDecrease
            End If

            If totalVolume > greatestTotalVol Then
                greatestVolTicker = Cells(i, 1).Value
                greatestTotalVol = totalVolume

                Cells(4, 15).Value = greatestVolTicker
                Cells(4, 16).Value = greatestTotalVol
            End If

            closePrice = 0
            openPrice = 0
            minDate = 0
            maxDate = 0
            totalVolume = 0
            count = count + 1
        End If
    Next i

    Range("K:K").NumberFormat = "0.00%"
    Range("P2:P3").NumberFormat = "0.00%"

End Sub

Sub LoopThroughSheets()
    Dim ws As Worksheet
    Application.ScreenUpdating = False

    For Each ws In Worksheets
        ws.Select
        Call stock_analyzer

    Next

    Application.ScreenUpdating = True
End Sub
