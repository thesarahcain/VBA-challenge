Sub CalculateYearlyChanges()

    Dim Ws As Worksheet
    Dim lastRow As Long
    Dim tickerColumn As Range
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

    ' Loop through all sheets
    For Each Ws In ThisWorkbook.Sheets
        ' Find the last row in column A
        lastRow = Ws.Cells(Ws.Rows.Count, "A").End(xlUp).Row
        ' Initialize variables for each sheet
        openingPrice = 0
        totalVolume = 0
        outputRow = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        greatestIncreaseTicker = ""
        greatestDecreaseTicker = ""
        greatestVolumeTicker = ""

        ' Label the columns
        Ws.Cells(1, "I").Value = "Ticker"
        Ws.Cells(1, "J").Value = "Yearly Change"
        Ws.Cells(1, "K").Value = "Percent Change"
        Ws.Cells(1, "L").Value = "Total Stock Volume"

        ' Loop through the data
        For i = 2 To lastRow
            ' Check if the ticker symbol changes
            If Ws.Cells(i, "A").Value <> Ws.Cells(i - 1, "A").Value Then
                ' Output results for the previous ticker
                Ws.Cells(outputRow, "I").Value = Ws.Cells(i - 1, "A").Value   'Ticker Symbol

                If openingPrice <> 0 Then
                    ' Calculate yearly change and percentage change
                    yearlyChange = closingPrice - openingPrice
                    percentageChange = yearlyChange / openingPrice * 100

                    ' Output results
                    Ws.Cells(outputRow, "J").Value = yearlyChange  'Yearly Change
                    Ws.Cells(outputRow, "K").Value = percentageChange & "%"  'Percentage Change
                    Ws.Cells(outputRow, "L").Value = totalVolume 'Total Stock Volume

                    ' Color code the "yearly Change" column
                    If yearlyChange > 0 Then
                        Ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0) 'Green for positive change
                    ElseIf yearlyChange < 0 Then
                        Ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0)  ' Red for negative change
                    End If

                    ' Check for greatest % increase, % decrease, and total volume
                    If percentageChange > greatestIncrease Then
                        greatestIncrease = percentageChange
                        greatestIncreaseTicker = Ws.Cells(outputRow, 9).Value
                    End If

                    If percentageChange < greatestDecrease Then
                        greatestDecrease = percentageChange
                        greatestDecreaseTicker = Ws.Cells(outputRow, 9).Value
                    End If

                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        greatestVolumeTicker = Ws.Cells(outputRow, 9).Value
                    End If

                    ' Reset variables for the next ticker
                    openingPrice = 0
                    totalVolume = 0
                    outputRow = outputRow + 1
                End If
            End If

            ' Update variables
            If openingPrice = 0 Then
                openingPrice = Ws.Cells(i, "C").Value  'Opening Price
            End If

            closingPrice = Ws.Cells(i, "F").Value  ' Closing Price
            totalVolume = totalVolume + Ws.Cells(i, "G").Value 'Volume
        Next i

        ' Output the stocks with the greatest % increase, % decrease, and total volume
        Ws.Cells(2, 14).Value = "Greatest % Increase"
        Ws.Cells(3, 14).Value = "Greatest % Decrease"
        Ws.Cells(4, 14).Value = "Greatest Total Volume"
        Ws.Cells(2, 15).Value = greatestIncreaseTicker
        Ws.Cells(3, 15).Value = greatestDecreaseTicker
        Ws.Cells(4, 15).Value = greatestVolumeTicker
        Ws.Cells(2, 16).Value = greatestIncrease
        Ws.Cells(3, 16).Value = greatestDecrease
        Ws.Cells(4, 16).Value = greatestVolume
    Next Ws

End Sub
