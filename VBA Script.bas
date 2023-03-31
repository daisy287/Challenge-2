Attribute VB_Name = "Module1"
Sub VBASript()
    
    Dim ws As Worksheet
    Dim Last As Long

    For Each ws In ThisWorkbook.Worksheets
        With ws

            Dim ticker As String
            Dim openingDay As Double
            Dim yearlyChange As Double
            Dim endDay As Double
            Dim lastRow As Long
            Dim i As Long
            Dim newRowCounter As Integer
            Dim percentChange As Double
            Dim tSV As Double

            Dim greatestDec As Double
            Dim greatestInc As Double
            Dim greatestVol As Double

            greatestDec = 0
            greatestInc = 0
            greatestVol = 0
            tSV = 0
            newRowCounter = 3
            ticker = .Cells(2, "A").Value
            lastRow = .Cells(Rows.Count, "A").End(xlUp).Row
            
            .Cells(2, "I").Value = ticker

            openingDay = .Cells(2, "C").Value
            
            .Cells(1, "I").Value = "Ticker"
            .Cells(1, "J").Value = "Yearly Change"
            .Cells(1, "K").Value = "Percent Change"
            .Cells(1, "L").Value = "Total Stock Volume"
            .Cells(1, "P").Value = "Ticker"
            .Cells(1, "Q").Value = "Value"

            .Cells(2, "O").Value = "Greatest % Increase"
            .Cells(3, "O").Value = "Greatest % Decrease"
            .Cells(4, "O").Value = "Greatest Total Volume"


            For i = 2 To lastRow

                If ticker <> .Cells(i, "A").Value Then
                    
                    endDay = .Cells(i - 1, "F").Value
                    yearlyChange = endDay - openingDay
                    .Cells(newRowCounter - 1, "J").Value = yearlyChange

                    If yearlyChange > 0 Then
                        .Cells(newRowCounter - 1, "J").Interior.Color = RGB(26, 255, 0)
                    Else
                        .Cells(newRowCounter - 1, "J").Interior.Color = RGB(255, 0, 0)
                    End If
                    
                    percentChange = (yearlyChange / openingDay)
                    .Cells(newRowCounter - 1, "K").Value = FormatPercent(percentChange)

                    If greatestInc < percentChange Then
                        greatestInc = percentChange
                        .Cells(2, "Q").Value = FormatPercent(greatestInc)
                        .Cells(2, "P").Value = ticker
                    End If

                    If greatestDec > percentChange Then
                        greatestDec = percentChange
                        .Cells(3, "Q").Value = FormatPercent(greatestDec)
                        .Cells(3, "P").Value = ticker
                    End If

                    openingDay = .Cells(i, "C").Value

                    If greatestVol < tSV Then
                        greatestVol = tSV
                        .Cells(4, "Q").Value = greatestVol
                        .Cells(4, "P").Value = ticker
                    End If

                    
                    .Cells(newRowCounter - 1, "L").Value = tSV
                    tSV = 0

                    ticker = .Cells(i, "A").Value
                    .Cells(newRowCounter, "I").Value = ticker
                    newRowCounter = newRowCounter + 1
                End If
                
                tSV = tSV + .Cells(i, "G").Value

                If i + 1 = lastRow Then
                    endDay = .Cells(i - 1, "F").Value
                    yearlyChange = endDay - openingDay

                    percentChange = (yearlyChange / openingDay)
                    .Cells(newRowCounter - 1, "K").Value = FormatPercent(percentChange)


                    If greatestInc < percentChange Then
                        greatestInc = percentChange
                        .Cells(2, "Q").Value = FormatPercent(greatestInc)
                        .Cells(2, "P").Value = ticker
                    End If

                    If greatestDec > percentChange Then
                        greatestDec = percentChange
                        .Cells(3, "Q").Value = FormatPercent(greatestDec)
                        .Cells(3, "P").Value = ticker
                    End If

                    If greatestVol < tSV Then
                        greatestVol = tSV
                        .Cells(4, "Q").Value = greatestVol
                        .Cells(4, "P").Value = ticker
                    End If

                    .Cells(newRowCounter - 1, "L").Value = tSV

                    .Cells(newRowCounter - 1, "J").Value = yearlyChange

                    If yearlyChange > 0 Then
                        .Cells(newRowCounter - 1, "J").Interior.Color = RGB(26, 255, 0)
                    Else
                        .Cells(newRowCounter - 1, "J").Interior.Color = RGB(255, 0, 0)
                    End If
                End If
            Next i
        End With
    Next ws
End Sub
