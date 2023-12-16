Attribute VB_Name = "Module1"
Sub CalculateStockData()
    Dim ws As Worksheet
    Dim yearSheets As Variant
    Dim sheetIndex As Integer
    Dim lastRow As Long, i As Long
    Dim startPrice As Double, endPrice As Double
    Dim totalVolume As Double, ticker As String
    Dim yearlyChange As Double, percentChange As Double
    Dim startRow As Long, outputRow As Long
    Dim maxIncrease As Double, maxDecrease As Double, maxVolume As Double
    Dim maxIncreaseTicker As String, maxDecreaseTicker As String, maxVolumeTicker As String

    ' Initialize the greatest values
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0

    ' Define the array of sheet names
    yearSheets = Array("2018", "2019", "2020")

    ' Loop through each sheet name in the array using a For loop
    For sheetIndex = LBound(yearSheets) To UBound(yearSheets)
        Set ws = ThisWorkbook.Sheets(yearSheets(sheetIndex))
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ticker = ""
        totalVolume = 0
        outputRow = 2 ' Start at row 2 for output

        ' Find the first non-empty row with data for opening price
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> "" And ticker = "" Then
                ticker = ws.Cells(i, 1).Value
                startPrice = ws.Cells(i, 3).Value
                Exit For
            End If
        Next i

        ' Assuming row 1 has headers, start the processing from the first data row
        startRow = 2

        ' Loop through all rows to process the data
        For i = startRow To lastRow
            ' Accumulate the total volume for the current ticker
            totalVolume = totalVolume + ws.Cells(i, 7).Value

            ' Check if we have reached a new ticker or the last row
            If ws.Cells(i + 1, 1).Value <> ticker Or i = lastRow Then
                ' Capture the closing price of the current ticker
                endPrice = ws.Cells(i, 6).Value

                ' Calculate yearly change and percent change
                yearlyChange = endPrice - startPrice
                If startPrice <> 0 Then
                    percentChange = (yearlyChange / startPrice) * 100
                Else
                    percentChange = 0
                End If

                ' Output the data for the current ticker
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = yearlyChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalVolume

                ' Apply color formatting based on yearly change
                If yearlyChange >= 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                Else
                    ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                End If

                ' Check for max and min percent change and max volume
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ticker
                ElseIf percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ticker
                End If

                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If

                ' Prepare for the next ticker
                outputRow = outputRow + 1

                ' If it's not the last row, set up for the next ticker
                If i <> lastRow Then
                    ticker = ws.Cells(i + 1, 1).Value
                    startPrice = ws.Cells(i + 1, 3).Value
                    totalVolume = 0
                End If
            End If
        Next i

        ' Formatting the output for better readability
        With ws
            .Columns("I:L").AutoFit
            .Range("J2:J" & outputRow - 1).NumberFormat = "0.00"
            .Range("K2:K" & outputRow - 1).NumberFormat = "0.00%"
        End With

        ' Output the greatest percent increase, greatest percent decrease, and greatest total volume
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(2, 17).Value = maxIncreaseTicker
        ws.Cells(2, 18).Value = maxIncrease
        
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(3, 17).Value = maxDecreaseTicker
        ws.Cells(3, 18).Value = maxDecrease
        
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Cells(4, 17).Value = maxVolumeTicker
        ws.Cells(4, 18).Value = maxVolume
    Next sheetIndex

    MsgBox "Stock data calculation is complete for all years."
End Sub

