Sub ticker_symbol()
Dim symbol As String

Dim totalStockVolume As Double
totalStockVolume = 0

Dim summaryTableRow As Integer
summaryTableRow = 2

Dim yearlyChange As Double
Dim percentChange As Double
Dim yearlyClosingPrice As Double

Dim lastRow As Integer
lastRow = Cells(RowCount, 1).End(xlUp).Row

For i = 2 To lastRow

Dim stockPriceCaptured As Boolean
If stockPriceCaptured = False Then
Dim yearlyOpeningPrice As Double
yearlyOpeningPrice = Cells(i, 3).Value
stockPriceCaptured = True
End If

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

symbol = Cells(i, 1).Value
yearlyClosingPrice = Cells(i, 6).Value
yearlyChange = yearlyClosingPrice - yearlyOpeningPrice
percentChange = yearlyChange / yearlyOpeningPrice
totalStockVolume = totalStockVolume + Cells(i, 7).Value

Range("I" & summaryTableRow).Value = symbol
Range("J" & summaryTableRow).Value = yearlyChange
Range("K" & summaryTableRow).Value = percentChange
Range("L" & summaryTableRow).Value = totalStockVolume

summaryTableRow = summaryTableRow + 1
totalStockVolume = 0
stockPriceCaptured = False
yearlyChange = 0

Else
totalStockVolume = totalStockVolume + Cells(i, 7).Value


End If
Next i

Dim formatting As Range
Dim yearlyChangeGreen As FormatCondition, yearlyChangeRed As FormatCondition
Set formatting = Range("J2:J4200")
formatting.FormatConditions.Delete
Set yearlyChangeGreen = formatting.FormatConditions.Add(xlCellValue, xlGreater, "0")
Set yearlyChangeRed = formatting.FormatConditions.Add(xlCellValue, xlLess, "0")
With yearlyChangeGreen
.Interior.Color = RGB(0, 255, 0)
End With
With yearlyChangeRed
.Interior.Color = RGB(255, 0, 0)
End With

End Sub
