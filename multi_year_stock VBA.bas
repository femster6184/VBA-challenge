Attribute VB_Name = "Module1"
Sub Multiple_year_stock()

For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

Dim TickerNext As Integer
Dim i As Long
Dim Ticker As String
Dim openprice As Double
Dim closeprice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double
Dim MaxPercent As Double
Dim MinPercent As Double
Dim MaxVolumeName As String
Dim MaxVolume As Double
Dim sheetname As String

openprice = ws.Cells(2, 3).Value
Ticker = " "
TickerNext = 1
YearlyChange = 0
PercentChange = 0
TotalStockVolume = 0

For i = 2 To LastRow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
TickerNext = TickerNext + 1
Ticker = ws.Cells(i, 1).Value
ws.Cells(TickerNext, "I").Value = Ticker
closeprice = ws.Cells(i, 6).Value
YearlyChange = closeprice - openprice
ws.Cells(TickerNext, "J").Value = YearlyChange

TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
ws.Cells(TickerNext, "L").Value = TotalStockVolume
PercentChange = (YearlyChange / openprice)
ws.Cells(TickerNext, "K").Value = PercentChange
ws.Cells(TickerNext, "K").NumberFormat = "0.00%"
openprice = ws.Cells(i + 1, 3).Value
TickerNext = TickerNext + 1
TotalStockVolume = 0

Else
TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

End If


Next i

lastrowyearlychange = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lastrowyearlychange
If ws.Cells(i, 10) > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 10

Else
    ws.Cells(i, 10).Interior.ColorIndex = 3

End If


MaxPercent = 0
MinPercent = 0
MaxVolume = 0

ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

MaxPercent = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrowyearlychange)), ws.Range("K2:K" & lastrowyearlychange), 0)
ws.Range("O2") = ws.Cells(MaxPercent + 1, 9)
MinPercent = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrowyearlychange)), ws.Range("K2:K" & lastrowyearlychange), 0)
MaxVolume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrowyearlychange)), ws.Range("L2:L" & lastrowyearlychange), 0)
ws.Range("O3") = ws.Cells(MinPercent + 1, 9)
ws.Range("O4") = ws.Cells(MaxVolume + 1, 9)
ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & lastrowyearlychange))
ws.Range("P2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastrowyearlychange)) * 100
ws.Range("P3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastrowyearlychange)) * 100

Next i

sheetname = ActiveSheet.Name
Worksheets(sheetname).Columns("I:L").AutoFit
Worksheets(sheetname).Columns("N:P").AutoFit
Next ws

End Sub
