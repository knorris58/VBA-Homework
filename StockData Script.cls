VERSION 1.0 CLASS

sub multipleYearStockData():

For Each ws In Worksheets

Dim i As Double
Dim Ticker As String
Dim totalVolume As Double
Dim Table As Integer
Dim lastRow As Double
Table = 2
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
totalVolume = 0

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Total Stock Volume"

For i = 2 To lastRow
 If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 
      Ticker = ws.Cells(i, 1).Value
      totalVolume = totalVolume + ws.Cells(i, 7).Value
      ws.Range("I" & Table).Value = Ticker
      ws.Range("J" & Table).Value = totalVolume
      
      Table = Table + 1
      
      totalVolume = 0

    Else
      totalVolume = totalVolume + ws.Cells(i, 7).Value

    End If

  Next i

Next ws

End Sub

