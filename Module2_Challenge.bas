Attribute VB_Name = "Module1"
Sub RunAllWorksheets()
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate
Call TickerSymbol
Call QuartChange
Call TotalStockVolume
Call Min_MaxPercentage
Next ws


End Sub

    Sub TickerSymbol()

 'create a variable to store ticker symbol
 Dim tckr As String
 
 'create variable to hold counter
 Dim x As Long
 
 'add title to column 9
 Cells(1, 9).Value = "Ticker"
 
 'establish last row
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

 'create a loop to go through each row and store tckr name
 For x = 2 To lastrow
     tckr = Cells(x, 1).Value
 
 'input the ticker symbol in each row
If Cells(x, 9).Value <> tckr Then
Cells(x, 9).Value = tckr
End If

 Next x

End Sub

Sub QuartChange()
'create variable to hold counter
Dim x As Long

 'create variable to hold Opening price
 Dim open_price As Double
 
 'create variable to hold Closing price
 Dim close_price As Double
 
 'create variable to hold the quarterly change
 Dim quart_change As Double
 
 'define last row
 lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 
 'add title to column 10
 Cells(1, 10).Value = "Quarterly Change"
 Columns("J").ColumnWidth = 15
 
 'add title to column 11
 Cells(1, 11).Value = "Percent Change"
 Columns("k").ColumnWidth = 15
 
 'add title to column 12
 Cells(1, 12).Value = "Total Stock Volume"
 Columns("L").ColumnWidth = 15
 
'loop through each row
For x = 2 To lastrow

' store pertinent variables from each column
open_price = Cells(x, 3).Value
close_price = Cells(x, 6).Value
quart_change = close_price - open_price

'perform function to calculate percentage change
If Cells(x, 11).Value = "" Then
Cells(x, 11).Value = ((close_price - open_price) / open_price)
Cells(x, 11).NumberFormat = "0.00%"

'format quarterly change colors
If Cells(x, 10).Value = "" Then
Cells(x, 10).Value = quart_change

If quart_change > 0 Then
Cells(x, 10).Interior.ColorIndex = 4

ElseIf quart_change < 0 Then
Cells(x, 10).Interior.ColorIndex = 3

End If

End If

End If
Next x

End Sub

Sub TotalStockVolume()
'create variable to hold counter
Dim x As Long

'create variable to hold stock volume
Dim StockVolume As Double
StockVolume = 0
'define last row
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'create loop to go through each row
For x = 2 To lastrow

'add previous volume to the next row
If Cells(x, 1).Value = Cells(x + 1, 1).Value Then
StockVolume = StockVolume + Cells(x, 7).Value
Cells(x, 12).Value = StockVolume

'start over when new ticker appears
Else
StockVolume = StockVolume + Cells(x, 7).Value
Cells(x, 12).Value = StockVolume
StockVolume = 0

End If


Next x

End Sub
Sub Min_MaxPercentage()

' create variable to hold counter
 Dim x As Long

'define last row
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'format summary table
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Columns("N").ColumnWidth = 20

'create variable to define percentage increase
Dim percent_up As Double
percent_up = WorksheetFunction.Max(Range("K2:K" & lastrow))
Cells(2, 16).Value = percent_up
Cells(2, 16).NumberFormat = "0.00%"

'create variable to define percentage decrease
Dim percent_down As Double
percent_down = WorksheetFunction.Min(Range("K2:K" & lastrow))
Cells(3, 16).Value = percent_down
Cells(3, 16).NumberFormat = "0.00%"

'create variable to define greatest total volume
Dim greatest_volume As Double
greatest_volume = WorksheetFunction.Max(Range("L2:L" & lastrow))
Cells(4, 16).Value = greatest_volume

'loop
For x = 2 To lastrow

If Cells(x, 11).Value = percent_up Then
Cells(2, 15).Value = Cells(x, 1).Value
End If
If Cells(x, 11).Value = percent_down Then
Cells(3, 15).Value = Cells(x, 1).Value
End If
If Cells(x, 12).Value = greatest_volume Then
Cells(4, 15).Value = Cells(x, 9).Value
End If

Next x

End Sub


