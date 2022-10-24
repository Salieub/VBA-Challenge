Sub Alphabetical_Testing()
'Loop through worksheets

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

ws.Activate
'
Dim ticker As String

Dim yearly_change As Double

Dim percent_change As Double

Dim total_stock_volume As Long

Dim open_price As Double

Dim close_price As Double

Dim total_ticker As Integer

Dim unique_ticker As Integer

total_stock_volume = 0
percent_change = 0
yearly_change = 0
unique_ticker = 2

'Create columns
Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "yearly change"
Cells(1, 11).Value = "percent change"
Cells(1, 12).Value = "total stock volume"
Last_row = Cells(Rows.Count, 1).End(xlUp).Row

'ticker symbol
For i = 2 To Last_row
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
ticker = Cells(i, 1).Value
Cells(unique_ticker, 9).Value = ticker
unique_ticker = unique_ticker + 1



End If
Next i
Next ws
End Sub
