Attribute VB_Name = "Module1"
Sub tickertotal()

'Create summary table.

Range("I1").Value = "Ticker Symbol"
Range("j1").Value = "Ticker Total"

'Set variables for ticker symbols and total stock volume

Dim ticker_symbol As String
Dim stock_volume As Double

'Set stock volume starting value

stock_volume = 0

'Set summary table row

Dim summary_table_row As Double

summary_table_row = 2

'Set code to run in all active worksheets

Dim ws As Worksheet

For Each ws In Worksheets

'Determine last row of table

Dim lastRow As Double

lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through stock data

For j = 2 To lastRow

'Add volumes for each ticker symbol

If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
ticker_symbol = ws.Cells(j, 1).Value
stock_volume = stock_volume + ws.Cells(j, 7).Value


'Print the ticker symbol in the summary table

Range("I" & summary_table_row).Value = ticker_symbol

'Print the stock volume to the summary table

Range("J" & summary_table_row).Value = stock_volume

'Add one to the summary table row

summary_table_row = summary_table_row + 1

'Reset stock volume

stock_volume = 0

End If

Next j

Next ws

End Sub
