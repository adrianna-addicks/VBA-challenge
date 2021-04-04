Attribute VB_Name = "Module1"
Sub ticker_project()
'Creating variables to house data for ticker project
Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim volume As Double
Dim open_price As Double
Dim close_price As Double
Dim Worksheetname As String
Dim WS As Worksheet
Dim lastrow As Double
Dim year_close As Double
Dim year_open As Integer

For Each WS In Worksheets

'Creating headers for all worksheets
WS.Cells(1, 9).Value = "ticker"
WS.Cells(1, 10).Value = "yearly change"
WS.Cells(1, 11).Value = "percent change"
WS.Cells(1, 12).Value = "total stock volume"

'Creating table to house information and setting the row for table
Dim Summary_table_row As Integer
'Start with row 2, add a row in each loop event
  Summary_table_row = 2
  
  Dim open_row_number As Double
 open_row_number = 2

'Returns the data in the last cell
lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through the ticker symbols
For i = 2 To lastrow
vol = WS.Cells(i, 7).Value + vol
'Pulling ticker info

If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then

ticker = WS.Cells(i, 1).Value
open_row_number = 2
year_open = WS.Cells(open_row_number, 3).Value
year_close = WS.Cells(i, 6).Value
yearly_change = year_close - year_open
'Attempting to handle avoiding dividing zero error

If WS.Cells(open_row_number, 3).Value <> 0 Then
percent_change = (WS.Cells(i, 6) - WS.Cells(open_row_number, 3)) / WS.Cells(open_row_number, 3)
Else
percent_change = ""
End If
volume = volume + WS.Cells(i, 7).Value
 WS.Cells(Summary_table_row, 12).Value = vol
 
 WS.Cells(Summary_table_row, 9).Value = ticker
 WS.Cells(Summary_table_row, 10).Value = yearly_change
WS.Cells(Summary_table_row, 11).Value = percent_change
 WS.Cells(Summary_table_row, 12).Value = vol
 Summary_table_row = Summary_table_row + 1
 vol = 0
End If
Next i
Next WS
End Sub


