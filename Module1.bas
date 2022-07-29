Attribute VB_Name = "Module1"
Sub stocks()

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Dim stock_name As String
Dim stock_open As Double
Dim stock_close As Double
Dim stock_vol As Double
Dim yearly_change As Double
Dim percent_change As Double

stock_open = 0
stock_close = 0
stock_vol = 0
yearly_change = 0
percent_change = 0

Dim summary_table_row As Integer
summary_table_row = 2

For i = 2 To lastrow

If stock_open = 0 Then
    stock_open = Cells(i, 3)
End If

    
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    stock_name = Cells(i, 1).Value
    stock_close = Cells(i, 6)
    stock_vol = stock_vol + Cells(i, 7)
    yearly_change = stock_close - stock_open
    percent_change = yearly_change / stock_open
    
    Range("I" & summary_table_row).Value = stock_name
    Range("J" & summary_table_row).Value = yearly_change
    Range("K" & summary_table_row).Value = Format(percent_change, "percent")
    Range("L" & summary_table_row).Value = stock_vol
    
If percent_change > 0 Then
    Cells(summary_table_row, 10).Interior.ColorIndex = 4
Else
    Cells(summary_table_row, 10).Interior.ColorIndex = 3
End If

summary_table_row = summary_table_row + 1
stock_vol = 0
stock_open = 0

Else:
    stock_vol = stock_vol + Cells(i, 7)
End If

Next i

End Sub

