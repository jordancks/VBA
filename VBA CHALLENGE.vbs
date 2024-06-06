Attribute VB_Name = "Module1"
Sub VBAChallenge()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets



ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quaterly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Valume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "GreatestTotal Valume"


Dim ticker_name As String

Dim total_stock As Double
total_stock = 0


Dim Quarterly_change As Double
Dim Open_price As Double
Dim closing_price As Double
Dim percent_change As Double


Dim table_row As Long
table_row = 2

Dim lastrow As Long

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim i As Long
For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker_name = ws.Cells(i, 1).Value
total_stock = total_stock + ws.Cells(i, 7).Value

ws.Range("I" & table_row).Value = ticker_name

ws.Range("L" & table_row).Value = total_stock

table_row = table_row + 1

total_stock = 0

Open_price = ws.Cells(i, 3).Value
closing_price = ws.Cells(i, 6).Value
Quarterly_change = closing_price - Open_price

ws.Range("J" & table_row).Value = Quarterly_change

percent_change = Quarterly_change / Open_price * 100

ws.Range("K" & table_row).Value = percent_change
ws.Range("K" & table_row).NumberFormat = "0.00%"


Else


total_stock = total_stock + ws.Cells(i, 7).Value


End If

If ws.Range("J" & table_row).Value > 0 Then
ws.Range("J" & table_row).Interior.ColorIndex = 4

ElseIf ws.Range("J" & table_row).Value < 0 Then
ws.Range("J" & table_row).Interior.ColorIndex = 3

Else
ws.Range("J" & table_row).Interior.ColorIndex = 2

End If

Dim maxincreaseticker As String
Dim maxdecreaseticker As String
Dim maxvolumeticker As String
Dim maxincrease As Double
Dim maxdecrease As Double
Dim maxvolume As Double

maxincrease = 0
maxdecrease = 0
maxvolume = 0

If ws.Cells(i, 11).Value > maxincrease Then
maxincrease = ws.Cells(i, 11)
maxincreaseticker = ws.Cells(i, 9)
ws.Range("Q2") = maxincrease
ws.Range("P2") = maxincreaseticker
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("P2").NumberFormat = "0.00%"

End If

If ws.Cells(i, 11).Value < maxdecrease Then
maxdecrease = ws.Cells(i, 11)
maxdecreaseticker = ws.Cells(i, 9)
ws.Range("Q3") = maxdecrease
ws.Range("P3") = maxidereaseticker
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("P3").NumberFormat = "0.00%"

End If

If ws.Cells(i, 12).Value > maxvolume Then
maxvolume = ws.Cells(i, 12)
maxvolumeticker = ws.Cells(i, 9)
ws.Range("Q4") = maxvolume
ws.Range("P4") = maxvolumeticker

End If

Next i


Next ws



End Sub



