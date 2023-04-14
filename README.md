# VBA-Challenge

Sub ticker()
 
'set labels for new cells
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Stock Volume"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greates total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'prepare dimensions
Dim ticker As String
Dim total_change As Long
Dim percentchange As Double
Dim summary_row As Integer
Dim days As Double
Dim RowCount As Long
Dim tickertotal As Long
Dim start As Double
Dim yearend As Double


'values
tickertotal = 0
total_change = 0
summary_row = 2
percentchange = 0

'Row count
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'loop for calculating new data
For i = 2 To RowCount

'same or not
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'setting ticker column
ticker = Cells(i, 1).Value

'calculating change for year, i am not sure how to get the value of the lowest cell value in C and highest cell value in F
total_change = total_change + (Cells(i, 6).Value - Cells(i, 3).Value)

'calculate percent change
percentchange = (total_change / Cells(1, 6).Value) * 100

'Calculate volume
tickertotal = tickertotal + Cells(i, 7).Value

'put in cells
Range("I" & summary_row).Value = ticker
Range("J" & summary_row).Value = total_change
Range("K", summary_row).Value = percentchange
Range("L" & summary_row).Value = tickertotal


summary_row = summary_row + 1
tickertotal = 0

Else

tickertotal = tickertotal + Cells(i, 7).Value


End If

Next i

If Range("J").Value > 1 Then
Cells(i, J).Interior.ColorIndex = 4

Else
Cells(i, J).Interior.ColorIndex = 3

End If



End Sub
