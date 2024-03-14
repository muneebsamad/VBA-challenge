
Sub Stock_data()

Dim w As Long

    Application.ScreenUpdating = False

    For w = 1 To Worksheets.Count
    Sheets(w).Select


    Application.ScreenUpdating = True
'Assign columns
Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
Range("P1:Q1").Value = Array("Ticker", "Value")
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"


'Assign variables
Dim Ticker As String
Dim ClosingPrice As Double
Dim YearlyChange As Double
Dim OpeningPrice As Double
Dim StockName As String
Dim StockVolume As Double
Dim TotalVolume As Double
Dim SummaryRow As Integer
Dim PercentChange As Double
Dim IncTicker As String
Dim DecTicker As String
Dim VolTicker As String
Dim rng As Range
Dim MaxValue As Variant
Dim MinValue As Variant
Dim GreatestVolume As Double
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double

'Initialize IncTicker
GreatestIncrease = 0
IncTicker = " "

'Initialize DecTicker
GreatestDecrease = 0
DecTicker = " "

'Initialize GreatestVolume
GreatestVolume = 0
VolTicker = " "

'Start summary table

SummaryRow = 2
lastRow = Cells(Rows.Count, 1).End(xlUp).Row


'Loop data
For i = 2 To lastRow

'Check if new ticker
If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
'Set Ticker
Ticker = Cells(i, 1).Value
'Set Opening Price
OpeningPrice = Cells(i, 3).Value
'ResetVolume
TotalVolume = 0
End If
'Add Total Volume
TotalVolume = TotalVolume + Cells(i, 7).Value
'Check last row of ticker
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
'Set Closing price
ClosingPrice = Cells(i, 6).Value
'Calc yearly change
YearlyChange = ClosingPrice - OpeningPrice
'Calc percentage change
    If OpeningPrice <> 0 Then
    PercentChange = (YearlyChange / OpeningPrice)
    Else
    PercentChange = 0
    End If
'Check IncTicker
    If PercentChange > GreatestIncrease Then
    GreatestIncrease = PercentChange
    IncTicker = Cells(i, 1).Value
    End If
'Check DecTicker
    If PercentChange < GreatestDecrease Then
    GreatestDecrease = PercentChange
    DecTicker = Cells(i, 1).Value
    End If
'Check VolTicker
    If TotalVolume > GreatestVolume Then
    GreatestVolume = TotalVolume
    VolTicker = Cells(i, 1).Value
    End If
'Display results
Range("I" & SummaryRow).Value = Ticker
Range("L" & SummaryRow).Value = TotalVolume
Range("J" & SummaryRow).Value = YearlyChange
Range("J" & SummaryRow).NumberFormat = "$#,##0.00"
Range("K" & SummaryRow).Value = PercentChange
Range("K" & SummaryRow).NumberFormat = "0.00%"
Range("P2").Value = IncTicker
Range("P3").Value = DecTicker
Range("P4").Value = VolTicker
'Add colors
    If YearlyChange > 0 Then
    Range("J" & SummaryRow).Interior.Color = vbGreen
    Else
    Range("J" & SummaryRow).Interior.Color = vbRed
    End If
    If PercentChange > 0 Then
    Range("K" & SummaryRow).Interior.Color = vbGreen
    Else
    Range("K" & SummaryRow).Interior.Color = vbRed
    End If
'Add summary
SummaryRow = SummaryRow + 1

Else

End If

Next i

'Set greatest values
Set rng = Range("K:K")
MaxValue = Application.WorksheetFunction.Max(rng)
Range("Q2").Value = MaxValue
Range("Q2").NumberFormat = "0.00%"

Set rng = Range("K:K")
MinValue = Application.WorksheetFunction.Min(rng)
Range("Q3").Value = MinValue
Range("Q3").NumberFormat = "0.00%"

Set rng = Range("L:L")
MaxValue = Application.WorksheetFunction.Max(rng)
Range("Q4").Value = MaxValue
Range("J:Q").Columns.AutoFit             'Autofits Columns
Next w
End Sub