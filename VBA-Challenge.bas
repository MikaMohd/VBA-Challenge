Sub StockSort()

For Sheet = 1 To ThisWorkbook.Worksheets.Count
Worksheets(Sheet).Activate

Dim TickerCount As Integer
Dim YROpeningVal
Dim YRClosingVal
Dim EndofTable As Long


YROpeningVal = 0
YRClosingVal = 0
TickerCount = 2
EndofTable = Range("A1").End(xlDown).Row

If Cells(EndofTable, 1).Value <> 0 Then
    For Col = 1 To 7
        Cells(EndofTable + 1, Col).Value = 0
    Next Col
End If

For Row = 2 To (EndofTable + 1)

    If IsNumeric(Cells(Row, 7).Value) = True Then
        SumVolume = SumVolume + Cells(Row - 1, 7).Value
    End If
       
    If Cells(Row, 1).Value <> Cells(Row - 1, 1).Value = True Then
           
        If Cells(Row, 1).Value <> 0 Then
            Cells(TickerCount, 9).Value = Cells(Row, 1).Value
        End If
           
        If IsNumeric(Cells(Row - 1, 6).Value) = True Then
            YRClosingVal = Cells(Row - 1, 6).Value
            Cells(TickerCount - 1, 10).Value = YRClosingVal - YROpeningVal
            Cells(TickerCount - 1, 11).Value = YRClosingVal / YROpeningVal - 1
            Cells(TickerCount - 1, 12).Value = SumVolume
        End If
           
        SumVolume = 0
        TickerCount = TickerCount + 1
        YROpeningVal = Cells(Row, 3).Value
    End If

Next Row

For Col = 1 To 7
    Cells(EndofTable + 1, Col).Clear
Next Col

Cells(1, 9).Value = "Tickers"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Range("I:I").ColumnWidth = 9
Range("J:K").ColumnWidth = 15
Range("L:L").ColumnWidth = 19

Range("K:K").NumberFormat = "0.00%"
Range("L:L").NumberFormat = "###,###,###,###,###,###"

Dim LTFormat As FormatCondition
Dim GTFormat As FormatCondition
Dim YearlyValues As Range
Set YearlyValues = Range("J2:J" & TickerCount)

YearlyValues.FormatConditions.Delete
Set LTFormat = YearlyValues.FormatConditions.Add(xlCellValue, xlLess, "0")
Set GTFormat = YearlyValues.FormatConditions.Add(xlCellValue, xlGreater, "0")

With LTFormat
    .Interior.Color = RGB(255, 180, 180)
End With
With GTFormat
    .Interior.Color = RGB(180, 255, 180)
End With
   
SumVolume = Empty
TickerCount = Empty
YROpeningVal = Empty
YRClosingVal = Empty
EndofTable = Empty


Dim GI As Double, GD As Double, GV As Double
Dim GIT As String, GDT As String, GVT As String
   
Dim EndofTable2 As Long
EndofTable2 = Range("I1").End(xlDown).Row

For Table2Row = 2 To EndofTable2
    If Cells(Table2Row, 11).Value > GI Then
    GI = Cells(Table2Row, 11).Value
    GIT = Cells(Table2Row, 9).Value
    End If
    
    If Cells(Table2Row, 11).Value < GD Then
    GD = Cells(Table2Row, 11).Value
    GDT = Cells(Table2Row, 9).Value
    End If
    
    If Cells(Table2Row, 12).Value > GV Then
    GV = Cells(Table2Row, 12).Value
    GVT = Cells(Table2Row, 9).Value
    End If
Next Table2Row

Cells(2, 16).Value = GIT
Cells(3, 16).Value = GDT
Cells(4, 16).Value = GVT

Cells(2, 17).Value = GI
Cells(3, 17).Value = GD
Cells(4, 17).Value = GV

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

Range("O:O").ColumnWidth = 20
Range("P:Q").ColumnWidth = 17

Range("Q2:Q3").NumberFormat = "0.00%"
Range("Q4").NumberFormat = "###,###,###,###,###,###"

GI = Empty
GD = Empty
GV = Empty
GIT = Empty
GDT = Empty
GVT = Empty

Next Sheet

MsgBox ("Here are the results")

End Sub