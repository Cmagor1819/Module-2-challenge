# Module-2-challenge
Got the following lines of code from a peer/group of peers
 Dim row As Double
    Dim column As Integer
    row = 2
    column = 1
     OpenPrice = Cells(2, column + 2).Value
      ClosePrice = Cells(i, column + 5).Value
      QuarterlyChange = ClosePrice - OpenPrice
      PercentChange = QuarterlyChange / OpenPrice
       OpenPrice = Cells(i + 1, column + 2).Value
       Quarterly_change_last_row = ws.Cells(Rows.Count, 9).End(xlUp).row
       For k = 2 To Quarterly_change_last_row
    If Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Quarterly_change_last_row)) Then
        Cells(2, 16).Value = Cells(k, 9).Value
        Cells(2, 17).Value = Cells(k, 12).Value
        Cells(2, 17).NumberFormat = "0.00%"
    ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Min(ws.Range("L2:L" & Quarterly_change_last_row)) Then
        Cells(3, 16).Value = Cells(k, 9).Value
        Cells(3, 17).Value = Cells(k, 12).Value
        Cells(3, 17).NumberFormat = "0.00%"
    ElseIf Cells(k, column + 9).Value = Application.WorksheetFunction.Max(ws.Range("J2:J" & Quarterly_change_last_row)) Then
        Cells(4, 16).Value = Cells(k, 9).Value
        Cells(4, 17).Value = Cells(k, 10).Value
        End If
        Next k
 
