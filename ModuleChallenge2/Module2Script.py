Sub Ticker()
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
    'Declare the variables
    Dim i As Long
    Dim Ticker As String
    Dim TotalVolume As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim LastRow As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    
    TotalVolume = 0
    
    'Find the last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).row
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"
    ws.Range("K1").Value = "Quarterly Change"
    ws.Range("L1").Value = "Percent Change"
    
    Dim row As Double
    Dim column As Integer
    row = 2
    column = 1
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'set initial price
     OpenPrice = Cells(2, column + 2).Value
    
    'Loop through all ticker symbols
    For i = 2 To LastRow
    
    'Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Setting ticker name
    Ticker = Cells(i, 1).Value
    
    'setting close price
    ClosePrice = Cells(i, column + 5).Value
    
    'calculate total volume
    TotalVolume = TotalVolume + Cells(i, 7).Value
    
    'Print the ticker in summary table
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    
    'Print volume in summary table
    ws.Range("J" & Summary_Table_Row).Value = TotalVolume
    
    'Calculate the quarterly change
    QuarterlyChange = ClosePrice - OpenPrice
    
    'Print the quarterly change in summary table
    ws.Range("K" & Summary_Table_Row).Value = QuarterlyChange
    
    'Calculate percent change
    PercentChange = QuarterlyChange / OpenPrice
    
    'Print percent change to summary table
    ws.Range("L" & Summary_Table_Row).Value = PercentChange
    ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
    
    'Add one to summary table row
    Summary_Table_Row = Summary_Table_Row + 1
    
    'reset open price to next ticker
    OpenPrice = Cells(i + 1, column + 2).Value
    
    'Reset volume
    TotalVolume = 0
    
    Else
    
    'Add to volume
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
    'Print volume to summary table
    ws.Range("J" & Summary_Table_Row).Value = Volume
    
    End If
    
    Next i
    
    'find the last row of ticker column
    Quarterly_change_last_row = ws.Cells(Rows.Count, 9).End(xlUp).row
    
    'Set the cell colors for quarterly change
    For j = 2 To Quarterly_change_last_row
        If (Cells(j, 11).Value > 0 Or Cells(j, 11).Value = 0) Then
            Cells(j, 11).Interior.ColorIndex = 10
            ElseIf Cells(j, 11).Value < 0 Then
            Cells(j, 11).Interior.ColorIndex = 3
            End If
    Next j
    
    'set cell colors for percent change
    For l = 2 To Quarterly_change_last_row
        If (Cells(l, 12).Value > 0 Or Cells(l, 12).Value = 0) Then
        Cells(l, 12).Interior.ColorIndex = 10
        ElseIf Cells(l, 12).Value < 0 Then
        Cells(l, 12).Interior.ColorIndex = 3
        End If
        Next l
        
    
    'set the ticker, value, greatest% inc, % dec, and greatest total volume headers
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Inc"
    Cells(3, 15).Value = "Greatest % Dec"
    Cells(4, 15).Value = "Greatest total volume"
    
    'find the highest value of each ticker
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
    
    Next ws
    
End Sub
