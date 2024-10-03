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
