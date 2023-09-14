Sub VbaChallange()
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Dim LastRow As Long
    Dim LastTicker As String
    Dim TotalSum As Double
    Dim HighestPercent As Double
    Dim HighestPercentIndex As Long
    Dim LowestPercent As Double
    Dim LowestPercentIndex As Long
    Dim HighestVolume As Double
    Dim HighestVolumeIndex As Long
    Dim TotalCount As Double
    Dim TotalCountIndex As Long
    Dim Counter As Long
    Dim InitialOpen As Double
    Dim FinalClose As Double
    Dim PriceChange As Double
    Dim PercentChange As Double
    
    ' Set the starting worksheet as the currently active worksheet
    Set starting_ws = ActiveSheet
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Activate the current worksheet in the loop
        ws.Activate
        
        ' Set headers for data columns in the current worksheet
        Cells(1, 9).Value = "Stock Symbol"
        Cells(1, 10).Value = "Annual Change"
        Cells(1, 11).Value = "Percentage Change"
        Cells(1, 12).Value = "Total Trading Volume"
        Cells(1, 15).Value = "Stock Symbol"
        Cells(1, 16).Value = "Value"
        Cells(2, 14).Value = "Largest Percentage Increase"
        Cells(3, 14).Value = "Largest Percentage Decrease"
        Cells(4, 14).Value = "Largest Total Volume"
        
        ' Auto-fit columns for better visibility
        Columns("I:L").AutoFit
        Columns("N:N").AutoFit
        Columns("O:O").AutoFit
        
        ' Find the last row with data in the current worksheet
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        LastTicker = ""
        TotalSum = 0
        Counter = 1
        
        ' Loop through rows in the current worksheet
        For Row = 2 To LastRow
            If Cells(Row, 1).Value = LastTicker Then
                ' Calculate the total trading volume for the same stock symbol
                TotalSum = TotalSum + Cells(Row, 7)
            ElseIf Counter = 1 Then
                ' Initialize values for the first row of a new stock symbol
                Counter = Counter + 1
                LastTicker = Cells(Row, 1).Value
                Cells(Counter, 9).Value = LastTicker
                InitialOpen = Cells(Row, 3).Value
            Else
                ' Calculate values for subsequent rows of the same stock symbol
                Counter = Counter + 1
                LastTicker = Cells(Row, 1).Value
                Cells(Counter, 9).Value = LastTicker
                FinalClose = Cells(Row - 1, 6).Value
                PriceChange = FinalClose - InitialOpen
                PercentChange = IIf(InitialOpen = 0, 0, PriceChange / InitialOpen)
                
                ' Track the highest percentage increase and decrease
                If PercentChange > HighestPercent Then
                    HighestPercent = PercentChange
                    HighestPercentIndex = Counter
                ElseIf PercentChange < LowestPercent Then
                    LowestPercent = PercentChange
                    LowestPercentIndex = Counter
                End If
                
                ' Track the stock with the largest total trading volume
                If TotalSum > TotalCount Then
                    TotalCount = TotalSum
                    TotalCountIndex = Counter
                End If
                
                ' Populate data in columns 10 (Annual Change) and 11 (Percentage Change)
                Cells(Counter - 1, 10) = PriceChange
                Cells(Counter - 1, 11) = Format(PercentChange, "0.00%")
                
                ' Color cells in column 10 based on price change direction
                If PriceChange > 0 Then
                    Cells(Counter - 1, 10).Interior.ColorIndex = 4 ' Green
                ElseIf PriceChange < 0 Then
                    Cells(Counter - 1, 10).Interior.ColorIndex = 3 ' Red
                End If
                
                ' Update InitialOpen and TotalSum for the next stock symbol
                InitialOpen = Cells(Row, 3).Value
                Cells(Counter - 1, 12).Value = TotalSum
                TotalSum = 0
            End If
        Next Row
        
        ' Populate summary data in columns 15 (Stock Symbol) and 16 (Value)
        Cells(2, 15).Value = Cells(HighestPercentIndex - 1, 9).Value
        Cells(3, 15).Value = Cells(LowestPercentIndex - 1, 9).Value
        Cells(4, 15).Value = Cells(TotalCountIndex - 1, 9).Value
        Cells(2, 16).Value = Format(HighestPercent, "0.00%")
        Cells(3, 16).Value = Format(LowestPercent, "0.00%")
        Cells(4, 16).Value = TotalCount
    Next ws
    
    ' Activate the starting worksheet
    starting_ws.Activate
End Sub
