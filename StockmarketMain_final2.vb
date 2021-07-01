Sub stockmarket()

    Dim ws As Worksheet
    For Each ws In Worksheets
        
        Dim EndRow As Long
        Dim SummaryTableRow As Integer
        Dim TickerName As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As Double
        SummaryTableRow = 2
        TotalStockVolume = 0

        'find the end of the rows
        
        EndRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        EndRow1 = ws.Cells(Rows.Count, 9).End(xlUp).Row

        'define the cell names
        
        ws.Cells(1, 9).Value = "Ticker Name"
        ws.Cells(1, 10).Value = "Opening Price"
        ws.Cells(1, 11).Value = "Closing Price"
        ws.Cells(1, 12).Value = "Yearly Change"
        ws.Cells(1, 13).Value = "Percent Change"
        ws.Cells(1, 14).Value = "Total Stock Volume"

        'this code gives us the opening price for the very first ticker
        ws.Cells(2, 10).Value = ws.Cells(2, 3).Value

        'for loop to find for the change of ticker names
        
        For i = 2 To EndRow

            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1).Value Then

            TickerName = ws.Cells(i, 1).Value
            OpenPrice = ws.Cells(i + 1, 3).Value
            ClosePrice = ws.Cells(i, 6).Value
            
            'move the calculated values and ticker names to designated columns
            ws.Range("I" & SummaryTableRow).Value = TickerName
            ws.Range("K" & SummaryTableRow).Value = ClosePrice
            ws.Range("J" & SummaryTableRow + 1).Value = OpenPrice
            
            'sum up the total stock volume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            ws.Range("N" & SummaryTableRow).Value = TotalStockVolume
                
            'adding one to summary table row in order to move to the next cell down
            SummaryTableRow = SummaryTableRow + 1
            TotalStockVolume = 0
            
            Else
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

        End If
        Next i


        'another for loop to calculate Yearly Change and Percentage change
        For j = 2 To EndRow1

            OpenPrice = ws.Cells(j, 10).Value
            ClosePrice = ws.Cells(j, 11).Value
            YearlyChange = OpenPrice - ClosePrice
            PercentChange = Round(1 - (ClosePrice / OpenPrice), 2)

            ws.Range("L" & j).Value = YearlyChange
            ws.Range("M" & j).Value = PercentChange
            
            ws.Range("M" & j).NumberFormat = "0.00%"

            SummaryTableRow = SummaryTableRow + 1
            
        ' conditional to highlight the positives green and negatives red
        
            If YearlyChange < 0 Then
            ws.Range("L" & j).Interior.ColorIndex = 3
            Else
            ws.Range("L" & j).Interior.ColorIndex = 4
            
            End If
        Next j
        
    'name the cells
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("O1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    'find max, min of Percent Changes and max of Total Stock Volumes
    
    ws.Range("Q2") = "%" & WorksheetFunction.max(ws.Range("M2:M" & EndRow)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("M2:M" & EndRow)) * 100
    ws.Range("Q4") = WorksheetFunction.max(ws.Range("N2:N" & EndRow))

    ' use match function to find the Ticker names for max,min% and max total stock volume

    GreatestIncrease = WorksheetFunction.Match(WorksheetFunction.max(ws.Range("M2:M" & EndRow)), ws.Range("M2:M" & EndRow), 0)
    GreatestDecrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("M2:M" & EndRow)), ws.Range("M2:M" & EndRow), 0)
    GreatestTotalVolume = WorksheetFunction.Match(WorksheetFunction.max(ws.Range("N2:N" & EndRow)), ws.Range("N2:N" & EndRow), 0)

    ' show ticker names in following cells
    
    ws.Range("P2") = ws.Cells(GreatestIncrease, 9)
    ws.Range("P3") = ws.Cells(GreatestDecrease, 9)
    ws.Range("P4") = ws.Cells(GreatestTotalVolume, 9)

    Next ws

End Sub





