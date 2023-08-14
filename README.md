# VBA-challenge
VBA scripting in Excel Module 2

Sub VBAModule2()

    'Loop through all worksheets in Worksheets
    For Each ws In Worksheets
        
        'Create the six columns and three rows
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Determine the last row and count all rows
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        totalvolume = 0
        initializer = 2
        Index = 0
        
        For Row = 2 To lastRow
        
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
                
                ' 0 + value of total stock at the rowth row
                totalvolume = totalvolume + ws.Cells(Row, 7).Value
                yearlychange = ws.Cells(Row, 6).Value - ws.Cells(initializer, 3).Value
                percentchange = yearlychange / ws.Cells(initializer, 3).Value
                
                'Outputs four columns but only for each specific ticker as shown with Index, which is incremented
                ws.Range("I" & 2 + Index) = ws.Cells(Row, 1).Value
                ws.Range("J" & 2 + Index).Value = yearlychange
                ws.Range("J" & 2 + Index).NumberFormat = "0.00"
                ws.Range("K" & 2 + Index).Value = percentchange
                ws.Range("K" & 2 + Index).NumberFormat = "0.00%"
                ws.Range("L" & 2 + Index).Value = totalvolume
                
                If yearlychange > 0 Then
                    ws.Range("J" & 2 + Index).Interior.ColorIndex = 4
                ElseIf yearlychange < 0 Then
                    ws.Range("J" & 2 + Index).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & 2 + Index).Interior.ColorIndex = 0
                End If
                
                initializer = Row + 1
                totalvolume = 0 'resets the total volume back to zero for different ticker
                Index = Index + 1
                
                'accounts for if the total is neither positive nor negative
                If totalvolume = 0 Then
                    ws.Range("I" & 2 + Index).Value = Cells(Row, 1).Value
                    ws.Range("J" & 2 + Index).Value = 0
                    ws.Range("K" & 2 + Index).Value = "0.00%"
                    ws.Range("L" & 2 + Index).Value = 0
                End If
                
            Else
                'if the ticker of current and next rows are the same, still add to the total and increment
                totalvolume = totalvolume + ws.Cells(Row, 7).Value
            End If
            
        Next Row
        
        ws.Range("Q2") = Application.WorksheetFunction.Max(ws.Range("K2", "K" & lastRow)) * 100 & "%"
        ws.Range("Q3") = Application.WorksheetFunction.Min(ws.Range("K2", "K" & lastRow)) * 100 & "%"
        ws.Range("Q4") = Application.WorksheetFunction.Max(ws.Range("L2", "L" & lastRow))
        
        Max = Application.WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2", "K" & lastRow)), ws.Range("K2", "K" & lastRow), 0)
        Min = Application.WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2", "K" & lastRow)), ws.Range("K2", "K" & lastRow), 0)
        totalstock = Application.WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2", "L" & lastRow)), ws.Range("L2", "L" & lastRow), 0)
        
        ws.Range("P2") = ws.Cells(Max + 1, 9)
        ws.Range("P3") = ws.Cells(Min + 1, 9)
        ws.Range("P4") = ws.Cells(totalstock + 1, 9)
        
    Next ws
    
End Sub
