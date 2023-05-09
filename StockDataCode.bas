Attribute VB_Name = "Module1"
Sub YearStock():
For Each ws In Worksheets
    'Creating Column Values
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    RowIndex = 2
    
    x = 2
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
            'Ticker
            ws.Cells(RowIndex, 9).Value = ws.Cells(i, 1).Value
            
            'YearlyChange
            YearlyChange = ws.Cells(i, 6).Value - ws.Cells(x, 3).Value
            ws.Cells(RowIndex, 10).Value = YearlyChange
                If ws.Cells(RowIndex, 10).Value < 0 Then
                    ws.Cells(RowIndex, 10).Interior.ColorIndex = 3
                ElseIf ws.Cells(RowIndex, 10).Value > 0 Then
                    ws.Cells(RowIndex, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(RowIndex, 10).Interior.ColorIndex = 6
            
                End If
            'Percent Change Calculations
            PercentChange = YearlyChange / ws.Cells(x, 3).Value
            ws.Cells(RowIndex, 11).Value = PercentChange
            ws.Cells(RowIndex, 11).NumberFormat = "0.00%"
            
            'Total Volume
            ws.Cells(RowIndex, 12) = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(i, 7), ws.Cells(x, 7)))
            
           
        RowIndex = RowIndex + 1
        x = i + 1
        
        End If
        
    Next i

'Creating table
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    'initial values
    GreatInc = ws.Cells(2, 11).Value
    GreatDec = ws.Cells(2, 11).Value
    GreatVol = ws.Cells(2, 12).Value
    GreatIncIndex = 0
    GreatDecIndex = 0
    GreatVolIndex = 0
    LastRowTable = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To LastRowTable
        If ws.Cells(i, 11) > GreatInc Then
            GreatInc = ws.Cells(i, 11).Value
            ws.Range("Q2") = GreatInc
            ws.Range("P2") = ws.Cells(i, 9)
       
        ElseIf ws.Cells(i, 11) < GreatDec Then
            GreatDec = ws.Cells(i, 11)
            ws.Range("Q3") = GreatDec
            ws.Range("P3") = ws.Cells(i, 9)
     
    
        ElseIf ws.Cells(i, 12) > GreatVol Then
            GreatVol = ws.Cells(i, 12)
            ws.Range("Q4") = GreatVol
            ws.Range("P4") = ws.Cells(i, 9)
        End If
    Next i
            
    'Formatting Cell Values
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
     
    
Next ws
End Sub
