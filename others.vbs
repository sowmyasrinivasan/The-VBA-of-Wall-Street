Sub others():
    Dim current As Worksheet
    Dim first_row As Double
    first_row = 2
    Dim max_percent As Double
    Dim min_percent As Double
    Dim max_vol As Double
    Dim last_row As Double
    For Each current In Worksheets
        last_row = Cells(Rows.Count, 12).End(xlUp).Row
        max_percent = Application.WorksheetFunction.Max(Range("l:l"))
        min_percent = Application.WorksheetFunction.Min(Range("l:l"))
        max_vol = Application.WorksheetFunction.Max(Range("m:m"))
        For i = first_row To last_row
            If Cells(i, 11).Value < 0 Then
                Cells(i, 11).Interior.ColorIndex = 3
            ElseIf Cells(i, 11).Value > 0 Then
                Cells(i, 11).Interior.ColorIndex = 4
            Else
            End If
            
            If Cells(i, 12).Value = max_percent Then
                Cells(2, 16).Value = Cells(i, 10).Value
                Cells(2, 17).Value = Cells(i, 12).Value
            ElseIf Cells(i, 12).Value = min_percent Then
                Cells(3, 16).Value = Cells(i, 10).Value
                Cells(3, 17).Value = Cells(i, 12).Value
            ElseIf Cells(i, 13).Value = max_vol Then
                Cells(4, 16).Value = Cells(i, 10).Value
                Cells(4, 17).Value = Cells(i, 13).Value
            Else
            End If
        Next i
    Next
    
                
                
        
            
    
End Sub
