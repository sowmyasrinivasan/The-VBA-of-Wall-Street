Sub stock_loop()
    Dim first_row As Integer
    first_row = 2
    Dim last_row As Double
    Dim insert_row As Double
    insert_row = 2
    Dim current As Worksheet
    Dim stock_vol As Double
    stock_vol = 0
    Dim open_value As Double
    open_value = Cells(2, 3).Value
    Dim close_value As Double
    For Each current In Worksheets
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        For i = first_row To last_row
            If Cells(i, 1).Value <> Cells(i + 1, 1) Then
                close_value = Cells(i, 6).Value
                Cells(insert_row, 10).Value = Cells(i, 1).Value
                Cells(insert_row, 11).Value = close_value - open_value
                If open_value <> 0 Then
                    Cells(insert_row, 12).Value = ((close_value - open_value) / open_value) * 100
                Else
                End If
                Cells(insert_row, 13).Value = stock_vol + Cells(i, 7).Value
                insert_row = insert_row + 1
                stock_vol = 0
                open_value = Cells(i + 1, 3).Value
            Else
                stock_vol = stock_vol + Cells(i, 7).Value
            End If
            Next i
        Next
        
                
                
End Sub

