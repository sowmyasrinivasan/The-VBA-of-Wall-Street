Sub stock():
    Dim row As Integer
    Dim volume As Double
    Dim op As Double
    Dim cl As Double
    row = 2
    volume = 0
    op = Cells(2, 3).Value
    For i = 2 To 70926
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            volume = volume + Cells(i, 7)
        Else
            volume = volume + Cells(i, 7)
            cl = Cells(i, 6).Value
            Cells(row, 9).Value = Cells(i, 1).Value
            Cells(row, 10).Value = op - cl
            Cells(row, 11).Value = ((op - cl) / op) * 100
            Cells(row, 12).Value = volume
            row = row + 1
            op = Cells(i + 1, 3).Value
            volume = 0
        End If
    Next i