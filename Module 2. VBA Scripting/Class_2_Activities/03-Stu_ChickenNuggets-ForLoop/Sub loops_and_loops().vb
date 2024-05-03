Sub loops_and_loops()
    Dim i As Integer
    For i = 1 To 10 
        Cells(i, 1).Value = "I will eat"
        Cells(i, 2).Value = i+10
        Cells(i, 3).Value = "Chicken Nuggeets"
    Next i
End Sub