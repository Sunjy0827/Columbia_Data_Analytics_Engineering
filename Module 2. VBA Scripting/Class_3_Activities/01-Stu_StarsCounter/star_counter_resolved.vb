Sub start_Counter()

Dim i As Integer
Dim j as Integer
Dim Starcounter as Integer

LastRow = cells(rows.count, "A").end(xlup).row

    For i = 2 To 51

        Starcounter = 0

        for j = 4 to 8
            If (Cells(i, j).Value = "Full-Star") then
            starcounter = starcounter + 1
            
            End if
        next j

    Cells(i, 9).value = Starcounter

    Next i

End Sub