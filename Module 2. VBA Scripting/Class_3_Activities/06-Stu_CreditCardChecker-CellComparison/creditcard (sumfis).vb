Sub creditcard()

Dim TotalCharged As Integer
TotalCharged = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim CCType As String

For i = 2 To 101

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    CCType = Cells(i, 1).Value
    TotalCharged = TotalCharged + Cells(i, 3).value
    Range("G" & Summary_Table_Row).Value = CCType
    Range("H" & Summary_Table_Row).Value = TotalCharged
    Summary_Table_Row = Summary_Table_Row + 1
    TotalCharged = 0

    Else

    TotalCharged = TotalCharged + Cells(i, 3).value
    
    End If
    
    Next i

End Sub