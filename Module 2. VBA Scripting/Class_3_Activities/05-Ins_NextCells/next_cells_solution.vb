sub next_cell()

Dim State_Total As Integer
State_total = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim State_Name As String

For i = 2 to 80

    If cells(i+1,1).Value <> Cells(i,1).value then

    State_Name = Cells(i,1).value

    State_total = State_total + 1

    Range("C" & Summary_Table_Row).Value = State_Name

    Range("D" & Summary_Table_Row).Value = State_Total

    Summary_Table_Row = Summary_Table_Row + 1

    State_total = 0

    Else

    State_total = 0

end sub