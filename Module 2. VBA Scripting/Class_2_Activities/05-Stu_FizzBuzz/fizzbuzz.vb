' instructor code
Sub conditional_loops()

 For i = 1 to 10
    if cells(i,2).value mod 2 = 0 then
        cells(i, 2).value = "Even Row"
    Else
        Cells(i, 2).Value = "Odd Row"
    End If
 Next i

End sub

Sub Fizz_Buzz()

For i = 1 To 99
    If Cells(i, 1).Value Mod 3 = 0 And Cells(i, 1).Value Mod 5 = 0 Then
        Cells(i, 2).Value = "FizzBuzz"
    ElseIf Cells(i, 1).Value Mod 3 = 0 Then
        Cells(i, 2).Value = "Fizz"
    ElseIf Cells(i, 1).Value Mod 5 = 0 Then
        Cells(i, 2).Value = "Buzz"
    End If
Next i

End Sub


Sub Fizz_Buzz()
    Dim i As Integer
    For i = 1 To 99
        If Cells(i, 1).Value = "" Then Exit For ' Check if cell is empty
        If Cells(i, 1).Value Mod 3 = 0 And Cells(i, 1).Value Mod 5 = 0 Then
            Cells(i, 2).Value = "FizzBuzz"
        ElseIf Cells(i, 1).Value Mod 3 = 0 Then
            Cells(i, 2).Value = "Fizz"
        ElseIf Cells(i, 1).Value Mod 5 = 0 Then
            Cells(i, 2).Value = "Buzz"
        End If
    Next i
End Sub