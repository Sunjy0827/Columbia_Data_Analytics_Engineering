sub cryptoken()

Dim ShibaInu As Integer
Dim Dogecoin As Integer

ShibaInu = 0
Dogecoin = 0

for i = 1 to 6
    for j = 1 to 7
        if cells(i,j).value = "Shiba Inu" then
        
        ShibaInu = ShibaInu + 1

        Elseif cells(i,j).value = "Dogecoin" then

        Dogecoin = Dogecoin + 1
        End If
    Next j
Next i

range("I2").value = ShibaInu
range("I5").value = Dogecoin

end sub



Sub CountABC()
    Dim countA As Integer
    Dim countB As Integer
    Dim countC As Integer
    Dim i As Integer
    Dim j As Integer
    
    ' Initialize counts
    countA = 0
    countB = 0
    countC = 0
    
    ' Loop through each cell in the specified range
    For i = 1 To 6 ' Rows
        For j = 1 To 7 ' Columns
            ' Check if the cell value is A, B, or C
            If Cells(i, j).Value = "A" Then
                countA = countA + 1
            ElseIf Cells(i, j).Value = "B" Then
                countB = countB + 1
            ElseIf Cells(i, j).Value = "C" Then
                countC = countC + 1
            End If
        Next j
    Next i
    
    ' Output the counts
    MsgBox "Count of A: " & countA & vbCrLf & _
           "Count of B: " & countB & vbCrLf & _
           "Count of C: " & countC
    


End Sub