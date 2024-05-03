Attribute VB_Name = "Moduel1"
Sub CellsAndRanges():

' Inserting Data Via Cells
Cells(2,1).value = "Cat"
' Inserting Data Via Ranges
Range("F1").value = "I"
'Inserting Data Across Ranges
Range("F5:F7").value = 5


End Sub