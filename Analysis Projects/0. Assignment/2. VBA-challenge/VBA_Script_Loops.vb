sub multipleyearstockdata ()

Dim ticker As String
Dim open As Double
Dim close As Double
Dim volume As Long
Dim lastRow As Long
DIm Summary_Table_Row AS String

volume = 0
LastRow = cells(rows.count, "A").end(xlup).row
Summary_Table_Row = 2


For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    ticker = Cells(i, 1).Value
    volume = volume + Cells(i, 7).value
    Range("I" & Summary_Table_Row).Value = ticker
    Rnage("J" & Summary_Table_Row).value = volume
    Range("K" & Summary_Table_Row).Value = volume
    Range("L" & Summary_Table_Row).Value = volume
    Summary_Table_Row = Summary_Table_Row + 1
    volume = 0

    End if

next i

end sub




for i = 2 to lastrow
    ticker = cells(i, 1).value
    openprice = cells(i,3).value
    closeprice = cells(i,6).value
    volume = cells(i,7).value

    cells(i,10).value = ticker
    cells(i,11).value = closeprice - openprice
    cells(i,12).value = volume

Next i

end sub
Sub ExtractUniqueValues()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim dict As Object
    Dim outputCol As Integer
    Dim i As Integer

    Set ws = ThisWorkbook.Sheets("Sheet1")  ' Modify the sheet name as needed
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row  ' Determine the last row in column A
    Set dict = CreateObject("Scripting.Dictionary")  ' Create a new dictionary

    ' Loop through each cell in column A
    For Each cell In ws.Range("A1:A" & lastRow)
        If Not dict.exists(cell.Value) And Not IsEmpty(cell.Value) Then
            dict.Add cell.Value, Nothing
        End If
    Next cell

    ' Output the unique values in another column, let's say column B (column index 2)
    outputCol = 2  ' Change this to the column number where you want to output the unique values
    i = 1  ' Start row for output

    ' Place each unique value in the output column
    For Each Key In dict.Keys
        ws.Cells(i, outputCol).Value = Key
        i = i + 1
    Next Key

    MsgBox "Unique values have been extracted."
End Sub
