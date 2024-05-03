Sub CellsAndRanges():
    
    Range("A1").Value = "Rook"
    Range("B1").Value = "Knight"
    Range("C1").Value = "Bishop"
    Range("D1").Value = "Queen"
    Range("E1").Value = "King"
    Range("F1").Value = "Bishop"
    Range("G1").Value = "Knight"
    Range("H1").Value = "Rook"

    Range("A2:H2").Value = "Pawn"
    Range("A7:H7").Value = "Pawn"

    Range("A1").Value = "Rook"
    Range("B8").Value = "Knight"
    Range("C8").Value = "Bishop"
    Range("D8").Value = "Queen"
    Range("E8").Value = "King"
    Range("F8").Value = "Bishop"
    Range("G8").Value = "Knight"
    Range("H8").Value = "Rook"
    
End Sub