Sub ChessBoard ():
    
    Cells(1,1).Value = "Rook"
    Cells(1,2).Value = "Knight"
    Cells(1,3).Value = "Bishop"
    Cells(1,4).Value = "Queen"
    Cells(1,5).Value = "King"
    Cells(1,6).Value = "Bishop"
    Cells(1,7).Value = "Knight"
    Cells(1,8).Value = "Rook"

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