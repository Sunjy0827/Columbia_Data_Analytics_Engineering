sub checkboard ()

    for i = 1 to 8
        for j = 1 to 8
        if cells(i,j).value mod 2 = 0 then        
        cells(i,j).Interior.ColorIndex = 1
        elseif cells(i,j).value mod 2 <> 0 then
        cells(i,j).Interior.ColorIndex = 3
        end if
        Next j
    Next i

end sub