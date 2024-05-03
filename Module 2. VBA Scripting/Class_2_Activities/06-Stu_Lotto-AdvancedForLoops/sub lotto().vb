sub lotto()

Dim first_place as long
Dim second_place as long
Dim third_place as long

first_place = cells(2, 8).value
second_place = cells(3, 8).value
third_place = cells(4, 8).value


For i  = 2 to 1001
 if cells(i, 3).value = first_place then

    msgbox "COngratulations" + Cells(i, 1).value

    celLs(2, 6).value = cells(i, 1).value
    cells(2, 7).value = cells(i, 2).value

    elseif cells(i, 3).value = second_place then
  
    CelLs(3, 6).value = cells(i, 1).value
    cells(3, 7).value = cells(i, 2).value

    elseif cells(i, 3).value = third_place then
  
    CelLs(4, 6).value = cells(i, 1).value
    cells(4, 7).value = cells(i, 2).value
 end if
next i
end sub