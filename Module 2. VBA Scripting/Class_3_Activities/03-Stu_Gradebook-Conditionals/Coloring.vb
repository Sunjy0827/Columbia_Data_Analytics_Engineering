sub color()

Dim grade As Integer
grade = cells(2,2).value

If grade >= 90 then 
    cells(2,4) = "A"
    cells(2,4).Interior.ColorIndex = 4
    cells(2,3) = "Pass"
 elseif grade >= 80 then 
    cells(2,4) = "B"
    cells(2,4).Interior.ColorIndex = 4
    cells(2,3) = "Pass"
 elseif grade >= 70 then 
    cells(2,4) = "c"
    cells(2,4).Interior.ColorIndex = 6
    cells(2,3) = "Warning"
 elseif grade < 70 then
    cells(2,4) = "F"
    cells(2,4).Interior.ColorIndex = 3
    cells(2,3) = "Fail"
 End If
end sub

sub Reset_button():

Cells(12,2).value = Cells(2,2).value

Cells(2,2).value = ""
Cells(2,3).value = ""

clear = 

* If the score is 90 or higher:
  * Add an "A" in the letter grade cell.
  * Fill the Pass/Warning/Fail cell with the color green.
  * Add the text “Pass” to the Pass/Warning/Fail cell.

* If the score is greater than or equal to 80 and less than 90:
  * Add a "B" in the letter grade cell.
  * Fill the Pass/Warning/Fail cell with the color green.
  * Add the text “Pass” to the Pass/Warning/Fail cell.


* If the score is greater than or equal to 70 and less than 80:
  * Add a "C" in the letter grade cell.
  * Fill the Pass/Warning/Fail cell with the color yellow.
  * Add the text "Warning" to the Pass/Warning/Fail cell.

* Finally, if the score is below 70:
  * Add an “F” in the letter grade cell.
  * Fill the Pass/Warning/Fail cell with the color red.
  * Add the text “Fail” to the Pass/Warning/Fail cell.
'