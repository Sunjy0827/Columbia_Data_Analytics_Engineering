sub forloop()
' create a variable to hold the counter
dim i as integer
' iterate through the rows placing a value of 1 throughout
for i =1 to 20
' iterate through the rows placing a value of 1 throughout
    cells(i, 1).value = i
' places increasing value 
    cells(i, 2).value = i*2
    cells(i, 3).value = i*3
next i

end sub