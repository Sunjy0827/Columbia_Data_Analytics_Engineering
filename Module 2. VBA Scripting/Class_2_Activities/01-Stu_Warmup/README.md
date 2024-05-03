# Budget Checker

In this activity, you'll write a VBA script to run a budget checker in Excel.

## Instructions

* There are three parts to this problem.

  * Part 1: Calculate the total amount after adding in the fee, and enter the value in the "Total" cell.

  * Part 2: Create a message box to alert the user if the total amount, including the fee, is within or over budget.

  * Part 3 (Challenge): If the total is over budget, correct the price so it fits within the user’s budget. Be sure to round down!

    * For example: If the user's budget is 100 and the fees are 15%, the max price should be 86.

## Hints

* Break up the problem into smaller steps.

* Look at old code! You got this!

—

© 2022 edX Boot Camps LLC. Confidential and Proprietary. All Rights Reserved.

Sub SimpleArrays():
    
    ' String Splitting Example
    ' ------------------------------------------
    Dim Words() As String
    Dim Shakespeare As String
    Shakespeare = "To be or not to be. That is the question"

    ' Break apart the Shakespeare quote into individual words
    Words = Split(Shakespeare, " ")

    ' Print individual word
    MsgBox (Words(5))

End Sub

Sub Variables():

    ' Basic String Variable
    ' ----------------------------------------
    Dim name As String
    name = "Gandalf"

    MsgBox (name)

    ' Basic String Concatenation (Combination)
    ' ----------------------------------------
    Dim title As String
    title = "The Great"

    Dim fullname As String
    fullname = name + " " + title

    MsgBox (fullname)

    ' Basic Integer, Double, Long Variables
    ' ----------------------------------------
    Dim age1 As Integer
    Dim age2 As Integer
    age1 = 5
    age2 = 10

    Dim price As Double
    Dim tax As Double
    price = 19.99
    tax = 0.05

    Dim lightspeed As Long
    lightspeed = 299792458

    ' Basic Numeric manipulation
    ' ----------------------------------------
    MsgBox (age1 + age2)
    Cells(1, 1).Value = price * (1 + tax)

    ' String, Numeric Combination (Casting)
    ' ----------------------------------------
    MsgBox ("I am " + Str(age1) + " years old.")

    ' Booleans
    ' ----------------------------------------
    Dim money_grows_on_trees As Boolean
    money_grows_on_trees = False
    
    Dim number_1 As Integer
    number_1 = Range("C1").Value
    
    Range("E1").Value = number_1
    
    Range("E2").Value = Range("C2").Value
    
    
    
    

End Sub