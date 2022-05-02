$Debug
'------------------------------------------------------------------------------------------------
'
' This is a test program. This can be used as a template for writing new source files
' Copyright (c) 2022, Samuel Gomes
'
'------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------
$NoPrefix
Option Explicit
Option ExplicitArray
Option Base 1
'$Static
$Resize:Smooth
'------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------
' Program entry point
'------------------------------------------------------------------------------------------------

Dim num1 As Single
Dim num2 As Single
Dim userselection1 As String
Dim useranswer As String

Cls
Print "Hi! this is a simple calculator to do arithmetic operations. thanks for seeing this!"
Print
Do
    Input ; "enter first number"; num1
    Print
    Input ; "enter second number"; num2
    Print
    Print "A) addition"
    Print "S) subtraction"
    Print "D) Division "
    Print "M) multiplication "
    Print "P) power"
    Print
    Input ; "enter arithmetic operator"; userselection1
    Print
    userselection1 = LCase$(userselection1)

    Select Case userselection1
        Case "a"
            Print num1 + num2
        Case "s"
            Print num1 - num2
        Case "d"
            Print num1 / num2
        Case "m"
            Print num1 * num2
        Case "p"
        print num1 ^ num2
    End Select

    Input "do you want to try again (y/n)", useranswer
    useranswer = UCase$(useranswer)

Loop While useranswer = "Y"



