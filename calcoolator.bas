$DEBUG
'------------------------------------------------------------------------------------------------
' This is a test program
'------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------
OPTION EXPLICIT
OPTION BASE 1
'$STATIC
'------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------
' Program entry point
'------------------------------------------------------------------------------------------------
DIM num1 AS SINGLE
DIM num2 AS SINGLE
DIM userselection1 AS STRING
DIM useranswer AS STRING

CLS
PRINT "Hi! this is a simple calculator to do arithmetic operations. thanks for seeing this!"
PRINT
DO
    INPUT ; "enter first number"; num1
    PRINT
    INPUT ; "enter second number"; num2
    PRINT
    PRINT "A) addition"
    PRINT "S) subtraction"
    PRINT "D) Division "
    PRINT "M) multiplication "
    PRINT "P) power"
    PRINT
    INPUT ; "enter arithmetic operator"; userselection1
    PRINT
    userselection1 = LCASE$(userselection1)

    SELECT CASE userselection1
        CASE "a"
            PRINT num1 + num2
        CASE "s"
            PRINT num1 - num2
        CASE "d"
            PRINT num1 / num2
        CASE "m"
            PRINT num1 * num2
        CASE "p"
            PRINT num1 ^ num2
    END SELECT

    INPUT "do you want to try again (y/n)", useranswer
    useranswer = UCASE$(useranswer)

LOOP WHILE useranswer = "Y"

END
