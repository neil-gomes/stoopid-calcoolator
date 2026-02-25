$CONSOLE:ONLY
OPTION _EXPLICIT

DIM expression AS STRING
DIM result AS DOUBLE

IF _COMMANDCOUNT > 0 THEN
    expression = COMMAND$
    DIM idx AS LONG: idx = 1
    ParseAddSub expression, idx, result
    PRINT result
    SYSTEM
END IF

PRINT
PRINT "======================================"
PRINT "     Calcoolator Expression REPL      "
PRINT "======================================"
PRINT
PRINT "Operations: +, -, *, /, ^ (power)"
PRINT "Functions:  abs, sin, cos, tan, log, exp, sqrt"
PRINT "Constants:  pi, e"
PRINT "  Examples: 5+3, sin(pi/2), abs(-10), 2*pi*5"
PRINT "  Type 'exit' or 'quit' to exit"
PRINT

DO
    INPUT "calc> ", expression

    IF LCASE$(expression) = "exit" _ORELSE LCASE$(expression) = "quit" THEN
        EXIT DO
    END IF

    expression = LTRIM$(RTRIM$(expression))
    IF LEN(expression) = 0 THEN
        _CONTINUE
    END IF

    DIM idx2 AS LONG
    idx2 = 1
    ParseAddSub expression, idx2, result
    PRINT "=> "; result

    PRINT
LOOP

PRINT
PRINT "Thank you for using Calcoolator!"
PRINT

SYSTEM

SUB ParseAddSub (expr AS STRING, idx AS LONG, result AS DOUBLE)
    DIM leftVal AS DOUBLE
    DIM rightVal AS DOUBLE

    ParseMultiplyDivide expr, idx, leftVal
    result = leftVal
    
    DO WHILE idx <= LEN(expr)
        SkipSpaces expr, idx
        IF idx > LEN(expr) THEN EXIT DO
        IF ASC(expr, idx) = 43 THEN
            idx = idx + 1
            ParseMultiplyDivide expr, idx, rightVal
            result = result + rightVal
        ELSEIF ASC(expr, idx) = 45 _ANDALSO idx > 1 THEN
            idx = idx + 1
            ParseMultiplyDivide expr, idx, rightVal
            result = result - rightVal
        ELSE
            EXIT DO
        END IF
    LOOP
END SUB

SUB ParseMultiplyDivide (expr AS STRING, idx AS LONG, result AS DOUBLE)
    DIM leftVal AS DOUBLE
    DIM rightVal AS DOUBLE
    
    ParsePower expr, idx, leftVal
    result = leftVal
    
    DO WHILE idx <= LEN(expr)
        SkipSpaces expr, idx
        IF idx > LEN(expr) THEN EXIT DO
        IF ASC(expr, idx) = 42 THEN
            idx = idx + 1
            ParsePower expr, idx, rightVal
            result = result * rightVal
        ELSEIF ASC(expr, idx) = 47 THEN
            idx = idx + 1
            ParsePower expr, idx, rightVal
            IF rightVal <> 0 THEN
                result = result / rightVal
            ELSE
                PRINT "Error: Division by zero"
            END IF
        ELSE
            EXIT DO
        END IF
    LOOP
END SUB

SUB ParsePower (expr AS STRING, idx AS LONG, result AS DOUBLE)
    DIM leftVal AS DOUBLE
    DIM rightVal AS DOUBLE
    
    ParseUnary expr, idx, leftVal
    result = leftVal
    
    IF idx <= LEN(expr) THEN
        SkipSpaces expr, idx
    END IF
    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 94 THEN
        idx = idx + 1
        ParsePower expr, idx, rightVal
        result = leftVal ^ rightVal
    END IF
END SUB

SUB ParseUnary (expr AS STRING, idx AS LONG, result AS DOUBLE)
    SkipSpaces expr, idx
    
    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 45 THEN
        idx = idx + 1
        ParsePrimary expr, idx, result
        result = -result
    ELSEIF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 43 THEN
        idx = idx + 1
        ParsePrimary expr, idx, result
    ELSE
        ParsePrimary expr, idx, result
    END IF
END SUB

SUB ParsePrimary (expr AS STRING, idx AS LONG, result AS DOUBLE)
    DIM innerVal AS DOUBLE, testStr AS STRING
    
    SkipSpaces expr, idx
    
    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
        idx = idx + 1
        ParseAddSub expr, idx, result
        SkipSpaces expr, idx
        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 41 THEN
            idx = idx + 1
        END IF
    ELSE
        testStr = LCASE$(MID$(expr, idx, 4))
        
        IF testStr = "sqrt" THEN
            idx = idx + 4
            SkipSpaces expr, idx
            IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
                idx = idx + 1
                ParseAddSub expr, idx, innerVal
                result = SQR(ABS(innerVal))
                IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 41 THEN
                    idx = idx + 1
                END IF
            END IF
        ELSE
            testStr = LCASE$(MID$(expr, idx, 3))
            SELECT CASE testStr
                CASE "sin"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
                        idx = idx + 1
                        ParseAddSub expr, idx, innerVal
                        result = SIN(innerVal)
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 41 THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "cos"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
                        idx = idx + 1
                        ParseAddSub expr, idx, innerVal
                        result = COS(innerVal)
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 41 THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "tan"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
                        idx = idx + 1
                        ParseAddSub expr, idx, innerVal
                        result = TAN(innerVal)
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 41 THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "log"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
                        idx = idx + 1
                        ParseAddSub expr, idx, innerVal
                        IF innerVal > 0 THEN
                            result = LOG(innerVal)
                        ELSE
                            result = 0
                        END IF
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 41 THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "exp"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
                        idx = idx + 1
                        ParseAddSub expr, idx, innerVal
                        result = EXP(innerVal)
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 41 THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "abs"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
                        idx = idx + 1
                        ParseAddSub expr, idx, innerVal
                        result = ABS(innerVal)
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 41 THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE ELSE
                    IF MID$(LCASE$(expr), idx, 2) = "pi" THEN
                        result = 3.14159265358979
                        idx = idx + 2
                    ELSEIF LEFT$(LCASE$(expr), 1) = "e" _ANDALSO (idx + 1 > LEN(expr) _ORELSE NOT ((ASC(expr, idx + 1) >= 48 _ANDALSO ASC(expr, idx + 1) <= 57) _ORELSE ASC(expr, idx + 1) = 46)) THEN
                        result = 2.71828182845905
                        idx = idx + 1
                    ELSE
                        ParseNumber expr, idx, result
                    END IF
            END SELECT
        END IF
    END IF
END SUB

SUB ParseNumber (expr AS STRING, idx AS LONG, result AS DOUBLE)
    DIM numStr AS STRING
    
    SkipSpaces expr, idx

    WHILE idx <= LEN(expr) _ANDALSO ((ASC(expr, idx) >= 48 _ANDALSO ASC(expr, idx) <= 57) _ORELSE ASC(expr, idx) = 46)
        numStr = numStr + CHR$(ASC(expr, idx))
        idx = idx + 1
    WEND

    IF LEN(numStr) > 0 THEN
        result = VAL(numStr)
    ELSE
        result = 0
    END IF
END SUB

SUB SkipSpaces (expr AS STRING, idx AS LONG)
    WHILE idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 32
        idx = idx + 1
    WEND
END SUB
