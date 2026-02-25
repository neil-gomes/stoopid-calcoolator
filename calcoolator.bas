$CONSOLE:ONLY
OPTION _EXPLICIT

DIM expression AS STRING
DIM result AS DOUBLE

IF _COMMANDCOUNT > 0 THEN
    expression = COMMAND$
    DIM idx AS LONG: idx = 1
    result = ParseAddSub#(expression, idx)
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
    result = ParseAddSub#(expression, idx2)
    PRINT "=> "; result

    PRINT
LOOP

PRINT
PRINT "Thank you for using Calcoolator!"
PRINT

SYSTEM

FUNCTION ParseAddSub# (expr AS STRING, idx AS LONG)
    DIM result AS DOUBLE

    result = ParseMultiplyDivide#(expr, idx)
    
    DO WHILE idx <= LEN(expr)
        SkipSpaces expr, idx
        IF idx > LEN(expr) THEN EXIT DO
        IF ASC(expr, idx) = 43 THEN
            idx = idx + 1
            result = result + ParseMultiplyDivide#(expr, idx)
        ELSEIF ASC(expr, idx) = 45 _ANDALSO idx > 1 THEN
            idx = idx + 1
            result = result - ParseMultiplyDivide#(expr, idx)
        ELSE
            EXIT DO
        END IF
    LOOP
    
    ParseAddSub# = result
END FUNCTION

FUNCTION ParseMultiplyDivide# (expr AS STRING, idx AS LONG)
    DIM result AS DOUBLE
    DIM rightVal AS DOUBLE
    
    result = ParsePower#(expr, idx)
    
    DO WHILE idx <= LEN(expr)
        SkipSpaces expr, idx
        IF idx > LEN(expr) THEN EXIT DO
        IF ASC(expr, idx) = 42 THEN
            idx = idx + 1
            rightVal = ParsePower#(expr, idx)
            result = result * rightVal
        ELSEIF ASC(expr, idx) = 47 THEN
            idx = idx + 1
            rightVal = ParsePower#(expr, idx)
            IF rightVal <> 0 THEN
                result = result / rightVal
            ELSE
                PRINT "Error: Division by zero"
            END IF
        ELSE
            EXIT DO
        END IF
    LOOP
    
    ParseMultiplyDivide# = result
END FUNCTION

FUNCTION ParsePower# (expr AS STRING, idx AS LONG)
    DIM result AS DOUBLE
    DIM rightVal AS DOUBLE
    
    result = ParseUnary#(expr, idx)
    
    IF idx <= LEN(expr) THEN
        SkipSpaces expr, idx
    END IF
    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 94 THEN
        idx = idx + 1
        rightVal = ParsePower#(expr, idx)
        result = result ^ rightVal
    END IF
    
    ParsePower# = result
END FUNCTION

FUNCTION ParseUnary# (expr AS STRING, idx AS LONG)
    DIM result AS DOUBLE
    
    SkipSpaces expr, idx
    
    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 45 THEN
        idx = idx + 1
        result = -ParsePrimary#(expr, idx)
    ELSEIF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 43 THEN
        idx = idx + 1
        result = ParsePrimary#(expr, idx)
    ELSE
        result = ParsePrimary#(expr, idx)
    END IF
    
    ParseUnary# = result
END FUNCTION

FUNCTION ParsePrimary# (expr AS STRING, idx AS LONG)
    DIM result AS DOUBLE
    DIM innerVal AS DOUBLE
    DIM testStr AS STRING
    
    SkipSpaces expr, idx
    
    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
        idx = idx + 1
        result = ParseAddSub#(expr, idx)
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
                innerVal = ParseAddSub#(expr, idx)
                result = SQR(ABS(innerVal))
                SkipSpaces expr, idx
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
                        innerVal = ParseAddSub#(expr, idx)
                        result = SIN(innerVal)
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 41 THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "cos"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        result = COS(innerVal)
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 41 THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "tan"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        result = TAN(innerVal)
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 41 THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "log"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        IF innerVal > 0 THEN
                            result = LOG(innerVal)
                        ELSE
                            result = 0
                        END IF
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 41 THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "exp"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        result = EXP(innerVal)
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 41 THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "abs"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 40 THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        result = ABS(innerVal)
                        SkipSpaces expr, idx
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
                        result = ParseNumber#(expr, idx)
                    END IF
            END SELECT
        END IF
    END IF
    
    ParsePrimary# = result
END FUNCTION

FUNCTION ParseNumber# (expr AS STRING, idx AS LONG)
    DIM numStr AS STRING
    DIM result AS DOUBLE
    
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
    
    ParseNumber# = result
END FUNCTION

SUB SkipSpaces (expr AS STRING, idx AS LONG)
    WHILE idx <= LEN(expr) _ANDALSO ASC(expr, idx) = 32
        idx = idx + 1
    WEND
END SUB
