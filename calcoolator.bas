$CONSOLE:ONLY
OPTION _EXPLICIT

DIM expression AS STRING, idx AS LONG

IF _COMMANDCOUNT > 0 THEN
    expression = COMMAND$
    IF NOT ValidateParentheses(expression) THEN
        PRINT "Error: Unbalanced parentheses"
        SYSTEM
    END IF
    idx = 1
    PRINT ParseAddSub#(expression, idx)
    SYSTEM
END IF

PRINT
PRINT "======================================"
PRINT "     Calcoolator Expression REPL      "
PRINT "======================================"
PRINT
PRINT "Operations: +, -, *, /, ^ (power)"
PRINT "Functions:  abs, sin, cos, tan (radians), log (base 10), ln, exp, sqrt"
PRINT "Conversion: d2r (degrees to radians), r2d (radians to degrees)"
PRINT "Constants:  pi, e"
PRINT "  Examples: sin(pi/2), sin(d2r(90)), log(100), e^2, 2*pi"
PRINT "  Type 'exit' or 'quit' to exit"
PRINT

DO
    INPUT "calc> ", expression

    IF LCASE$(expression) = "exit" _ORELSE LCASE$(expression) = "quit" THEN
        EXIT DO
    END IF

    expression = _TRIM$(expression)
    IF LEN(expression) = 0 THEN
        _CONTINUE
    END IF

    IF NOT ValidateParentheses(expression) THEN
        PRINT "Error: Unbalanced parentheses"
        _CONTINUE
    END IF

    idx = 1
    PRINT "=> "; ParseAddSub#(expression, idx)

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
        IF ASC(expr, idx) = _ASC_PLUS THEN
            idx = idx + 1
            result = result + ParseMultiplyDivide#(expr, idx)
        ELSEIF ASC(expr, idx) = _ASC_MINUS _ANDALSO idx > 1 THEN
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
        IF ASC(expr, idx) = _ASC_ASTERISK THEN
            idx = idx + 1
            rightVal = ParsePower#(expr, idx)
            result = result * rightVal
        ELSEIF ASC(expr, idx) = _ASC_FORWARDSLASH THEN
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
    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_CARET THEN
        idx = idx + 1
        rightVal = ParsePower#(expr, idx)
        result = result ^ rightVal
    END IF
    
    ParsePower# = result
END FUNCTION

FUNCTION ParseUnary# (expr AS STRING, idx AS LONG)
    DIM result AS DOUBLE
    
    SkipSpaces expr, idx
    
    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_MINUS THEN
        idx = idx + 1
        result = -ParsePrimary#(expr, idx)
    ELSEIF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_PLUS THEN
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
    
    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_LEFTBRACKET THEN
        idx = idx + 1
        result = ParseAddSub#(expr, idx)
        SkipSpaces expr, idx
        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_RIGHTBRACKET THEN
            idx = idx + 1
        END IF
    ELSE
        testStr = LCASE$(MID$(expr, idx, 4))
        
        IF testStr = "sqrt" THEN
            idx = idx + 4
            SkipSpaces expr, idx
            IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_LEFTBRACKET THEN
                idx = idx + 1
                innerVal = ParseAddSub#(expr, idx)
                result = SQR(ABS(innerVal))
                SkipSpaces expr, idx
                IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_RIGHTBRACKET THEN
                    idx = idx + 1
                END IF
            END IF
        ELSE
            testStr = LCASE$(MID$(expr, idx, 3))
            
            IF MID$(LCASE$(expr), idx, 2) = "ln" THEN
                testStr = "ln"
            END IF
            
            SELECT CASE testStr
                CASE "sin"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_LEFTBRACKET THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        result = SIN(innerVal)
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_RIGHTBRACKET THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "cos"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_LEFTBRACKET THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        result = COS(innerVal)
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_RIGHTBRACKET THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "tan"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_LEFTBRACKET THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        result = TAN(innerVal)
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_RIGHTBRACKET THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "log"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_LEFTBRACKET THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        IF innerVal > 0 THEN
                            result = LOG(innerVal) / LOG(10)
                        ELSE
                            result = 0
                        END IF
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_RIGHTBRACKET THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "ln"
                    idx = idx + 2
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_LEFTBRACKET THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        IF innerVal > 0 THEN
                            result = LOG(innerVal)
                        ELSE
                            result = 0
                        END IF
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_RIGHTBRACKET THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "exp"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_LEFTBRACKET THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        result = EXP(innerVal)
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_RIGHTBRACKET THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "abs"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_LEFTBRACKET THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        result = ABS(innerVal)
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_RIGHTBRACKET THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "d2r"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_LEFTBRACKET THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        result = _D2R(innerVal)
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_RIGHTBRACKET THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE "r2d"
                    idx = idx + 3
                    SkipSpaces expr, idx
                    IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_LEFTBRACKET THEN
                        idx = idx + 1
                        innerVal = ParseAddSub#(expr, idx)
                        result = _R2D(innerVal)
                        SkipSpaces expr, idx
                        IF idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_RIGHTBRACKET THEN
                            idx = idx + 1
                        END IF
                    END IF
                CASE ELSE
                    IF MID$(LCASE$(expr), idx, 2) = "pi" THEN
                        result = _PI
                        idx = idx + 2
                    ELSEIF ASC(expr, idx) = ASC("e") _ANDALSO (idx + 1 > LEN(expr) _ORELSE NOT ((ASC(expr, idx + 1) >= ASC("0") _ANDALSO ASC(expr, idx + 1) <= ASC("9")) _ORELSE ASC(expr, idx + 1) = _ASC_FULLSTOP)) THEN
                        result = _E
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

    WHILE idx <= LEN(expr) _ANDALSO ((ASC(expr, idx) >= ASC("0") _ANDALSO ASC(expr, idx) <= ASC("9")) _ORELSE ASC(expr, idx) = _ASC_FULLSTOP)
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
    WHILE idx <= LEN(expr) _ANDALSO ASC(expr, idx) = _ASC_SPACE
        idx = idx + 1
    WEND
END SUB

FUNCTION ValidateParentheses (expr AS STRING)
    DIM i AS LONG
    DIM depth AS LONG
    
    depth = 0
    FOR i = 1 TO LEN(expr)
        IF ASC(expr, i) = _ASC_LEFTBRACKET THEN
            depth = depth + 1
        ELSEIF ASC(expr, i) = _ASC_RIGHTBRACKET THEN
            depth = depth - 1
            IF depth < 0 THEN
                ValidateParentheses = 0
                EXIT FUNCTION
            END IF
        END IF
    NEXT i
    
    IF depth <> 0 THEN
        ValidateParentheses = 0
    ELSE
        ValidateParentheses = -1
    END IF
END FUNCTION
