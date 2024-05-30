CLS

DIM ids(4) AS STRING
DIM outlets(4) AS STRING

ids(0) = "3544"
outlets(0) = "BEM"
ids(1) = "3684"
outlets(1) = "WWP"
ids(2) = "3685"
outlets(2) = "CWP"
ids(3) = "3686"
outlets(3) = "JRP"
ids(4) = "3687"
outlets(4) = "J8"
ids(5) = "3702"
outlets(5) = "NP"
ids(6) = "3703"
outlets(6) = "BJ"
ids(7) = "3705"
outlets(7) = "AMK"
ids(8) = "3706"
outlets(8) = "CLM"
ids(9) = "3707"
outlets(9) = "COP"
ids(10) = "3708"
outlets(10) = "DTE"
ids(11) = "3709"
outlets(11) = "IMM"
ids(12) = "3711"
outlets(12) = "NEX"
ids(13) = "3712"
outlets(13) = "PZS"
ids(14) = "3713"
outlets(14) = "SPC"
ids(15) = "3714"
outlets(15) = "SUN"
ids(16) = "3715"
outlets(16) = "TAM"
ids(17) = "3717"
outlets(17) = "TP"
ids(18) = "3718"
outlets(18) = "WM"
ids(x) = "0000"
outlets(x) = "xx"
ids(x) = "0000"
outlets(x) = "xx"
ids(x) = "0000"
outlets(x) = "xx"
ids(x) = "0000"
outlets(x) = "xx"
ids(x) = "0000"
outlets(x) = "xx"
ids(x) = "0000"
outlets(x) = "xx"
ids(x) = "0000"
outlets(x) = "xx"
ids(x) = "0000"
outlets(x) = "xx"
ids(x) = "0000"
outlets(x) = "xx"


INPUT "Date of file? (YYYYMMDD format, eg, 20221101)", FileDate$
INPUT "Directory?", FileDir$

OPEN "C:\" + FileDir$ + "\tx_" + FileDate$ + "_consolidated.csv" FOR OUTPUT AS #1

FOR i = 0 TO 3 STEP 1
    FileName$ = "C:\" + FileDir$ + "\tx_" + FileDate$ + "_" + ids(i) + ".csv"
    PRINT FileName$
    OPEN FileName$ FOR INPUT AS #2
    IF LOF(2) = 0 THEN
        PRINT outlets(i) + " OUTLET HAS NO DATA"
        KILL FileName$
    ELSE
        DO WHILE NOT EOF(2)
            INPUT #2, COLDATE$, COLTIME$, COLNUM$, COLITEM$, COLCAT$, COLQTY$, COLTYPE$
            InputStr$ = ""
            InputStr$ = InputStr$ + CHR$(34) + COLDATE$ + CHR$(34) + ","
            InputStr$ = InputStr$ + CHR$(34) + COLTIME$ + CHR$(34) + ","
            InputStr$ = InputStr$ + CHR$(34) + COLNUM$ + CHR$(34) + ","
            InputStr$ = InputStr$ + CHR$(34) + COLITEM$ + CHR$(34) + ","
            InputStr$ = InputStr$ + CHR$(34) + COLCAT$ + CHR$(34) + ","
            InputStr$ = InputStr$ + CHR$(34) + COLQTY$ + CHR$(34) + ","
            InputStr$ = InputStr$ + CHR$(34) + COLTYPE$ + CHR$(34) + ","
            InputStr$ = InputStr$ + CHR$(34) + outlets(i) + CHR$(34)
            PRINT #1, InputStr$
        LOOP
        CLOSE #2
    END IF
NEXT

CLOSE #1
