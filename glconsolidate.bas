CLS

DIM ids(21) AS STRING
DIM outlets(21) AS STRING

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
ids(7) = "3704"
outlets(7) = "PWP"
ids(8) = "3705"
outlets(8) = "AMK"
ids(9) = "3706"
outlets(9) = "CLM"
ids(10) = "3707"
outlets(10) = "COP"
ids(11) = "3708"
outlets(11) = "DTE"
ids(12) = "3709"
outlets(12) = "IMM"
ids(13) = "3711"
outlets(13) = "NEX"
ids(14) = "3712"
outlets(14) = "PZS"
ids(15) = "3713"
outlets(15) = "SPC"
ids(16) = "3714"
outlets(16) = "SUN"
ids(17) = "3715"
outlets(17) = "TAM"
ids(18) = "3716"
outlets(18) = "TP"
ids(19) = "3717"
outlets(19) = "TP"
ids(20) = "3718"
outlets(20) = "WM"
ids(21) = "18002"
outlets(21) = "VCT"

INPUT "Date of file? (YYYYMMDD format, eg, 20221101)", FileDate$
INPUT "Directory?", FileDir$

OPEN "C:\" + FileDir$ + "\gl_" + FileDate$ + "_consolidated.csv" FOR OUTPUT AS #1

FOR i = 0 TO 21 STEP 1
    FileName$ = "C:\" + FileDir$ + "\gl_" + FileDate$ + "_" + ids(i) + ".csv"
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
