CLS

DIM ids(4) AS STRING
DIM outlets(4) AS STRING

ids(0) = "1122"
outlets(0) = "TUAS"
ids(1) = "1123"
outlets(1) = "PASIR RIS"
ids(2) = "1124"
outlets(2) = "BALESTIER"
ids(3) = "1125"
outlets(3) = "LITTLE INDIA"


INPUT "Date of file? (YYYYMMDD format, eg, 20221101)", FileDate$
INPUT "Directory?", FileDir$

OPEN "C:\" + FileDir$ + "\gl_" + FileDate$ + "_consolidated.csv" FOR OUTPUT AS #1

FOR i = 0 TO 20 STEP 1
    FileName$ = "C:\" + FileDir$ + "\gl_" + FileDate$ + "_" + ids(i) + ".csv"
    PRINT FileName$
    OPEN FileName$ FOR INPUT AS #2
    IF LOF(2) = 0 THEN
        PRINT outlets(i) + " OUTLET HAS NO DATA"
        KILL FileName$
    ELSE
        DO WHILE NOT EOF(2)
            INPUT #2, l1$, l2$, l3$, l4$, l5$
            InputStr$ = ""
            InputStr$ = InputStr$ + CHR$(34) + l1$ + CHR$(34) + ","
            InputStr$ = InputStr$ + CHR$(34) + l2$ + CHR$(34) + ","
            InputStr$ = InputStr$ + CHR$(34) + l3$ + CHR$(34) + ","
            InputStr$ = InputStr$ + CHR$(34) + l4$ + CHR$(34) + ","
            InputStr$ = InputStr$ + CHR$(34) + l5$ + CHR$(34) + ","
            InputStr$ = InputStr$ + CHR$(34) + outlets(i) + CHR$(34)
            PRINT #1, InputStr$
        LOOP
        CLOSE #2
    END IF
NEXT

CLOSE #1

