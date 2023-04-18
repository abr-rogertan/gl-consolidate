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

