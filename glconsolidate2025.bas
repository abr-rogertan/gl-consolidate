Cls

Dim ids(26) As String
Dim outlets(26) As String

ids(0) = "3544"
outlets(0) = "BEM"
ids(1) = "3684"
outlets(1) = "WWP"
ids(2) = "3685"
outlets(2) = "CWP"
ids(3) = "3686"
outlets(3) = "JP"
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
ids(21) = "17422"
outlets(21) = "AT2"
ids(21) = "18002"
outlets(21) = "VCT"
ids(21) = "18222"
outlets(21) = "PRM"
ids(22) = "21043"
outlets(22) = "TGP"
ids(23) = "18002"
outlets(23) = "VCT"
ids(24) = "17422"
outlets(24) = "AT2"
ids(25) = "21783"
outlets(25) = "TM3"
ids(26) = "25286"
outlets(26) = "STP"

Input "Date of file? (YYYYMMDD format, eg, 20221101)", FileDate$
Input "Directory?", FileDir$

Open "C:\" + FileDir$ + "\gl_" + FileDate$ + "_consolidated.csv" For Output As #1

For i = 0 To 26 Step 1
    FileName$ = "C:\" + FileDir$ + "\gl_" + FileDate$ + "_" + ids(i) + ".csv"
    Print FileName$
    Open FileName$ For Input As #2
    If LOF(2) = 0 Then
        Print outlets(i) + " OUTLET HAS NO DATA"
        Kill FileName$
    Else
        Do While Not EOF(2)
            Input #2, COL1$, COL2$, COL3$, COL4$, COL5$, COL6$, COL7$, COL8$, COL9$, COL10$, COL11$, COL12$, COL13$
            InputStr$ = ""
            InputStr$ = InputStr$ + Chr$(34) + COL1$ + Chr$(34) + ","
            InputStr$ = InputStr$ + Chr$(34) + COL2$ + Chr$(34) + ","
            InputStr$ = InputStr$ + Chr$(34) + COL3$ + Chr$(34) + ","
            InputStr$ = InputStr$ + Chr$(34) + COL4$ + Chr$(34) + ","
            InputStr$ = InputStr$ + Chr$(34) + COL5$ + Chr$(34) + ","
            InputStr$ = InputStr$ + Chr$(34) + COL6$ + Chr$(34) + ","
            InputStr$ = InputStr$ + Chr$(34) + COL7$ + Chr$(34) + ","
            InputStr$ = InputStr$ + Chr$(34) + COL8$ + Chr$(34) + ","
            InputStr$ = InputStr$ + Chr$(34) + COL9$ + Chr$(34) + ","
            InputStr$ = InputStr$ + Chr$(34) + COL10$ + Chr$(34) + ","
            InputStr$ = InputStr$ + Chr$(34) + COL11$ + Chr$(34) + ","
            InputStr$ = InputStr$ + Chr$(34) + COL12$ + Chr$(34) + ","
            InputStr$ = InputStr$ + Chr$(34) + COL13$ + Chr$(34)
            Print #1, InputStr$
        Loop
        Close #2
    End If
Next

Close #1

