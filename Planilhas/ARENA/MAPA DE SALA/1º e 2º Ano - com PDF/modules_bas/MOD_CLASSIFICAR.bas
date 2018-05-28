Attribute VB_Name = "MOD_CLASSIFICAR"
Sub ORDENA_TURMA()
Attribute ORDENA_TURMA.VB_Description = "Macro gravada em 15/08/2007 por renato"
Attribute ORDENA_TURMA.VB_ProcData.VB_Invoke_Func = " \n14"
' ORDENAR POR TURMA
Sheets("BD").Activate
    Range(Cells(1, 1), Cells(Range("D65000").End(xlUp).Row, 5)).Select
    Selection.Sort Key1:=Range("C1"), Order1:=xlAscending, Key2:=Range("D1") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
Sheets("CONFIG").Activate
End Sub

Sub ORDENA_TURMA_RELATORIO_1()
Sheets("Rel-Turma").Activate
    Range("B12:E" & Range("J65000").End(xlUp).Row).Select
    Selection.Sort Key1:=Range("C13"), Order1:=xlAscending, Key2:=Range("E13" _
        ), Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
End Sub
Sub ORDENA_TURMA_RELATORIO_2()
Sheets("Rel-Turma").Activate
    Range("I12:J" & Range("J65000").End(xlUp).Row).Select
    Selection.Sort Key1:=Range("I13"), Order1:=xlAscending, Key2:=Range("J13" _
        ), Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
End Sub
Sub ORDENA_SALA_RELATORIO_1()
Sheets("Rel-Sala").Activate
    Range("B12:E" & Range("D65000").End(xlUp).Row).Select
    Selection.Sort Key1:=Range("D13"), Order1:=xlAscending, Key2:=Range("C13" _
        ), Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
End Sub

