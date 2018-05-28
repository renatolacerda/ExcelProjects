Attribute VB_Name = "ARENA_ORDENA_SALA"
Sub ORDENA_SALA()
Sheets("BD").Activate
    Range(Cells(1, 1), Cells(Range("D65000").End(xlUp).Row, 5)).Select
    Selection.Sort Key1:=Range("E1"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
Sheets("CONFIG").Activate
End Sub

