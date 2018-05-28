Attribute VB_Name = "Módulo1"
Sub ORGANIZA_POR_SALA()
Attribute ORGANIZA_POR_SALA.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ORGANIZA_POR_SALA Macro
'
'
Columns("A:E").Select
    ActiveWorkbook.Worksheets("BD").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BD").Sort.SortFields.Add Key:=Range("A1:A" & Sheets("BD").Range("D65000").End(xlUp).Row), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BD").Sort
        .SetRange Range("A1:E" & Sheets("BD").Range("D65000").End(xlUp).Row)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
