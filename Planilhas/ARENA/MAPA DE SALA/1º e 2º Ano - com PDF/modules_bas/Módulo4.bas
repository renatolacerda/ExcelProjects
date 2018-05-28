Attribute VB_Name = "Módulo4"
Sub ORDENA_TURMA3()
Attribute ORDENA_TURMA3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ORDENA_TURMA3 Macro
'

'
    Columns("A:E").Select
    ActiveWorkbook.Worksheets("BD").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BD").Sort.SortFields.Add Key:=Range("D1:D564"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("BD").Sort.SortFields.Add Key:=Range("E1:E564"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BD").Sort
        .SetRange Range("A1:E564")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub todasVisiveis()

For x = 1 To Worksheets.count

    Sheets(x).Visible = xlSheetVisible

Next
End Sub
