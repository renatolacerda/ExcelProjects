Attribute VB_Name = "ARENA_RELATORIO_SALA"
Sub RELATORIO_SALA_SEGUNDA_LISTA()
    ActiveSheet.PageSetup.PrintArea = "$C$3:$J$" & Range("B65000").End(xlUp).Row
    'DESOCULTA
    Columns("A:C").Select
    Selection.EntireColumn.Hidden = False '.ColumnWidth = 11.43
    Columns("B:B").Select
    Selection.NumberFormat = "0000000"
    'ActiveSheet.PageSetup.PrintArea = "$B$3:$E$" & Range("B65000").End(xlUp).Row
    'OCULTA
    'Columns("B:B").Select
    'Selection.EntireColumn.Hidden = True
End Sub
