Attribute VB_Name = "ARENA_AREA_IMPRESSAO"
Sub AREA_IMPRESSAO_SALA()
    'Range("B1:E" & Range("B65000").End(xlUp).Row).Select
    ActiveSheet.PageSetup.PrintArea = "$B$1:$E$" & Range("B65000").End(xlUp).Row
End Sub
Sub AREA_IMPRESSAO_TURMA1()
    ActiveSheet.PageSetup.PrintArea = "$B$1:$E$" & Range("B65000").End(xlUp).Row
End Sub
Sub AREA_IMPRESSAO_TURMA2()
    ActiveSheet.PageSetup.PrintArea = "$H$1:$J$" & Range("B65000").End(xlUp).Row
End Sub
Sub AREA_IMPRESSAO_TURMA3()
    ActiveSheet.PageSetup.PrintArea = "$B$3:$E$" & Range("B65000").End(xlUp).Row
    'OCULTA
    Columns("B:B").Select
    Selection.EntireColumn.Hidden = True
    
    
End Sub
Sub AREA_IMPRESSAO_TURMA4()
    ActiveSheet.PageSetup.PrintArea = "$C$3:$J$" & Range("B65000").End(xlUp).Row
    'DESOCULTA
    Columns("A:C").Select
    Selection.EntireColumn.Hidden = False '.ColumnWidth = 11.43
    Columns("B:B").Select
    Selection.NumberFormat = "0000000"
End Sub
