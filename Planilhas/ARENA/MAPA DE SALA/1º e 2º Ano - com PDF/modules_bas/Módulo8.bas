Attribute VB_Name = "Módulo8"

Sub formata_sl3_novo()
Attribute formata_sl3_novo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' formata_sl3_novo Macro
'

'
    Range("AK24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W19:AG21").Select
    ActiveSheet.Paste
    Range("AF23:AG33").Select
    ActiveSheet.Paste
End Sub
Sub formata_sl4_novo()
Attribute formata_sl4_novo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' formata_sl4_novo Macro
'

'
    Columns("W:X").Select
    Selection.ClearContents
    Selection.EntireColumn.Hidden = True
    Range("AK28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AF31:AG33").Select
    ActiveSheet.Paste
End Sub
Sub formata_sl5_novo()
Attribute formata_sl5_novo.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("AL27").Select
    Selection.Copy
    Range("AF15:AG17").Select
    ActiveSheet.Paste
    Range("W19:AG21").Select
    ActiveSheet.Paste
    Range("W23:AG25").Select
    ActiveSheet.Paste
    Range("AF27:AG29").Select
    ActiveSheet.Paste
End Sub
Sub formata_sl2_novo()
Attribute formata_sl2_novo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' formata_sl2_novo Macro
'

    Columns("X:AF").Select
    Selection.EntireColumn.Hidden = False
    Columns("AC:AE").Select
    Selection.EntireColumn.Hidden = True
    Range("T15:U33").Select
    Selection.Copy
    Range("Z15:AA16").Select
    ActiveSheet.Paste
    Range("AF15:AG30").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("AK23").Select
    Selection.Copy
    Range("AF15:AG29").Select
    ActiveSheet.Paste
End Sub
Sub formata_sl1_novo()
Attribute formata_sl1_novo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' formata_sl1_novo Macro
'

'
    Range("AL24").Select
    Selection.Copy
    Range("AF15:AG17").Select
    ActiveSheet.Paste
    Range("W19:AG25").Select
    Range("AF19").Activate
    ActiveSheet.Paste
    Range("AF27:AG29").Select
    ActiveSheet.Paste
End Sub

