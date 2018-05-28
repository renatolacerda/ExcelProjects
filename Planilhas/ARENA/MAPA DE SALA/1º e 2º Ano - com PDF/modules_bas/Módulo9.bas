Attribute VB_Name = "Módulo9"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("D17").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=COUNTA(BD!R[-16]C[-2]:R[182]C[-2])"
    Range("D18").Select
End Sub
