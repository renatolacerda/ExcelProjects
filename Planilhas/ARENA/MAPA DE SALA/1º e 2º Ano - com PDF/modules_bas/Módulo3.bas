Attribute VB_Name = "Módulo3"
Sub mudaParaEsquerda()
Attribute mudaParaEsquerda.VB_ProcData.VB_Invoke_Func = " \n14"
'
' mudaParaEsquerda Macro
'

'
    Range("AL13:AM13").Select
    Selection.Copy
    Range("E13").Select
    ActiveSheet.Paste
    Range("AL13:AM13").Select
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
    Selection.ClearContents
    Range("AN13").Select
    Selection.AutoFill Destination:=Range("AL13:AN13"), Type:=xlFillDefault
    Range("AL13:AN13").Select
    Range("AR17").Select
End Sub
