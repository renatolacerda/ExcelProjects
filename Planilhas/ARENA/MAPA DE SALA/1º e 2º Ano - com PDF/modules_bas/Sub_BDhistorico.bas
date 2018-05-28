Attribute VB_Name = "Sub_BDHISTORICO"
Sub HISTORICO_BD()
   Application.ScreenUpdating = False
   Application.CutCopyMode = False
    
    
    Dim p As Worksheet
    Dim H As Worksheet
    
    Set p = Sheets("BD")
    Set H = Sheets("BD-HISTORICO")
    
    If IsEmpty(H.Range("L2")) Then H.Range("L2") = 0
        
If H.Range("K2") > H.Range("L2") Then
    
    'P.Activate
    BDLINHA = UltimaLinha(p, 2)
    
    p.Range("A1", "E" & BDLINHA).Copy
    
    ULTIMA_LINHA_HISTORICO = UltimaLinha(H, 2)
    
    If ULTIMA_LINHA_HISTORICO <> 1 Then ULTIMA_LINHA_HISTORICO = ULTIMA_LINHA_HISTORICO + 1
        
    H.Activate
    Worksheets("BD-HISTORICO").Range("A" & ULTIMA_LINHA_HISTORICO).Select
    
    Selection.PasteSpecial xlPasteValues
    
    H.Columns("A:E").EntireColumn.AutoFit
    
    H.Range("F" & ULTIMA_LINHA_HISTORICO & ":F" & UltimaLinha(H, 2)) = Now()
    
    H.Columns("F:F").EntireColumn.AutoFit
        
    H.Range("L2") = H.Range("L2") + 1
    ElseIf H.Range("K2") <= H.Range("L2") Then
        H.Columns("A:F").ClearContents
        H.Range("L2") = 0
End If
    Application.ScreenUpdating = True
    Application.CutCopyMode = True
End Sub
