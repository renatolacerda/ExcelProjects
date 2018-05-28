Attribute VB_Name = "ARENA_HIDE_PLANILHAS"
Sub HIDE_PLANILHAS()
    For L = 3 To UltimaLinha(Sheets("CONFIG"), 3)
        visiveis = visiveis & ";" & Sheets("config").Cells(L, 3)
    Next
    
    visiveis = visiveis & ";CONFIG;BD;CONFIG-QTD;CONFIG-SALAS;Rel-Turma;Rel-Sala"
    
    If Left(visiveis, 1) = ";" Then visiveis = Mid(visiveis, 2, 999)
    
    v = Split(visiveis, ";")
    
    For x = 1 To Sheets.count
    
        visivel = 0
        For L = 0 To UBound(v)
            If v(L) = Sheets(x).name Then visivel = 1: Exit For
        Next
        If visivel = 1 Then
            Sheets(x).Visible = True
            visivel = 0
        Else
            Sheets(x).Visible = False
        End If
        
    Next
End Sub
