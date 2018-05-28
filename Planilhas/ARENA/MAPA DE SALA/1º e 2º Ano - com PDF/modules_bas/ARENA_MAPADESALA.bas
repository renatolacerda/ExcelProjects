Attribute VB_Name = "ARENA_MAPADESALA"
Sub MAPA_DE_SALA()
    Application.ScreenUpdating = False
    
    Dim c As Worksheet
    
    Set c = Sheets("CONFIG")
    
    ' verifica se o tipo de modelo está vazio e adiciona ao modelo na coluna E
    For L = 3 To UltimaLinha(c, 3)
    
        c.Range("E" & L).FormulaR1C1 = "=IF(RC[-2]="""","""",SALAS(RC[-2]))"
    
    Next
            
    ' MISTURA OS ALUNOS - RANDOM
    If MsgBox("Você deseja misturar os alunos?", vbYesNo) = vbYes Then
        Call MAPA
    End If
    
    Application.ScreenUpdating = False
        
    DELETA_PLANILHAS ' DELETA AS PLANILHAS SE EXISTIREM
    
    FRM_MAPADESALA.Show 1
    
    CRIA_PLANILHAS
    ' ESTABELECE ONDE FICA CADA ALUNO
    ' CRIA_ESPACOS
    ' DISTRIBUIR_ALUNOS
        DISTRIBUIR_ALUNOS
    MsgBox "Processo finalizado...", vbInformation
    
    Sheets("CONFIG").Activate
    
    HIDE_PLANILHAS
    
    Sheets("backup").Range("B1") = "SIM"
        
    Application.ScreenUpdating = True
    
    ThisWorkbook.Save
    
End Sub
