Attribute VB_Name = "ARENA_AJUSTE_CARTEIRAS"
Sub AJUSTE_CARTEIRAS()
    'Dim CONFIG As Worksheet,
    Dim BD As Worksheet, CONFIG_SALAS As Worksheet
    Dim REDUZIR_ALUNOS As Integer, uLin As Integer
    Dim RETIRAR_QTD_ALUNOS As Integer, ADICIONAR_CARTEIRAS As Integer
    
    'Set CONFIG = Sheets("CONFIG-QTD")
    Set BD = Sheets("BD")
    Set CONFIG_SALAS = Sheets("CONFIG-SALAS")
    
    uLin = UltimaLinha(CONFIG_SALAS, 2)
    TOTAL_CARTEIRAS_CONFIG = WorksheetFunction.Sum(CONFIG_SALAS.Range("C2:C" & uLin))
    
    uLin = UltimaLinha(BD, 2)
    TOTAL_ALUNOS_BD = WorksheetFunction.CountA(BD.Range("B1:B" & uLin))
    
    If TOTAL_CARTEIRAS_CONFIG > TOTAL_ALUNOS_BD Then
        RETIRAR_QTD_ALUNOS = TOTAL_CARTEIRAS_CONFIG - TOTAL_ALUNOS_BD
repetir:
        uLin = UltimaLinha(CONFIG_SALAS, 2)
        For linha = 2 To uLin
            CONFIG_SALAS.Range("C" & linha).Value = CONFIG_SALAS.Range("C" & linha).Value - 1
            RETIRAR_QTD_ALUNOS = RETIRAR_QTD_ALUNOS - 1
            If RETIRAR_QTD_ALUNOS = 0 Then Exit For
        Next
        If RETIRAR_QTD_ALUNOS > 0 Then GoTo repetir
        
    ElseIf TOTAL_CARTEIRAS_CONFIG < TOTAL_ALUNOS_BD Then
        'MENSAGEM A QTD DE ALUNOS É MAIOR QUE A QTD DE CARTEIRAS
        MsgBox "A QTD DE ALUNOS É MAIOR QUE A QTD DE CARTEIRAS COLOCAR MAIS CARTEIRAS", vbCritical
        ADICIONAR_CARTEIRAS = TOTAL_ALUNOS_BD - TOTAL_CARTEIRAS_CONFIG
        
        MsgBox "AJUSTAR MANUALMENTE A QUANTIDADE DE CARTEIRAS NA ABA [CONFIG-SALAS]", vbInformation
        
        Sheets("CONFIG-QTD").Activate
        
        Exit Sub
        
    End If
    
    
End Sub
