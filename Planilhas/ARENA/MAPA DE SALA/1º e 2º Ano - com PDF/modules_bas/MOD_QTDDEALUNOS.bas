Attribute VB_Name = "MOD_QTDDEALUNOS"
Public CONTAR
Sub LEVANTAMENTO_DE_DADOS()
PLAN_ATUAL = ActiveSheet.name
Sheets("CONFIG").Activate
Sheets("BD").Range("D:F").ClearContents
CONTA_ALUNOS_DB ' OK!
QTD_TURMAS ' OK!
MEDIA ' OK!
QTD_SALAS ' OK!
'VAGAS_POR_SALA
Permutation
ORDENA_TURMA
ENTURMA_ALUNOS_POR_SALA
Sheets(PLAN_ATUAL).Activate
End Sub
Sub MEDIA()
For ACHA_COLUNA = 1 To Range("IV1").End(xlToLeft).Column
    If Cells(1, ACHA_COLUNA) = "TURMA" Then
        Exit For
    End If
Next
Range(Cells(2, ACHA_COLUNA + 2), Cells(65000, ACHA_COLUNA + 2)).ClearContents
For x = 2 To Cells(65000, ACHA_COLUNA).End(xlUp).Row ' SOMA POR GRUPO (VARRE TODAS AS TURMAS)
    Select Case Left(Cells(x, ACHA_COLUNA), 1) ' faz a média dos valores estabelecendo o teto para poder criar os espaços
        Case 1
            total1 = total1 + Cells(x, ACHA_COLUNA + 1)
            qtd1 = qtd1 + 1
        Case 2
            total1 = total1 + Cells(x, ACHA_COLUNA + 1)
            qtd1 = qtd1 + 1
        Case 3
            total3 = total3 + Cells(x, ACHA_COLUNA + 1)
            qtd3 = qtd3 + 1
    End Select
Next
    'total1 = total1 / qtd1
    total3 = total3 / qtd3
    'If Int(total1) <> total1 Then total1 = Int(total1) + 1
    If Int(total3) <> total3 Then total3 = Int(total3) + 1
For x = 2 To Cells(65000, ACHA_COLUNA).End(xlUp).Row ' ESTABELECE (QTD-ALUNOS-POR-SALA)
    Select Case Left(Cells(x, ACHA_COLUNA), 1)
        Case 1
           Cells(x, ACHA_COLUNA + 2) = total1
        Case 2
           Cells(x, ACHA_COLUNA + 2) = total1
        Case 3
            Cells(x, ACHA_COLUNA + 2) = total3
    End Select
Next
End Sub
Sub CONTA_ALUNOS_DB()
d = "BD"
c = "CONFIG"
For ACHA_COLUNA = 1 To Range("IV1").End(xlToLeft).Column
    If Cells(1, ACHA_COLUNA) = "TOTAL-BD" Then
        C_TOTALBD = ACHA_COLUNA
    End If
    If Cells(1, ACHA_COLUNA) = "TURMA" Then
        C_TURMA = ACHA_COLUNA
    End If
Next
Sheets(c).Range(Cells(2, C_TOTALBD), Cells(65000, C_TOTALBD)).ClearContents
For L = 1 To Sheets(d).Range("A65000").End(xlUp).Row
    For L2 = Sheets(c).Cells(65000, C_TURMA).End(xlUp).Row To 2 Step -1 ' VARRE AS TURMAS
        If Sheets(c).Cells(L2, C_TURMA) = Sheets(d).Cells(L, 3) Then ' SE A TURMA FOR IGUAL
            Sheets(c).Cells(L2, C_TURMA + 1) = Sheets(c).Cells(L2, C_TURMA + 1) + 1
            'Exit For
        End If
    Next
Next
End Sub
Sub QTD_ALUNOS_P_SALA()
For ACHA_COLUNA = 1 To Range("IV1").End(xlToLeft).Column
    If Cells(1, ACHA_COLUNA) = "TURMA" Then
        C_TURMA = ACHA_COLUNA
        Exit For
    End If
Next
C_QTD_ALUNOS_POR_SALA = CTURMA + 4

End Sub
Sub QTD_TURMAS()
' CONFIG SALAS
' PREENCHE A QTD DE TURMAS CONFORME A TABELA PREENCHIDA DO COORDENADOR (DADOS NA COLUNA A)
For ACHA_COLUNA = 1 To Range("IV1").End(xlToLeft).Column
    If Cells(2, ACHA_COLUNA) = "TURMAS" Then
        Exit For
    End If
Next
Range(Cells(3, ACHA_COLUNA + 1), Cells(65000, ACHA_COLUNA + 1)).ClearContents
For L = 3 To Cells(65000, ACHA_COLUNA).End(xlUp).Row
    MSPLIT = Split(Cells(L, 1), ";")
    For x = 0 To UBound(MSPLIT)
        Cells(L, 2) = Cells(L, 2) + 1
    Next
Next
End Sub
Sub ENTURMA_ALUNOS_POR_SALA()
' VARRE A COLUNA "A"
' CRIAR UM SPLIT USANDO FOR PARA VARRER TODOS OS VALORES
' COPIA O VALOR SALA
' VERIFICA A QTD QUE VAI SER ANEXADA NA SALA NA TABELA AO LADO
' PROCURA A QTD E COLOCA DO LADO DO NOME DO ALUNO NO BD
Sheets("CONFIG").Activate
Sheets("BD").Range("E:E").ClearContents
For ACHA_COLUNA = 1 To Range("IV1").End(xlToLeft).Column ' ACHA AS COLUNAS
    If Cells(1, ACHA_COLUNA) = "TURMA" Then
        C_TURMA = ACHA_COLUNA
    End If
    If Cells(2, ACHA_COLUNA) = "TURMAS" Then
        TODAS_TURMA = ACHA_COLUNA
    End If
Next
For C1 = 2 To Cells(65000, C_TURMA).End(xlUp).Row ' VARRE A COLUNA I (TURMA)
    For C2 = 3 To Cells(65000, TODAS_TURMA).End(xlUp).Row ' VARRE A COLUNA A (TURMAS - CONFIG-SALAS)
        MY_TURMA = Split(Cells(C2, TODAS_TURMA), ";")
        For x = 0 To UBound(MY_TURMA)
            If MY_TURMA(x) Like "*" & Cells(C1, C_TURMA) & "*" Then
                sala = Cells(C2, TODAS_TURMA + 2)
                qtd = Cells(C1, C_TURMA + 4)
                CONTAR = 0
                Call PREENCHE_TURMA(sala, qtd, Cells(C1, C_TURMA))
            End If
        Next
    Next
Next
End Sub
Sub PREENCHE_TURMA(sala, qtd, TURMA)
For x = 1 To Sheets("BD").Range("B65000").End(xlUp).Row
    If Sheets("BD").Cells(x, 3) = TURMA And Sheets("BD").Cells(x, 5) = "" Then
        Sheets("BD").Cells(x, 5) = sala
        CONTAR = CONTAR + 1
        If CONTAR = qtd Then Exit Sub
    End If
Next
End Sub

Sub zOld______ENTURMA_ALUNOS_POR_SALA()
Sheets("CONFIG").Activate
Sheets("BD").Range("E:E").ClearContents
For ACHA_COLUNA = 1 To Range("IV1").End(xlToLeft).Column ' ACHA AS COLUNAS
    If Cells(1, ACHA_COLUNA) = "TURMA" Then
        C_TURMA = ACHA_COLUNA
    End If
    If Cells(2, ACHA_COLUNA) = "TURMAS" Then
        TODAS_TURMA = ACHA_COLUNA
    End If
Next
'================================== VARRE A TABELA DE TURMA
For L1 = 2 To Cells(65000, C_TURMA).End(xlUp).Row
QTD_ALUNOS_POR_SALA = Cells(L1, C_TURMA + 4) ' PEGA A QTD DE ALUNOS POR SALA
'================================== VARRE A CONFIG-SALAS
    For L2 = 3 To Cells(65000, TODAS_TURMA).End(xlUp).Row
        If Cells(L2, TODAS_TURMA) Like "*" & Cells(L1, C_TURMA) & "*" Then
        If Cells(L1, C_TURMA) = "1C" Then
            parar = 1
        End If
        sala = Cells(L2, TODAS_TURMA + 2)
        ' encontra o primeiro aluno da turma
        LIN_FINAL = Sheets("BD").Range("E65000").End(xlUp).Row
        If Cells(L1, C_TURMA) <> Sheets("BD").Range("C" & LIN_FINAL) Then
            MsgBox "ERRO!! VERIFICAR A DISTRIBUIÇÃO DAS SALAS.", vbInformation ': Exit Sub
        End If
        
        If LIN_FINAL <> 1 Then LIN_FINAL = LIN_FINAL + 1
        FIM = LIN_FINAL + QTD_ALUNOS_POR_SALA - 1
            For PSALA = LIN_FINAL To FIM ' PREENCHE SE A SALA SE A TURMA FOR IGUAL
                If Sheets("BD").Cells(PSALA, 3) = Cells(L1, C_TURMA) Then
                    Sheets("BD").Cells(PSALA, 5) = sala
                End If
            Next
        End If
    Next
Next
End Sub

Sub zOLD____ALUNOS_POR_SALA()
On Error GoTo F2
Dim ALUNOS_POR_SALA
Sheets("CONFIG").Activate
Sheets("BD").Range("E:E").ClearContents
For ACHA_COLUNA = 1 To Range("IV1").End(xlToLeft).Column ' ACHA AS COLUNAS
    If Cells(1, ACHA_COLUNA) = "TURMA" Then
        C_TURMA = ACHA_COLUNA
    End If
    If Cells(2, ACHA_COLUNA) = "TURMAS" Then
        TODAS_TURMA = ACHA_COLUNA
    End If
Next
' VARRE A TABELA DE TURMA
For L1 = 2 To Cells(65000, C_TURMA).End(xlUp).Row
ALUNOS_POR_SALA = Cells(L1, C_TURMA + 4) ' PEGA A QTD DE ALUNOS POR SALA
    ' VARRE A CONFIG-SALAS
    For L2 = 3 To Cells(65000, TODAS_TURMA).End(xlUp).Row
        If Cells(L2, TODAS_TURMA) Like "*" & Cells(L1, C_TURMA) & "*" Then
        
            'MYSPLIT = Split(Cells(L2, TODAS_TURMA), ";")
            'For X = 0 To UBound(MYSPLIT) ' VERIFICA SE EXISTE A TURMA
            '    If MYSPLIT(X) = Cells(L1, C_TURMA) Then
                    sala = Cells(L2, TODAS_TURMA + 2)
                    LIN_FINAL = Sheets("BD").Range("E65000").End(xlUp).Row
                    'If LIN_FINAL = 1 Then
                    If LIN_FINAL <> 1 Then LIN_FINAL = LIN_FINAL + 1
                    FIM = LIN_FINAL + ALUNOS_POR_SALA - 1
                        For PSALA = LIN_FINAL To FIM ' PREENCHE SE A SALA SE A TURMA FOR IGUAL
                            If Sheets("BD").Cells(PSALA, 3) = Cells(L1, C_TURMA) Then
                                Sheets("BD").Cells(PSALA, 5) = sala
                            End If
                        Next
                'End If
            'Next
        End If
    Next
Next
Exit Sub
F2:
MsgBox "Erro na quantidade de alunos, verificar a base de dados com a qtd estabelecida(Turmas)", vbInformation
End Sub
Sub QTD_SALAS()
For c = 1 To Range("IV1").End(xlToLeft).Column
            If Cells(1, c) = "TURMA" Then COLUNA = c: Exit For
Next
Range(Cells(2, COLUNA + 3), Cells(65000, COLUNA + 3)).ClearContents
For L = 3 To Range("A65000").End(xlUp).Row
    CONTA_TURMAS = Split(Range("A" & L), ";")
    For x = 0 To UBound(CONTA_TURMAS)
        ' ACHA COLUNA
        
        
        For X2 = 2 To Cells(65000, COLUNA).End(xlUp).Row
            If Cells(X2, COLUNA) = CONTA_TURMAS(x) Then
                Cells(X2, COLUNA + 3) = Cells(X2, COLUNA + 3) + 1: Exit For
            End If
        Next
    Next
Next
End Sub
Sub CALL_RELATORIO()
    FRM_RELATORIO.Show 0
End Sub
