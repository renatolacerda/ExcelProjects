Attribute VB_Name = "old_DISTRIBUIR_ALUNOS_1"
'Sub quebra_galho()
'DELETA_PLANILHAS ' DELETA AS PLANILHAS SE EXISTIREM
'FRM_MAPADESALA.Show 1
'CRIA_PLANILHAS
'CRIA_ESPACOS
'DISTRIBUIR_ALUNOS
'End Sub
Sub old_DISTRIBUIR_ALUNOS()

    ATUAL = "BD"
    Sheets(ATUAL).Activate
    
    ORDENA_SALA
    
    Sheets(ATUAL).Activate
    
    Range("F:F").ClearContents
    
    'c = 2
    'For LK = 3 To Sheets("CONFIG").Range("C65000").End(xlUp).Row
    '    If Sheets("CONFIG").Cells(LK, 1) Like "*" & Left(OPCAO_DE_MAPA, 1) & "*" Then
    '        For Linha = 1 To Range("A65000").End(xlUp).Row
    '        If Cells(Linha, c + 4) <> "ENTURMADO" Then
    '            MSPLIT = Split(Sheets("CONFIG").Cells(LK, 1), ";")
    '                For x = 0 To UBound(MSPLIT)
    '                    If MSPLIT(x) = Cells(Linha, 3) Then
    '                        Call DOit(Cells(Linha, c), Cells(Linha, c + 1), Cells(Linha, c + 3))
    '                        Sheets(ATUAL).Activate
    '                        Cells(Linha, c + 4).Select
    '                        Cells(Linha, c + 4) = "ENTURMADO"
    '                    End If
    '                Next
    '        End If
    '        Next
    '    End If
    'Next
End Sub
Sub DOit(ALUNO, TURMA, sala)
Application.DisplayAlerts = False
s = sala
'If S = "Sala 9" Then S = "Sala 8": SALA = "Sala 8"
Sheets(s).Activate
'If SALA = "Auditorio" Then LIN = 15: COL = 5: LIN_MAX = 71: COL_MAX = 30: M = 0: GoTo XX
'If Right(SALA, 2) <> "27" Then LIN = 15: COL = 5: LIN_MAX = 50: COL_MAX = 21: M = 0: GoTo XX
'If SALA = "Sala 27" Then LIN = 15: COL = 5: LIN_MAX = 38: COL_MAX = 26: M = 0: GoTo XX
For x = 1 To 10
    If sala = "Sala " & x Then LIN = 15: COL = 5: LIN_MAX = 31: COL_MAX = 34: M = 0: GoTo XX
Next
XX:
For L = LIN To LIN_MAX Step 4
    For c = COL To COL_MAX Step 3
        If Cells(L, c) = "" And Cells(L + 2, c) = TURMA Then
            Cells(L, c) = ALUNO
            Exit Sub
        End If
    Next
Next

'NAOACHOU = 1
'For L = LIN To LIN_MAX Step 4
    'For C = COL To COL_MAX Step 3
        ''If Right(SALA, 2) <> 27 Then
            ''If C = 8 Or C = 14 Or C = 20 Then
            ''    GoTo X2
            ''End If

            
            
            'If Cells(L, C) = "" And Cells(L + 2, C) = TURMA Then
                'Cells(L, C) = ALUNO
                'NAOACHOU = 0
                'Exit Sub
            'End If
        ''End If
''X2:
    'Next
'Next
Application.DisplayAlerts = True
End Sub
