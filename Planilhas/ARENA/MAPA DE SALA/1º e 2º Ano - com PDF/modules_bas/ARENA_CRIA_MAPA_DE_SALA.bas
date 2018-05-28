Attribute VB_Name = "ARENA_CRIA_MAPA_DE_SALA"
Public CONTA As Integer, M As Integer, VALOR_MAX_MATRIX As Integer, FIM_DO_PROCESSO As Boolean, OPCAO_DE_MAPA, QTD_ATUAL, DESC_QTD_ATUAL, SALASUPERLOTADA As Boolean

Sub CORRECAO_POR_SALA()
If Sheets("CONFIG").Range("E1") = True Then
ALUNOS_POR_SALA = InputBox("QUAL A QUANTIDADE QUE DESEJA RETIRAR DAS OUTRAS TURMAS? (PADRÃO = 7)")
If ALUNOS_POR_SALA = "" Then ALUNOS_POR_SALA = 7
Sheets("BD").Activate
ORGANIZA_POR_SALA
MATRIX = "1A;1B;1C;1D;1E;2A;2B;2C;2D"
MYSPLIT = Split(MATRIX, ";")
For k = 0 To UBound(MYSPLIT)
    For CONTA = 1 To ALUNOS_POR_SALA
        For x = 1 To Sheets("BD").Range("B65000").End(xlUp).Row
            If MYSPLIT(k) = Sheets("BD").Range("C" & x) And Sheets("BD").Range("E" & x) <> "Auditorio" Then
                Sheets("BD").Range("E" & x) = "Auditorio"
                Exit For
            End If
        Next
    Next
Next
Sheets("CONFIG").Activate
End If
End Sub
Sub VERIFICA_QTD_MAX_POR_SALA()
SALASUPERLOTADA = False
For s = 3 To Sheets("CONFIG").Range("C65000").End(xlUp).Row
sala = Sheets("CONFIG").Cells(s, 3)
    For CONTA = 1 To Sheets("BD").Range("A65000").End(xlUp).Row
        If Sheets("BD").Cells(CONTA, 5) = sala Then
            SOMA = SOMA + 1
        End If
    Next
    If SOMA > 40 Then MsgBox "A SALA ESTÁ SUPERLOTADA!" & Chr(13) & "SALA: " & sala, vbInformation: SALASUPERLOTADA = True: 'Exit Sub
    SOMA = 0
Next
End Sub
Sub CORRECAO_LEVANTAMENTO_DE_DADOS()
Sheets("CONFIG").Activate
' ACHA A COLUNA (INDICE)
For ACHA_COLUNA = 1 To Range("IV1").End(xlToLeft).Column
    If Cells(1, ACHA_COLUNA) = "TURMA" Then
        c = ACHA_COLUNA: Exit For
    End If
Next
' COLOCA O ALUNO NA SALA CERTA
INI = 0
For IND = 2 To Range("H65000").End(xlUp).Row
    'If Cells(IND, C) > 0 Then ' QTD DE ALUNOS SEM SALA
        TURMA = Cells(IND, c)
        'QTD = Cells(IND, C)
        'ACHAR O ALUNO PENDENTE E COLOCÁ-LO NA SALA
        For x = 1 To Sheets("BD").Range("A65000").End(xlUp).Row
            If Sheets("BD").Cells(x, 5) = "" Then
                If Sheets("BD").Cells(x, 3) = TURMA Then
                Call ACHA_SALA_MAIS_VAZIA(IND, c, TURMA) ' VERIFICA A SALA MAIS VAZIA
                    Sheets("BD").Cells(x, 5) = DESC_QTD_ATUAL
                    'MCONT = MCONT + 1
                    'If MCONT = QTD Then Exit For
                    'Exit For
                End If
            End If
        Next
    'End If
Next
End Sub
Sub ACHA_SALA_MAIS_VAZIA(IND, c, TURMA)
' ACHA SALA MAIS VAZIA
    For k = 3 To Sheets("CONFIG").Range("C65000").End(xlUp).Row
        If Cells(k, 1) Like "*" & TURMA & "*" Then
            sala = Cells(k, 3)
            For KK = 1 To Sheets("BD").Range("A65000").End(xlUp).Row
                If Sheets("BD").Cells(KK, 5) = sala And Sheets("BD").Cells(KK, 3) = TURMA Then
                    QTD_SOMA = QTD_SOMA + 1
                End If
            Next
            If IsEmpty(QTD_MENOR) Then QTD_MENOR = QTD_SOMA: DESC_QTD_ATUAL = sala
            If QTD_SOMA >= QTD_MENOR Then
            Else
                QTD_MENOR = QTD_SOMA: DESC_QTD_ATUAL = sala
            End If
            QTD_SOMA = 0
        End If
    Next
    
    QTD_ATUAL = QTD_MENOR

End Sub
Sub CORRECAO_ALUNOS_POR_SALA()
Sheets("CONFIG").Activate
For x = 2 To Range("N65000").End(xlUp).Row
    V1 = Left(Cells(x, 8), 1)
    If Cells(x, 14) < 0 Then
        
    End If
Next
End Sub
Sub CRIA_PLANILHAS()
    Dim Q As New Worksheet
    Set Q = Sheets("CONFIG-SALAS")
    
    'VARRE A PLANILHA DE CONFIGURAÇÃO DE SALA
    For L = 2 To UltimaLinha(Q, 2)
        ' PEGA O MODELO DE SALA
        valor = Q.Cells(L, 2)
        ' SE NÃO FOR NULO MOSTRA A SALA
        If Not IsNull(valor) Then Sheets(valor).Visible = True
    Next
    
    ' VARRE A PLANILHA DE CONFIGURAÇÃO
    Sheets("CONFIG").Activate
    For L = 3 To Range("C65000").End(xlUp).Row
        
        'PEGA O NOME DO MODELO
        valor = Sheets("CONFIG").Cells(L, 5)
            
        'FAZ UMA COPIA DA PLANILHA MODELO E RENOMEIA
        Sheets(valor).Copy AFTER:=Worksheets("Rel-Sala")
        Sheets(valor & " (2)").Select
        Sheets(valor & " (2)").name = Sheets("CONFIG").Cells(L, 3)
        ActiveSheet.Shapes("WordArt 1").Select
        Selection.ShapeRange.TextEffect.Text = "Mapa - " & Sheets("CONFIG").Cells(L, 3) & " - " & Sheets("CONFIG").Cells(4, 6)
    
    Next
    '-=-=-=-=-=-=-=-=-= ((finalizando o processo)) -=-=-=-=-=-=-=-=-=
    MsgBox "Mapas Criados!!!", vbInformation

Call SHOWALL
Set Q = Sheets("CONFIG-SALAS")
For L = 2 To UltimaLinha(Q, 2)
NOME = Q.Cells(L, 2)
    Sheets(NOME).Visible = False
Next

End Sub
Sub DELETA_PLANILHAS()
inicio:
For WS = 1 To Sheets.count
    If Left(UCase(Sheets(WS).name), 4) = "SALA" Or Left(UCase(Sheets(WS).name), 4) = "AUDI" Then
    Application.DisplayAlerts = False
        Sheets(WS).Delete
    Application.DisplayAlerts = True
    GoTo inicio
    End If
Next
Sheets(1).Activate
End Sub
Sub CRIA_ESPACOS()
Sheets("CONFIG").Activate
For LK = 3 To Range("C65000").End(xlUp).Row
    If Sheets("CONFIG").Cells(LK, 1) Like "*" & Left(OPCAO_DE_MAPA, 1) & "*" Then
        Sheets("CONFIG").Activate
        MATRIX_TURMAS = Sheets("CONFIG").Range("A" & LK)
        If Range("E" & LK) = "N" Then LIN = 15: COL = 5: LIN_MAX = 50: COL_MAX = Sheets("MAPA - N").Range("IV13").End(xlToLeft).Column - 3: M = 0
        If Range("E" & LK) = "S27" Then LIN = 15: COL = 5: LIN_MAX = 39: COL_MAX = 26: M = 0
        If Range("E" & LK) = "Auditorio" Then LIN = 15: COL = 5: LIN_MAX = 39: COL_MAX = 26: M = 0
        sala = Range("C" & LK)
        Sheets(sala).Activate
        MATRIX = Split(MATRIX_TURMAS, ";")
        
        VALOR_MAX_MATRIX = UBound(MATRIX) + 1
        t = 0
        CONTA = 1
        FIM_DO_PROCESSO = False
        For c = COL To COL_MAX Step 3
            For L = LIN To LIN_MAX Step 4
            'If CONTA = 9 Then CONTA = 1: Exit For
            If M = UBound(MATRIX) + 1 Then M = 0
            If FIM_DO_PROCESSO = True Then GoTo KNEXT
                Call REL_MAPA_DE_SALA(Left(MATRIX_TURMAS, 1), MATRIX(M), L, c, Sheets("CONFIG").Range("E" & LK))
            Next
        Next
    End If
KNEXT:
Next

End Sub
Sub CRIA_ESPACOS_old()
'PREENCHE MATRIX
Sheets("CONFIG").Activate
For LK = 3 To Range("C65000").End(xlUp).Row
    If Sheets("CONFIG").Cells(LK, 1) Like "*" & Left(OPCAO_DE_MAPA, 1) & "*" Then
        Sheets("CONFIG").Activate
        MATRIX_TURMAS = Sheets("CONFIG").Range("A" & LK)
        If Range("E" & LK) = "N" Then LIN = 15: COL = 5: LIN_MAX = 43: COL_MAX = Sheets("MAPA - N").Range("IV13").End(xlToLeft).Column - 3: M = 0
        If Range("E" & LK) = "S27" Then LIN = 16: COL = 5: LIN_MAX = 32: COL_MAX = 26: M = 0
        sala = Range("C" & LK)
        Sheets(sala).Activate
        MATRIX = Split(MATRIX_TURMAS, ";")
        
        VALOR_MAX_MATRIX = UBound(MATRIX)
        t = 0
        CONTA = 1
        FIM_DO_PROCESSO = False
        For c = COL To COL_MAX Step 3
            For L = LIN To LIN_MAX Step 4
            'If CONTA = 9 Then CONTA = 1: Exit For
            If FIM_DO_PROCESSO = True Then GoTo KNEXT
                Call REL_MAPA_DE_SALA(Left(MATRIX_TURMAS, 1), MATRIX(M), L, c, Sheets("CONFIG").Range("E" & LK))
            Next
        Next
    End If
KNEXT:
Next
End Sub
Sub REL_MAPA_DE_SALA(ANOS, valor, linha, COLUNA, sala)
If ANOS = 2 Or ANOS = 1 Then ANOS = "ESPECIAL"
Select Case ANOS
    Case 3
    
        Select Case sala
            Case "N"
            ' loop até acabar
                Cells(linha + 2, COLUNA) = valor
                M = M + 1: CONTA = CONTA + 1
                If VALOR_MAX_MATRIX + 1 = M Then M = 0
            Case "S27"
            ' preenche diferente em loop
            
            x = AQUI
        End Select
    Case "ESPECIAL"
        Select Case sala
            Case "N"
            ' preenche a 1st fila e depois copia
                Cells(linha + 2, COLUNA) = valor
                M = M + 1: CONTA = CONTA + 1
                If VALOR_MAX_MATRIX = M Then M = 0
                ' COPIA VALORES
                If CONTA = 99999 Then
                    Range("E35:F45").Copy
                    Range("H15").PasteSpecial xlPasteAll
                    Range("E15:F33").Copy
                    Range("H27").PasteSpecial xlPasteAll
                    '==================
                    Range("H35:I45").Copy
                    Range("K15").PasteSpecial xlPasteAll
                    Range("H15:I33").Copy
                    Range("K27").PasteSpecial xlPasteAll
                    '==================
                    Range("K35:L45").Copy
                    Range("N15").PasteSpecial xlPasteAll
                    Range("K15:L33").Copy
                    Range("N27").PasteSpecial xlPasteAll
                    '==================
                    Range("N35:O45").Copy
                    Range("Q15").PasteSpecial xlPasteAll
                    Range("N15:O33").Copy
                    Range("Q27").PasteSpecial xlPasteAll
                    '==================
                    Range("Q35:R45").Copy
                    Range("T15").PasteSpecial xlPasteAll
                    Range("Q15:R33").Copy
                    Range("T27").PasteSpecial xlPasteAll
                    '==================
                    Range("T35:U45").Copy
                    Range("W15").PasteSpecial xlPasteAll
                    Range("T15:U33").Copy
                    Range("W27").PasteSpecial xlPasteAll
                    '==================
                    Range("W35:X45").Copy
                    Range("Z15").PasteSpecial xlPasteAll
                    Range("W15:X33").Copy
                    Range("Z27").PasteSpecial xlPasteAll
                    ' FINALIZA O PROCESSO
                    FIM_DO_PROCESSO = True
                End If
            Case Else '"S27"
            ' preenche a 1st fila na horizontal e depois copia
                Cells(linha + 2, COLUNA) = valor
                M = M + 1: CONTA = CONTA + 1
                If VALOR_MAX_MATRIX = M Then M = 0
        End Select
End Select
End Sub
Sub SHOWALL()
    For x = 1 To Sheets.count
        Sheets(x).Visible = True
    Next
End Sub
Sub HideIt()
    For L = 2 To UltimaLinha(ActiveSheet, 1)
        PLAN = Range("B" & L)
        If Not IsEmpty(PLAN) Then
            Sheets(PLAN).Visible = False
        End If
    Next
    
    Sheets("backup").Visible = False
    Sheets("ANOTAÇÕES").Visible = False
    Sheets("arena-3").Visible = False
    Sheets("arena-4").Visible = False
    
End Sub

Sub getMAX()
Dim sala As Worksheet
Dim qtd As Worksheet
Set sala = Sheets("CONFIG-SALAS")
Set qtd = Sheets("CONFIG-QTD")

qtd.Range("A:D").ClearContents
Sheets("CONFIG").Range("C3:C" & UltimaLinha(Sheets("CONFIG"), 3)).Copy

qtd.Range("a1").PasteSpecial xlPasteValues

For L = 1 To UltimaLinha(qtd, 1)
    s = qtd.Cells(L, 1)
    For ll = 2 To UltimaLinha(sala, 1)
        If s = sala.Cells(ll, 1) Then
            qtd.Range("b" & L).Value = sala.Range("C" & ll): Exit For
        End If
    Next
Next

LIN = UltimaLinha(Sheets("BD"), 1)
qtd.Range("e" & UltimaLinha(qtd, 2) - 1) = "Total - BD: "
qtd.Range("f" & UltimaLinha(qtd, 2) - 1).FormulaR1C1 = "=COUNTA(BD!R1C[-2]:R[" & LIN & "]C[-2])"

qtd.Range("e" & UltimaLinha(qtd, 2) + 1) = "Total: "
qtd.Range("f" & UltimaLinha(qtd, 2) + 1).FormulaR1C1 = "=SUM(R1C[-4]:R[-1]C[-4])"
End Sub
