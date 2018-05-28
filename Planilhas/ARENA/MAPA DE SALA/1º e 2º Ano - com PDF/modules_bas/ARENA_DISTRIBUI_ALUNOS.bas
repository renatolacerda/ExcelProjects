Attribute VB_Name = "ARENA_DISTRIBUI_ALUNOS"
Option Explicit
Sub DISTRIBUIR_ALUNOS()
    Dim d As Worksheet
    Dim s As Worksheet
    Set d = Sheets("BD")
    Dim NOME As String, TURMA As String, sala As String
    Dim colNOME As Integer, colTURMA As Integer, colSALA As Integer, colNomeSala As Integer, colNomeTurma As Integer, LinhaBD As Integer, LinhaSala As Integer, ultlinha As Integer
    Dim ss
    
    colNOME = 2: colTURMA = 3: colSALA = 5
    colNomeSala = 37
    colNomeTurma = 38
    
    'VARRE TODAS AS LINHAS DO DB
    ultlinha = UltimaLinha(d, 2)
    For LinhaBD = 1 To ultlinha
    'If LinhaBD = 29 Then
    '    ss = 1
    'End If
        sala = d.Cells(LinhaBD, colSALA)
        Set s = Sheets(sala)
        's.Select
        s.Activate
        LinhaSala = UltimaLinha(s, colNomeSala) + 1
            'NOME
                NOME = d.Cells(LinhaBD, colNOME)
            'TURMA
                TURMA = d.Cells(LinhaBD, colTURMA)
            'SALA
                sala = d.Cells(LinhaBD, colSALA)
            
            'MsgBox NOME & " / " & TURMA & " / " & SALA
            
            's.Cells(LinhaSala, colNomeSala).Select
            s.Cells(LinhaSala, colNomeSala) = NOME
            s.Cells(LinhaSala, colNomeTurma) = TURMA
        
    'Next
    
    ' salas
    
'    Dim linha
    
'    For linha = 3 To UltimaLinha(Sheets("CONFIG"), 3)
'        sala = Sheets("CONFIG").Cells(linha, 3)
'        Set s = Sheets(sala)
        
'        For L = 1 To UltimaLinha(d, 2)
'            linha_ini
'            linha_fim
'        Next
    Next
    
    
End Sub
Sub AtribuiAlunoCadeira()
Dim s As Worksheet
Dim LIN, LIN_MAX, COL, COL_MAX, colNOME As Integer, colTURMA As Integer, LISTA, L As Integer, c As Integer
Dim ALUNO, TURMA

Set s = ActiveSheet

LIN = s.Range("AL6") ' linha inicial
COL = s.Range("AL7") 'coluna inicial
LIN_MAX = s.Range("AL8") 'linha maxima
COL_MAX = s.Range("AL9") 'coluna maxima

colNOME = 37
colTURMA = colNOME + 1

For LISTA = 14 To UltimaLinha(s, colNOME)
    
    ALUNO = s.Cells(LISTA, colNOME)
    TURMA = s.Cells(LISTA, colTURMA)
    
    For L = LIN To LIN_MAX Step 4
        For c = COL To COL_MAX Step 3
            If s.Cells(L, c) = "" And s.Cells(L + 2, c) = TURMA Then
                s.Cells(L, c) = ALUNO
                s.Cells(LISTA, colNOME) = ""
                s.Cells(LISTA, colTURMA) = ""
                GoTo OUTROALUNO
            End If
        Next
    Next
OUTROALUNO:
Next
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
