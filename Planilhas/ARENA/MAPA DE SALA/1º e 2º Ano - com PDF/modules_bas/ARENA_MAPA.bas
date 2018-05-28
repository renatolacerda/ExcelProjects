Attribute VB_Name = "ARENA_MAPA"
Dim p As Worksheet
Dim Q As Worksheet
Dim H As Worksheet
Sub MAPA()
'criar histórico para não repetir alunos na mesma sala
'HISTORICO_BD
AJUSTE_CARTEIRAS
RANDOM_ALUNOS
QTDS_SALAS
End Sub
Sub old__QTDS_SALAS()
Dim NOME As String, H_TURMAS As String, ARR_SALAS As String
Set p = Sheets("BD")
Set Q = Sheets("CONFIG-QTD")
Dim TODASSALAS()
'QTD =
ReDim TODASSALAS(0 To 10)

Q.Range("C" & UltimaLinha(Q, 3)).ClearContents

For L = 1 To UltimaLinha(p, 2) ' EM BD
    NOME = p.Cells(L, 2)
    H_TURMAS = GET_HISTORICO(NOME) ' acha as turmas que o aluno já esteve
    ARR_HISTORICO = Split(H_TURMAS, ";")
    For s = 1 To UltimaLinha(Q, 1) 'EM CONFIG-QTD
        For t = 0 To UBound(ARR_HISTORICO)
            If ARR_HISTORICO(t) = Q.Cells(s, 1) Then
            Else
                If ARR_SALAS = "" Then
                    ARR_SALAS = Q.Cells(s, 1): Exit For
                Else
                    ARR_SALAS = ARR_SALAS & ";" & Q.Cells(s, 1): Exit For
                End If
            End If
        Next
    
    
        'mySALAS = Q.Range("A1:A" & UltimaLinha(Q, 2))
        'For s = 1 To UltimaLinha(Q, 1) 'EM CONFIG-QTD
        '    If Q.Cells(s, 1) <> MYSPLIT(T) Then
        '        If mySALAS = "" Then
        '            mySALAS = MYSPLIT(T)
        '        Else
        '            mySALAS = mySALAS & ";" & MYSPLIT(T)
        '        End If
                'P.Cells(L, 5) = Q.Cells(s, 1): Exit For
        '    End If
        'Next
    Next
Next
End Sub
Function GET_HISTORICO(ALUNO As String) As String
Dim H_TURMAS As String
Set H = Sheets("BD-HISTORICO")

For L = 1 To UltimaLinha(H, 2)
    If H.Cells(L, 2) = ALUNO Then
        If H_TURMAS = "" Then
            H_TURMAS = H.Cells(L, 5)
        Else
            H_TURMAS = H_TURMAS & ";" & H.Cells(L, 5)
        End If
    End If
Next
    GET_HISTORICO = H_TURMAS
    
End Function

Sub QTDS_SALAS()
Dim p As Worksheet

Call getMAX

Sheets("BD").Activate
Set p = Sheets("CONFIG-QTD")

LIN_INI = 0

For x = 1 To UltimaLinha(p, 1)
    qtd = p.Range("B" & x)
    sala = p.Range("A" & x)
    If x = 1 Then
        Sheets("BD").Range(Cells(1, 5), Cells(LIN_INI + qtd, 5)) = sala
    Else
        Sheets("BD").Range(Cells(LIN_INI, 5), Cells(LIN_INI + qtd, 5)) = sala
    End If
    LIN_INI = LIN_INI + qtd
Next
Sheets("CONFIG").Activate
End Sub
Sub config_Turmas()
Dim c As Worksheet
Set c = Sheets("CONFIG")
Dim L As Integer

For L = 2 To UltimaLinha(c, 8)
    If L = 2 Then
        MATRIX = c.Range("I" & L)
    Else
        MATRIX = MATRIX & ";" & c.Range("I" & L)
    End If
Next

MATRIX = MATRIX & ";"

For L = 3 To UltimaLinha(c, 3)
    c.Cells(L, 1) = MATRIX
    c.Cells(L, 2) = 6
Next

End Sub
