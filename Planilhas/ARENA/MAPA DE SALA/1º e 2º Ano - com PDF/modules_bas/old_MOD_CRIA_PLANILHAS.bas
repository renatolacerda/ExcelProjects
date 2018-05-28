Attribute VB_Name = "old_MOD_CRIA_PLANILHAS"
Sub Cria_ANOS()
'==================================================
For linha = 1 To Range("CONFIGURA플O!AB60000").End(xlUp).Row
    If Left(Range("CONFIGURA플O!AB" & linha), 1) = Left(Range("$K$6"), 1) Then
        Range("CONFIGURA플O!AD" & Range("AD60000").End(xlUp).Row + 1) = Range("CONFIGURA플O!AB" & linha)
    End If
Next
'==================================================
ReDim TURMA(1 To Range("CONFIGURA플O!Z60000").End(xlUp).Row)
CONTA = 1
For linha = 1 To Range("CONFIGURA플O!Z60000").End(xlUp).Row
    If linha = 1 Then
    TURMA(linha) = Range("CONFIGURA플O!Z" & linha)
    'CONTA = CONTA + 1
    Else
        If TURMA(linha) <> Range("CONFIGURA플O!Z" & linha) Then
            TURMA(linha) = Range("CONFIGURA플O!Z" & linha)
            'CONTA = CONTA + 1
        End If
    End If
Next

Worksheets("MODELO-ANO").Visible = True

For p = 1 To (linha - 1)
    Worksheets("MODELO-ANO").Select
    Application.ScreenUpdating = False
    For WS = 1 To Worksheets.count
        If TURMA(p) = Worksheets(WS).name Then
            GoTo NPROXIMO
        End If
    Next
    Plan3.Copy AFTER:=Worksheets("MODELO-ANO")
    Worksheets("MODELO-ANO (3)").Select
    Worksheets("MODELO-ANO (3)").name = TURMA(p)
    Application.ScreenUpdating = True
    Worksheets(TURMA(p)).Activate
NPROXIMO:
Next

Worksheets("MODELO-ANO").Visible = False
End Sub
Sub Cria_SALAS()
ReDim MSALA(1 To Range("CONFIGURA플O!AB60000").End(xlUp).Row)
CONTA = 1
For linha = 1 To Range("CONFIGURA플O!AB60000").End(xlUp).Row
    If linha = 1 Then
    MSALA(linha) = Range("CONFIGURA플O!AB" & linha)
    'CONTA = CONTA + 1
    Else
        If MSALA(linha) <> Range("CONFIGURA플O!AB" & linha) Then
            MSALA(linha) = Range("CONFIGURA플O!AB" & linha)
            'CONTA = CONTA + 1
        End If
    End If
Next

Worksheets("MODELO-SALA").Visible = True

For p = 1 To (linha - 1)
    Worksheets("MODELO-SALA").Select
    Application.ScreenUpdating = False
    For WS = 1 To Worksheets.count
        If MSALA(p) = Worksheets(WS).name Then
            GoTo NPROXIMO
        End If
    Next
    Plan8.Copy AFTER:=Plan8
    Worksheets("MODELO-SALA (2)").Select
    Worksheets("MODELO-SALA (2)").name = MSALA(p)
    Application.ScreenUpdating = True
    Worksheets(MSALA(p)).Activate
NPROXIMO:
Next

Worksheets("MODELO-SALA").Visible = False
End Sub
