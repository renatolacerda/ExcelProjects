Attribute VB_Name = "MOD_DISTRIBUIR"
Sub distribui_sala9()
sala = ActiveSheet.name
Worksheets(sala).Activate
r = Range("D14:AJ34")
For L = 14 To Range("BL65000").End(xlUp).Row
TURMA = Range("BM" & L)
ALUNO = Range("BL" & L)
    For ll = 15 To 42
        For c = 6 To 34
            If (TURMA = Cells(ll, c)) And Cells(ll - 2, c) = "" Then
                Cells(ll - 2, c) = ALUNO
                Cells(L, 56) = ""
                Cells(L, 57) = ""
                GoTo FIM
            End If
        Next
    Next
FIM:
Next
ACHA_FALTANTES_SL89
End Sub
Sub distribui_sala8()
sala = ActiveSheet.name
Worksheets(sala).Activate
For L = 14 To Range("BD65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("BE" & L)
ALUNO = Range("BD" & L)
    For ll = 15 To 34 'MAPA DE SALA
        For c = 33 To 51
            If (TURMA = Cells(ll, c)) And Cells(ll - 2, c) = "" Then
                Cells(ll - 2, c) = ALUNO
                Cells(L, 64) = ""
                Cells(L, 65) = ""
                GoTo FIM
            End If
        Next
    Next
FIM:
Next
End Sub
Sub distribui_sala7()
Dim BD As Worksheet
Set BD = Sheets("BD")
sala = ActiveSheet.name

' copia os dados para a planilha atual
If MsgBox("Deseja importar os nomes da base de dados?", vbYesNo) = vbYes Then
'LIMPA INTERVALO DAS CADEIRAS
    Worksheets(sala).Range("W14:AP65000").Clear

    For L = 1 To BD.Range("B65000").End(xlUp).Row
        If BD.Range("E" & L) = sala Then
        LIN = Worksheets(sala).Range("W65000").End(xlUp).Row + 1
            Worksheets(sala).Range("W" & LIN) = BD.Range("B" & L) 'NOME
            Worksheets(sala).Range("X" & LIN) = BD.Range("C" & L) 'TURMA
        End If
    Next
End If
'enturma
'SALA = ActiveSheet.Name
Worksheets(sala).Activate
For L = 14 To Range("W65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("X" & L)
ALUNO = Range("W" & L)
    For ll = 15 To 37 'MAPA DE SALA
        For c = 5 To 17
            If (TURMA = Cells(ll, c)) And Cells(ll - 2, c) = "" Then
                Cells(ll - 2, c) = ALUNO
                Cells(L, 23) = ""
                Cells(L, 24) = ""
                GoTo FIM
            End If
        Next
    Next
FIM:
Next
End Sub
Sub distribui_sala6()
Dim BD As Worksheet
Set BD = Sheets("BD")
sala = ActiveSheet.name

' copia os dados para a planilha atual
If MsgBox("Deseja importar os nomes da base de dados?", vbYesNo) = vbYes Then
'LIMPA INTERVALO DAS CADEIRAS
    Worksheets(sala).Range("AO14:AP65000").Clear

    For L = 1 To BD.Range("B65000").End(xlUp).Row
        If BD.Range("E" & L) = sala Then
        LIN = Worksheets(sala).Range("AO65000").End(xlUp).Row + 1
            Worksheets(sala).Range("AO" & LIN) = BD.Range("B" & L) 'NOME
            Worksheets(sala).Range("AP" & LIN) = BD.Range("C" & L) 'TURMA
        End If
    Next
End If
'enturma
sala = ActiveSheet.name
Worksheets(sala).Activate
For L = 14 To Range("AO65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("AP" & L)
ALUNO = Range("AO" & L)
    For ll = 15 To 43 'MAPA DE SALA
        For c = 5 To 35
            If (TURMA = Cells(ll, c)) And Cells(ll - 2, c) = "" Then
                Cells(ll - 2, c) = ALUNO
                Cells(L, 42) = ""
                Cells(L, 41) = ""
                GoTo FIM
            End If
        Next
    Next
FIM:
Next

'ACHA ESPAÇOS VAZIOS E ADICIONA OS NOMES SEM ESPAÇO
sala = ActiveSheet.name
Worksheets(sala).Activate
For L = 14 To Range("AO65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("AP" & L)
ALUNO = Range("AO" & L)
    For ll = 15 To 43 'MAPA DE SALA
        For c = 5 To 35
            If TURMA <> "" And ALUNO <> "" Then
                If (Len(Cells(ll, c)) = 2 And Cells(ll - 2, c) = "") Then
                    Cells(ll - 2, c) = ALUNO
                    Cells(ll, c) = TURMA
                    Cells(L, 42) = ""
                    Cells(L, 41) = ""
                    GoTo FIM2
                End If
            End If
        Next
    Next
FIM2:
Next
End Sub
Sub distribui_sala27()
Dim BD As Worksheet
Set BD = Sheets("BD")
sala = ActiveSheet.name

If MsgBox("Deseja importar os nomes da base de dados?", vbYesNo) = vbYes Then
'LIMPA INTERVALO DAS CADEIRAS
    Worksheets(sala).Range("AK14:AL65000").Clear
    For L = 1 To BD.Range("B65000").End(xlUp).Row
        If BD.Range("E" & L) = sala Then
        LIN = Worksheets(sala).Range("Ak65000").End(xlUp).Row + 1
            Worksheets(sala).Range("Ak" & LIN) = BD.Range("B" & L) 'NOME
            Worksheets(sala).Range("Al" & LIN) = BD.Range("C" & L) 'TURMA
        End If
    Next
End If
'enturma
sala = ActiveSheet.name
Worksheets(sala).Activate
For L = 14 To Range("Ak65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("Al" & L)
ALUNO = Range("Ak" & L)
    For ll = 15 To 43 'MAPA DE SALA
        For c = 5 To 32
            If (TURMA = Cells(ll, c)) And Cells(ll - 2, c) = "" Then
                Cells(ll - 2, c) = ALUNO
                Cells(L, 37) = ""
                Cells(L, 38) = ""
                GoTo FIM
            End If
        Next
    Next
FIM:
Next
'ACHA ESPAÇOS VAZIOS E ADICIONA OS NOMES SEM ESPAÇO
sala = ActiveSheet.name
Worksheets(sala).Activate
For L = 14 To Range("Ak65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("Al" & L)
ALUNO = Range("Ak" & L)
    For ll = 15 To 43 'MAPA DE SALA
        For c = 5 To 35
            If TURMA <> "" And ALUNO <> "" Then
                If (Len(Cells(ll, c)) = 2 And Cells(ll - 2, c) = "") Then
                    Cells(ll - 2, c) = ALUNO
                    Cells(ll, c) = TURMA
                    Cells(L, 37) = ""
                    Cells(L, 38) = ""
                    GoTo FIM2
                End If
            End If
        Next
    Next
FIM2:
Next


End Sub
Sub distribui_sala5()
Dim BD As Worksheet
Set BD = Sheets("BD")
sala = ActiveSheet.name
' copia os dados para a planilha atual
If MsgBox("Deseja importar os nomes da base de dados?", vbYesNo) = vbYes Then
'LIMPA INTERVALO DAS CADEIRAS
    Worksheets(sala).Range("AR14:AS65000").Clear

    For L = 1 To BD.Range("B65000").End(xlUp).Row
        If BD.Range("E" & L) = sala Then
        LIN = Worksheets(sala).Range("AR65000").End(xlUp).Row + 1
            Worksheets(sala).Range("AR" & LIN) = BD.Range("B" & L) 'NOME
            Worksheets(sala).Range("AS" & LIN) = BD.Range("C" & L) 'TURMA
        End If
    Next
End If
'enturma
sala = ActiveSheet.name
Worksheets(sala).Activate
For L = 14 To Range("AR65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("AS" & L)
ALUNO = Range("AR" & L)
    For ll = 15 To 41 'MAPA DE SALA
        For c = 5 To 40
            If (TURMA = Cells(ll, c)) And Cells(ll - 2, c) = "" Then
                Cells(ll - 2, c) = ALUNO
                Cells(L, 45) = ""
                Cells(L, 44) = ""
                GoTo FIM
            End If
        Next
    Next
FIM:
Next

'ACHA ESPAÇOS VAZIOS E ADICIONA OS NOMES SEM ESPAÇO
sala = ActiveSheet.name
Worksheets(sala).Activate
For L = 14 To Range("AR65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("AS" & L)
ALUNO = Range("AR" & L)
    For ll = 15 To 41 'MAPA DE SALA
        For c = 5 To 40
            If TURMA <> "" And ALUNO <> "" Then
                If (Len(Cells(ll, c)) = 2 And Cells(ll - 2, c) = "") Then
                    Cells(ll - 2, c) = ALUNO
                    Cells(ll, c) = TURMA
                    Cells(L, 45) = ""
                    Cells(L, 44) = ""
                    GoTo FIM2
                End If
            End If
        Next
    Next
FIM2:
Next
End Sub
Sub distribui_sala4()

Dim BD As Worksheet
Set BD = Sheets("BD")
sala = ActiveSheet.name

' copia os dados para a planilha atual
If MsgBox("Deseja importar os nomes da base de dados?", vbYesNo) = vbYes Then
'LIMPA INTERVALO DAS CADEIRAS
    Worksheets(sala).Range("AK14:AL65000").Clear

    For L = 1 To BD.Range("B65000").End(xlUp).Row
        If BD.Range("E" & L) = sala Then
        LIN = Worksheets(sala).Range("AK65000").End(xlUp).Row + 1
            Worksheets(sala).Range("AK" & LIN) = BD.Range("B" & L) 'NOME
            Worksheets(sala).Range("AL" & LIN) = BD.Range("C" & L) 'TURMA
        End If
    Next
End If
'enturma
sala = ActiveSheet.name
Worksheets(sala).Activate
For L = 14 To Range("AK65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("AL" & L)
ALUNO = Range("AK" & L)
    For ll = 15 To 37 'MAPA DE SALA
        For c = 5 To 32
            If (TURMA = Cells(ll, c)) And Cells(ll - 2, c) = "" Then
                Cells(ll - 2, c) = ALUNO
                Cells(L, 37) = ""
                Cells(L, 38) = ""
                GoTo FIM
            End If
        Next
    Next
FIM:
Next

'ACHA ESPAÇOS VAZIOS E ADICIONA OS NOMES SEM ESPAÇO
sala = ActiveSheet.name
Worksheets(sala).Activate
For L = 14 To Range("Ak65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("AL" & L)
ALUNO = Range("AK" & L)
    For ll = 15 To 37 'MAPA DE SALA
        For c = 5 To 32
            If TURMA <> "" And ALUNO <> "" Then
                If (Len(Cells(ll, c)) = 2 And Cells(ll - 2, c) = "") Then
                    Cells(ll - 2, c) = ALUNO
                    Cells(ll, c) = TURMA
                    Cells(L, 37) = ""
                    Cells(L, 38) = ""
                    GoTo FIM2
                End If
            End If
        Next
    Next
FIM2:
Next
End Sub
Sub distribui_sala3()
Dim BD As Worksheet
Set BD = Sheets("BD")
sala = ActiveSheet.name

' copia os dados para a planilha atual
If MsgBox("Deseja importar os nomes da base de dados?", vbYesNo) = vbYes Then
'LIMPA INTERVALO DAS CADEIRAS
    Worksheets(sala).Range("Ak14:Al65000").Clear

    For L = 1 To BD.Range("B65000").End(xlUp).Row
        If BD.Range("E" & L) = sala Then
        LIN = Worksheets(sala).Range("Ak65000").End(xlUp).Row + 1
            Worksheets(sala).Range("Ak" & LIN) = BD.Range("B" & L) 'NOME
            Worksheets(sala).Range("Al" & LIN) = BD.Range("C" & L) 'TURMA
        End If
    Next
End If
'enturma

'SALA = ActiveSheet.Name
Worksheets(sala).Activate
For L = 14 To Range("AK65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("AL" & L)
ALUNO = Range("AK" & L)
    For ll = 15 To 42 'MAPA DE SALA
        For c = 5 To 32
            If (TURMA = Cells(ll, c)) And Cells(ll - 2, c) = "" Then
                Cells(ll - 2, c) = ALUNO
                Cells(L, 37) = ""
                Cells(L, 38) = ""
                GoTo FIM
            End If
        Next
    Next
FIM:
Next

'ACHA ESPAÇOS VAZIOS E ADICIONA OS NOMES SEM ESPAÇO
sala = ActiveSheet.name
Worksheets(sala).Activate
For L = 14 To Range("AK65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("AL" & L)
ALUNO = Range("AK" & L)
    For ll = 15 To 42 'MAPA DE SALA
        For c = 5 To 32
            If TURMA <> "" And ALUNO <> "" Then
                If (Len(Cells(ll, c)) = 2 And Cells(ll - 2, c) = "") Then
                    Cells(ll - 2, c) = ALUNO
                    Cells(ll, c) = TURMA
                    Cells(L, 37) = ""
                    Cells(L, 38) = ""
                    GoTo FIM2
                End If
            End If
        Next
    Next
FIM2:
Next
End Sub
Sub distribui_sala2()
Dim BD As Worksheet
Set BD = Sheets("BD")
sala = ActiveSheet.name

' copia os dados para a planilha atual
If MsgBox("Deseja importar os nomes da base de dados?", vbYesNo) = vbYes Then
'LIMPA INTERVALO DAS CADEIRAS
    Worksheets(sala).Range("AK14:AL65000").Clear

    For L = 1 To BD.Range("B65000").End(xlUp).Row
        If BD.Range("E" & L) = sala Then
        LIN = Worksheets(sala).Range("AK65000").End(xlUp).Row + 1
            Worksheets(sala).Range("AK" & LIN) = BD.Range("B" & L) 'NOME
            Worksheets(sala).Range("AL" & LIN) = BD.Range("C" & L) 'TURMA
        End If
    Next
End If
'enturma
'SALA = ActiveSheet.Name
Worksheets(sala).Activate
For L = 14 To Range("AK65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("AL" & L)
ALUNO = Range("AK" & L)
    For ll = 14 To 38 'MAPA DE SALA
        For c = 5 To 32
            If (TURMA = Cells(ll, c)) And Cells(ll - 2, c) = "" Then
                Cells(ll - 2, c) = ALUNO
                Cells(L, 37) = ""
                Cells(L, 38) = ""
                GoTo FIM
            End If
        Next
    Next
FIM:
Next
'ACHA ESPAÇOS VAZIOS E ADICIONA OS NOMES SEM ESPAÇO
sala = ActiveSheet.name
Worksheets(sala).Activate
For L = 14 To Range("Ak65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("Al" & L)
ALUNO = Range("Ak" & L)
    For ll = 15 To 43 'MAPA DE SALA
        For c = 5 To 35
            If TURMA <> "" And ALUNO <> "" Then
                If (Len(Cells(ll, c)) = 2 And Cells(ll - 2, c) = "") Then
                    Cells(ll - 2, c) = ALUNO
                    Cells(ll, c) = TURMA
                    Cells(L, 37) = ""
                    Cells(L, 38) = ""
                    GoTo FIM2
                End If
            End If
        Next
    Next
FIM2:
Next
End Sub
Sub distribui_sala21()
Dim BD As Worksheet
Set BD = Sheets("BD")
sala = ActiveSheet.name

' copia os dados para a planilha atual
If MsgBox("Deseja importar os nomes da base de dados?", vbYesNo) = vbYes Then
'LIMPA INTERVALO DAS CADEIRAS
    Worksheets(sala).Range("AN14:AO65000").Clear

    For L = 1 To BD.Range("B65000").End(xlUp).Row
        If BD.Range("E" & L) = sala Then
        LIN = Worksheets(sala).Range("AN65000").End(xlUp).Row + 1
            Worksheets(sala).Range("AN" & LIN) = BD.Range("B" & L) 'NOME
            Worksheets(sala).Range("AO" & LIN) = BD.Range("C" & L) 'TURMA
        End If
    Next
End If
'enturma
'SALA = ActiveSheet.Name
Worksheets(sala).Activate
For L = 14 To Range("AN65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("AO" & L)
ALUNO = Range("AN" & L)
    For ll = 14 To 38 'MAPA DE SALA
        For c = 5 To 32
            If (TURMA = Cells(ll, c)) And Cells(ll - 2, c) = "" Then
                Cells(ll - 2, c) = ALUNO
                Cells(L, 40) = ""
                Cells(L, 41) = ""
                GoTo FIM
            End If
        Next
    Next
FIM:
Next
'ACHA ESPAÇOS VAZIOS E ADICIONA OS NOMES SEM ESPAÇO
sala = ActiveSheet.name
Worksheets(sala).Activate
For L = 14 To Range("AN65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("AO" & L)
ALUNO = Range("AN" & L)
    For ll = 15 To 43 'MAPA DE SALA
        For c = 5 To 35
            If TURMA <> "" And ALUNO <> "" Then
                If (Len(Cells(ll, c)) = 2 And Cells(ll - 2, c) = "") Then
                    Cells(ll - 2, c) = ALUNO
                    Cells(ll, c) = TURMA
                    Cells(L, 40) = ""
                    Cells(L, 41) = ""
                    GoTo FIM2
                End If
            End If
        Next
    Next
FIM2:
Next
End Sub
Sub distribui_sala1()
Dim BD As Worksheet
Set BD = Sheets("BD")
sala = ActiveSheet.name

' copia os dados para a planilha atual
If MsgBox("Deseja importar os nomes da base de dados?", vbYesNo) = vbYes Then
'LIMPA INTERVALO DAS CADEIRAS
    Worksheets(sala).Range("AO15:AP65000").Clear

    For L = 1 To BD.Range("B65000").End(xlUp).Row
        If BD.Range("E" & L) = sala Then
        LIN = Worksheets(sala).Range("AO65000").End(xlUp).Row + 1
            Worksheets(sala).Range("AO" & LIN) = BD.Range("B" & L) 'NOME
            Worksheets(sala).Range("AP" & LIN) = BD.Range("C" & L) 'TURMA
        End If
    Next
End If
'enturma

'SALA = ActiveSheet.Name
Worksheets(sala).Activate
For L = 14 To Range("AO65000").End(xlUp).Row 'LISTA MAPA DE SALA
TURMA = Range("AP" & L)
ALUNO = Range("AO" & L)
    For ll = 15 To 38 'MAPA DE SALA
        For c = 5 To 36
            If (TURMA = Cells(ll, c)) And Cells(ll - 2, c) = "" Then
                Cells(ll - 2, c) = ALUNO
                Cells(L, 41) = ""
                Cells(L, 42) = ""
                GoTo FIM
            End If
        Next
    Next
FIM:
Next
End Sub
