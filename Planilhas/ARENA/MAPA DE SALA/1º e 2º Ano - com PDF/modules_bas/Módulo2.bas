Attribute VB_Name = "Módulo2"
Sub ACHA_FALTANTES_AUDITORIO()
Application.ScreenUpdating = False
r = Sheets("Auditorio").Range("E13:X65000")
For Each RESPOSTA In r
    For LIN = 14 To Sheets("Auditorio").Range("AB65000").End(xlUp).Row
    If RESPOSTA = "" Then Exit For
        If RESPOSTA = Sheets("Auditorio").Range("AB" & LIN) Then
            Sheets("Auditorio").Range("AB" & LIN) = ""
            Sheets("Auditorio").Range("AC" & LIN) = ""
            Exit For
        End If
    Next
Next
Application.ScreenUpdating = True
End Sub
Sub ACHA_FALTANTES_SL1()
Application.ScreenUpdating = False
r = Sheets("Sala 1").Range("E13:AK42")
For Each RESPOSTA In r
    For LIN = 14 To Sheets("Sala 1").Range("AO65000").End(xlUp).Row
    If RESPOSTA = "" Then Exit For
        If RESPOSTA = Sheets("Sala 1").Range("AO" & LIN) Then
            Sheets("Sala 1").Range("AO" & LIN) = ""
            Sheets("Sala 1").Range("AP" & LIN) = ""
            Exit For
        End If
    Next
Next
MsgBox "Fim...", vbInformation
Application.ScreenUpdating = True
End Sub
Sub ACHA_FALTANTES_SL2()
Application.ScreenUpdating = False
sala = "Sala 2"
r = Sheets(sala).Range("E13:AF38")
For Each RESPOSTA In r
    For LIN = 14 To Sheets(sala).Range("AK65000").End(xlUp).Row
    If RESPOSTA = "" Then Exit For
        If RESPOSTA = Sheets(sala).Range("AK" & LIN) Then
            Sheets(sala).Range("AK" & LIN) = ""
            Sheets(sala).Range("AL" & LIN) = ""
            Exit For
        End If
    Next
Next
MsgBox "Fim...", vbInformation
Application.ScreenUpdating = True
End Sub
Sub ACHA_FALTANTES_SL3()
Application.ScreenUpdating = False
sala = "Sala 3"
r = Sheets(sala).Range("E13:AF35")
For Each RESPOSTA In r
    For LIN = 14 To Sheets(sala).Range("AK65000").End(xlUp).Row
    If RESPOSTA = "" Then Exit For
        If RESPOSTA = Sheets(sala).Range("AK" & LIN) Then
            Sheets(sala).Range("AK" & LIN) = ""
            Sheets(sala).Range("AL" & LIN) = ""
            Exit For
        End If
    Next
Next
MsgBox "Fim...", vbInformation
Application.ScreenUpdating = True
End Sub
Sub ACHA_FALTANTES_SL4()
Application.ScreenUpdating = False
sala = "Sala 4"
r = Sheets(sala).Range("E13:AI39")
For Each RESPOSTA In r
    For LIN = 14 To Sheets(sala).Range("AK65000").End(xlUp).Row
    If RESPOSTA = "" Then Exit For
        If RESPOSTA = Sheets(sala).Range("AK" & LIN) Then
            Sheets(sala).Range("AK" & LIN) = ""
            Sheets(sala).Range("AL" & LIN) = ""
            Exit For
        End If
    Next
Next
MsgBox "Fim...", vbInformation
Application.ScreenUpdating = True
End Sub
Sub ACHA_FALTANTES_SL5()
Application.ScreenUpdating = False
sala = "Sala 5"
r = Sheets(sala).Range("E13:AN30")
For Each RESPOSTA In r
    For LIN = 14 To Sheets(sala).Range("AR65000").End(xlUp).Row
    If RESPOSTA = "" Then Exit For
        If RESPOSTA = Sheets(sala).Range("AR" & LIN) Then
            Sheets(sala).Range("AR" & LIN) = ""
            Sheets(sala).Range("AS" & LIN) = ""
            Exit For
        End If
    Next
Next
MsgBox "Fim...", vbInformation
Application.ScreenUpdating = True
End Sub
Sub ACHA_FALTANTES_SL6()
Application.ScreenUpdating = False
sala = "Sala 6"
r = Sheets(sala).Range("E13:AI34")
For Each RESPOSTA In r
    For LIN = 14 To Sheets(sala).Range("AO65000").End(xlUp).Row
    If RESPOSTA = "" Then Exit For
        If RESPOSTA = Sheets(sala).Range("AO" & LIN) Then
            Sheets(sala).Range("AO" & LIN) = ""
            Sheets(sala).Range("AP" & LIN) = ""
            Exit For
        End If
    Next
Next
MsgBox "Fim...", vbInformation
Application.ScreenUpdating = True
End Sub
Sub ACHA_FALTANTES_SL7()
Application.ScreenUpdating = False
sala = "Sala 7"
r = Sheets(sala).Range("E13:K38")
For Each RESPOSTA In r
    For LIN = 14 To Sheets(sala).Range("Q65000").End(xlUp).Row
    If RESPOSTA = "" Then Exit For
        If RESPOSTA = Sheets(sala).Range("Q" & LIN) Then
            Sheets(sala).Range("Q" & LIN) = ""
            Sheets(sala).Range("R" & LIN) = ""
            Exit For
        End If
    Next
Next
MsgBox "Fim...", vbInformation
Application.ScreenUpdating = True
End Sub
Sub ACHA_FALTANTES_SL89()
Application.ScreenUpdating = False
sala = "Sala 9"
r = Sheets(sala).Range("E13:AG33")
For Each RESPOSTA In r
    For LIN = 14 To Sheets(sala).Range("BL45000").End(xlUp).Row
    If RESPOSTA = "" Then Exit For
        If RESPOSTA = Sheets(sala).Range("BL" & LIN) Then
            Sheets(sala).Range("BL" & LIN) = ""
            Sheets(sala).Range("BM" & LIN) = ""
            Exit For
        End If
    Next
Next
MsgBox "Fim...", vbInformation
Application.ScreenUpdating = True
End Sub
Sub tira16sala5()
valor = 1
v = "F;E;D;C;B;A"
contador = 16
Worksheets("BD").Activate
MYVALOR = Split(v, ";")
FK:
For y = Worksheets("BD").Range("A65000").End(xlUp).Row To 1 Step -1
    If Worksheets("BD").Range("E" & y) = "Sala 5" Then
            IND = valor - 1
            If IND = 6 Then IND = 1
            If IND = 7 Then IND = 2
            If valor = 7 Then valor = 1
            If valor = 5 Then valor = 6
        If valor <> 5 And contador <> 0 And Worksheets("BD").Range("C" & y) = "3" & MYVALOR(IND) Then
            Worksheets("BD").Range("E" & y).Select
            Worksheets("BD").Range("E" & y) = "Sala " & valor
            valor = valor + 1
            contador = contador - 1: GoTo FK
        End If
    End If
Next
Worksheets("CONFIG").Activate
End Sub
Sub tira16sala5parasala7()
valor = 7
v = "F;E;D;C;B;A"
V1 = 0
contador = 16
Worksheets("BD").Activate
MYVALOR = Split(v, ";")
FK2:
For y = Worksheets("BD").Range("A65000").End(xlUp).Row To 1 Step -1
    If Worksheets("BD").Range("E" & y) = "Sala 5" Then
            If V1 = 6 Then V1 = 0
            IND = V1
            If IND = 6 Then IND = 1
            If IND = 7 Then IND = 2
        If valor <> 5 And contador <> 0 And Worksheets("BD").Range("C" & y) = "3" & MYVALOR(IND) Then
            Worksheets("BD").Range("E" & y).Select
            Worksheets("BD").Range("E" & y) = "Sala " & valor
            valor = 7: V1 = V1 + 1
            contador = contador - 1: GoTo FK2
        End If
    End If
Next
For WS = 1 To Sheets.count
    If UCase(Sheets(WS).name) = "SALA 7" Then
        Application.DisplayAlerts = False
        Sheets("SALA 7").Delete
        Application.DisplayAlerts = True
        Exit For
    End If
Next
Sheets("MAPA - SL7").Visible = True
    MAPAN5.Copy AFTER:=Worksheets("Rel-Sala")
    Worksheets("MAPA - SL7 (2)").Select
    Worksheets("MAPA - SL7 (2)").name = "Sala 7"
    ActiveSheet.Shapes("WordArt 1").Select
    Selection.ShapeRange.TextEffect.Text = "Mapa - " & "Sala 7"
Sheets("MAPA - SL7").Visible = False

Worksheets("CONFIG").Activate
End Sub
Sub tirar_alunos_E1()
Select Case Range("E1")
Case True
    Range("n1") = False
End Select
End Sub
Sub tirar_alunos_N1()
Select Case Range("N1")
Case True
    Range("E1") = False
End Select
End Sub

