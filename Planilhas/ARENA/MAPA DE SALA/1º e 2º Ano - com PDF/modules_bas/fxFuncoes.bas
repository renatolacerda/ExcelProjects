Attribute VB_Name = "fxFuncoes"
Function TIRA_ACENTOS(N)
N = Replace(N, "  ", " ")
N = Replace(N, "   ", " ")
N = Replace(N, "    ", " ")
N = UCase(N)
For CONTA = 1 To Len(N)
    valor = Mid(N, CONTA, 1)
    Select Case UCase(valor)
        Case "Á"
            LETRA = "A"
        Case "Â"
            LETRA = "A"
        Case "À"
            LETRA = "A"
        Case "Ä"
            LETRA = "A"
        Case "Ã"
            LETRA = "A"
        Case "É"
            LETRA = "E"
        Case "Ê"
            LETRA = "E"
        Case "È"
            LETRA = "E"
        Case "Ë"
            LETRA = "E"
        Case "Í"
            LETRA = "I"
        Case "Î"
            LETRA = "I"
        Case "Ì"
            LETRA = "I"
        Case "Ï"
            LETRA = "I"
        Case "Ó"
            LETRA = "O"
        Case "Ô"
            LETRA = "O"
        Case "Ò"
            LETRA = "O"
        Case "Ö"
            LETRA = "O"
        Case "Õ"
            LETRA = "O"
        Case "Ú"
            LETRA = "U"
        Case "Û"
            LETRA = "U"
        Case "Ù"
            LETRA = "U"
        Case "Ü"
            LETRA = "U"
        Case "Ç"
            LETRA = "C"
        Case "Ñ"
            LETRA = "N"
        Case "'"
            LETRA = " "
        Case Else
            LETRA = Mid(N, CONTA, 1)
    End Select
    NOME = NOME & LETRA
    'If InStr(1, NOME, "(ITA)") Then
    '    NOME = Left(NOME, InStr(1, NOME, "(ITA)") - 1)
    '    If Right(NOME, 1) = " " Then NOME = Left(NOME, Len(NOME) - 1)
    'End If
Next
TIRA_ACENTOS = NOME
End Function
Public Function UltimaColuna(NomeDaPlanilha As String, linha As Integer, COLUNA As Integer)
Dim PLAN As Worksheet
Set PLAN = Sheets(NomeDaPlanilha)
    UltimaColuna = PLAN.Cells(linha, COLUNA).End(xlToLeft).Column
End Function

Public Function UltimaLinha(PLAN As Worksheet, COLUNA As Integer)
    UltimaLinha = PLAN.Cells(65000, COLUNA).End(xlUp).Row
End Function
Public Function GetDados(PLAN As Worksheet, linha As Integer, COLUNA As Integer)
    GetDados = TIRA_ACENTOS(PLAN.Cells(linha, COLUNA).Value)
End Function

Public Function GetDados_Pesquisa(PLAN As Worksheet, linha As Integer, pesquisa As Variant)
    For Each c In PLAN.Cells
        If UCase(c) = UCase(pesquisa) Then
            COLUNA = c.Column
            Exit For
        End If
    Next
    'Plan.Range(Cells(linha, coluna), Cells(linha, coluna)).Select
    GetDados_Pesquisa = PLAN.Cells(linha, COLUNA).Value
End Function
Public Function AchaColuna(PLAN As Worksheet, valor As Variant) As Integer
    For Each c In PLAN.Cells
        If UCase(c) = UCase(valor) Then
            COLUNA = c.Column
            Exit For
        End If
    Next
    AchaColuna = COLUNA
End Function
Public Function FORMATAR(NomeDaPlanilha As String, COLUNA As Integer, NOMEFORMATO As String, LINHAINICIAL As Integer, COLUNAINICIAL As Integer)
Dim PLAN As Worksheet
Set PLAN = Sheets(NomeDaPlanilha)
Dim ultlinha, ultcoluna
    ultlinha = UltimaLinha(PLAN, COLUNA)
    ultcoluna = UltimaColuna(PLAN.name, LINHAINICIAL - 1, 200)
    PLAN.Range(NOMEFORMATO).Copy
    PLAN.Range(Cells(LINHAINICIAL, COLUNAINICIAL), Cells(ultlinha, ultcoluna)).Select
    Selection.PasteSpecial xlPasteFormats
End Function
