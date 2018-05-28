Attribute VB_Name = "fxFuncoes"
Function TIRA_ACENTOS(N)
N = Replace(N, "  ", " ")
N = Replace(N, "   ", " ")
N = Replace(N, "    ", " ")
N = UCase(N)
For CONTA = 1 To Len(N)
    valor = Mid(N, CONTA, 1)
    Select Case UCase(valor)
        Case "�"
            LETRA = "A"
        Case "�"
            LETRA = "A"
        Case "�"
            LETRA = "A"
        Case "�"
            LETRA = "A"
        Case "�"
            LETRA = "A"
        Case "�"
            LETRA = "E"
        Case "�"
            LETRA = "E"
        Case "�"
            LETRA = "E"
        Case "�"
            LETRA = "E"
        Case "�"
            LETRA = "I"
        Case "�"
            LETRA = "I"
        Case "�"
            LETRA = "I"
        Case "�"
            LETRA = "I"
        Case "�"
            LETRA = "O"
        Case "�"
            LETRA = "O"
        Case "�"
            LETRA = "O"
        Case "�"
            LETRA = "O"
        Case "�"
            LETRA = "O"
        Case "�"
            LETRA = "U"
        Case "�"
            LETRA = "U"
        Case "�"
            LETRA = "U"
        Case "�"
            LETRA = "U"
        Case "�"
            LETRA = "C"
        Case "�"
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
