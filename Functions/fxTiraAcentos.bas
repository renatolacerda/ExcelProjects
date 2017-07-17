Attribute VB_Name = "fxTiraAcentos"
Function TIRA_ACENTOS(N)
N = Replace(N, "  ", " ")
N = Replace(N, "   ", " ")
N = Replace(N, "    ", " ")
N = UCase(N)
For conta = 1 To Len(N)
    valor = Mid(N, conta, 1)
    Select Case UCase(valor)
        Case "�"
            Letra = "A"
        Case "�"
            Letra = "A"
        Case "�"
            Letra = "A"
        Case "�"
            Letra = "A"
        Case "�"
            Letra = "A"
        Case "�"
            Letra = "E"
        Case "�"
            Letra = "E"
        Case "�"
            Letra = "E"
        Case "�"
            Letra = "E"
        Case "�"
            Letra = "I"
        Case "�"
            Letra = "I"
        Case "�"
            Letra = "I"
        Case "�"
            Letra = "I"
        Case "�"
            Letra = "O"
        Case "�"
            Letra = "O"
        Case "�"
            Letra = "O"
        Case "�"
            Letra = "O"
        Case "�"
            Letra = "O"
        Case "�"
            Letra = "U"
        Case "�"
            Letra = "U"
        Case "�"
            Letra = "U"
        Case "�"
            Letra = "U"
        Case "�"
            Letra = "C"
        Case "�"
            Letra = "N"
        Case "'"
            Letra = " "
        Case Else
            Letra = Mid(N, conta, 1)
    End Select
    nome = nome & Letra
Next
TIRA_ACENTOS = nome
End Function

Function TiraAcento(Palavra)
    CAcento = "�����������������������������������������������"
    SAcento = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
    Texto = ""
    If Palavra <> "" Then
        For X = 1 To Len(Palavra)
            Letra = Mid(Palavra, X, 1)
            Pos_Acento = InStr(CAcento, Letra)
            
            If Pos_Acento > 0 Then
                Letra = Mid(SAcento, Pos_Acento, 1)
            End If
            
            Texto = Texto & Letra
        Next
    TiraAcento = Texto
End If
End Function

Function VerificaPalavra(atributo)

Dim i
Dim id
Dim Auxiliar
Dim Resultado

Auxiliar = Split(atributo, " ", -1, vbBinaryCompare)

For i = LBound(Auxiliar) To UBound(Auxiliar)
    Resultado = Resultado & " " & TiraAcento(Auxiliar(i))
Next

VerificaPalavra = Trim(Resultado)

End Function

